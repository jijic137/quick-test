#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Token 接口快速测试工具
======================
功能：读取持久化测试用例 → 逐条执行 → 输出单个 CSV 结果文件
格式：严格对齐 template.xlsx，报告 + 用例放在同一个文件，方便阅读

用法（命令行）：
  python token_test.py -u https://api.example.com/v1/chat/completions -k sk-xxx -m gpt-4o

  python token_test.py -u $URL -k $KEY -m deepseek-v4 -c 8 -n 40 -o my_result.csv

  python token_test.py --help   查看所有参数

可复现：相同的参数 + 相同的用例定义 = 相同的测试流程
持久化：用例为 JSON 文件，可纳入版本管理；配置可通过命令行传入或 test_config.json 提供默认值
"""

import argparse
import csv
import json
import os
import sys
import time
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Windows 终端编码兼容 ────────────────────────────────────
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ── 第三方依赖检测 ──────────────────────────────────────────
try:
    import requests
except ImportError:
    print("[ERROR] 缺少 requests 库，请运行: python -m pip install requests")
    sys.exit(1)

# ── 路径常量 ────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent
CONFIG_PATH   = BASE_DIR / "test_config.json"
CASES_PATH    = BASE_DIR / "test_cases.json"
RESULTS_DIR   = BASE_DIR / "results"                  # 每次运行的详情输出目录
RESULTS_MD_DIR = BASE_DIR / "result_md"              # 验收报告输出目录
RESULTS_MD_DIR = BASE_DIR / "result_md"              # 验收报告输出目录
SUMMARY_PATH  = BASE_DIR / "test_summary.csv"         # 累积汇总文件（每次追加一行）

# 北京时间
TZ_BJ = timezone(timedelta(hours=8))

# ── 线程安全存储 ────────────────────────────────────────────
_stats_lock = threading.Lock()
_req_counter = [0]  # 全局请求计数器（用 list 避免 nonlocal）

# ── 基准测试参考文本（足够长，用于生成不同长度的输入）──
_BENCH_BASE_TEXT = (
    "You are a senior software engineer with expertise in system design, "
    "distributed systems, databases, networking, security, and machine learning. "
    "You provide clear, well-structured technical explanations with practical examples.\n\n"
    "The following is a technical reference covering key computer science topics:\n\n"
    "## System Design Principles\n"
    "Scalability: horizontal vs vertical scaling. Horizontal scaling adds more machines "
    "to distribute load (scale out), while vertical scaling adds more resources to a "
    "single machine (scale up). Stateless services are easier to scale horizontally. "
    "Stateful services require careful partitioning and replication strategies.\n\n"
    "Availability: measured in nines (99.9% = 8.76h downtime/year, 99.99% = 52min/year). "
    "Achieved through redundancy, failover mechanisms, health checks, circuit breakers, "
    "and graceful degradation. Active-active and active-passive are common HA patterns.\n\n"
    "Consistency: strong consistency guarantees all reads see the latest write (CP systems). "
    "Eventual consistency allows temporary inconsistencies for better availability (AP systems). "
    "Consensus algorithms (Paxos, Raft) provide strong consistency in distributed systems.\n\n"
    "## Database Design\n"
    "SQL databases (PostgreSQL, MySQL) provide ACID transactions and rich querying. "
    "NoSQL databases (MongoDB, Cassandra) offer flexible schemas and horizontal scaling. "
    "Indexing strategies: B-tree for range queries, hash for equality, GiST for geospatial. "
    "Query optimization: EXPLAIN ANALYZE, proper indexing, avoiding N+1 queries, connection pooling.\n\n"
    "## API Design\n"
    "REST: resources as URLs, HTTP methods as actions. GraphQL: client-specified queries. "
    "gRPC: high-performance RPC with Protocol Buffers. Rate limiting: token bucket, "
    "leaky bucket, fixed/sliding window. Authentication: JWT, OAuth 2.0, API keys, mTLS.\n\n"
    "## Networking\n"
    "TCP: connection-oriented, reliable delivery, flow control, congestion control. "
    "UDP: connectionless, low latency, no guarantees. HTTP/1.1: persistent connections, "
    "pipelining. HTTP/2: multiplexing, header compression, server push. HTTP/3: QUIC, "
    "improved multiplexing. DNS: hierarchical naming, caching, load balancing via DNS.\n\n"
    "## Security\n"
    "Defense in depth: multiple layers of security controls. Principle of least privilege: "
    "minimal permissions required. Common vulnerabilities: SQL injection, XSS, CSRF, SSRF. "
    "Mitigations: parameterized queries, output encoding, CSRF tokens, input validation.\n\n"
    "## Observability\n"
    "Three pillars: logging (structured event records), metrics (numeric time-series), "
    "tracing (request flow across services). Key metrics: latency, throughput, error rate, "
    "saturation. SLO/SLI/SLA: service level objectives, indicators, agreements.\n\n"
    "## CI/CD\n"
    "Continuous Integration: frequent merges, automated builds and tests. Continuous "
    "Delivery: automated deployment pipeline. Deployment strategies: blue-green, canary, "
    "rolling updates, feature flags. Infrastructure as Code: Terraform, Pulumi, CloudFormation.\n\n"
)

_GEN_BENCH_SYSTEM = {}  # 缓存不同长度的 system prompt


def _gen_benchmark_messages(target_input_tokens: int, user_question: str = None):
    """生成指定输入长度的 messages，用于压测和流式测试。
    system prompt 填充到 (target_input_tokens - user_tokens) 左右。
    返回 (messages, estimated_input_tokens)。
    """
    if user_question is None:
        user_question = (
            "Based on the technical reference provided above, please write a comprehensive "
            "summary covering the most important concepts across all sections. Include "
            "specific technical details and best practices. Write about 250-300 words."
        )
    user_tokens = max(50, len(user_question.split()) * 3 // 2)  # 粗略估算

    system_target = max(500, target_input_tokens - user_tokens)

    # 缓存 key 取最接近的 500 token 档位
    cache_key = (system_target // 500) * 500
    if cache_key in _GEN_BENCH_SYSTEM:
        system_text = _GEN_BENCH_SYSTEM[cache_key]
    else:
        # 重复填充到目标长度（英文 ~0.75 token/word ≈ 4 chars/token）
        needed_chars = system_target * 4
        repeats = max(1, needed_chars // len(_BENCH_BASE_TEXT) + 1)
        system_text = _BENCH_BASE_TEXT * repeats
        _GEN_BENCH_SYSTEM[cache_key] = system_text

    return [
        {"role": "system", "content": system_text},
        {"role": "user", "content": user_question},
    ]


def _gen_context_padding(target_tokens: int, chars_per_token: float = 4.0) -> str:
    """生成约 target_tokens 长的英文 padding 文本，用于上下文窗口测试。

    英文 ~4 chars/token，用 _BENCH_BASE_TEXT 重复填充，估算比中文"测"准确得多。
    返回 (padding_text, estimated_tokens)。
    """
    needed_chars = int(target_tokens * chars_per_token)
    repeats = max(1, needed_chars // len(_BENCH_BASE_TEXT) + 1)
    text = (_BENCH_BASE_TEXT * repeats)[:needed_chars]
    estimated_tokens = len(text) / chars_per_token
    return text, int(estimated_tokens)


# ── 工具函数 ────────────────────────────────────────────────

def _gen_request_id() -> str:
    """生成唯一请求 ID，用于链路追踪。"""
    _req_counter[0] += 1
    ts = datetime.now(TZ_BJ).strftime("%Y%m%d%H%M%S%f")
    return f"req_{ts}_{_req_counter[0]:06d}"


def _error_category(error_msg: str, status_code: int = None) -> str:
    """将错误归类到有限的桶里，便于统计分布。"""
    if error_msg == "请求超时":
        return "请求超时"
    if error_msg and error_msg.startswith("连接失败"):
        return "连接失败"
    if error_msg.startswith("未收到任何 token"):
        return "流式无输出"
    if status_code and status_code >= 400:
        return f"HTTP {status_code}"
    if status_code and status_code < 400 and error_msg:
        return f"HTTP {status_code}: {error_msg[:40]}"
    if error_msg:
        return error_msg[:60]
    return "未知错误"


def _chunk_keys(chunk: dict) -> dict:
    """提取 chunk 的结构信息（仅 key 和类型），用于流式调试。"""
    info = {}
    for k, v in chunk.items():
        if k == "choices" and isinstance(v, list) and v:
            delta = v[0].get("delta", {}) if isinstance(v[0], dict) else {}
            info["choices[0].delta"] = list(delta.keys()) if delta else "(empty)"
            info["choices[0].finish_reason"] = v[0].get("finish_reason") if isinstance(v[0], dict) else None
        elif isinstance(v, dict):
            info[k] = list(v.keys())
        elif isinstance(v, list):
            info[k] = f"[{len(v)}]"
        else:
            info[k] = type(v).__name__
    return info



def _error_category(error_msg: str, status_code: int = None) -> str:
    if error_msg == "请求超时": return "请求超时"
    if error_msg and error_msg.startswith("连接失败"): return "连接失败"
    if error_msg.startswith("未收到任何 token"): return "流式无输出"
    if status_code and status_code >= 400: return f"HTTP {status_code}"
    if status_code and status_code < 400 and error_msg: return f"HTTP {status_code}: {error_msg[:40]}"
    if error_msg: return error_msg[:60]
    return "未知错误"


def _chunk_keys(chunk: dict) -> dict:
    info = {}
    for k, v in chunk.items():
        if k == "choices" and isinstance(v, list) and v:
            delta = v[0].get("delta", {}) if isinstance(v[0], dict) else {}
            info["choices[0].delta"] = list(delta.keys()) if delta else "(empty)"
            info["choices[0].finish_reason"] = v[0].get("finish_reason") if isinstance(v[0], dict) else None
        elif isinstance(v, dict): info[k] = list(v.keys())
        elif isinstance(v, list): info[k] = f"[{len(v)}]"
        else: info[k] = type(v).__name__
    return info

def _percentiles(data: list, *ps) -> dict:
    """计算百分位值。data 为空时返回 {p: 0 for p in ps}。"""
    if not data:
        return {f"p{p}": 0 for p in ps}
    sorted_data = sorted(data)
    n = len(sorted_data)
    result = {}
    for p in ps:
        idx = int((p / 100.0) * (n - 1))
        result[f"p{p}"] = round(sorted_data[min(idx, n - 1)], 3)
    return result


def _std(data: list) -> float:
    """计算样本标准差。"""
    if len(data) < 2:
        return 0.0
    mean = sum(data) / len(data)
    return (sum((x - mean) ** 2 for x in data) / (len(data) - 1)) ** 0.5


# ---------------------------------------------------------------------------
# 参数解析
# ---------------------------------------------------------------------------
def _preprocess_argv() -> list:
    """
    预处理命令行参数：将 -k / --key 后的空格分隔部分合并为一个 token，
    解决 `-k x-api-key: 1232113`（无引号含空格）被 shell 拆成两个参数的问题。
    同时兼容 `-k "x-api-key: 1232113"` 和 `-k sk-plain-key`。
    """
    raw = sys.argv[1:]
    if not raw:
        return raw

    flags_needing_join = {"-k", "--key"}
    merged = []
    i = 0
    while i < len(raw):
        arg = raw[i]
        if arg in flags_needing_join and i + 1 < len(raw):
            # 收集后续非 flag 的 token 合并到 key 值
            parts = []
            j = i + 1
            while j < len(raw) and raw[j] not in flags_needing_join and not raw[j].startswith("-"):
                parts.append(raw[j])
                j += 1
            merged.append(arg)
            merged.append(" ".join(parts))
            i = j
        else:
            merged.append(arg)
            i += 1
    return merged


def parse_args():
    """解析命令行参数，未提供的参数从 test_config.json 取默认值。"""
    # 先从配置文件读取默认值
    defaults = _load_config_defaults()

    parser = argparse.ArgumentParser(
        description="Token 接口快速测试工具 — 执行测试用例并输出单个 CSV",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python token_test.py -u https://api.openai.com/v1/chat/completions -k sk-xxx -m gpt-4o
  python token_test.py -u $URL -k $KEY -m deepseek-v4 -c 8 -n 40 -t 120
  python token_test.py -u $URL -k $KEY -m gpt-4o --cases my_cases.json -o result.csv

参数未提供时，自动从 test_config.json 读取默认值。
        """,
    )
    parser.add_argument(
        "-u", "--url",
        default=defaults["url"],
        help="API 接口地址 (默认: 来自 test_config.json)",
    )
    parser.add_argument(
        "-k", "--key",
        default=defaults["key"],
        help="API Key (默认: 来自 test_config.json)",
    )
    parser.add_argument(
        "-m", "--model",
        default=defaults["model"],
        help="模型名称 (默认: 来自 test_config.json)",
    )
    parser.add_argument(
        "-c", "--concurrency",
        type=int,
        default=defaults["concurrency"],
        help="压测并发数 (默认: 来自 test_config.json)",
    )
    parser.add_argument(
        "-n", "--requests",
        type=int,
        default=defaults["total_requests"],
        help="压测总请求数 (默认: 来自 test_config.json)",
    )
    parser.add_argument(
        "-t", "--timeout",
        type=int,
        default=defaults["timeout"],
        help="单次请求超时秒数 (默认: 来自 test_config.json)",
    )
    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=None,
        help="输出 CSV 路径 (默认: 自动生成时间戳文件名 test_output_YYYYMMDD_HHMMSS.csv)",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=SUMMARY_PATH,
        help=f"累积汇总 CSV 路径，每次运行追加一行 (默认: {SUMMARY_PATH.name})",
    )
    parser.add_argument(
        "--cases",
        type=Path,
        default=CASES_PATH,
        help=f"测试用例 JSON 路径 (默认: {CASES_PATH.name})",
    )
    parser.add_argument(
        "--context-tokens",
        type=_parse_token_count,
        default=defaults["context_tokens"],
        help="上下文验收阈值, 支持 k/m 后缀。例: 128k, 1m, 512000 (默认: 512000)",
    )
    parser.add_argument(
        "--input-tokens",
        type=_parse_token_count,
        default=defaults["input_tokens"],
        help="压测/流式测试的输入 token 数，支持 k/m 后缀。例: 1k, 8k, 50k (默认: 1k)",
    )
    parser.add_argument(
        "--probe-context",
        action="store_true",
        default=False,
        help="启用渐进式上下文探测：从小量递增直到找到实际上限（较慢但精确）",
    )
    parser.add_argument(
        "--g-start",
        type=int,
        default=100,
        help="梯度测试起始并发数 (默认: 100)",
    )
    parser.add_argument(
        "--g-step",
        type=int,
        default=200,
        help="梯度测试步长 (默认: 200)",
    )
    parser.add_argument(
        "--g-max",
        type=int,
        default=None,
        help="梯度测试最大并发数 (默认: 取 -c 的值)",
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=50,
        help="实际并发线程数上限 (默认: 50)",
    )
    parser.add_argument(
        "--platform",
        type=str,
        default="",
        help="平台信息备注（如 GPU: A100×8），写入报告头部",
    )
    parser.add_argument(
        "--io-benchmark",
        action="store_true",
        default=False,
        help="独立开关：仅运行 TC-10 分输入输出档位性能测试（忽略 test_cases.json 其他用例）",
    )
    parser.add_argument(
        "--io-tiers",
        type=str,
        default="",
        help=(
            "覆盖 TC-10 档位，格式 in:out 逗号分隔，例: 1k:300,8k:1000,32k:2000 "
            "(支持 k/m 后缀)。不传则使用 test_cases.json 中 TC-10 的 tiers"
        ),
    )
    parser.add_argument(
        "--io-concurrency",
        type=int,
        default=None,
        help="TC-10 每档并发数（默认: test_cases.json 中的值或 benchmark.concurrency）",
    )
    parser.add_argument(
        "--io-requests",
        type=int,
        default=None,
        help="TC-10 每档总请求数（默认: test_cases.json 中的值或 10）",
    )
    parser.add_argument(
        "--io-max-context",
        type=_parse_token_count,
        default=None,
        help=(
            "TC-10 最大上下文长度上限（支持 k/m 后缀，如 256k / 700k）。"
            "混合分布模式：超过此值的请求跳过不测。"
            "朴素扫描模式：决定扫描终点（默认 380K），如 --io-max-context 700k 则扫到 ~693K。"
            "不传则使用默认值。"
        ),
    )
    parser.add_argument(
        "--io-step",
        type=_parse_token_count,
        default=10000,
        help="朴素扫描模式步长（支持 k/m 后缀）。默认 10k。例: 5k, 20k",
    )
    parser.add_argument(
        "--naive-io-tier",
        action="store_true",
        default=False,
        help=(
            "朴素扫描模式：输入从 10K 开始，步长 10K 递增至上限（默认 380K，"
            "--io-max-context 可调）。输出按档位固定：≤50K→0.2K, 50~80K→0.6K, "
            "80~160K→1.3K, >160K→7K。默认串行逐条发出；--io-concurrency N (N>1) 切换为"
            "每步重复 N 条全并发取均值。可与 --io-benchmark 联用，也可在完整测试中替换 TC-10。"
        ),
    )

    return parser.parse_args(_preprocess_argv())


def _load_config_defaults() -> dict:
    """从 test_config.json 加载默认值，文件不存在则返回内置默认值。"""
    builtin = {
        "url": "",
        "key": "",
        "model": "",
        "concurrency": 4,
        "total_requests": 20,
        "timeout": 60,
        "context_tokens": 512000,
        "input_tokens": 1000,  # 1k
    }
    if not CONFIG_PATH.exists():
        return builtin
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        return builtin

    api = cfg.get("api", {})
    bench = cfg.get("benchmark", {})
    ctx = cfg.get("context_test", {})

    return {
        "url": api.get("url", builtin["url"]),
        "key": api.get("key", builtin["key"]),
        "model": api.get("model", builtin["model"]),
        "concurrency": bench.get("concurrency", builtin["concurrency"]),
        "total_requests": bench.get("total_requests", builtin["total_requests"]),
        "timeout": api.get("timeout", builtin["timeout"]),
        "context_tokens": builtin["context_tokens"],  # argparse 只用单值，多值从 build_cfg 读取
        "input_tokens": bench.get("input_tokens", builtin["input_tokens"]),
    }


def _parse_token_count(value: str) -> int:
    """解析 --context-tokens 参数，支持后缀 k(=x1000) m(=x1000000)。"""
    value = value.strip().lower()
    if value.endswith("k"):
        return int(float(value[:-1]) * 1000)
    elif value.endswith("m"):
        return int(float(value[:-1]) * 1000 * 1000)
    return int(value)


def parse_io_tiers(raw: str) -> list:
    """解析 --io-tiers 字符串，格式 in:out 逗号分隔，支持 k/m 后缀。
    例: "1k:300,8k:1000,32k:2000" → [{'name':'1k:300',...}, ...]
    """
    tiers = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            continue
        in_s, out_s = part.split(":", 1)
        in_tok = _parse_token_count(in_s)
        out_tok = _parse_token_count(out_s)
        tiers.append({
            "name": f"{in_s.strip()}:{out_s.strip()}",
            "input_tokens": in_tok,
            "output_tokens": out_tok,
        })
    return tiers


def _build_naive_sweep(args, max_context=None) -> tuple:
    """构造朴素串行梯度扫描配置，返回 (sweep_list, label_string)。

    输入从 10K 开始，步长 10K，递增至 max_input。
    默认上限 380K，可通过 --io-max-context 指定（如 256K / 700K）。
    输出按输入所在档位固定：
      - 10K~50K   → 0.2K output  (P50)
      - 60K~80K   → 0.6K output  (AVG)
      - 90K~160K  → 1.3K output  (P90)
      - 170K~380K → 7K output    (P99)
      - >380K     → 7K output    (P99 延伸，不继续增加)
    """
    step_size = getattr(args, 'io_step', 10000) or 10000
    input_start = step_size  # 起始值 = 步长
    input_step = step_size
    input_max = (max_context - 7000) if max_context else 380000  # 预留输出空间

    # 档位输出映射: (input_upper_bound, output_tokens, label)
    tier_map = [
        (50000, 200, "P50"),
        (80000, 600, "AVG"),
        (160000, 1300, "P90"),
        (float("inf"), 7000, "P99"),
    ]

    sweep = []
    names = []
    current = input_start
    seq = 1
    while current <= input_max:
        # 确定当前输入属于哪个档位
        out_tok = 200
        tier_label = "P50"
        for bound, out, label in tier_map:
            if current <= bound:
                out_tok = out
                tier_label = label
                break

        sweep.append({
            "seq": seq,
            "input_tokens": current,
            "output_tokens": out_tok,
            "tier": tier_label,
        })
        names.append(f"{tier_label}({current//1000}K→{out_tok/1000:.1f}K)")
        current += input_step
        seq += 1

    label_str = f"串行扫描 {len(sweep)} 步: {input_start//1000}K→{input_max//1000}K, 步长{input_step//1000}K, 并发=1"
    return sweep, label_str


def _resolve_tier_concurrency(tier: dict, args_concurrency, bench_concurrency) -> int:
    """解析单档并发数：CLI 显式 > 档位内置默认 > benchmark 全局默认。"""
    if args_concurrency is not None:
        return args_concurrency
    return tier.get("_default_concurrency", bench_concurrency)


# ---------------------------------------------------------------------------
# 数据加载
# ---------------------------------------------------------------------------
def load_json(path: Path) -> dict:
    """加载 JSON 文件，出错时友好退出。"""
    if not path.exists():
        print(f"[ERROR] 找不到文件: {path}")
        sys.exit(1)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        print(f"[ERROR] JSON 解析失败 ({path.name}): {e}")
        sys.exit(1)


def resolve_template(text: str, variables: dict) -> str:
    """把模板中的 {{key}} 替换为实际值。"""
    for k, v in variables.items():
        text = text.replace(f"{{{{{k}}}}}", str(v))
    return text

# ---------------------------------------------------------------------------
# HTTP 请求封装
# ---------------------------------------------------------------------------
def api_request(
    url: str,
    key: str,
    model: str,
    messages: list,
    max_tokens: int = 50,
    timeout: int = 60,
    stream: bool = False,
) -> dict:
    """
    发送一次 Chat Completions 请求，返回统一结构：
      {
        "ok": bool,
        "status_code": int | None,
        "data": dict | None,
        "error": str | None,
        "latency": float,             # 秒
        "usage": {"prompt": int, "completion": int, "total": int} | None,
      }
    """
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Request-Id": _gen_request_id(),
    }
    payload = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "stream": stream,
    }
    result = {
        "ok": False,
        "status_code": None,
        "data": None,
        "error": None,
        "latency": 0.0,
        "usage": None,
    }

    t0 = time.perf_counter()
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
        result["latency"] = round(time.perf_counter() - t0, 3)
        result["status_code"] = resp.status_code

        try:
            body = resp.json()
        except Exception:
            body = {"_raw": resp.text[:500]}
        result["data"] = body

        if resp.status_code == 200:
            result["ok"] = True
            if "usage" in body:
                u = body["usage"]
                result["usage"] = {
                    "prompt": int(u.get("prompt_tokens", 0)),
                    "completion": int(u.get("completion_tokens", 0)),
                    "total": int(u.get("total_tokens", 0)),
                }
        else:
            err_msg = body.get("error", {})
            if isinstance(err_msg, dict):
                result["error"] = err_msg.get("message", resp.text[:300])
            else:
                result["error"] = str(err_msg) or resp.text[:300]
    except requests.exceptions.Timeout:
        result["latency"] = round(time.perf_counter() - t0, 3)
        result["error"] = "请求超时"
    except requests.exceptions.ConnectionError as e:
        result["latency"] = round(time.perf_counter() - t0, 3)
        result["error"] = f"连接失败: {e}"
    except Exception as e:
        result["latency"] = round(time.perf_counter() - t0, 3)
        result["error"] = f"未知错误: {e}"

    return result


def api_request_stream(url: str, key: str, model: str, messages: list,
                       max_tokens: int = 300, timeout: int = 120) -> dict:
    """
    流式请求，返回逐 token 时间戳用于 prefill/decode 分阶段测量。
    返回:
      {
        "ok": bool, "status_code": int|None,
        "ttft": float,              # Time To First Token (prefill 耗时)
        "total_latency": float,     # 总耗时
        "token_times": [float, ...],# 每个 token 到达的相对时间（从请求发出算）
        "completion_tokens": int,
        "prompt_tokens": int,
        "total_tokens": int,
        "error": str|None,
      }
    """
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Request-Id": _gen_request_id(),
    }
    payload = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "stream": True,
        "stream_options": {"include_usage": True},
    }
    result = {
        "ok": False, "status_code": None,
        "ttft": 0.0, "total_latency": 0.0,
        "token_times": [],
        "completion_tokens": 0, "prompt_tokens": 0, "total_tokens": 0,
        "cached_tokens": 0,
        "error": None,
    }

    t_start = time.perf_counter()
    try:
        resp = requests.post(url, headers=headers, json=payload,
                             stream=True, timeout=timeout)
        result["status_code"] = resp.status_code

        if resp.status_code != 200:
            result["error"] = f"HTTP {resp.status_code}: {resp.text[:300]}"
            result["total_latency"] = round(time.perf_counter() - t_start, 3)
            return result

        first_token = None
        first_chunk_time = None   # 第一个有效 SSE chunk 时间（用于兜底 TTFT）
        last_token_time = t_start
        completion_count = 0
        _debug_chunks = []        # 兜底触发时打印前几个 chunk 帮助定位

        for line in resp.iter_lines(decode_unicode=True):
            if not line or not line.startswith("data: "):
                continue
            data_str = line[6:].strip()
            if data_str == "[DONE]":
                break
            try:
                chunk = json.loads(data_str)
            except json.JSONDecodeError:
                continue

            # 记录前几个 chunk 用于兜底诊断
            if len(_debug_chunks) < 5:
                _debug_chunks.append(data_str[:500])

            now = time.perf_counter()

            # 提取 delta
            choices = chunk.get("choices", [])
            delta = choices[0].get("delta", {}) if choices else {}

            # 记录首个有效 SSE chunk 时间（任何非空 delta 字段都算）
            if first_chunk_time is None and delta:
                first_chunk_time = now

            # 首 token 时间 — 检查 content / reasoning_content / reasoning (思考模型)
            content_val = delta.get("content")
            reasoning_val = delta.get("reasoning_content") or delta.get("reasoning")
            has_ct = (content_val is not None and content_val != "") or (reasoning_val is not None and reasoning_val != "")
            if has_ct:
                if first_token is None:
                    first_token = now
                last_token_time = now
                completion_count += 1
                result["token_times"].append(round(now - t_start, 4))
            elif first_token is None and any(v for k, v in delta.items() if k != "role" and v):
                first_token = now  # 非标准 delta 字段兜底

            # usage 通常在最后一块返回
            if "usage" in chunk and chunk["usage"]:
                u = chunk["usage"]
                result["prompt_tokens"] = int(u.get("prompt_tokens", 0))
                result["completion_tokens"] = int(u.get("completion_tokens", 0))
                result["total_tokens"] = int(u.get("total_tokens", 0))
                ptd = u.get("prompt_tokens_details") or u.get("prompt_tokens_detail") or {}
                result["cached_tokens"] = int(ptd.get("cached_tokens", 0))

        result["total_latency"] = round(time.perf_counter() - t_start, 3)

        # usage 兜底恢复
        uc = result.get("completion_tokens", 0)
        if completion_count == 0 and uc > 0:
            result["completion_tokens"] = uc
            # 诊断打印：API 未按标准流式格式返回
            print(f"\n         [DEBUG SSE] completion_count=0, usage.completion_tokens={uc}")
            print(f"         前 {len(_debug_chunks)} 个 chunk 结构:")
            for i, c in enumerate(_debug_chunks):
                try:
                    j = json.loads(c)
                    # 只打印 keys 结构和 choices, 不打印完整 content
                    keys = list(j.keys())
                    choices_info = []
                    for ch in j.get("choices", []):
                        delta_keys = list(ch.get("delta", {}).keys()) if ch.get("delta") else "(无delta)"
                        choices_info.append(f"finish={ch.get('finish_reason')}, delta_keys={delta_keys}")
                    print(f"           [{i+1}] top_keys={keys} | choices=[{' / '.join(choices_info)}]")
                except Exception:
                    print(f"           [{i+1}] (json解析失败) raw={c[:200]}")
            print("")
            if first_chunk_time is not None:
                # 用首个 chunk 到达时间作为 TTFT（次优但真实）
                result["ttft"] = round(first_chunk_time - t_start, 3)
                result["ttft_method"] = "first_chunk"
            else:
                result["ttft"] = 0.001
                result["ttft_method"] = "fallback"
            completion_count = uc
        elif first_token:
            result["ttft"] = round(first_token - t_start, 3)
            result["ttft_method"] = "content"
            if result["completion_tokens"] == 0:
                result["completion_tokens"] = completion_count
        else:
            result["error"] = "未收到任何 token"

        result["ok"] = (first_token is not None and first_token != t_start) or (result.get("status_code") == 200 and uc > 0)

    except requests.exceptions.Timeout:
        result["total_latency"] = round(time.perf_counter() - t_start, 3)
        result["error"] = "请求超时"
    except requests.exceptions.ConnectionError as e:
        result["total_latency"] = round(time.perf_counter() - t_start, 3)
        result["error"] = f"连接失败: {e}"
    except Exception as e:
        result["total_latency"] = round(time.perf_counter() - t_start, 3)
        result["error"] = f"未知错误: {e}"

    return result


def query_model_info(url: str, key: str, model: str, timeout: int = 30) -> dict:
    """
    查询 /v1/models 获取模型声明的上下文长度等信息。
    返回 {"ok": bool, "context_length": int|None, "raw": dict|None, "error": str|None}
    """
    result = {"ok": False, "context_length": None, "raw": None, "error": None}

    # 从 chat completions URL 推导 models URL
    models_url = url.rstrip("/").replace("/chat/completions", "/models")
    if not models_url.endswith("/models"):
        models_url = url.rstrip("/") + "/models"

    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    try:
        resp = requests.get(models_url, headers=headers, timeout=timeout)
        if resp.status_code == 200:
            data = resp.json()
            result["ok"] = True
            result["raw"] = data

            # 尝试提取目标模型的 context_length
            models_list = data.get("data", [data] if isinstance(data, dict) else [])
            for m in models_list:
                if isinstance(m, dict) and m.get("id") == model:
                    cl = m.get("context_length") or m.get("max_context_length") or m.get("context_window")
                    if cl:
                        result["context_length"] = int(cl)
                    result["raw"] = m
                    break
            # 没找到精确匹配，返回第一个模型的 context_length 作为参考
            if result["context_length"] is None and models_list:
                first = models_list[0] if isinstance(models_list[0], dict) else {}
                cl = first.get("context_length") or first.get("max_context_length") or first.get("context_window")
                if cl:
                    result["context_length"] = int(cl)
        else:
            result["error"] = f"HTTP {resp.status_code}: {resp.text[:300]}"
    except Exception as e:
        result["error"] = str(e)

    return result


def probe_context_limit(cfg: dict, params: dict) -> dict:
    """
    渐进式探测实际上下文上限（二分搜索）。
    从小量开始，逐步翻倍直到失败，然后在最后区间二分精确定位。
    返回 {"max_tokens": int, "passed": bool, "attempts": list, "error": str|None}
    """
    api = cfg["api"]
    target = cfg.get("context_test", {}).get("target_tokens", params.get("target_tokens", 512000))
    chars_per_token = cfg.get("context_test", {}).get("chars_per_token_estimate", 4.0)
    timeout = cfg.get("context_test", {}).get("timeout") or max(api.get("timeout", 60), 120)
    prompt_template = params.get(
        "prompt_template",
        "请总结以下文本的开头3个字和结尾3个字，用英文回复: {padding}"
    )

    attempts = []
    low, high = 1000, target  # 从 1000 tokens 开始
    last_success = 0
    found_limit = None

    print(f"\n         [探测] 开始渐进式上下文探测 (timeout={timeout}s)...")

    # 阶段 1: 翻倍递增，找到上界
    test_size = low
    while test_size <= high:
        padding_text, _ = _gen_context_padding(test_size, chars_per_token)
        messages = [{"role": "user", "content": prompt_template.format(padding=padding_text)}]

        t0 = time.perf_counter()
        result = api_request(
            url=api["url"], key=api["key"], model=api["model"],
            messages=messages, max_tokens=50, timeout=timeout,
        )
        elapsed = round(time.perf_counter() - t0, 2)
        status = "PASS" if result["ok"] else "FAIL"
        err = result.get("error", "")[:80] if not result["ok"] else ""
        print(f"         [探测] {test_size:>8} tokens -> {status} ({elapsed}s){' | '+err if err else ''}")

        attempts.append({
            "size": test_size,
            "ok": result["ok"],
            "elapsed": elapsed,
            "status_code": result.get("status_code"),
            "error": result.get("error", "")[:200],
        })

        if result["ok"]:
            last_success = test_size
            if test_size >= high:
                found_limit = test_size
                break
            test_size *= 2
            if test_size > high:
                test_size = high  # 最后一跳精确到 target
        else:
            # 失败了，在当前区间二分
            high = test_size
            low = last_success
            break

    # 阶段 2: 二分精确定位（如果翻倍阶段找到了失败点）
    if found_limit is None and low < high:
        print(f"         [探测] 二分定位: {low} ~ {high}")
        while low < high - 500:  # 精度到 500 tokens
            mid = (low + high) // 2
            padding_text, _ = _gen_context_padding(mid, chars_per_token)
            messages = [{"role": "user", "content": prompt_template.format(padding=padding_text)}]

            t0 = time.perf_counter()
            result = api_request(
                url=api["url"], key=api["key"], model=api["model"],
                messages=messages, max_tokens=50, timeout=timeout,
            )
            elapsed = round(time.perf_counter() - t0, 2)
            status = "PASS" if result["ok"] else "FAIL"
            err = result.get("error", "")[:80] if not result["ok"] else ""
            print(f"         [探测] {mid:>8} tokens -> {status} ({elapsed}s){' | '+err if err else ''}")

            attempts.append({
                "size": mid,
                "ok": result["ok"],
                "elapsed": elapsed,
                "status_code": result.get("status_code"),
                "error": result.get("error", "")[:200],
            })

            if result["ok"]:
                low = mid
                last_success = mid
            else:
                high = mid

        found_limit = last_success if last_success > 0 else low

    if found_limit is None:
        found_limit = last_success

    passed = found_limit is not None and found_limit >= target

    return {
        "passed": passed,
        "found_limit": found_limit,
        "target": target,
        "attempts": attempts,
    }


# ---------------------------------------------------------------------------
# 配置字典构建（统一 CLI args 和 CFG 为兼容的 dict）
# ---------------------------------------------------------------------------
def build_cfg(args) -> dict:
    """将 argparse 结果构建为与旧 cfg dict 兼容的结构。"""
    # 从配置文件读取 context_test 相关默认值
    ctx_defaults = {}
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                j = json.load(f)
            ctx_defaults = j.get("context_test", {})
        except Exception:
            pass

    # target_tokens 支持 int 或 int[]，统一转成列表
    raw_target = ctx_defaults.get("target_tokens", args.context_tokens)
    if isinstance(raw_target, list):
        target_list = raw_target
    else:
        target_list = [raw_target]
    # CLI 显式传入时覆盖配置文件
    if args.context_tokens != raw_target if isinstance(raw_target, int) else args.context_tokens not in raw_target:
        target_list = [args.context_tokens]

    return {
        "api": {
            "url": args.url,
            "key": args.key,
            "model": args.model,
            "timeout": args.timeout,
        },
        "benchmark": {
            "concurrency": args.concurrency,
            "total_requests": args.requests,
            "input_tokens": args.input_tokens,
            "max_workers": args.max_workers,
            "gradient": {
                "start": args.g_start,
                "step": args.g_step,
                "max": args.g_max or args.concurrency,
            },
        },
        "context_test": {
            "target_tokens": target_list,
            "chars_per_token_estimate": ctx_defaults.get("chars_per_token_estimate", 4.0),
            "timeout": ctx_defaults.get("timeout"),
        },
    }

# ---------------------------------------------------------------------------
# 各测试方法
# ---------------------------------------------------------------------------

def test_connectivity(cfg: dict, case: dict) -> dict:
    """TC-01: 接口连通性 — 单次简单请求。"""
    api = cfg["api"]
    params = case.get("params", {})
    result = api_request(
        url=api["url"],
        key=api["key"],
        model=api["model"],
        messages=params.get("messages", [{"role": "user", "content": "Hi"}]),
        max_tokens=params.get("max_tokens", 50),
        timeout=api.get("timeout", 60),
    )
    return {
        "case_id": case["id"],
        "passed": result["ok"] and result["usage"] is not None,
        "detail": {
            "status_code": result["status_code"],
            "latency": result["latency"],
            "usage": result["usage"],
            "error": result["error"],
        },
    }


def test_tps_benchmark(cfg: dict, case: dict) -> dict:
    """TC-02: TPS 压测 — 并发梯度 + Decode TPS/TPM。"""
    api = cfg["api"]
    params = case.get("params", {})
    input_tokens = cfg["benchmark"].get("input_tokens", 1000)
    max_tokens = min(500, max(params.get("max_tokens", 300), input_tokens // 3))
    messages = _gen_benchmark_messages(input_tokens)
    timeout = api.get("timeout", 60)
    gradient = cfg["benchmark"].get("gradient", {})

    def _run_bench(concurrency: int, total: int):
        stats = {"ok": 0, "fail": 0,
                 "total_ttft": 0.0, "total_decode_time": 0.0, "total_latency": 0.0,
                 "total_prompt_tokens": 0, "total_completion_tokens": 0,
                 "total_tokens": 0, "errors": []}

        def _worker(_idx):
            varied_msgs = [dict(m) for m in messages]
            varied_msgs[-1]["content"] += f"\n\n[req:{_idx}]"
            r = api_request_stream(url=api["url"], key=api["key"], model=api["model"],
                                    messages=varied_msgs, max_tokens=max_tokens,
                                    timeout=timeout)
            with _stats_lock:
                if r["ok"]:
                    stats["ok"] += 1
                    ttft = r.get("ttft", 0)
                    total_lat = r.get("total_latency", 0)
                    decode_t = total_lat - ttft if total_lat > ttft else 0
                    stats["total_ttft"] += ttft
                    stats["total_decode_time"] += decode_t
                    stats["total_latency"] += total_lat
                    stats["total_prompt_tokens"] += r.get("prompt_tokens", 0)
                    stats["total_completion_tokens"] += r.get("completion_tokens", 0)
                    stats["total_tokens"] += r.get("total_tokens", 0)
                else:
                    stats["fail"] += 1
                    if r["error"] and len(stats["errors"]) < 20:
                        stats["errors"].append(r["error"])
            return r

        t0 = time.perf_counter()
        with ThreadPoolExecutor(max_workers=concurrency) as pool:
            futures = []
            for i in range(total):
                futures.append(pool.submit(_worker, i))
                time.sleep(0.02)
            for f in as_completed(futures):
                f.result()
        elapsed = time.perf_counter() - t0

        success_rate = stats["ok"] / total * 100 if total else 0
        # 总 TPS = 总 token / 墙上时间
        tps_tokens = stats["total_tokens"] / elapsed if elapsed > 0 else 0
        # Decode TPS = 输出 token / decode 时间（total_latency - TTFT）
        decode_tps = (stats["total_completion_tokens"] / stats["total_decode_time"]
                      if stats["total_decode_time"] > 0 else 0)
        # Prefill TPS = 输入 token / TTFT
        prefill_tps = (stats["total_prompt_tokens"] / stats["total_ttft"]
                       if stats["total_ttft"] > 0 else 0)
        avg_latency = stats["total_latency"] / stats["ok"] if stats["ok"] else 0
        avg_ttft = stats["total_ttft"] / stats["ok"] if stats["ok"] else 0

        error_counts = {}
        for err in stats["errors"]:
            cat = _error_category(err, None)
            error_counts[cat] = error_counts.get(cat, 0) + 1

        return {"concurrency": concurrency, "total": total,
                "elapsed": round(elapsed, 2), "ok": stats["ok"], "fail": stats["fail"],
                "success_rate": f"{success_rate:.1f}%",
                "tps_tokens": f"{tps_tokens:.1f}", "tpm_tokens": f"{tps_tokens * 60:.1f}",
                "tps_per_c": round(tps_tokens / concurrency, 1) if concurrency else 0,
                "decode_tps": f"{decode_tps:.1f}", "decode_tpm": f"{decode_tps * 60:.1f}",
                "prefill_tps": f"{prefill_tps:.1f}",
                "avg_latency": round(avg_latency, 3), "avg_ttft": round(avg_ttft, 3),
                "total_tokens": stats["total_tokens"],
                "total_prompt_tokens": stats["total_prompt_tokens"],
                "total_completion_tokens": stats["total_completion_tokens"],
                "error_counts": error_counts}

    g_start = gradient["start"]
    g_step = gradient["step"]
    g_max = gradient["max"]
    levels = []

    print(f"         [梯度] {g_start} → {g_max} (步长 {g_step})")
    for c in range(g_start, g_max + 1, g_step):
        print(f"         [{c}] 并发={c} 请求={c} ... ", end="", flush=True)
        lr = _run_bench(c, c)
        levels.append(lr)
        print(f"ok={lr['ok']} fail={lr['fail']} TPS={lr['tps_tokens']} tok/s Decode={lr['decode_tps']} tok/s")
        if c + g_step <= g_max:
            time.sleep(1)

    total_ok = sum(l["ok"] for l in levels)
    total_fail = sum(l["fail"] for l in levels)
    total_tokens_all = sum(l["total_tokens"] for l in levels)
    best = max(levels, key=lambda l: float(l["tps_tokens"]))
    passed = total_ok > 0

    return {"case_id": case["id"], "passed": passed,
            "detail": {"mode": "gradient", "levels": levels,
                       "ok": total_ok, "fail": total_fail,
                       "total_tokens": total_tokens_all,
                       "tps_tokens": best["tps_tokens"], "tpm_tokens": best["tpm_tokens"],
                       "decode_tps": best["decode_tps"], "decode_tpm": best["decode_tpm"],
                       "best_concurrency": best["concurrency"]}}
def test_tpm_calc(cfg: dict, case: dict, prev_result: dict = None) -> dict:
    """TC-03: TPM 换算 — 基于 TC-02 的 TPS 结果验证 TPM = TPS x 60。"""
    passed = False
    detail = {"tps": "N/A", "tpm": "N/A", "formula_ok": False}

    if prev_result and prev_result.get("detail"):
        d = prev_result["detail"]
        tps_str = d.get("tps_tokens", "0")
        tpm_str = d.get("tpm_tokens", "0")
        try:
            tps = float(tps_str)
            tpm = float(tpm_str)
            detail["tps"] = f"{tps:.1f}"
            detail["tpm"] = f"{tpm:.1f}"
            expected_tpm = tps * 60
            detail["formula_ok"] = abs(tpm - expected_tpm) < 0.01 * expected_tpm
            passed = detail["formula_ok"]
        except (ValueError, TypeError):
            detail["error"] = "无法解析 TPS/TPM 数值"
    else:
        detail["error"] = "依赖 TC-02 结果，请先运行 TC-02"

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": detail,
    }


def test_context_limit(cfg: dict, case: dict, probe_mode: bool = False,
                       model_info: dict = None) -> dict:
    """TC-04: 上下文窗口验收 — 支持多档位阶梯测试（128k/256k/456k 等）。

    对 target_tokens 列表中的每个档位逐一测试，记录每个档位的 pass/fail。
    通过条件（观测模式）：任一档位通过即可 — 帮助定位实际支持的上限。
    """
    api = cfg["api"]
    params = case.get("params", {})
    raw_target = cfg.get("context_test", {}).get("target_tokens", params.get("target_tokens", 512000))
    # 统一为列表，降序（从大到小测，成功即停）
    targets = sorted(
        raw_target if isinstance(raw_target, list) else [raw_target],
        reverse=True,
    )
    chars_per_token = cfg.get("context_test", {}).get("chars_per_token_estimate", 4.0)
    timeout = cfg.get("context_test", {}).get("timeout") or max(api.get("timeout", 60), 120)

    # ── 渐进式探测模式（使用最大目标值）──
    if probe_mode:
        max_target = max(targets)
        # 临时修改 cfg 中的 target_tokens 为单值给 probe_context_limit 用
        probe_cfg = {
            "api": api,
            "context_test": {"target_tokens": max_target, "chars_per_token_estimate": chars_per_token},
        }
        probe_result = probe_context_limit(probe_cfg, params)
        found = probe_result.get("found_limit", 0)
        return {
            "case_id": case["id"],
            "passed": probe_result["passed"],
            "detail": {
                "mode": "probe",
                "targets": targets,
                "found_limit": found,
                "attempts_count": len(probe_result.get("attempts", [])),
                "attempts": probe_result.get("attempts", []),
            },
        }

    # ── 快速模式：优先用模型声明判断 ──
    declared = None
    if model_info and model_info.get("ok") and model_info.get("context_length"):
        declared = model_info["context_length"]

    if declared is not None:
        passed = declared >= max(targets)
        return {
            "case_id": case["id"],
            "passed": passed,
            "detail": {
                "mode": "declared",
                "targets": targets,
                "declared_context_length": declared,
                "passed": passed,
                "reason": (
                    f"模型声明上下文长度 {declared} >= {max(targets)}，通过"
                    if passed else
                    f"模型声明上下文长度 {declared} < {max(targets)}，不满足要求"
                ),
            },
        }

    # ── 阶梯式多档位测试 ──
    prompt_template = params.get(
        "prompt_template",
        "请总结以下文本的开头3个字和结尾3个字，用英文回复: {padding}"
    )
    tier_results = []

    print(f"         [阶梯测试] 共 {len(targets)} 个档位: {[f'{t//1000}k' for t in targets]}")

    for tier_idx, target in enumerate(targets):
        padding_text, estimated_input = _gen_context_padding(target, chars_per_token)
        messages = [
            {"role": "user", "content": prompt_template.format(padding=padding_text)}
        ]

        print(f"         [{tier_idx+1}/{len(targets)}] 测试 {target//1000}k ({estimated_input:,} tokens est, {len(padding_text):,} chars) ... ", end="", flush=True)

        t0 = time.perf_counter()
        result = api_request(
            url=api["url"], key=api["key"], model=api["model"],
            messages=messages,
            max_tokens=params.get("max_output_tokens", 50),
            timeout=timeout,
        )
        elapsed = round(time.perf_counter() - t0, 2)

        # ── 失败原因分析（复用现有逻辑）──
        failure_reason = None
        is_context_error = False
        is_timeout = False
        is_channel_error = False

        if not result["ok"]:
            status_code = result.get("status_code")
            error_msg = result.get("error") or ""
            err_lower = error_msg.lower()

            is_channel = any(
                kw in err_lower
                for kw in ["渠道不存在", "可用渠道", "no available channel", "no channel", "retry"]
            )
            if status_code and status_code in (500, 502, 503, 504) and is_channel:
                is_channel_error = True
                failure_reason = (
                    f"Relay 渠道不可用 (HTTP {status_code}): {error_msg[:200]}。"
                    "可能原因: 1) 长上下文请求处理时间过长，后端实例被健康检查摘除; "
                    "2) 分组下所有渠道均已下线或超载; 3) relay 重试耗尽仍未找到可用后端。"
                )
            elif error_msg == "请求超时":
                is_timeout = True
                failure_reason = (
                    f"请求超时 ({elapsed}s, timeout={timeout}s)。"
                    "可能原因: 1) 模型处理长上下文耗时超过超时设置; "
                    "2) 网关/代理层超时; 3) 模型不支持此长度导致无响应。"
                )
            elif status_code in (400, 413, 422):
                is_context_error = True
                failure_reason = f"服务端明确拒绝 (HTTP {status_code}): {error_msg[:200]}"
            elif status_code in (429,):
                failure_reason = f"被限流 (HTTP 429): {error_msg[:200]}"
            elif status_code in (500, 502, 503, 504):
                failure_reason = f"服务端错误 (HTTP {status_code}): {error_msg[:200]}"
            else:
                failure_reason = f"未知错误 (HTTP {status_code}): {error_msg[:200]}"

            is_context_error = is_context_error or any(
                kw in err_lower
                for kw in ["context", "token", "length", "limit", "maximum", "exceed", "too long", "truncat"]
            )

        tier_ok = result["ok"]
        status = "PASS" if tier_ok else "FAIL"
        usage_str = ""
        if result.get("usage"):
            u = result["usage"]
            usage_str = f" | prompt={u.get('prompt', '?')} tokens"
        err_str = f" | {failure_reason[:120]}" if failure_reason else ""
        print(f"{status} ({elapsed}s){usage_str}{err_str}")

        tier_results.append({
            "target": target,
            "estimated_input_tokens": estimated_input,
            "ok": tier_ok,
            "status_code": result["status_code"],
            "latency": elapsed,
            "usage": result["usage"],
            "error": result.get("error"),
            "failure_reason": failure_reason,
            "is_context_error": is_context_error,
            "is_timeout": is_timeout,
            "is_channel_error": is_channel_error,
        })

        # 成功即停，不再测更小的档
        if tier_ok:
            remaining = targets[tier_idx + 1:]
            if remaining:
                print(f"         -> {target//1000}k 通过，跳过更小档位 {[f'{t//1000}k' for t in remaining]}")
            break

    # ── 汇总 ──
    all_passed = all(t["ok"] for t in tier_results)
    any_passed = any(t["ok"] for t in tier_results)
    passed_count = sum(1 for t in tier_results if t["ok"])

    return {
        "case_id": case["id"],
        "passed": any_passed,  # 观测模式：任一档位通过即可
        "detail": {
            "mode": "multi_tier",
            "targets": targets,
            "declared": declared,
            "all_passed": all_passed,
            "any_passed": any_passed,
            "passed_count": f"{passed_count}/{len(targets)}",
            "tiers": tier_results,
        },
    }


def test_auth_failure(cfg: dict, case: dict) -> dict:
    """TC-05: 鉴权失败 — 无 Key 直接通过，有 Key 测无效 Key 拒绝。"""
    api = cfg["api"]
    key = api.get("key", "").strip()
    if not key or key == "sk-your-api-key-here":
        return {"case_id": case["id"], "passed": True,
                "detail": {"status_code": None, "expected_codes": "N/A（无需鉴权）",
                           "error": None, "note": "未提供 API Key，服务不需要鉴权"}}
    key = api.get("key", "").strip()
    if not key or key == "sk-your-api-key-here":
        return {"case_id": case["id"], "passed": True,
                "detail": {"status_code": None, "expected_codes": "N/A（无需鉴权）",
                           "error": None, "note": "未提供 API Key，服务不需要鉴权"}}

    result = api_request(
        url=api["url"],
        key="sk-invalid-key-for-testing",
        model=api["model"],
        messages=[{"role": "user", "content": "Hi"}],
        max_tokens=10,
        timeout=api.get("timeout", 60),
    )

    passed = result["status_code"] in (401, 403)

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "status_code": result["status_code"],
            "expected_codes": "401 / 403",
            "error": result["error"],
            "note": "使用无效 Key 发起请求，预期被拒绝" if passed else "未能触发预期的鉴权拒绝",
        },
    }


def test_rate_limit(cfg: dict, case: dict) -> dict:
    """TC-06: 限流/超时 — 高频并发，观察 429/超时行为。"""
    api = cfg["api"]
    params = case.get("params", {})
    concurrency = params.get("concurrency", 20)
    total = params.get("total_requests", 100)
    timeout = params.get("timeout", 5)

    stats = {
        "total": total,
        "concurrency": concurrency,
        "ok": 0,
        "fail": 0,
        "rate_limited": 0,
        "timeout_count": 0,
        "auth_error": 0,
        "other_error": 0,
        "errors": [],
    }

    def _worker(_idx):
        r = api_request(
            url=api["url"], key=api["key"], model=api["model"],
            messages=[{"role": "user", "content": "Hi"}],
            max_tokens=10, timeout=timeout,
        )
        with _stats_lock:
            if r["ok"]:
                stats["ok"] += 1
            else:
                stats["fail"] += 1
                code = r.get("status_code") or 0
                if code == 429:
                    stats["rate_limited"] += 1
                elif r.get("error") == "请求超时":
                    stats["timeout_count"] += 1
                elif code in (401, 403):
                    stats["auth_error"] += 1
                else:
                    stats["other_error"] += 1
                if r.get("error") and len(stats["errors"]) < 10:
                    stats["errors"].append(f"[{code}] {r['error']}")
        return r

    t_start = time.perf_counter()
    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = [pool.submit(_worker, i) for i in range(total)]
        for f in as_completed(futures):
            pass

    elapsed = time.perf_counter() - t_start
    passed = (stats["rate_limited"] + stats["timeout_count"]) > 0 or stats["ok"] > 0

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "elapsed": round(elapsed, 2),
            "ok": stats["ok"],
            "fail": stats["fail"],
            "rate_limited_429": stats["rate_limited"],
            "timeout": stats["timeout_count"],
            "auth_error": stats["auth_error"],
            "other_error": stats["other_error"],
            "errors": stats["errors"][:5],
            "note": "限流/超时被正确记录，不影响整体统计" if passed else "测试异常",
        },
    }

# ---------------------------------------------------------------------------
# TC-07: 工具调用验证
# ---------------------------------------------------------------------------

def test_tool_calling(cfg: dict, case: dict) -> dict:
    """TC-07: 完整工具调用闭环验证。
    第 1 步：模型调用工具 → 校验 tool_calls 和 arguments
    第 2 步：模拟工具返回 → 模型消化结果并给出最终回复
    """
    api = cfg["api"]
    params = case.get("params", {})
    tools = params.get("tools", [])
    max_tokens = params.get("max_tokens", 200)
    timeout = api.get("timeout", 60)

    headers = {
        "Authorization": f"Bearer {api['key']}",
        "Content-Type": "application/json",
        "Request-Id": _gen_request_id(),
    }

    # ── 第 1 步：发起带 tool 的请求 ──
    step1_msgs = [dict(m) for m in params.get("messages", [])]
    t1 = time.perf_counter()
    try:
        resp1 = requests.post(api["url"], headers=headers, json={
            "model": api["model"],
            "messages": step1_msgs,
            "tools": tools,
            "max_tokens": max_tokens,
            "stream": False,
        }, timeout=timeout)
        latency1 = round(time.perf_counter() - t1, 3)
    except Exception as e:
        return _tool_result(case["id"], False, {
            "step": 1, "error": str(e), "verdict": f"第 1 步请求失败: {e}",
        })

    body1 = {}
    try: body1 = resp1.json()
    except: pass

    choices1 = body1.get("choices", [])
    msg1 = choices1[0].get("message", {}) if choices1 else {}
    tool_calls = msg1.get("tool_calls", [])
    has_tool_calls = len(tool_calls) > 0

    if not has_tool_calls:
        content = msg1.get("content", "")
        return _tool_result(case["id"], False, {
            "step": 1, "status_code": resp1.status_code, "latency_step1": latency1,
            "has_tool_calls": False, "error": f"未返回 tool_calls: {content[:150]}",
            "verdict": "第 1 步: 模型未触发工具调用",
            "usage_step1": body1.get("usage"),
        })

    tc = tool_calls[0]
    tool_name = tc.get("function", {}).get("name", "unknown")
    args_str = tc.get("function", {}).get("arguments", "{}")
    args_valid = False
    arguments = None
    try:
        arguments = json.loads(args_str)
        args_valid = isinstance(arguments, dict) and len(arguments) > 0
    except json.JSONDecodeError:
        return _tool_result(case["id"], False, {
            "step": 1, "status_code": resp1.status_code, "latency_step1": latency1,
            "has_tool_calls": True, "tool_name": tool_name, "arguments_raw": args_str[:200],
            "verdict": f"第 1 步: tool_call 参数非合法 JSON",
            "usage_step1": body1.get("usage"),
        })

    if not args_valid:
        return _tool_result(case["id"], False, {
            "step": 1, "status_code": resp1.status_code, "latency_step1": latency1,
            "has_tool_calls": True, "tool_name": tool_name, "arguments": arguments,
            "verdict": "第 1 步: tool_call 参数为空",
            "usage_step1": body1.get("usage"),
        })

    # ── 第 2 步：模拟工具执行结果，发送回模型 ──
    tool_result = {"temperature": "26°C", "condition": "晴", "humidity": "45%"}
    step2_msgs = step1_msgs + [
        {"role": "assistant", "content": None, "tool_calls": tool_calls},
        {"role": "tool", "tool_call_id": tc.get("id", "call_1"),
         "content": json.dumps(tool_result, ensure_ascii=False)},
    ]

    t2 = time.perf_counter()
    try:
        resp2 = requests.post(api["url"], headers=headers, json={
            "model": api["model"],
            "messages": step2_msgs,
            "max_tokens": max_tokens,
            "stream": False,
        }, timeout=timeout)
        latency2 = round(time.perf_counter() - t2, 3)
    except Exception as e:
        return _tool_result(case["id"], False, {
            "step": 2, "tool_name": tool_name, "arguments": arguments,
            "error": str(e), "verdict": f"第 2 步请求失败: {e}",
        })

    body2 = {}
    try: body2 = resp2.json()
    except: pass

    choices2 = body2.get("choices", [])
    msg2 = choices2[0].get("message", {}) if choices2 else {}
    content2 = msg2.get("content", "") or ""

    # 检查最终回复是否引用了工具结果
    used_tool_result = any(kw in content2 for kw in ["26", "晴", "45", "温度", "weather", "天气"])

    passed = resp2.status_code == 200 and len(content2) > 0

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "status_code": resp2.status_code,
            "latency": round(latency1 + latency2, 3),
            "has_tool_calls": True,
            "tool_name": tool_name,
            "arguments_valid": args_valid,
            "arguments": arguments,
            "tool_result": tool_result,
            "final_response": content2[:200],
            "used_tool_result": used_tool_result,
            "verdict": (
                f"闭环成功: {tool_name}({json.dumps(arguments, ensure_ascii=False)}) "
                f"→ 模拟结果 → 模型回复: {content2[:100]}"
                if passed else f"闭环失败: 第 2 步 HTTP {resp2.status_code}"
            ),
            "usage_step1": body1.get("usage"),
            "usage_step2": body2.get("usage"),
        },
    }


def _tool_result(cid: str, passed: bool, detail: dict) -> dict:
    return {"case_id": cid, "passed": passed, "detail": detail}


# ---------------------------------------------------------------------------
# TC-08: 流式性能测试 — prefill / decode 分阶段测量
# ---------------------------------------------------------------------------

def test_streaming_benchmark(cfg: dict, case: dict) -> dict:
    """流式请求，分别统计 prefill/decode 阶段指标 + 百分位 + ITL 检查 + 缓存追踪。"""
    api = cfg["api"]
    params = case.get("params", {})
    concurrency = params.get("concurrency", 5)
    total = params.get("total_requests", 10)
    input_tokens = cfg["benchmark"].get("input_tokens", 1000)
    max_tokens = min(500, max(params.get("max_tokens", 300), input_tokens // 3))
    messages = _gen_benchmark_messages(input_tokens)

    # 每请求指标收集（用于百分位计算）
    per_req_ttft = []        # prefill 延迟
    per_req_tpot = []         # decode 每 token 延迟 (ms)
    per_req_decode_tps = []   # decode tok/s
    per_req_itl_max = []      # 最大包间延迟 (ms)
    per_req_cached = []       # 缓存命中 token 数

    stats = {
        "total": total, "concurrency": concurrency,
        "ok": 0, "fail": 0,
        "total_ttft": 0.0,
        "total_decode_time": 0.0,
        "total_prompt_tokens": 0,
        "total_completion_tokens": 0,
        "total_cached_tokens": 0,
        "total_itl": 0.0, "itl_count": 0,
        "itl_over_500ms": 0,            # 超过 500ms 的 ITL 次数
        "itl_over_500ms_requests": 0,   # 有超 500ms ITL 的请求数
        "errors": [],
    }

    t_start = time.perf_counter()

    def _worker(_idx):
        varied_msgs = [dict(m) for m in messages]
        varied_msgs[-1]["content"] += f"\n\n[req:{_idx}]"
        r = api_request_stream(
            url=api["url"], key=api["key"], model=api["model"],
            messages=varied_msgs, max_tokens=max_tokens,
            timeout=api.get("timeout", 120),
        )
        with _stats_lock:
            if r["ok"]:
                stats["ok"] += 1
                ttft = r.get("ttft", 0)
                total_lat = r.get("total_latency", 0)
                decode_time = total_lat - ttft if total_lat > ttft else 0
                completion = r.get("completion_tokens", 0)

                stats["total_ttft"] += ttft
                stats["total_decode_time"] += decode_time
                stats["total_prompt_tokens"] += r.get("prompt_tokens", 0)
                stats["total_completion_tokens"] += completion
                stats["total_cached_tokens"] += r.get("cached_tokens", 0)

                # 收集百分位数据
                per_req_ttft.append(ttft)
                tpot_ms = round(decode_time / completion * 1000, 2) if completion > 0 else 0
                per_req_tpot.append(tpot_ms)
                decode_tps = round(completion / decode_time, 1) if decode_time > 0 else 0
                per_req_decode_tps.append(decode_tps)

                # ITL 统计
                times = r.get("token_times", [])
                req_itl_max = 0
                req_itl_over = 0
                if len(times) >= 2:
                    for i in range(1, len(times)):
                        itl = times[i] - times[i-1]
                        itl_ms = round(itl * 1000, 1)
                        stats["total_itl"] += itl
                        stats["itl_count"] += 1
                        if itl > 0.5:  # 500ms
                            stats["itl_over_500ms"] += 1
                            req_itl_over += 1
                        req_itl_max = max(req_itl_max, itl)
                per_req_itl_max.append(round(req_itl_max * 1000, 1))
                if req_itl_over > 0:
                    stats["itl_over_500ms_requests"] += 1

                # 缓存
                per_req_cached.append(r.get("cached_tokens", 0))
            else:
                stats["fail"] += 1
                if r.get("error") and len(stats["errors"]) < 10:
                    stats["errors"].append(r["error"])
        return r

    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = [pool.submit(_worker, i) for i in range(total)]
        for f in as_completed(futures):
            f.result()

    elapsed = time.perf_counter() - t_start

    # ── 汇总指标 ──
    n_ok = stats["ok"] if stats["ok"] > 0 else 1
    ttft_avg = round(stats["total_ttft"] / n_ok, 3)
    decode_time_avg = round(stats["total_decode_time"] / n_ok, 3)
    total_lat_avg = round(ttft_avg + decode_time_avg, 3)

    total_prompt = stats["total_prompt_tokens"]
    total_completion = stats["total_completion_tokens"]

    prefill_tps = round(total_prompt / stats["total_ttft"], 1) if stats["total_ttft"] > 0 else 0
    decode_tps = round(total_completion / stats["total_decode_time"], 1) if stats["total_decode_time"] > 0 else 0

    itl_avg = round(stats["total_itl"] / stats["itl_count"] * 1000, 1) if stats["itl_count"] > 0 else 0
    tpot_avg = round(stats["total_decode_time"] / total_completion * 1000, 1) if total_completion > 0 else 0

    cache_hit_rate = round(stats["total_cached_tokens"] / (stats["total_prompt_tokens"] + 1) * 100, 1)

    # ── 百分位 ──
    ttft_pct = _percentiles(per_req_ttft, 50, 75, 90, 99)
    tpot_pct = _percentiles(per_req_tpot, 50, 99)
    itl_pct = _percentiles(per_req_itl_max, 50, 90, 99)
    itl_over_req_pct = round(stats["itl_over_500ms_requests"] / n_ok * 100, 1)

    ok_rate = stats["ok"] / total * 100 if total else 0
    ok_rate = stats["ok"] / total * 100 if total else 0
    passed = stats["ok"] > 0

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "elapsed": round(elapsed, 2),
            "reliable": ok_rate >= 50,
            "reliable": ok_rate >= 50,
            "ok": stats["ok"],
            "fail": stats["fail"],
            # 平均值
            "ttft_avg": ttft_avg,
            "total_lat_avg": total_lat_avg,
            "prefill_tps": prefill_tps,
            "decode_tps": decode_tps,
            "total_prompt_tokens": total_prompt,
            "total_completion_tokens": total_completion,
            "itl_avg_ms": itl_avg,
            "tpot_avg_ms": tpot_avg,
            # 百分位
            "ttft_p50": ttft_pct["p50"],
            "ttft_p75": ttft_pct["p75"],
            "ttft_p90": ttft_pct["p90"],
            "ttft_p99": ttft_pct["p99"],
            "tpot_p50": tpot_pct["p50"],
            "tpot_p99": tpot_pct["p99"],
            "itl_max_p50_ms": itl_pct["p50"],
            "itl_max_p90_ms": itl_pct["p90"],
            "itl_max_p99_ms": itl_pct["p99"],
            # 缓存 & 包间延迟
            "cache_hit_rate": f"{cache_hit_rate}%",
            "total_cached_tokens": stats["total_cached_tokens"],
            "itl_over_500ms_count": stats["itl_over_500ms"],
            "itl_over_500ms_requests_pct": f"{itl_over_req_pct}%",
            "errors": stats["errors"][:5],
        },
    }


# ---------------------------------------------------------------------------
# TC-09: 前缀缓存命中验证
# ---------------------------------------------------------------------------

_CACHE_SYSTEM_TEXT = (
    "You are an AI assistant with deep knowledge of software engineering, "
    "distributed systems, machine learning, and cloud architecture. "
    "The following is a comprehensive technical reference document.\n\n"
    "## 1. Distributed Systems\n"
    "A distributed system is a collection of independent computers that appears "
    "as a single coherent system. The CAP theorem: Consistency, Availability, "
    "Partition Tolerance — pick two. Raft consensus: leader election, log "
    "replication, safety. Paxos: proposers, acceptors, learners, two phases.\n"
    "## 2. Load Balancing\n"
    "Round Robin, Least Connections, IP Hash, Weighted variants. Layer 4 (TCP) "
    "vs Layer 7 (HTTP). Health checks: active and passive.\n"
    "## 3. Caching\n"
    "Cache-Aside, Read-Through, Write-Through, Write-Behind. Eviction: LRU, "
    "LFU, TTL, FIFO. Invalidation: time-based, event-driven, version-based.\n"
    "## 4. Databases\n"
    "B-Tree indexes for range queries. Hash indexes for point lookups. Sharding: "
    "range, hash, directory, geo. Replication: master-slave, master-master, "
    "synchronous vs asynchronous.\n"
    "## 5. Message Queues\n"
    "Kafka: high-throughput, partition-based. RabbitMQ: AMQP, flexible routing. "
    "SQS: managed, standard/FIFO. Event patterns: notification, event-carried "
    "state, event sourcing, CQRS.\n"
    "## 6. API Design\n"
    "REST: nouns for resources, HTTP methods for actions. Rate limiting: token "
    "bucket, fixed window, sliding window. Auth: API keys, JWT, OAuth 2.0, mTLS.\n"
    "## 7. Observability\n"
    "Three pillars: logging, metrics, tracing. RED: Rate, Errors, Duration. "
    "USE: Utilization, Saturation, Errors. Four Golden Signals: latency, "
    "traffic, errors, saturation. Alert on symptoms, not causes.\n"
    "## 8. Security\n"
    "Defense in depth. Principle of least privilege. Zero trust architecture. "
    "SQL injection prevention: parameterized queries. XSS: output encoding. "
    "CSRF: tokens. Authentication vs authorization. Encryption at rest and "
    "in transit. Certificate management and rotation.\n"
    "## 9. CI/CD\n"
    "Continuous Integration: frequent merges, automated builds and tests. "
    "Continuous Delivery: automated deployment to staging. Continuous "
    "Deployment: automated production deployment. Blue-green, canary, "
    "rolling deployments. Feature flags for gradual rollouts.\n"
    "## 10. Performance\n"
    "Amdahl's Law: speedup limited by serial portion. Gustafson's Law: "
    "larger problems benefit more from parallelism. Little's Law: L = λW. "
    "Percentile-based SLOs: p50, p95, p99. Tail latency amplification "
    "in microservices. Connection pooling and HTTP keep-alive.\n"
    "[End Reference]\n"
)

_GEN_CACHE_SYSTEM = ""


def _cache_system_prompt(target_tokens: int = 3000) -> str:
    """生成长 system prompt（约 target_tokens token），用于缓存命中测试。"""
    global _GEN_CACHE_SYSTEM
    if _GEN_CACHE_SYSTEM:
        return _GEN_CACHE_SYSTEM
    # 英文 ~0.75 token/word ≈ 4 chars/token，重复直到够长
    needed_chars = target_tokens * 4
    repeats = max(1, needed_chars // len(_CACHE_SYSTEM_TEXT) + 1)
    _GEN_CACHE_SYSTEM = _CACHE_SYSTEM_TEXT * repeats
    return _GEN_CACHE_SYSTEM


def test_cache_hit(cfg: dict, case: dict) -> dict:
    """TC-09: 缓存命中 TPM 测试。

    1. 发 1 次请求预热（填充前缀缓存）
    2. 并发发送 N 个请求（相同长 system prompt + 不同的 user message 后缀）
    3. 统计缓存命中率 + 缓存下的 TPM

    通过条件: 有缓存命中（cached_tokens > 0）且请求成功。
    """
    api = cfg["api"]
    params = case.get("params", {})
    concurrency = params.get("concurrency") or 4
    total = params.get("total_requests") or 20
    max_tokens = params.get("max_tokens", 50)
    target_tokens = params.get("system_prompt_tokens", 5000)
    timeout = api.get("timeout", 120)

    system_text = _cache_system_prompt(target_tokens)
    base_user = params.get("messages", [{"role": "user", "content": "Summarize."}])[0].get("content", "Summarize.")

    # ── 预热：填充前缀缓存 ──
    messages_warm = [
        {"role": "system", "content": system_text},
        {"role": "user", "content": base_user},
    ]
    warm = api_request_stream(
        url=api["url"], key=api["key"], model=api["model"],
        messages=messages_warm, max_tokens=max_tokens, timeout=timeout,
    )

    # ── 并发压测 ──
    stats = {
        "total": total, "concurrency": concurrency,
        "ok": 0, "fail": 0,
        "total_latency": 0.0,
        "total_prompt_tokens": 0,
        "total_completion_tokens": 0,
        "total_cached_tokens": 0,
        "total_ttft": 0.0,
        "errors": [],
    }

    t_start = time.perf_counter()

    def _worker(_idx):
        # 相同 system prompt + 唯一 user 后缀（确保前缀可缓存但未尾不可）
        varied_user = base_user + f"\n\n[req:{_idx}]"
        varied_msgs = [
            {"role": "system", "content": system_text},
            {"role": "user", "content": varied_user},
        ]
        r = api_request_stream(
            url=api["url"], key=api["key"], model=api["model"],
            messages=varied_msgs, max_tokens=max_tokens, timeout=timeout,
        )
        with _stats_lock:
            if r["ok"]:
                stats["ok"] += 1
                stats["total_latency"] += r.get("total_latency", 0)
                stats["total_ttft"] += r.get("ttft", 0)
                stats["total_prompt_tokens"] += r.get("prompt_tokens", 0)
                stats["total_completion_tokens"] += r.get("completion_tokens", 0)
                stats["total_cached_tokens"] += r.get("cached_tokens", 0)
            else:
                stats["fail"] += 1
                if r.get("error") and len(stats["errors"]) < 10:
                    stats["errors"].append(r["error"])
        return r

    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = [pool.submit(_worker, i) for i in range(total)]
        for f in as_completed(futures):
            f.result()

    elapsed = time.perf_counter() - t_start

    # ── 汇总 ──
    n_ok = stats["ok"] if stats["ok"] > 0 else 1
    total_tokens = stats["total_prompt_tokens"] + stats["total_completion_tokens"]
    tpm = round(total_tokens / elapsed * 60, 0) if elapsed > 0 else 0
    ttft_avg = round(stats["total_ttft"] / n_ok, 3)
    avg_latency = round(stats["total_latency"] / n_ok, 3)

    # 缓存命中率 (prompt 中包含 system 的缓存命中)
    prompt_total = stats["total_prompt_tokens"] if stats["total_prompt_tokens"] > 0 else 1
    cache_hit_rate = round(stats["total_cached_tokens"] / prompt_total * 100, 1)
    cache_hit = stats["total_cached_tokens"] > 0

    # 预热 vs 缓存后的 TTFT 对比
    warm_ttft = warm.get("ttft", 0)
    ttft_reduction = round((warm_ttft - ttft_avg) / warm_ttft * 100, 1) if warm_ttft > 0 else 0

    if cache_hit:
        verdict = (
            f"缓存命中率={cache_hit_rate}%, "
            f"TPM={tpm:.0f} tok/min, "
            f"TTFT cold={warm_ttft}s→avg={ttft_avg}s "
            f"(降{ttft_reduction}%), "
            f"成功{stats['ok']}/{stats['total']}"
        )
    else:
        verdict = (
            f"未命中, TPM={tpm:.0f} tok/min, "
            f"TTFT cold={warm_ttft}s/avg={ttft_avg}s, "
            f"成功{stats['ok']}/{stats['total']}"
        )

    return {
        "case_id": case["id"],
        "passed": stats["ok"] > 0 and cache_hit,
            "reliable": (stats["ok"] / total * 100 >= 50) if total else False,
        "detail": {
            "reliable": (stats["ok"] / total * 100 >= 50) if total else False,
            "warmup": {
                "ok": warm["ok"],
                "ttft": warm["ttft"],
                "prompt_tokens": warm.get("prompt_tokens", 0),
                "cached_tokens": warm.get("cached_tokens", 0),
            },
            "elapsed": round(elapsed, 2),
            "ok": stats["ok"],
            "fail": stats["fail"],
            "concurrency": concurrency,
            "total_requests": total,
            "tpm_tokens": f"{tpm:.0f}",
            "cache_hit_rate_pct": cache_hit_rate,
            "total_cached_tokens": stats["total_cached_tokens"],
            "total_prompt_tokens": stats["total_prompt_tokens"],
            "ttft_avg": ttft_avg,
            "ttft_reduction_pct": f"{ttft_reduction}%",
            "avg_latency": avg_latency,
            "errors": stats["errors"][:5],
            "verdict": verdict,
        },
    }


# ---------------------------------------------------------------------------
# TC-10: 分输入输出档位性能测试
# ---------------------------------------------------------------------------

def test_io_sweep_benchmark(cfg: dict, case: dict) -> dict:
    """TC-10 朴素扫描模式：输入从 10K 开始，步长 10K 递增，输出按档位固定。

    两种模式由 params.repetitions 控制：
      - repetitions=1（串行）：38 步逐条发出，测纯输入→性能曲线
      - repetitions>1（并发）：所有 38×N 条请求在同一秒内全部发出，
        每步取 N 条重复的平均值，测「并发争抢下」各输入长度的性能衰减

    params 结构:
      {
        "sweep": [                   # 扫描列表（必填）
          {"seq": 1, "input_tokens": 10000, "output_tokens": 200, "tier": "P50"},
          ...
        ],
        "repetitions": 1,            # 每步重复数（1=串行, >1=全并发）
        "max_output": 8192,
      }
    """
    api = cfg["api"]
    params = case.get("params", {})
    sweep = params.get("sweep", [])
    if not sweep:
        return {
            "case_id": case["id"],
            "passed": False,
            "detail": {"error": "params.sweep 为空，请提供扫描列表", "sweep": []},
        }

    timeout = api.get("timeout", 120)
    max_out = int(params.get("max_output", 8192))
    reps_original = int(params.get("repetitions", 1))
    total_steps = len(sweep)
    # reps=1: 串行; reps>1: 每步 N 条全并发
    if reps_original > 1:
        mode_label = f"按档并发(每步×{reps_original}, 共{total_steps * reps_original}条, P50→AVG→P90→P99依次发出)"
        total_requests = total_steps * reps_original
        is_concurrent = True
        actual_reps = reps_original
    else:
        mode_label = "串行(并发=1)"
        total_requests = total_steps
        is_concurrent = False
        actual_reps = 1

    print(f"         [朴素扫描] {total_steps} 步, 输入 {sweep[0]['input_tokens']//1000}K→"
          f"{sweep[-1]['input_tokens']//1000}K, 步长 10K, {mode_label}")

    # ── 构建所有请求任务列表 ──
    # 每条任务: (unique_step_index, repetition_index, seq_label)
    tasks = []
    for step in sweep:
        in_tok = step["input_tokens"]
        out_tok = min(step["output_tokens"], max_out)
        tier_label = step.get("tier", "?")
        seq = step["seq"]
        for rep in range(actual_reps):
            tasks.append({
                "step_seq": seq,
                "rep": rep,
                "input_tokens": in_tok,
                "output_tokens": out_tok,
                "tier": tier_label,
                "label": f"{seq}-{rep}" if actual_reps > 1 else str(seq),
            })

    # ── 单请求执行函数 ──
    def _do_one(task):
        in_tok = task["input_tokens"]
        out_tok = task["output_tokens"]
        messages = _gen_benchmark_messages(in_tok, user_question=(
            f"Write a detailed technical answer of about {max(200, out_tok // 2)} words "
            f"covering the topic above. Be thorough and structured. [sweep:{task['label']}]"
        ))
        varied = [dict(m) for m in messages]
        varied[-1]["content"] += f"\n\n[sweep:{task['label']}]"
        r = api_request_stream(
            url=api["url"], key=api["key"], model=api["model"],
            messages=varied, max_tokens=out_tok, timeout=timeout,
        )
        ttft_method = r.get("ttft_method", "")
        rec = {
            "step_seq": task["step_seq"],
            "rep": task["rep"],
            "input_tokens": in_tok,
            "output_tokens": out_tok,
            "tier": task["tier"],
            "status": "ok" if r["ok"] else "fail",
            "ttft": 0.0,
            "ttft_method": ttft_method,
            "decode_tps": 0.0,
            "tpot_ms": 0.0,
            "total_latency": 0.0,
            "completion_tokens": 0,
            "error": r.get("error", ""),
        }
        if r["ok"]:
            ttft = r.get("ttft", 0)
            total_lat = r.get("total_latency", 0)
            decode_d = total_lat - ttft if total_lat > ttft else 0
            completion = r.get("completion_tokens", 0)
            rec["ttft"] = round(ttft, 3)
            rec["total_latency"] = round(total_lat, 3)
            rec["completion_tokens"] = completion
            rec["decode_tps"] = round(completion / decode_d, 1) if decode_d > 0 else 0
            rec["tpot_ms"] = round(decode_d / completion * 1000, 2) if completion > 0 else 0
        return rec

    # ── 执行 ──
    t_start = time.perf_counter()
    raw_results = []

    if not is_concurrent:
        # 串行模式
        for task in tasks:
            raw_results.append(_do_one(task))
            seq = task["step_seq"]
            if True:  # 每条打印
                pct = seq * 100 // total_steps
                last = raw_results[-1]
                ttft_mark = "⚠️" if last.get("ttft_method") == "first_chunk" else ""
                print(f"         [{seq}/{total_steps}] {pct}% | "
                      f"in={last['input_tokens']//1000}K {last['tier']} | "
                      f"{'✅' if last['status']=='ok' else '❌'} "
                      f"TTFT={last['ttft']}s{ttft_mark} | Decode={last['decode_tps']}t/s")
    else:
        # 并发模式：按步依次并发（每步 N 条同时发出，不同步串行）
        step_tasks = {}
        for t in tasks:
            step_tasks.setdefault(t["step_seq"], []).append(t)
        for seq in range(1, total_steps + 1):
            st = step_tasks.get(seq, [])
            if not st:
                continue
            n = len(st)
            in_k = st[0]["input_tokens"] // 1000
            tier = st[0]["tier"]
            print(f"         [{seq}/{total_steps}] {seq * 100 // total_steps}% | "
                  f"in={in_k}K {tier}, 同时发出 {n} 条 ... ", end="", flush=True)
            t_step = time.perf_counter()
            with ThreadPoolExecutor(max_workers=n) as pool:
                futures = [pool.submit(_do_one, t) for t in st]
                for f in as_completed(futures):
                    raw_results.append(f.result())
            step_elapsed = round(time.perf_counter() - t_step, 2)
            # 本步摘要
            step_recs = [r for r in raw_results if r["step_seq"] == seq and r["status"] == "ok"]
            fail_n = n - len(step_recs)
            avg_ttft = round(sum(r["ttft"] for r in step_recs) / len(step_recs), 3) if step_recs else 0
            avg_decode = round(sum(r["decode_tps"] for r in step_recs) / len(step_recs), 1) if step_recs else 0
            ttft_mark = "⚠️" if any(r.get("ttft_method") == "first_chunk" for r in step_recs) else ""
            print(f"耗时{step_elapsed}s | ok={len(step_recs)}/fail={fail_n} | "
                  f"avg TTFT={avg_ttft}s{ttft_mark} | avg Decode={avg_decode}t/s")
        # 按 step_seq 排序以保持输出有序
        raw_results.sort(key=lambda r: (r["step_seq"], r["rep"]))

    elapsed = round(time.perf_counter() - t_start, 2)

    # ── 聚合：每步取 N 条重复的平均值 ──
    from collections import defaultdict
    step_groups = defaultdict(list)
    for r in raw_results:
        step_groups[r["step_seq"]].append(r)

    records = []
    ok_count = 0
    fail_count = 0
    for seq in sorted(step_groups.keys()):
        group = step_groups[seq]
        ok_recs = [r for r in group if r["status"] == "ok"]
        fail_recs = [r for r in group if r["status"] != "ok"]
        first = group[0]

        if ok_recs:
            ok_count += 1
            avg_ttft = round(sum(r["ttft"] for r in ok_recs) / len(ok_recs), 3)
            avg_decode_tps = round(sum(r["decode_tps"] for r in ok_recs) / len(ok_recs), 1)
            avg_tpot = round(sum(r["tpot_ms"] for r in ok_recs) / len(ok_recs), 1)
            avg_total_lat = round(sum(r["total_latency"] for r in ok_recs) / len(ok_recs), 3)
            avg_completion = int(sum(r["completion_tokens"] for r in ok_recs) / len(ok_recs))
            # 并发模式下额外统计标准差
            ttft_std = round(_std([r["ttft"] for r in ok_recs]), 3) if len(ok_recs) > 1 else 0
            decode_std = round(_std([r["decode_tps"] for r in ok_recs]), 1) if len(ok_recs) > 1 else 0
        else:
            fail_count += 1
            avg_ttft = avg_decode_tps = avg_tpot = avg_total_lat = ttft_std = decode_std = 0
            avg_completion = 0

        records.append({
            "seq": seq,
            "input_tokens": first["input_tokens"],
            "output_tokens": first["output_tokens"],
            "tier": first["tier"],
            "status": "ok" if ok_recs else "fail",
            "repetitions": len(group),
            "ok_reps": len(ok_recs),
            "fail_reps": len(fail_recs),
            "ttft": avg_ttft,
            "ttft_std": ttft_std,
            "decode_tps": avg_decode_tps,
            "decode_tps_std": decode_std,
            "tpot_ms": avg_tpot,
            "total_latency": avg_total_lat,
            "completion_tokens": avg_completion,
            "error": fail_recs[0]["error"] if fail_recs else "",
        })

    # ── 分档位汇总 ──
    tier_summary = {}
    for rec in records:
        t = rec["tier"]
        if t not in tier_summary:
            tier_summary[t] = {"ok": 0, "fail": 0, "ttfts": [], "decode_tpss": [], "tpots": []}
        if rec["status"] == "ok":
            tier_summary[t]["ok"] += 1
            tier_summary[t]["ttfts"].append(rec["ttft"])
            tier_summary[t]["decode_tpss"].append(rec["decode_tps"])
            tier_summary[t]["tpots"].append(rec["tpot_ms"])
        else:
            tier_summary[t]["fail"] += 1

    tier_stats = []
    for t_name in ["P50", "AVG", "P90", "P99"]:
        ts = tier_summary.get(t_name)
        if not ts:
            continue
        n = max(ts["ok"], 1)
        tier_stats.append({
            "tier": t_name,
            "ok": ts["ok"], "fail": ts["fail"],
            "ttft_avg": round(sum(ts["ttfts"]) / n, 3) if ts["ttfts"] else 0,
            "ttft_max": round(max(ts["ttfts"]), 3) if ts["ttfts"] else 0,
            "decode_tps_avg": round(sum(ts["decode_tpss"]) / n, 1) if ts["decode_tpss"] else 0,
            "tpot_avg_ms": round(sum(ts["tpots"]) / n, 1) if ts["tpots"] else 0,
        })

    passed = ok_count > 0
    print(f"         [朴素扫描] 完成: {ok_count}/{total_steps} 步成功, "
          f"{fail_count} 步失败, 共 {total_requests} 条请求, 耗时 {elapsed}s")

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "mode": "io_sweep",
            "repetitions": reps_original,
            "total_steps": total_steps,
            "total_requests": total_requests,
            "ok": ok_count,
            "fail": fail_count,
            "elapsed": elapsed,
            "sweep": records,
            "tier_stats": tier_stats,
            "verdict": (
                f"{mode_label}: {ok_count}/{total_steps} 步成功, "
                f"共 {total_requests} 条请求, 耗时 {elapsed}s"
            ),
        },
    }


def test_io_tier_benchmark(cfg: dict, case: dict) -> dict:
    """TC-10: 输入/输出分档位性能测试。

    对 params.tiers 中定义的多个 (input_tokens, output_tokens) 档位，
    分别用流式请求测量 TTFT / Decode TPS / TPOT / ITL 等指标，
    输出每个档位的吞吐与延迟对比，帮助定位「长输入 + 长输出」组合下的拐点。

    params 结构:
      {
        "concurrency": 5,            # 每档并发数（可选，默认取 benchmark.concurrency）
        "total_requests": 8,         # 每档总请求数（可选，默认 10）
        "tiers": [                   # 档位列表（必填）
          {"name": "1k-in/0.3k-out", "input_tokens": 1000, "output_tokens": 300},
          {"name": "8k-in/1k-out",   "input_tokens": 8000, "output_tokens": 1000},
          {"name": "32k-in/2k-out",  "input_tokens": 32000, "output_tokens": 2000}
        ]
      }

    通过条件: 所有档位均有成功请求（ok > 0）。
    """
    api = cfg["api"]
    params = case.get("params", {})
    tiers = params.get("tiers", [])
    if not tiers:
        return {
            "case_id": case["id"],
            "passed": False,
            "detail": {
                "error": "params.tiers 为空，请至少定义一个 (input_tokens, output_tokens) 档位",
                "tiers": [],
            },
        }

    global_concurrency = params.get("concurrency")  # None 表示各档用内置默认
    bench_concurrency = cfg["benchmark"].get("concurrency", 4)
    total = params.get("total_requests") or 10
    timeout = api.get("timeout", 120)

    tier_results = []

    tier_names = ", ".join(
        t.get("name") or f"{t.get('input_tokens')}in/{t.get('output_tokens')}out"
        for t in tiers
    )
    print(f"         [IO档位] 共 {len(tiers)} 档: {tier_names}")

    for tier_idx, tier in enumerate(tiers):
        in_tok = int(tier.get("input_tokens", 1000))
        out_tok = min(int(tier.get("output_tokens", 300)), 8192)
        name = tier.get("name", f"{in_tok}in/{out_tok}out")
        # 每档并发：CLI 显式值 > 档位内置默认 > benchmark 全局默认
        tier_concurrency = _resolve_tier_concurrency(tier, global_concurrency, bench_concurrency)

        messages = _gen_benchmark_messages(in_tok, user_question=(
            f"Write a detailed technical answer of about {max(200, out_tok // 2)} words "
            f"covering the topic above. Be thorough and structured. [tier:{name}]"
        ))

        # 每档指标的收集容器
        per_ttft, per_tpot, per_decode_tps = [], [], []
        stats = {
            "ok": 0, "fail": 0,
            "total_ttft": 0.0, "total_decode_time": 0.0,
            "total_prompt_tokens": 0, "total_completion_tokens": 0,
            "total_tokens": 0, "total_itl": 0.0, "itl_count": 0,
            "itl_over_500ms": 0, "errors": [],
        }

        def _worker(_idx):
            varied = [dict(m) for m in messages]
            varied[-1]["content"] += f"\n\n[req:{_idx}]"
            r = api_request_stream(
                url=api["url"], key=api["key"], model=api["model"],
                messages=varied, max_tokens=out_tok, timeout=timeout,
            )
            with _stats_lock:
                if r["ok"]:
                    stats["ok"] += 1
                    ttft = r.get("ttft", 0)
                    total_lat = r.get("total_latency", 0)
                    decode_t = total_lat - ttft if total_lat > ttft else 0
                    completion = r.get("completion_tokens", 0)
                    stats["total_ttft"] += ttft
                    stats["total_decode_time"] += decode_t
                    stats["total_prompt_tokens"] += r.get("prompt_tokens", 0)
                    stats["total_completion_tokens"] += completion
                    stats["total_tokens"] += r.get("total_tokens", 0)

                    per_ttft.append(ttft)
                    per_tpot.append(round(decode_t / completion * 1000, 2) if completion > 0 else 0)
                    per_decode_tps.append(round(completion / decode_t, 1) if decode_t > 0 else 0)

                    times = r.get("token_times", [])
                    if len(times) >= 2:
                        for i in range(1, len(times)):
                            itl = times[i] - times[i - 1]
                            stats["total_itl"] += itl
                            stats["itl_count"] += 1
                            if itl > 0.5:
                                stats["itl_over_500ms"] += 1
                else:
                    stats["fail"] += 1
                    if r.get("error") and len(stats["errors"]) < 10:
                        stats["errors"].append(r["error"])

        print(f"         [{tier_idx+1}/{len(tiers)}] {name} (in={in_tok}, out≤{out_tok}, "
              f"并发={tier_concurrency}, 请求={total}) ... ", end="", flush=True)

        t_start = time.perf_counter()
        with ThreadPoolExecutor(max_workers=tier_concurrency) as pool:
            futures = [pool.submit(_worker, i) for i in range(total)]
            for f in as_completed(futures):
                f.result()
        elapsed = round(time.perf_counter() - t_start, 2)

        n_ok = stats["ok"] if stats["ok"] > 0 else 1
        ttft_avg = round(stats["total_ttft"] / n_ok, 3)
        decode_time_avg = round(stats["total_decode_time"] / n_ok, 3)
        total_lat_avg = round(ttft_avg + decode_time_avg, 3)
        decode_tps = round(stats["total_completion_tokens"] / stats["total_decode_time"], 1) if stats["total_decode_time"] > 0 else 0
        prefill_tps = round(stats["total_prompt_tokens"] / stats["total_ttft"], 1) if stats["total_ttft"] > 0 else 0
        tpot_avg = round(stats["total_decode_time"] / stats["total_completion_tokens"] * 1000, 1) if stats["total_completion_tokens"] > 0 else 0
        itl_avg = round(stats["total_itl"] / stats["itl_count"] * 1000, 1) if stats["itl_count"] > 0 else 0

        ttft_pct = _percentiles(per_ttft, 50, 90, 99)
        tpot_pct = _percentiles(per_tpot, 50, 99)

        status = "PASS" if stats["ok"] > 0 else "FAIL"
        print(f"{status} (耗时{elapsed}s, "
              f"TTFT_avg={ttft_avg}s, Decode={decode_tps}tok/s, "
              f"TPOT_avg={tpot_avg}ms, ITL>500ms={stats['itl_over_500ms']}次)")

        tier_results.append({
            "name": name,
            "input_tokens": in_tok,
            "output_tokens": out_tok,
            "concurrency": tier_concurrency,
            "total": total,
            "ok": stats["ok"],
            "fail": stats["fail"],
            "elapsed": elapsed,
            "ttft_avg": ttft_avg,
            "ttft_p50": ttft_pct["p50"],
            "ttft_p90": ttft_pct["p90"],
            "ttft_p99": ttft_pct["p99"],
            "total_lat_avg": total_lat_avg,
            "prefill_tps": prefill_tps,
            "decode_tps": decode_tps,
            "tpot_avg_ms": tpot_avg,
            "tpot_p50_ms": tpot_pct["p50"],
            "tpot_p99_ms": tpot_pct["p99"],
            "itl_avg_ms": itl_avg,
            "itl_over_500ms": stats["itl_over_500ms"],
            "total_tokens": stats["total_tokens"],
            "total_prompt_tokens": stats["total_prompt_tokens"],
            "total_completion_tokens": stats["total_completion_tokens"],
            "errors": stats["errors"][:5],
        })

    all_ok = all(t["ok"] > 0 for t in tier_results)
    passed = all_ok

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "mode": "io_tier",
            "concurrency": "per_tier",
            "total_requests": total,
            "tiers": tier_results,
            "passed_count": f"{sum(1 for t in tier_results if t['ok'] > 0)}/{len(tier_results)}",
            "verdict": (
                f"分档位性能测试完成: {sum(1 for t in tier_results if t['ok'] > 0)}/{len(tier_results)} 档成功, "
                f"各档 Decode TPS 见详情"
                if passed else
                f"存在失败档位({sum(1 for t in tier_results if t['ok'] == 0)} 档 ok=0)"
            ),
        },
    }


def _gen_io_mix_requests(n: int, in_anchors: list, out_anchors: list) -> list:
    """按分位数锚点线性插值生成 n 条 (input_tokens, output_tokens) 请求。

    anchors: [(percentile_rank 0~1, tokens), ...] 升序
    """
    def lerp(anchors, r):
        if r <= anchors[0][0]:
            return anchors[0][1]
        if r >= anchors[-1][0]:
            return anchors[-1][1]
        for i in range(len(anchors) - 1):
            r0, v0 = anchors[i]
            r1, v1 = anchors[i + 1]
            if r0 <= r <= r1:
                return v0 + (v1 - v0) * (r - r0) / (r1 - r0)
        return anchors[-1][1]

    reqs = []
    for i in range(1, n + 1):
        r = (i - 0.5) / n
        inp = int(round(lerp(in_anchors, r)))
        outp = int(round(lerp(out_anchors, r)))
        reqs.append({"seq": i, "input_tokens": max(1, inp), "output_tokens": max(1, outp)})
    return reqs


def test_io_mix_benchmark(cfg: dict, case: dict) -> dict:
    """TC-10 混合分布模式：按线上流量分位数生成 N 条请求（输入/输出长度），
    逐条下发流式请求，测量每条 TTFT/Decode TPS/TPOT/ITL，最后输出整体分位数，
    用于对照你的 SLA 指标。

    params 结构:
      {
        "concurrency": 5,            # 并发数（可选，默认 benchmark.concurrency）
        "total_requests": 100,       # 总请求数（可选，默认 100）
        "in_anchors": [[0.0,2000],[0.5,50000],[0.9,160000],[0.99,380000],[1.0,380000]],
        "out_anchors": [[0.0,50],[0.5,200],[0.9,1300],[0.99,7000],[1.0,7000]],
        "max_output": 8192           # 单请求输出上限保护（可选）
      }
    """
    api = cfg["api"]
    params = case.get("params", {})
    n = int(params.get("total_requests") or 100)
    concurrency = int(params.get("concurrency") or cfg["benchmark"].get("concurrency", 4))
    timeout = api.get("timeout", 120)
    max_out = int(params.get("max_output", 8192))
    max_ctx = params.get("max_context")

    # 默认锚点：对齐用户给定的分位数
    in_anchors = params.get("in_anchors") or [
        [0.0, 2000], [0.5, 50000], [0.9, 160000], [0.99, 380000], [1.0, 380000],
    ]
    out_anchors = params.get("out_anchors") or [
        [0.0, 50], [0.5, 200], [0.9, 1300], [0.99, 7000], [1.0, 7000],
    ]

    reqs = _gen_io_mix_requests(n, in_anchors, out_anchors)

    # ── 最大上下文过滤：输入长度超过上限的请求跳过（不测）──
    max_ctx = params.get("max_context")
    skipped = 0
    if max_ctx:
        kept = []
        for req in reqs:
            # 预留输出 token 余量，避免拼上输出后超上下文
            if req["input_tokens"] + req["output_tokens"] <= max_ctx:
                kept.append(req)
            else:
                skipped += 1
                req["status"] = "skipped"
                req["skip_reason"] = f"输入+输出超出上下文上限 {max_ctx}"
                req["ttft"] = 0.0
                req["decode_tps"] = 0.0
                req["tpot_ms"] = 0.0
                req["itl_max_ms"] = 0.0
                req["completion_tokens"] = 0
        reqs = kept
    print(f"         [IO混合] 共 {n} 条请求（输入/输出按分位数线性插值）"
          + (f"，跳过 {skipped} 条（超上下文上限 {max_ctx}）" if skipped else ""))

    def _worker(req):
        in_tok = req["input_tokens"]
        out_tok = min(req["output_tokens"], max_out)
        messages = _gen_benchmark_messages(in_tok, user_question=(
            f"Write a detailed technical answer of about {max(200, out_tok // 2)} words "
            f"covering the topic above. Be thorough and structured. [seq:{req['seq']}]"
        ))
        varied = [dict(m) for m in messages]
        varied[-1]["content"] += f"\n\n[req:{req['seq']}]"
        r = api_request_stream(
            url=api["url"], key=api["key"], model=api["model"],
            messages=varied, max_tokens=out_tok, timeout=timeout,
        )
        rec = {
            "seq": req["seq"], "input_tokens": in_tok, "output_tokens": out_tok,
            "status": "ok" if r["ok"] else "fail",
            "error": r.get("error"),
            "ttft": 0.0, "decode_tps": 0.0, "tpot_ms": 0.0, "itl_max_ms": 0.0,
            "completion_tokens": 0,
        }
        if r["ok"]:
            ttft = r.get("ttft", 0)
            total_lat = r.get("total_latency", 0)
            decode_d = total_lat - ttft if total_lat > ttft else 0
            completion = r.get("completion_tokens", 0)
            rec["ttft"] = ttft
            rec["completion_tokens"] = completion
            rec["decode_tps"] = round(completion / decode_d, 1) if decode_d > 0 else 0
            rec["tpot_ms"] = round(decode_d / completion * 1000, 2) if completion > 0 else 0
            times = r.get("token_times", [])
            rec["itl_max_ms"] = round((max(times) - min(times)) * 1000, 1) if len(times) >= 2 else 0
        return rec

    t_start = time.perf_counter()
    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = [pool.submit(_worker, req) for req in reqs]
        for f in as_completed(futures):
            f.result()
    elapsed = round(time.perf_counter() - t_start, 2)

    per_req = [f.result() for f in futures]
    per_req.sort(key=lambda x: x["seq"])

    ok_recs = [r for r in per_req if r["status"] == "ok"]
    skipped_recs = [r for r in per_req if r["status"] == "skipped"]
    fail_count = len(per_req) - len(ok_recs) - len(skipped_recs)

    ttfts = [r["ttft"] for r in ok_recs]
    decode_tpss = [r["decode_tps"] for r in ok_recs]
    tpots = [r["tpot_ms"] for r in ok_recs]
    itl_maxs = [r["itl_max_ms"] for r in ok_recs]

    ttft_pct = _percentiles(ttfts, 50, 75, 90, 99)
    decode_pct = _percentiles(decode_tpss, 50, 99)
    tpot_pct = _percentiles(tpots, 50, 75, 90, 99)
    itl_pct = _percentiles(itl_maxs, 50, 90, 99)

    passed = len(ok_recs) > 0
    verdict = (
        f"混合分布 {n} 条请求完成: 成功 {len(ok_recs)}/{len(per_req) if not skipped_recs else n}, "
        f"跳过 {len(skipped_recs)}/{n}（超上下文）"
        if skipped_recs else
        f"混合分布 {n} 条请求完成: 成功 {len(ok_recs)}/{n}, "
        f"TTFT p50={ttft_pct['p50']}s p99={ttft_pct['p99']}s, "
        f"Decode p50={decode_pct['p50']}tok/s, TPOT p99={tpot_pct['p99']}ms"
    )

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "mode": "io_mix",
            "total_requests": n,
            "concurrency": concurrency,
            "elapsed": elapsed,
            "ok": len(ok_recs),
            "fail": fail_count,
            "skipped": len(skipped_recs),
            "max_context": max_ctx,
            "in_anchors": in_anchors,
            "out_anchors": out_anchors,
            "ttft_p50": ttft_pct["p50"], "ttft_p75": ttft_pct["p75"],
            "ttft_p90": ttft_pct["p90"], "ttft_p99": ttft_pct["p99"],
            "decode_tps_p50": decode_pct["p50"], "decode_tps_p99": decode_pct["p99"],
            "tpot_p50_ms": tpot_pct["p50"], "tpot_p75_ms": tpot_pct["p75"],
            "tpot_p90_ms": tpot_pct["p90"], "tpot_p99_ms": tpot_pct["p99"],
            "itl_max_p50_ms": itl_pct["p50"], "itl_max_p90_ms": itl_pct["p90"], "itl_max_p99_ms": itl_pct["p99"],
            "requests": per_req,
            "verdict": verdict,
        },
    }


# ---------------------------------------------------------------------------
# CSV 输出 — 单文件：报告区 + 用例表
# ---------------------------------------------------------------------------

def write_single_csv(results: list, cases: list, cfg: dict, output_path: Path) -> None:
    """
    输出单个 CSV，分成上下两部分：
      【报告区】2 列 — 对齐 template.xlsx「测试报告」sheet
      【空行分隔】
      【用例表】7 列 — 对齐 template.xlsx「测试用例」sheet

    与模板结构一一对应，方便在一份文件中阅读。
    """
    rows = []

    # ═══════════════════════════════════════════════════════════
    # 第一部分：报告区（2 列：标签 | 值）
    # ═══════════════════════════════════════════════════════════
    api = cfg["api"]

    # 汇总数据
    total_cases = len(results)
    passed_cases = sum(1 for r in results if r.get("passed"))
    failed_cases = total_cases - passed_cases
    total_tokens = 0
    total_latency = 0.0
    latency_count = 0
    context_passed = "否"
    context_error = ""
    estimated_input = "N/A"

    for r in results:
        d = r.get("detail", {})
        if isinstance(d, dict):
            if "total_tokens" in d:
                total_tokens += int(d["total_tokens"])
            if "latency" in d:
                try:
                    total_latency += float(d["latency"])
                    latency_count += 1
                except (ValueError, TypeError):
                    pass
            if r.get("case_id") == "TC-04":
                if d.get("mode") == "multi_tier":
                    context_passed = d.get("passed_count", "0/?")
                    tiers_detail = ", ".join(
                        f"{t['target']//1000}k={'✓' if t.get('ok') else '✗'}"
                        for t in d.get("tiers", [])
                    )
                    context_error = tiers_detail
                    estimated_input = ", ".join(
                        f"{t.get('estimated_input_tokens', '?')}"
                        for t in d.get("tiers", [])
                    )
                else:
                    context_passed = "是" if r.get("passed") else "否"
                    if d.get("failure_reason"):
                        context_error = str(d["failure_reason"])[:200]
                    elif d.get("error"):
                        context_error = str(d["error"])[:200]
                    estimated_input = str(d.get("estimated_input_tokens", "N/A"))

    avg_latency = round(total_latency / latency_count, 3) if latency_count else "N/A"
    success_rate = f"{passed_cases / total_cases * 100:.1f}%" if total_cases else "N/A"
    test_time = datetime.now(TZ_BJ).strftime("%Y-%m-%d %H:%M:%S")
    conclusion = "通过" if failed_cases == 0 else f"{failed_cases}/{total_cases} 未通过"

    tps_val = "N/A"
    tpm_val = "N/A"
    decode_tps_val = "N/A"
    decode_tpm_val = "N/A"
    for r in results:
        if r.get("case_id") == "TC-02" and r.get("detail"):
            d = r["detail"]
            tps_val = d.get("tps_tokens", "N/A")
            tpm_val = d.get("tpm_tokens", "N/A")
            decode_tps_val = d.get("decode_tps", "N/A")
            decode_tpm_val = d.get("decode_tpm", "N/A")
            decode_tps_val = d.get("decode_tps", "N/A")
            decode_tpm_val = d.get("decode_tpm", "N/A")

    key_display = api.get("key", "")
    if len(key_display) > 20:
        key_display = key_display[:20] + "..."

    report_section = [
        ["Token 接口测试报告", ""],
        ["（由 token_test.py 自动生成）", ""],
        ["", ""],
        ["被测接口 URL", api.get("url", "")],
        ["API Key", key_display],
        ["模型名", api.get("model", "")],
        ["TPS (tokens/秒)", tps_val],
        ["TPM (tokens/分钟)", tpm_val],
        ["请求成功率", success_rate],
        ["生成 Token 总数", str(total_tokens)],
        ["平均延迟 (秒)", str(avg_latency)],
        ["压测并发数", str(cfg.get("benchmark", {}).get("concurrency", ""))],
        ["压测总请求数", str(cfg.get("benchmark", {}).get("total_requests", ""))],
        ["上下文验收档位 (tokens)", str(cfg.get("context_test", {}).get("target_tokens", []))],
        ["各档位实测 Token 估算", str(estimated_input)],
        ["上下文测试结果", context_passed],
        ["各档位详情", context_error or "无"],
        ["测试时间", test_time],
        ["测试结论", conclusion],
    ]
    rows.extend(report_section)

    # ═══════════════════════════════════════════════════════════
    # 分隔行
    # ═══════════════════════════════════════════════════════════
    rows.append(["", ""])

    # ═══════════════════════════════════════════════════════════
    # 第二部分：用例表（7 列）
    # ═══════════════════════════════════════════════════════════
    result_map = {r["case_id"]: r for r in results}
    variables = {
        "concurrency": cfg.get("benchmark", {}).get("concurrency", "N/A"),
        "total_requests": cfg.get("benchmark", {}).get("total_requests", "N/A"),
    }

    table_section = [["编号", "用例名称", "测试类型", "输入/条件", "预期结果", "实际结果", "是否通过"]]

    for case in cases:
        cid = case["id"]
        result = result_map.get(cid, {})

        input_str = resolve_template(case.get("input", ""), variables)
        expected_str = resolve_template(case.get("expected", ""), variables)

        actual = "未执行"
        passed_str = "未执行"

        if result:
            detail = result.get("detail", {})
            passed = result.get("passed", False)
            passed_str = "通过" if passed else "失败"

            if cid == "TC-01":
                actual = (
                    f"HTTP {detail.get('status_code')}, "
                    f"延迟 {detail.get('latency', 'N/A')}s, "
                    f"Usage tokens: {(detail.get('usage') or {}).get('total', 'N/A')}"
                )
                if not passed and detail.get("error"):
                    actual += f", 错误: {detail['error']}"
            elif cid == "TC-02":
                if detail.get("mode") == "gradient":
                    lvs = detail.get("levels", [])
                    best_c = detail.get("best_concurrency", "?")
                if detail.get("mode") == "gradient":
                    lvs = detail.get("levels", [])
                    best_c = detail.get("best_concurrency", "?")
                    actual = (f"梯度{lvs[0]["concurrency"]}→{lvs[-1]["concurrency"]}({len(lvs)}级), "
                              f"最高TPS={detail.get("tps_tokens")}tok/s(并发={best_c}), "
                              f"Decode={detail.get("decode_tps")}tok/s, "
                              f"总成功/失败={detail.get("ok")}/{detail.get("fail")}")
                else:
                    actual = (
                        f"成功{detail.get("ok")}/{detail.get("ok",0)+detail.get("fail",0)}, "
                        f"TPS={detail.get("tps_tokens")}tok/s, "
                        f"TPM={detail.get("tpm_tokens")}tok/min, "
                        f"Decode={detail.get("decode_tps","N/A")}tok/s, "
                        f"平均延迟{detail.get("avg_latency")}s"
                    )
                    ec = detail.get("error_counts", {})
                    if ec:
                        parts = [f"{c}×{n}" for c, n in sorted(ec.items(), key=lambda x: -x[1])[:3]]
                        actual += f", 失败:{", ".join(parts)}"
            elif cid == "TC-03":
                actual = (
                    f"TPS={detail.get('tps')}, TPM={detail.get('tpm')}, "
                    f"公式验证: {'通过' if detail.get('formula_ok') else '失败'}"
                )
                if detail.get("error"):
                    actual += f", {detail['error']}"
            elif cid == "TC-04":
                mode = detail.get("mode", "")
                if mode == "declared":
                    actual = (
                        f"模型声明 {detail.get('declared_context_length')} tokens, "
                        f"阈值 {detail.get('targets')}, {detail.get('reason', '')}"
                    )
                elif mode == "probe":
                    actual = (
                        f"探测上限: {detail.get('found_limit')} tokens, "
                        f"目标 {detail.get('targets')}, "
                        f"共 {detail.get('attempts_count', 0)} 次尝试"
                    )
                elif mode == "multi_tier":
                    parts = [f"通过 {detail.get('passed_count')} 档"]
                    for t in detail.get("tiers", []):
                        status = "✓" if t.get("ok") else "✗"
                        parts.append(f"{t['target']//1000}k:{status}")
                    actual = ", ".join(parts)
                else:
                    actual = (
                        f"HTTP {detail.get('status_code')}, "
                        f"延迟 {detail.get('latency', 'N/A')}s, "
                    )
                    if detail.get("failure_reason"):
                        actual += f"{detail['failure_reason'][:150]}"
                    elif detail.get("usage"):
                        actual += f"usage={detail['usage'].get('total', 'N/A')} tokens"
                    else:
                        actual += f"error: {(detail.get('error') or 'N/A')[:100]}"
            elif cid == "TC-05":
                actual = (
                    f"HTTP {detail.get('status_code')}, "
                    f"预期 401/403, {'符合预期' if passed else '不符合预期'}"
                )
                if detail.get("note"):
                    actual += f", {detail['note']}"
            elif cid == "TC-06":
                actual = (
                    f"OK={detail.get('ok')}, 429={detail.get('rate_limited_429')}, "
                    f"超时={detail.get('timeout')}, "
                    f"耗时 {detail.get('elapsed')}s"
                )
                if detail.get("note"):
                    actual += f", {detail['note']}"
            elif cid == "TC-07":
                actual = (
                    f"闭环{'通过' if passed else '失败'}, "
                    f"tool={detail.get('tool_name') or 'N/A'}, "
                    f"args={json.dumps(detail.get('arguments', {}), ensure_ascii=False)[:60]}, "
                    f"结果已利用={'是' if detail.get('used_tool_result') else '否'}, "
                    f"HTTP {detail.get('status_code')}, "
                    f"延迟 {detail.get('latency', 'N/A')}s"
                )
            elif cid == "TC-08":
                actual = (
                    f"成功 {detail.get('ok')}/{detail.get('ok',0)+detail.get('fail',0)}, "
                    f"TTFT avg={detail.get('ttft_avg')}s P50={detail.get('ttft_p50')}s P99={detail.get('ttft_p99')}s, "
                    f"Prefill={detail.get('prefill_tps')} tok/s, "
                    f"Decode={detail.get('decode_tps')} tok/s, "
                    f"TPOT avg={detail.get('tpot_avg_ms')}ms P50={detail.get('tpot_p50')}ms P99={detail.get('tpot_p99')}ms, "
                    f"ITL avg={detail.get('itl_avg_ms')}ms, "
                    f"ITL>500ms={detail.get('itl_over_500ms_count')}次({detail.get('itl_over_500ms_requests_pct')}请求), "
                    f"Cache命中={detail.get('cache_hit_rate')}"
                )
            elif cid == "TC-09":
                warmup = detail.get("warmup", {}) or {}
                actual = (
                    f"缓存命中率: {detail.get('cache_hit_rate_pct', 'N/A')}%, "
                    f"TPM(缓存): {detail.get('tpm_tokens', 'N/A')} tok/min, "
                    f"TTFT cold={warmup.get('ttft', 'N/A')}s→avg={detail.get('ttft_avg', 'N/A')}s, "
                    f"cached={detail.get('total_cached_tokens', 0)}/{detail.get('total_prompt_tokens', 0)} tokens, "
                    f"成功{detail.get('ok', 0)}/{detail.get('total_requests', 0)}"
                )
                if detail.get("verdict"):
                    actual += f" | {detail['verdict'][:120]}"
            elif cid == "TC-10":
                mode = detail.get("mode", "io_tier")
                if mode == "io_mix":
                    d = detail
                    actual = (
                        f"混合 {d.get('total_requests')} 条: 成功{d.get('ok')}/{d.get('total_requests')}, "
                        f"TTFT p50={d.get('ttft_p50')}s p90={d.get('ttft_p90')}s p99={d.get('ttft_p99')}s, "
                        f"Decode p50={d.get('decode_tps_p50')}tok/s p99={d.get('decode_tps_p99')}tok/s, "
                        f"TPOT p50={d.get('tpot_p50_ms')}ms p99={d.get('tpot_p99_ms')}ms, "
                        f"ITLmax p99={d.get('itl_max_p99_ms')}ms"
                    )
                else:
                    tiers = detail.get("tiers", [])
                    parts = [f"通过 {detail.get('passed_count', '?')} 档"]
                    for t in tiers:
                        status = "✓" if t.get("ok", 0) > 0 else "✗"
                        parts.append(
                            f"{t.get('name')}:{status} TTFT={t.get('ttft_avg')}s "
                            f"Decode={t.get('decode_tps')}tok/s TPOT={t.get('tpot_avg_ms')}ms"
                        )
                    actual = "; ".join(parts)
            else:
                actual = json.dumps(detail, ensure_ascii=False)

        table_section.append([cid, case["name"], case["type"], input_str, expected_str, actual, passed_str])

    rows.extend(table_section)

    # ── 写入 ──
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

    print(f"[OK] 结果已写入: {output_path}")
# ---------------------------------------------------------------------------
# HTML 验收报告模板 — 从 Markdown 转换并注入美观样式
# ---------------------------------------------------------------------------

def write_markdown_report(results, cases, cfg, args, output_path, test_time):
    """生成简洁 Markdown 验收报告。"""
    api = cfg["api"]
    bench = cfg.get("benchmark", {})
    grad = bench.get("gradient", {})
    inp = bench.get("input_tokens", 1000)
    out_max = min(500, max(300, inp // 3))
    tc = sum(1 for r in results if r.get("passed"))
    fc = len(results) - tc
    sr = f"{tc / len(results) * 100:.1f}%" if results else "N/A"
    cn = "全部通过" if fc == 0 else f"{fc}/{len(results)} 未通过"
    cmd = "python token_test.py " + " ".join(f'"{a}"' if " " in a else a for a in sys.argv[1:])

    lines = []
    lines.append(f"# Token 接口测试报告 — {test_time} — {'✅' if fc==0 else '❌'}{cn}")
    lines.append(f"```bash\n{cmd}\n```")
    plat = getattr(args, 'platform', '')
    plat_str = f" | **平台**：{plat}" if plat else ""
    lines.append(f"**模型**：{api.get('model', '')} | **超时**：{api.get('timeout', '')}s | **输入/输出**：~{inp}/≤{out_max}tok(TC-09:≤50) | **梯度**：{grad.get('start','')}→{grad.get('max','')}步长{grad.get('step','')}{plat_str} | **通过{tc}/{len(results)}**({sr})")
    lines.append("")

    rm = {r["case_id"]: r for r in results}
    for case in cases:
        cid = case["id"]
        r = rm.get(cid, {})
        if not r:
            lines.append(f"### {cid}{case['name']}⚪\n*未执行*\n")
            continue
        p = r.get("passed", False)
        d = r.get("detail", {}) or {}
        b = "✅" if p else "❌"
        lines.append(f"### {cid}{case['name']}{b}")

        if cid == "TC-01":
            u = d.get("usage") or {}
            lines.append(f"HTTP{d.get('status_code')}|{d.get('latency','N/A')}s|{u.get('total','N/A')}tokens\n")
        elif cid == "TC-02":
            lvs = d.get("levels", [])
            if lvs:
                best_c = d.get("best_concurrency", "?")
                at = lvs[0].get("total_tokens", 0) // max(lvs[0].get("ok", 1), 1) if lvs else 0
                lines.append(f"最高TPS**{d.get('tps_tokens')}**tok/s(并发={best_c})|Decode**{d.get('decode_tps')}**tok/s|均≈{at}tok/请求")
                lines.append("|并发|成功|失败|耗时|TPS|TPS/并|Dec TPS|Prefill TPS|总Tok|输入Tok|输出Tok|TTFT|延迟|错误|")
                lines.append("|------|------|------|------|-----|-----|---------|-----------|-----|-----|-----|----|------|------|")
                for lv in lvs:
                    ec = lv.get("error_counts", {})
                    es = ", ".join(f"{c}×{n}" for c, n in sorted(ec.items(), key=lambda x: -x[1])[:2]) or "-"
                    lines.append(f"|{lv['concurrency']}|{lv['ok']}|{lv['fail']}|{lv['elapsed']}s|{lv['tps_tokens']}|{lv.get('tps_per_c','N/A')}|{lv['decode_tps']}|{lv.get('prefill_tps','N/A')}|{lv['total_tokens']}|{lv['total_prompt_tokens']}|{lv['total_completion_tokens']}|{lv.get('avg_ttft','N/A')}s|{lv['avg_latency']}s|{es}|")
                lines.append("")
        elif cid == "TC-03":
            lines.append(f"TC-02 TPS={d.get('tps')}→TPM={d.get('tpm')}|TPM=TPS×60:{'✅正确' if d.get('formula_ok') else '❌不符'}\n")
        elif cid == "TC-04":
            mode = d.get("mode", "")
            if mode == "multi_tier":
                lines.append(f"通过{d.get('passed_count','?')}")
                lines.append("|档位|状态|延迟|原因|")
                lines.append("|------|------|------|------|")
                for t in d.get("tiers", []):
                    lines.append(f"|{t['target']//1000}k|{'✅' if t.get('ok') else '❌'}|{t.get('latency','')}s|{(t.get('failure_reason') or t.get('error') or '—')[:80]}|")
                lines.append("")
            elif mode in ("declared", "probe"):
                lines.append(f"{d.get('declared_context_length','') or d.get('found_limit','')}tokens→{'✅' if p else '❌'}\n")
            else:
                lines.append(f"HTTP{d.get('status_code')}|{(d.get('failure_reason') or d.get('error') or '无')[:120]}\n")
        elif cid == "TC-05":
            lines.append("无需鉴权，跳过\n" if d.get("status_code") is None else f"HTTP{d.get('status_code')}|{'✅预期' if p else '❌'}\n")
        elif cid == "TC-06":
            lines.append(f"OK={d.get('ok')}429={d.get('rate_limited_429')}超时={d.get('timeout')}耗时={d.get('elapsed')}s\n")
        elif cid == "TC-07":
            lines.append(f"{'✅' if p else '❌'}{d.get('tool_name') or 'N/A'}|HTTP{d.get('status_code')}|{d.get('latency','N/A')}s|{d.get('verdict','')[:120]}\n")
        elif cid == "TC-08":
            w = " ⚠️成功率<50%不可靠" if not d.get("reliable", True) else ""
            lines.append(f"ok={d.get('ok')}fail={d.get('fail')}|{d.get('elapsed')}s|入={d.get('total_prompt_tokens','N/A')}出={d.get('total_completion_tokens','N/A')}tok|TTFT avg={d.get('ttft_avg')}s P99={d.get('ttft_p99')}s|Decode={d.get('decode_tps')}tok/s|TPOT={d.get('tpot_avg_ms')}ms{w}")
            ec = d.get("error_counts", {})
            if ec: lines.append(f"失败:{', '.join(f'{c}×{n}' for c,n in sorted(ec.items(),key=lambda x:-x[1])[:3])}")
            lines.append("")
        elif cid == "TC-09":
            w = " ⚠️成功率<50%不可靠" if not d.get("reliable", True) else ""
            wu = d.get("warmup", {}) or {}
            lines.append(f"命中率{d.get('cache_hit_rate_pct','N/A')}%|TPM{d.get('tpm_tokens','N/A')}tok/min|prompt={d.get('total_prompt_tokens','N/A')}compl={d.get('total_completion_tokens','N/A')}|预热prompt={wu.get('prompt_tokens','N/A')}|TTFT预热={wu.get('ttft','N/A')}s→并发={d.get('ttft_avg','N/A')}s|ok={d.get('ok',0)}fail={d.get('fail',0)}{w}")
            ec = d.get("error_counts", {})
            if ec: lines.append(f"失败:{', '.join(f'{c}×{n}' for c,n in sorted(ec.items(),key=lambda x:-x[1])[:3])}")
            lines.append("")
        elif cid == "TC-10":
            mode = d.get("mode", "io_tier")
            if mode == "io_sweep":
                reps = d.get("repetitions", 1)
                mode_tag = f"×{reps}" if reps > 1 else ("全并发" if reps == 0 else "串行")
                lines.append(f"朴素扫描({mode_tag}){d.get('total_steps')}步 {d.get('total_requests','?')}条请求|成功{d.get('ok')}/{d.get('total_steps')}|耗时{d.get('elapsed','N/A')}s")
                lines.append("|#|in|out|档位|状态|TTFT|DecodeTPS|TPOT|")
                lines.append("|--|--|--|--|--|--|--|--|")
                for rq in d.get("sweep", []):
                    st = "✅" if rq.get("status") == "ok" else "❌"
                    lines.append(f"|{rq.get('seq')}|{rq.get('input_tokens')}|{rq.get('output_tokens')}|{rq.get('tier','?')}|{st}|{rq.get('ttft')}|{rq.get('decode_tps')}|{rq.get('tpot_ms')}|")
                lines.append("")
                if d.get("tier_stats"):
                    lines.append("|档位|ok|fail|TTFT avg|TTFT max|Decode avg|TPOT avg|")
                    lines.append("|------|------|------|------|------|------|------|")
                    for ts in d.get("tier_stats", []):
                        lines.append(f"|{ts['tier']}|{ts['ok']}|{ts['fail']}|{ts['ttft_avg']}s|{ts['ttft_max']}s|{ts['decode_tps_avg']}|{ts['tpot_avg_ms']}ms|")
                    lines.append("")
            elif mode == "io_mix":
                lines.append(f"混合{d.get('total_requests')}条|成功{d.get('ok')}/{d.get('total_requests')}|并发={d.get('concurrency','N/A')}|耗时{d.get('elapsed','N/A')}s")
                lines.append(f"**TTFT**: p50={d.get('ttft_p50')}s p75={d.get('ttft_p75')}s p90={d.get('ttft_p90')}s p99={d.get('ttft_p99')}s")
                lines.append(f"**Decode TPS**: p50={d.get('decode_tps_p50')} p99={d.get('decode_tps_p99')} | **TPOT**: p50={d.get('tpot_p50_ms')}ms p99={d.get('tpot_p99_ms')}ms")
                lines.append(f"**ITL max**: p50={d.get('itl_max_p50_ms')}ms p90={d.get('itl_max_p90_ms')}ms p99={d.get('itl_max_p99_ms')}ms")
                lines.append("|#|in|out|状态|TTFT|DecodeTPS|TPOT|ITLmax|")
                lines.append("|--|--|--|--|--|--|--|--|")
                for rq in d.get("requests", []):
                    st = "✅" if rq.get("status") == "ok" else "❌"
                    lines.append(f"|{rq.get('seq')}|{rq.get('input_tokens')}|{rq.get('output_tokens')}|{st}|{rq.get('ttft')}|{rq.get('decode_tps')}|{rq.get('tpot_ms')}|{rq.get('itl_max_ms')}|")
                lines.append("")
            else:
                lines.append(f"通过{d.get('passed_count','?')}|并发={d.get('concurrency','N/A')}请求/档={d.get('total_requests','N/A')}")
                lines.append("|档位|in/out|ok/fail|耗时|TTFT avg|TTFT p99|Decode TPS|TPOT avg|TPOT p99|ITL avg|ITL>500ms|")
                lines.append("|------|------|------|------|------|------|------|------|------|------|------|")
                for t in d.get("tiers", []):
                    lines.append(f"|{t.get('name','')}|{t.get('input_tokens','')}/{t.get('output_tokens','')}|{t.get('ok','')}/{t.get('fail','')}|{t.get('elapsed','')}s|{t.get('ttft_avg','')}s|{t.get('ttft_p99','')}s|{t.get('decode_tps','')}|{t.get('tpot_avg_ms','')}ms|{t.get('tpot_p99_ms','')}ms|{t.get('itl_avg_ms','')}ms|{t.get('itl_over_500ms','')}|")
                lines.append("")
        else:
            lines.append(f"```json\n{json.dumps(d, ensure_ascii=False, indent=2)}\n```")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Markdown 报告已写入: {output_path}")


def write_html_report(results, cases, cfg, args, output_path, test_time, model_info=None):
    """生成深色主题 Chart.js 仪表盘 HTML 报告，直接从 results 数据渲染。"""
    api = cfg["api"]
    rm = {r["case_id"]: r for r in results}

    # ── 汇总统计 ──
    total_cases = len(results)
    passed_cases = sum(1 for r in results if r.get("passed"))
    failed_cases = total_cases - passed_cases
    pass_rate = f"{passed_cases / total_cases * 100:.1f}%" if total_cases else "N/A"

    # ── TC-10 数据提取 ──
    tc10 = rm.get("TC-10", {})
    tc10d = tc10.get("detail", {}) or {}
    tc10_mode = tc10d.get("mode", "")
    tc10_requests = tc10d.get("requests", [])
    tc10_tiers = tc10d.get("tiers", [])
    tc10_sweep = tc10d.get("sweep", [])
    tc10_ok_recs = [r for r in tc10_requests if r.get("status") == "ok"]
    tc10_has_data = bool(tc10_requests or tc10_tiers or tc10_sweep)

    # ── TC-02 数据提取 ──
    tc02 = rm.get("TC-02", {})
    tc02d = tc02.get("detail", {}) or {}
    tc02_levels = tc02d.get("levels", [])
    tc02_has_data = bool(tc02_levels)

    # ── TC-10 输入/输出 token 分位数 ──
    in_tokens = sorted([r.get("input_tokens", 0) for r in tc10_ok_recs])
    out_tokens = sorted([r.get("completion_tokens", 0) for r in tc10_ok_recs])
    def _pct(sorted_data, p):
        if not sorted_data: return 0
        idx = int((p / 100.0) * (len(sorted_data) - 1))
        return sorted_data[min(idx, len(sorted_data) - 1)]
    def _avg(data):
        return round(sum(data) / len(data), 1) if data else 0

    # ── 构建 HTML ──
    now_str = datetime.now(TZ_BJ).strftime("%Y-%m-%d %H:%M:%S")
    model = api.get("model", "N/A")
    plat = getattr(args, 'platform', '') or 'N/A'
    cmd = "python token_test.py " + " ".join(
        f'"{a}"' if " " in a else a for a in sys.argv[1:])

    # 将结构化数据序列化为 JSON 供前端 Chart.js 使用
    report_data = {
        "testTime": test_time,
        "generatedAt": now_str,
        "model": model,
        "platform": plat,
        "command": cmd,
        "summary": {
            "totalCases": total_cases,
            "passed": passed_cases,
            "failed": failed_cases,
            "passRate": pass_rate,
        },
        "tc02": {
            "passed": tc02.get("passed", False),
            "levels": tc02_levels,
        } if tc02_has_data else None,
        "tc10": {
            "mode": tc10_mode,
            "passed": tc10.get("passed", False),
            "totalRequests": tc10d.get("total_requests", 0) or tc10d.get("total_steps", 0),
            "concurrency": tc10d.get("concurrency", 0),
            "elapsed": tc10d.get("elapsed", 0),
            "ok": tc10d.get("ok", 0),
            "fail": tc10d.get("fail", 0),
            "ttft_p50": tc10d.get("ttft_p50", 0),
            "ttft_p75": tc10d.get("ttft_p75", 0),
            "ttft_p90": tc10d.get("ttft_p90", 0),
            "ttft_p99": tc10d.get("ttft_p99", 0),
            "decode_tps_p50": tc10d.get("decode_tps_p50", 0),
            "decode_tps_p99": tc10d.get("decode_tps_p99", 0),
            "tpot_p50_ms": tc10d.get("tpot_p50_ms", 0),
            "tpot_p75_ms": tc10d.get("tpot_p75_ms", 0),
            "tpot_p90_ms": tc10d.get("tpot_p90_ms", 0),
            "tpot_p99_ms": tc10d.get("tpot_p99_ms", 0),
            "itl_max_p50_ms": tc10d.get("itl_max_p50_ms", 0),
            "itl_max_p90_ms": tc10d.get("itl_max_p90_ms", 0),
            "itl_max_p99_ms": tc10d.get("itl_max_p99_ms", 0),
            "in_p50": _pct(in_tokens, 50), "in_avg": _avg(in_tokens),
            "in_p90": _pct(in_tokens, 90), "in_p99": _pct(in_tokens, 99),
            "out_p50": _pct(out_tokens, 50), "out_avg": _avg(out_tokens),
            "out_p90": _pct(out_tokens, 90), "out_p99": _pct(out_tokens, 99),
            "requests": [{
                "seq": r.get("seq", i + 1),
                "input_tokens": r.get("input_tokens", 0),
                "output_tokens": r.get("output_tokens", 0),
                "status": r.get("status", "?"),
                "ttft": r.get("ttft", 0),
                "decode_tps": r.get("decode_tps", 0),
                "tpot_ms": r.get("tpot_ms", 0),
                "itl_max_ms": r.get("itl_max_ms", 0),
                "error": r.get("error", ""),
            } for i, r in enumerate(tc10_requests)],
            "tiers": [{
                "name": t.get("name", ""),
                "input_tokens": t.get("input_tokens", 0),
                "output_tokens": t.get("output_tokens", 0),
                "concurrency": t.get("concurrency", 0),
                "ok": t.get("ok", 0), "fail": t.get("fail", 0),
                "elapsed": t.get("elapsed", 0),
                "ttft_avg": t.get("ttft_avg", 0),
                "ttft_p99": t.get("ttft_p99", 0),
                "decode_tps": t.get("decode_tps", 0),
                "tpot_avg_ms": t.get("tpot_avg_ms", 0),
                "tpot_p99_ms": t.get("tpot_p99_ms", 0),
                "itl_avg_ms": t.get("itl_avg_ms", 0),
            } for t in tc10d.get("tiers", [])],
            "sweep": [{
                "seq": r.get("seq", i + 1),
                "input_tokens": r.get("input_tokens", 0),
                "output_tokens": r.get("output_tokens", 0),
                "tier": r.get("tier", "?"),
                "status": r.get("status", "?"),
                "ttft": r.get("ttft", 0),
                "ttft_std": r.get("ttft_std", 0),
                "decode_tps": r.get("decode_tps", 0),
                "decode_tps_std": r.get("decode_tps_std", 0),
                "tpot_ms": r.get("tpot_ms", 0),
                "total_latency": r.get("total_latency", 0),
                "repetitions": r.get("repetitions", 1),
                "ok_reps": r.get("ok_reps", 0),
                "fail_reps": r.get("fail_reps", 0),
            } for i, r in enumerate(tc10_sweep)],
            "total_requests": tc10d.get("total_requests", 0),
            "repetitions": tc10d.get("repetitions", 1),
            "tier_stats": [{
                "tier": ts.get("tier", ""),
                "ok": ts.get("ok", 0), "fail": ts.get("fail", 0),
                "ttft_avg": ts.get("ttft_avg", 0),
                "ttft_max": ts.get("ttft_max", 0),
                "decode_tps_avg": ts.get("decode_tps_avg", 0),
                "tpot_avg_ms": ts.get("tpot_avg_ms", 0),
            } for ts in tc10d.get("tier_stats", [])],
        } if tc10_has_data else None,
        "cases": [{
            "id": c["id"],
            "name": c.get("name", ""),
            "type": c.get("type", ""),
            "passed": rm.get(c["id"], {}).get("passed", False),
            "detail": rm.get(c["id"], {}).get("detail", {}),
        } for c in cases],
    }
    report_json = json.dumps(report_data, ensure_ascii=False, default=str)

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>LLM 推理服务测试指标看板 — {test_time}</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<script>window.Chart||document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.4/chart.umd.min.js"><\\/script>')</script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script>window.ChartDataLabels||document.write('<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"><\\/script>')</script>
<style>
:root{{
  --bg:#0B0D12; --panel:#12151C; --panel-border:#232838; --grid:#1E2330;
  --text:#E7E9F0; --muted:#7C8499; --cyan:#4FD8C4; --amber:#F0A868;
  --violet:#9A8CFF; --rose:#F0708A; --green:#6FD98C;
  --mono:'IBM Plex Mono','SFMono-Regular',Consolas,monospace;
  --sans:'Inter',-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;
}}
*{{box-sizing:border-box;}}
body{{
  margin:0; background: radial-gradient(1200px 500px at 12% -10%, rgba(79,216,196,0.07), transparent 60%),
    radial-gradient(1000px 500px at 92% 0%, rgba(154,140,255,0.06), transparent 60%), var(--bg);
  color:var(--text); font-family:var(--sans); padding:32px 40px 60px;
}}
header{{
  display:flex; justify-content:space-between; align-items:flex-end;
  border-bottom:1px solid var(--panel-border); padding-bottom:20px; margin-bottom:24px;
  flex-wrap:wrap; gap:16px;
}}
.brand-eyebrow{{font-family:var(--mono); font-size:11px; letter-spacing:.18em; color:var(--cyan);
  text-transform:uppercase; margin-bottom:6px;}}
h1{{margin:0; font-size:25px; font-weight:650; letter-spacing:-0.01em;}}
header .meta{{font-family:var(--mono); font-size:12px; color:var(--muted); text-align:right; line-height:1.7;}}
header .meta span{{color:var(--text);}}
.kpi-strip{{display:grid; grid-template-columns:repeat(7,1fr); gap:14px; margin-bottom:24px;}}
.kpi{{background:var(--panel); border:1px solid var(--panel-border); border-radius:10px;
  padding:14px 16px; position:relative;}}
.kpi .k{{font-family:var(--mono); font-size:10px; color:var(--muted); letter-spacing:.05em;
  text-transform:uppercase;}}
.kpi .v{{font-family:var(--mono); font-size:22px; font-weight:650; margin-top:6px;}}
.kpi .d{{font-size:11px; color:var(--muted); margin-top:3px;}}
.kpi.ok .v{{color:var(--green);}} .kpi.warn .v{{color:var(--amber);}} .kpi.bad .v{{color:var(--rose);}}
.grid{{display:grid; grid-template-columns:repeat(12,1fr); gap:18px;}}
.panel{{
  background:linear-gradient(180deg, rgba(255,255,255,0.015), rgba(255,255,255,0));
  background-color:var(--panel); border:1px solid var(--panel-border);
  border-radius:10px; padding:20px 22px; position:relative; overflow:hidden;
}}
.panel::before{{content:""; position:absolute; top:0; left:0; right:0; height:2px; opacity:.75;}}
.panel.c-cyan{{--bar:var(--cyan);}} .panel.c-amber{{--bar:var(--amber);}}
.panel.c-violet{{--bar:var(--violet);}} .panel.c-rose{{--bar:var(--rose);}} .panel.c-green{{--bar:var(--green);}}
.panel::before{{background:var(--bar);}}
.panel-head{{display:flex; justify-content:space-between; align-items:baseline; margin-bottom:4px;}}
.panel-title{{font-size:14px; font-weight:600; letter-spacing:.01em;}}
.panel-tag{{font-family:var(--mono); font-size:10.5px; color:var(--muted); letter-spacing:.06em;}}
.panel-sub{{font-family:var(--mono); font-size:11.5px; color:var(--muted); margin-bottom:14px;}}
.span-6{{grid-column:span 6;}} .span-12{{grid-column:span 12;}}
.chart-box{{position:relative; width:100%;}}
.h-260{{height:260px;}} .h-220{{height:220px;}}
.summary-section{{margin-bottom:24px;}}
table{{width:100%; border-collapse:collapse; font-family:var(--mono); font-size:11.5px;}}
thead th{{
  text-align:right; font-weight:600; color:var(--muted); font-size:10px;
  letter-spacing:.04em; text-transform:uppercase; padding:8px 10px;
  border-bottom:1px solid var(--panel-border);
  position:sticky; top:0; background:var(--panel);
}}
thead th:first-child, thead th:nth-child(2){{text-align:left;}}
tbody td{{padding:7px 10px; text-align:right; border-bottom:1px solid #171B26; color:var(--text);}}
tbody td:first-child, tbody td:nth-child(2){{text-align:left; color:var(--muted);}}
tbody tr:hover{{background:rgba(255,255,255,0.02);}}
.badge{{
  display:inline-block; padding:2px 8px; border-radius:10px; font-size:10px;
  font-family:var(--mono); letter-spacing:.03em;
}}
.badge.ok{{background:rgba(111,217,140,.12); color:var(--green); border:1px solid rgba(111,217,140,.3);}}
.badge.warn{{background:rgba(240,168,104,.12); color:var(--amber); border:1px solid rgba(240,168,104,.3);}}
.badge.bad{{background:rgba(240,112,138,.12); color:var(--rose); border:1px solid rgba(240,112,138,.3);}}
.pager{{display:flex; align-items:center; justify-content:flex-end; gap:10px; margin-top:14px;
  font-family:var(--mono); font-size:11.5px; color:var(--muted);}}
.pager button{{
  border:1px solid var(--panel-border); background:var(--panel); border-radius:6px;
  padding:5px 14px; font-family:var(--mono); font-size:11px; cursor:pointer; color:var(--text);
  letter-spacing:.03em;
}}
.pager button:hover:not(:disabled){{border-color:var(--cyan); color:var(--cyan);}}
.pager button:disabled{{opacity:.35; cursor:default;}}
footer{{margin-top:30px; font-family:var(--mono); font-size:11px; color:var(--muted);
  text-align:center; letter-spacing:.03em;}}
.cmd-box{{background:var(--panel); border:1px solid var(--panel-border); border-radius:8px;
  padding:12px 18px; margin:16px 0; font-family:var(--mono); font-size:11px; color:var(--muted);
  overflow-x:auto; white-space:pre-wrap; word-break:break-all;}}
@media (max-width:1100px){{.kpi-strip{{grid-template-columns:repeat(3,1fr);}}}}
@media (max-width:900px){{.span-6{{grid-column:span 12;}} body{{padding:20px;}}}}
</style>
</head>
<body>

<header>
  <div>
    <div class="brand-eyebrow">Inference Benchmark · Load Test Report</div>
    <h1>LLM 推理服务测试指标看板</h1>
  </div>
  <div class="meta">
    测试时间 <span>{test_time}</span><br>
    模型 <span>{escape_html(model)}</span> · 平台 <span>{escape_html(plat)}</span><br>
    生成时间 <span>{now_str}</span>
  </div>
</header>

<div class="cmd-box">$ {escape_html(cmd)}</div>

<div class="kpi-strip">
  <div class="kpi ok"><div class="k">用例通过</div><div class="v">{passed_cases}/{total_cases}</div><div class="d">通过率 {pass_rate}</div></div>
  <div class="kpi {"ok" if failed_cases == 0 else "bad"}"><div class="k">用例失败</div><div class="v">{failed_cases}</div><div class="d">{f"{failed_cases} 项未通过" if failed_cases else "全部通过 ✅"}</div></div>
  <div class="kpi">{_kpi_tc10_total(tc10d, tc10_mode)}</div>
  <div class="kpi">{_kpi_tc10_ok(tc10d, tc10_mode)}</div>
  <div class="kpi">{_kpi_tc10_ttft(tc10d, tc10_mode)}</div>
  <div class="kpi">{_kpi_tc10_tpot(tc10d, tc10_mode)}</div>
  <div class="kpi">{_kpi_tc10_decode(tc10d, tc10_mode)}</div>
</div>

<div class="grid">

  {_render_tc10_charts(tc10_has_data, tc10_mode)}

  {_render_tc02_chart(tc02_has_data)}

  <!-- 全部用例汇总表 -->
  <div class="panel c-cyan span-12 summary-section">
    <div class="panel-head">
      <div class="panel-title">全部测试用例汇总</div>
      <div class="panel-tag">ALL CASES · SUMMARY</div>
    </div>
    <div class="panel-sub">共 {total_cases} 项用例 · 通过 {passed_cases} · 失败 {failed_cases}</div>
    <div style="max-height:500px; overflow-y:auto;">
    <table>
      <thead><tr>
        <th>ID</th><th>名称</th><th>类型</th><th>状态</th><th>关键指标</th>
      </tr></thead>
      <tbody>
        {_render_summary_rows(cases, rm)}
      </tbody>
    </table>
    </div>
  </div>

  {_render_tc10_detail_table(tc10_has_data)}

</div>

<footer>INFERENCE BENCHMARK · 报告由 token_test.py 自动生成 · {now_str}</footer>

<script>
window.REPORT = {report_json};

(function() {{
  if (window.ChartDataLabels) {{ Chart.register(ChartDataLabels); }}
  Chart.defaults.font.family = "'IBM Plex Mono', monospace";
  Chart.defaults.color = '#7C8499';
  Chart.defaults.font.size = 11;
  Chart.defaults.set('plugins.datalabels', {{ display: false }});
  const gridColor = '#1E2330';
  const commonScales = (extra={{}}) => ({{
    x: {{ grid: {{ color: gridColor, drawTicks: false }}, border: {{ color: '#232838' }} }},
    y: {{ grid: {{ color: gridColor, drawTicks: false }}, border: {{ color: '#232838' }}, beginAtZero: true, ...extra }}
  }});
  function barGradient(ctx, top, bottom, h) {{
    const g = ctx.createLinearGradient(0, 0, 0, h || 260);
    g.addColorStop(0, top); g.addColorStop(1, bottom);
    return g;
  }}

  const R = window.REPORT;

  {_render_tc10_charts_js()}

  {_render_tc02_chart_js()}

  {_render_tc10_table_js()}

}})();
</script>

</body>
</html>'''

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] HTML 报告已写入: {output_path}")


# ── Dashboard HTML 辅助函数 ──

def escape_html(text):
    """HTML 转义。"""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _kpi_tc10_total(d, mode):
    if not d or not mode: return '<div class="k">TC-10 请求</div><div class="v">N/A</div><div class="d">未运行 IO 档位测试</div>'
    total = d.get("total_requests", 0)
    conc = d.get("concurrency", 0)
    return f'<div class="k">TC-10 请求数</div><div class="v">{total}</div><div class="d">并发={conc} · 耗时{d.get("elapsed",0)}s</div>'

def _kpi_tc10_ok(d, mode):
    if not d or not mode: return '<div class="k">TC-10 成功率</div><div class="v">N/A</div><div class="d">—</div>'
    ok_n = d.get("ok", 0)
    fail_n = d.get("fail", 0)
    total = ok_n + fail_n
    rate = f"{ok_n / total * 100:.1f}%" if total else "N/A"
    cls = "ok" if fail_n == 0 else ("warn" if fail_n <= total * 0.05 else "bad")
    return f'<div class="k {cls}">TC-10 成功率</div><div class="v">{rate}</div><div class="d">成功 {ok_n} · 失败 {fail_n}</div>'

def _kpi_tc10_ttft(d, mode):
    if not d or not mode: return '<div class="k">TTFT P50</div><div class="v">N/A</div><div class="d">—</div>'
    p50 = d.get("ttft_p50", 0)
    p99 = d.get("ttft_p99", 0)
    return f'<div class="k">TTFT P50</div><div class="v">{p50}s</div><div class="d">P99={p99}s</div>'

def _kpi_tc10_tpot(d, mode):
    if not d or not mode: return '<div class="k">TPOT P50</div><div class="v">N/A</div><div class="d">—</div>'
    p50 = d.get("tpot_p50_ms", 0)
    p99 = d.get("tpot_p99_ms", 0)
    return f'<div class="k">TPOT P50</div><div class="v">{p50}ms</div><div class="d">P99={p99}ms</div>'

def _kpi_tc10_decode(d, mode):
    if not d or not mode: return '<div class="k">Decode TPS</div><div class="v">N/A</div><div class="d">—</div>'
    p50 = d.get("decode_tps_p50", 0)
    p99 = d.get("decode_tps_p99", 0)
    return f'<div class="k">Decode TPS P50</div><div class="v">{p50}</div><div class="d">P99={p99} tok/s</div>'

def _render_tc10_charts(has_data, mode):
    if not has_data:
        return '''<!-- TC-10 未运行，跳过图表 -->
  <div class="panel c-cyan span-12">
    <div class="panel-head"><div class="panel-title">TC-10 分输入输出档位性能</div><div class="panel-tag">NOT RUN</div></div>
    <div class="panel-sub">未运行 TC-10 测试，无图表数据</div>
  </div>'''
    if mode == "io_sweep":
        return '''<!-- TC-10 朴素扫描曲线 -->
  <div class="panel c-amber span-12">
    <div class="panel-head"><div class="panel-title">朴素扫描 · TTFT 随输入增长曲线</div><div class="panel-tag">SWEEP · TTFT vs INPUT</div></div>
    <div class="panel-sub">X轴：输入 tokens (K) · Y轴：TTFT (s) · 按档位着色</div>
    <div class="chart-box h-260"><canvas id="chartSweepTTFT"></canvas></div>
  </div>
  <div class="panel c-green span-12">
    <div class="panel-head"><div class="panel-title">朴素扫描 · Decode TPS 随输入增长曲线</div><div class="panel-tag">SWEEP · DECODE_TPS vs INPUT</div></div>
    <div class="panel-sub">X轴：输入 tokens (K) · Y轴：Decode TPS (tok/s) · 按档位着色</div>
    <div class="chart-box h-260"><canvas id="chartSweepTPS"></canvas></div>
  </div>
  <div class="panel c-rose span-12">
    <div class="panel-head"><div class="panel-title">朴素扫描 · TPOT 随输入增长曲线</div><div class="panel-tag">SWEEP · TPOT vs INPUT</div></div>
    <div class="panel-sub">X轴：输入 tokens (K) · Y轴：TPOT (ms) · 按档位着色</div>
    <div class="chart-box h-260"><canvas id="chartSweepTPOT"></canvas></div>
  </div>'''
    elif mode == "io_mix":
        return '''<!-- Input Tokens 分位数 -->
  <div class="panel c-cyan span-6">
    <div class="panel-head"><div class="panel-title">Input Tokens 分位数</div><div class="panel-tag">INPUT_TOKENS · PERCENTILE</div></div>
    <div class="panel-sub" id="inSub">单位：tokens</div>
    <div class="chart-box h-260"><canvas id="chartInput"></canvas></div>
  </div>
  <!-- Output Tokens 分位数 -->
  <div class="panel c-violet span-6">
    <div class="panel-head"><div class="panel-title">Output Tokens 分位数</div><div class="panel-tag">OUTPUT_TOKENS · PERCENTILE</div></div>
    <div class="panel-sub" id="outSub">单位：tokens</div>
    <div class="chart-box h-260"><canvas id="chartOutput"></canvas></div>
  </div>
  <!-- TTFT 分位数 -->
  <div class="panel c-amber span-6">
    <div class="panel-head"><div class="panel-title">TTFT 首响时延分位数</div><div class="panel-tag">TTFT · PERCENTILE</div></div>
    <div class="panel-sub">单位：秒 · P50 / P75 / P90 / P99</div>
    <div class="chart-box h-260"><canvas id="chartTTFT"></canvas></div>
  </div>
  <!-- TPOT 分位数 -->
  <div class="panel c-rose span-6">
    <div class="panel-head"><div class="panel-title">TPOT Decode 速度分位数</div><div class="panel-tag">TPOT · PERCENTILE</div></div>
    <div class="panel-sub">单位：ms · P50 / P75 / P90 / P99（越低越好）</div>
    <div class="chart-box h-260"><canvas id="chartTPOT"></canvas></div>
  </div>'''
    else:
        return '''<!-- TC-10 档位模式图表 -->
  <div class="panel c-cyan span-12">
    <div class="panel-head"><div class="panel-title">TC-10 分档位性能对比</div><div class="panel-tag">IO_TIER · BENCHMARK</div></div>
    <div class="panel-sub">柱：Decode TPS (tok/s) · 线：TTFT avg (s) · 每档独立并发压测</div>
    <div class="chart-box h-260"><canvas id="chartTier"></canvas></div>
  </div>'''

def _render_tc02_chart(has_data):
    if not has_data: return ""
    return '''<!-- TC-02 并发梯度 -->
  <div class="panel c-green span-12">
    <div class="panel-head"><div class="panel-title">TC-02 并发梯度 · TPS 对比</div><div class="panel-tag">CONCURRENCY · TPS</div></div>
    <div class="panel-sub">柱：TPS (tok/s) · 线：Decode TPS (tok/s) · 不同并发级别</div>
    <div class="chart-box h-260"><canvas id="chartTC02"></canvas></div>
  </div>'''

def _render_summary_rows(cases, rm):
    rows = []
    for c in cases:
        cid = c["id"]
        r = rm.get(cid, {})
        passed = r.get("passed", False)
        d = r.get("detail", {}) or {}
        status_cls = "ok" if passed else "bad"
        status_text = "✅ 通过" if passed else "❌ 失败"
        # 提取关键指标摘要
        metric = _case_metric(cid, d)
        rows.append(f'''<tr>
          <td>{escape_html(cid)}</td>
          <td>{escape_html(c.get("name", ""))}</td>
          <td>{escape_html(c.get("type", ""))}</td>
          <td><span class="badge {status_cls}">{status_text}</span></td>
          <td>{escape_html(metric)}</td>
        </tr>''')
    return "\n".join(rows)

def _case_metric(cid, d):
    """从 detail 提取单行关键指标。"""
    if not d: return "未执行"
    if cid == "TC-01":
        return f"HTTP {d.get('status_code','?')} · {d.get('latency','?')}s"
    elif cid == "TC-02":
        return f"最高 TPS {d.get('tps_tokens','?')} tok/s · Decode {d.get('decode_tps','?')} tok/s"
    elif cid == "TC-03":
        return f"TPS→TPM {'✅' if d.get('formula_ok') else '❌'}"
    elif cid == "TC-04":
        mode = d.get("mode", "")
        if mode == "multi_tier":
            return f"通过 {d.get('passed_count','?')}"
        return f"{d.get('declared_context_length','?') or d.get('found_limit','?')} tokens"
    elif cid == "TC-05":
        return "跳过鉴权" if d.get("status_code") is None else f"HTTP {d.get('status_code')}"
    elif cid == "TC-06":
        return f"OK={d.get('ok',0)} · 429={d.get('rate_limited_429',0)}"
    elif cid == "TC-07":
        return f"{d.get('tool_name','?')} · HTTP {d.get('status_code','?')}"
    elif cid == "TC-08":
        return f"TTFT avg={d.get('ttft_avg','?')}s · Decode={d.get('decode_tps','?')} tok/s"
    elif cid == "TC-09":
        return f"命中率 {d.get('cache_hit_rate_pct','?')}% · TPM {d.get('tpm_tokens','?')}"
    elif cid == "TC-10":
        mode = d.get("mode", "")
        if mode == "io_sweep":
            reps = d.get("repetitions", 1)
            mode = "全并发" if reps == 0 else (f"×{reps}并发" if reps > 1 else "串行")
            return f"朴素扫描({mode}) {d.get('total_steps','?')}步 {d.get('total_requests','?')}条 · 耗时{d.get('elapsed','?')}s"
        if mode == "io_mix":
            return f"混合 {d.get('total_requests','?')}条 · TTFT p50={d.get('ttft_p50','?')}s · Decode p50={d.get('decode_tps_p50','?')}"
        else:
            return f"档位 {d.get('passed_count','?')} · 并发={d.get('concurrency','?')}"
    return json.dumps(d, ensure_ascii=False, default=str)[:120]

def _render_tc10_detail_table(has_data):
    if not has_data: return ""
    return '''<!-- TC-10 请求明细表 -->
  <div class="panel c-cyan span-12">
    <div class="panel-head">
      <div class="panel-title">TC-10 请求明细</div>
      <div class="panel-tag" id="tc10tableTag">RAW SAMPLES</div>
    </div>
    <div class="panel-sub" id="tc10tableSub">字段见下表，每页 20 条</div>
    <div style="max-height:520px; overflow-y:auto;">
    <table>
      <thead id="tc10thead">
        <tr>
          <th>#</th><th>input_tokens</th><th>output_tokens</th>
          <th>ttft (s)</th><th>decode_tps</th><th>tpot_ms</th><th>itl_max_ms</th><th>状态</th>
        </tr>
      </thead>
      <tbody id="tc10tbody"></tbody>
    </table>
    </div>
    <div class="pager">
      <button id="tc10prev">← 上一页</button>
      <span id="tc10page"></span>
      <button id="tc10next">下一页 →</button>
    </div>
  </div>'''

def _render_tc10_charts_js():
    return '''
  // ── TC-10 Input Tokens 分位数 ──
  if (R.tc10 && R.tc10.mode === "io_mix") {
    const d = R.tc10;
    document.getElementById("inSub").textContent =
      `P50=${d.in_p50.toLocaleString()} · avg=${d.in_avg.toLocaleString()} · P90=${d.in_p90.toLocaleString()} · P99=${d.in_p99.toLocaleString()} tokens`;
    new Chart(document.getElementById("chartInput"), {
      type: "bar",
      data: {
        labels: ["P50", "AVG", "P90", "P99"],
        datasets: [{
          data: [d.in_p50, d.in_avg, d.in_p90, d.in_p99],
          backgroundColor: c => barGradient(c.chart.ctx, "rgba(79,216,196,0.9)", "rgba(79,216,196,0.2)"),
          borderRadius: 4, borderSkipped: false, maxBarThickness: 52
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => " " + i.formattedValue + " tokens" } },
          datalabels: { display: true, anchor: "end", align: "top", offset: 2,
            color: "#9FA6B8", font: { size: 10 },
            formatter: v => v >= 1000 ? (v/1000).toFixed(1)+"K" : v }
        },
        scales: commonScales({ title: { display: true, text: "tokens", color: "#7C8499", font: { size: 10 } } })
      }
    });

    // ── TC-10 Output Tokens 分位数 ──
    document.getElementById("outSub").textContent =
      `P50=${d.out_p50.toLocaleString()} · avg=${d.out_avg.toLocaleString()} · P90=${d.out_p90.toLocaleString()} · P99=${d.out_p99.toLocaleString()} tokens`;
    new Chart(document.getElementById("chartOutput"), {
      type: "bar",
      data: {
        labels: ["P50", "AVG", "P90", "P99"],
        datasets: [{
          data: [d.out_p50, d.out_avg, d.out_p90, d.out_p99],
          backgroundColor: c => barGradient(c.chart.ctx, "rgba(154,140,255,0.9)", "rgba(154,140,255,0.2)"),
          borderRadius: 4, borderSkipped: false, maxBarThickness: 52
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => " " + i.formattedValue + " tokens" } },
          datalabels: { display: true, anchor: "end", align: "top", offset: 2,
            color: "#9FA6B8", font: { size: 10 },
            formatter: v => v >= 1000 ? (v/1000).toFixed(1)+"K" : v }
        },
        scales: commonScales({ title: { display: true, text: "tokens", color: "#7C8499", font: { size: 10 } } })
      }
    });

    // ── TC-10 TTFT 分位数 ──
    new Chart(document.getElementById("chartTTFT"), {
      type: "bar",
      data: {
        labels: ["P50", "P75", "P90", "P99"],
        datasets: [{
          label: "实测值",
          data: [d.ttft_p50, d.ttft_p75, d.ttft_p90, d.ttft_p99],
          backgroundColor: c => barGradient(c.chart.ctx, "rgba(240,168,104,0.9)", "rgba(240,168,104,0.2)"),
          borderRadius: 4, borderSkipped: false, maxBarThickness: 44
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => " " + i.formattedValue + "s" } },
          datalabels: { display: true, anchor: "end", align: "top", offset: 2,
            color: "#9FA6B8", font: { size: 10 }, formatter: v => v + "s" }
        },
        scales: commonScales({ title: { display: true, text: "秒 (s)", color: "#7C8499", font: { size: 10 } } })
      }
    });

    // ── TC-10 TPOT 分位数 ──
    new Chart(document.getElementById("chartTPOT"), {
      type: "bar",
      data: {
        labels: ["P50", "P75", "P90", "P99"],
        datasets: [{
          label: "实测值",
          data: [d.tpot_p50_ms, d.tpot_p75_ms, d.tpot_p90_ms, d.tpot_p99_ms],
          backgroundColor: c => barGradient(c.chart.ctx, "rgba(240,112,138,0.9)", "rgba(240,112,138,0.2)"),
          borderRadius: 4, borderSkipped: false, maxBarThickness: 44
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => " " + i.formattedValue + " ms" } },
          datalabels: { display: true, anchor: "end", align: "top", offset: 2,
            color: "#9FA6B8", font: { size: 10 }, formatter: v => v + "ms" }
        },
        scales: commonScales({ title: { display: true, text: "ms", color: "#7C8499", font: { size: 10 } } })
      }
    });
  }

  // ── TC-10 档位模式图表 ──
  if (R.tc10 && R.tc10.mode === "io_tier" && R.tc10.tiers && R.tc10.tiers.length > 0) {
    const tiers = R.tc10.tiers;
    const labels = tiers.map(t => t.name + " (并发" + (t.concurrency || "?") + ")");
    new Chart(document.getElementById("chartTier"), {
      type: "bar",
      data: {
        labels: labels,
        datasets: [
          {
            type: "bar", label: "Decode TPS (tok/s)", data: tiers.map(t => t.decode_tps),
            backgroundColor: c => barGradient(c.chart.ctx, "rgba(79,216,196,0.9)", "rgba(79,216,196,0.2)"),
            borderRadius: 4, borderSkipped: false, maxBarThickness: 36, yAxisID: "y"
          },
          {
            type: "line", label: "TTFT avg (s)", data: tiers.map(t => t.ttft_avg),
            fill: false, tension: 0.35, borderColor: "#F0A868", pointRadius: 4,
            pointBackgroundColor: "#F0A868", borderWidth: 2.5, yAxisID: "y1"
          }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: true, position: "top", align: "end",
            labels: { boxWidth: 10, boxHeight: 10, padding: 12 } },
          tooltip: { mode: "index", intersect: false },
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" } },
          y: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            beginAtZero: true, title: { display: true, text: "tok/s", color: "#7C8499", font: { size: 10 } } },
          y1: { position: "right", beginAtZero: true, grid: { drawOnChartArea: false },
            border: { color: "#232838" }, title: { display: true, text: "TTFT (s)", color: "#7C8499", font: { size: 10 } } }
        }
      }
    });
  }

  // ── TC-10 朴素串行扫描曲线 ──
  if (R.tc10 && R.tc10.mode === "io_sweep" && R.tc10.sweep && R.tc10.sweep.length > 0) {
    const sweep = R.tc10.sweep.filter(r => r.status === "ok");
    const labels = sweep.map(r => (r.input_tokens / 1000).toFixed(0) + "K");
    const tiers = sweep.map(r => r.tier);
    const tierColors = { "P50": "#4FD8C4", "AVG": "#9A8CFF", "P90": "#F0A868", "P99": "#F0708A" };
    const pointColors = tiers.map(t => tierColors[t] || "#7C8499");
    const hasConcurrent = R.tc10.repetitions > 1;

    // TTFT 曲线
    new Chart(document.getElementById("chartSweepTTFT"), {
      type: "line",
      data: {
        labels: labels,
        datasets: [{
          label: "TTFT (s)", data: sweep.map(r => r.ttft),
          fill: false, tension: 0.3, borderColor: "#F0A868",
          pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: pointColors,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => {
            const r = sweep[i.dataIndex];
            let tip = " " + r.tier + " · " + r.input_tokens.toLocaleString() + " tokens · TTFT=" + i.formattedValue + "s";
            if (hasConcurrent && r.repetitions > 1 && r.ttft_std > 0) tip += " · ±1σ=" + r.ttft_std.toFixed(2) + "s";
            return tip;
          }}},
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            title: { display: true, text: "Input tokens", color: "#7C8499", font: { size: 10 } } },
          y: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            beginAtZero: true, title: { display: true, text: "秒 (s)", color: "#7C8499", font: { size: 10 } } }
        }
      }
    });

    // Decode TPS 曲线
    new Chart(document.getElementById("chartSweepTPS"), {
      type: "line",
      data: {
        labels: labels,
        datasets: [{
          label: "Decode TPS (tok/s)", data: sweep.map(r => r.decode_tps),
          fill: false, tension: 0.3, borderColor: "#6FD98C",
          pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: pointColors,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => {
            const r = sweep[i.dataIndex];
            let tip = " " + r.tier + " · " + r.input_tokens.toLocaleString() + " tokens · Decode=" + i.formattedValue + " tok/s";
            if (hasConcurrent && r.repetitions > 1 && r.decode_tps_std > 0) tip += " · ±1σ=" + r.decode_tps_std.toFixed(1);
            return tip;
          }}},
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            title: { display: true, text: "Input tokens", color: "#7C8499", font: { size: 10 } } },
          y: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            beginAtZero: true, title: { display: true, text: "tok/s", color: "#7C8499", font: { size: 10 } } }
        }
      }
    });

    // TPOT 曲线
    new Chart(document.getElementById("chartSweepTPOT"), {
      type: "line",
      data: {
        labels: labels,
        datasets: [{
          label: "TPOT (ms)", data: sweep.map(r => r.tpot_ms),
          fill: false, tension: 0.3, borderColor: "#F0708A",
          pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: pointColors,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: i => {
            const r = sweep[i.dataIndex];
            let tip = " " + r.tier + " · " + r.input_tokens.toLocaleString() + " tokens · TPOT=" + i.formattedValue + "ms";
            if (hasConcurrent && r.repetitions > 1) tip += " · " + r.ok_reps + "/" + r.repetitions + " ok";
            return tip;
          }}},
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            title: { display: true, text: "Input tokens", color: "#7C8499", font: { size: 10 } } },
          y: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            beginAtZero: true, title: { display: true, text: "ms", color: "#7C8499", font: { size: 10 } } }
        }
      }
    });
  }'''

def _render_tc02_chart_js():
    return '''
  // ── TC-02 并发梯度 TPS 对比 ──
  if (R.tc02 && R.tc02.levels && R.tc02.levels.length > 0) {
    const lv = R.tc02.levels;
    const labels = lv.map(l => "并发=" + l.concurrency);
    const tpsData = lv.map(l => l.tps_tokens);
    const decodeData = lv.map(l => l.decode_tps);
    new Chart(document.getElementById("chartTC02"), {
      type: "bar",
      data: {
        labels: labels,
        datasets: [
          {
            type: "bar", label: "TPS (tok/s)", data: tpsData,
            backgroundColor: "rgba(111,217,140,0.18)", borderColor: "rgba(111,217,140,0.4)",
            borderWidth: 1, borderRadius: 3, maxBarThickness: 32, yAxisID: "y"
          },
          {
            type: "line", label: "Decode TPS (tok/s)", data: decodeData,
            fill: false, tension: 0.35, borderColor: "#6FD98C", pointRadius: 4,
            pointBackgroundColor: "#6FD98C", borderWidth: 2.5, yAxisID: "y1"
          }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: true, position: "top", align: "end",
            labels: { boxWidth: 10, boxHeight: 10, padding: 12 } },
          tooltip: { mode: "index", intersect: false },
          datalabels: { display: false }
        },
        scales: {
          x: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" } },
          y: { grid: { color: gridColor, drawTicks: false }, border: { color: "#232838" },
            beginAtZero: true, title: { display: true, text: "TPS (tok/s)", color: "#7C8499", font: { size: 10 } } },
          y1: { position: "right", beginAtZero: true, grid: { drawOnChartArea: false },
            border: { color: "#232838" }, title: { display: true, text: "Decode TPS", color: "#7C8499", font: { size: 10 } } }
        }
      }
    });
  }'''

def _render_tc10_table_js():
    return '''
  // ── TC-10 请求明细分页（io_mix 模式）──
  function renderTableCommon(reqs, tbody, pageInfo, prevBtn, nextBtn, renderRow) {
    const pageSize = 20;
    let page = 0;
    function render() {
      const slice = reqs.slice(page * pageSize, (page + 1) * pageSize);
      tbody.innerHTML = slice.map(renderRow).join("");
      const totalPages = Math.ceil(reqs.length / pageSize);
      pageInfo.textContent = `第 ${page + 1} / ${totalPages} 页 · 共 ${reqs.length} 条`;
      prevBtn.disabled = page === 0;
      nextBtn.disabled = page >= totalPages - 1;
    }
    prevBtn.onclick = () => { if (page > 0) { page--; render(); } };
    nextBtn.onclick = () => { if (page < Math.ceil(reqs.length / pageSize) - 1) { page++; render(); } };
    render();
  }

  const tbody = document.getElementById("tc10tbody");
  if (!tbody) {} // no table in this mode

  if (R.tc10 && R.tc10.requests && R.tc10.requests.length > 0) {
    renderTableCommon(
      R.tc10.requests,
      tbody,
      document.getElementById("tc10page"),
      document.getElementById("tc10prev"),
      document.getElementById("tc10next"),
      r => {
        const badgeCls = r.status === "ok" ? "ok" : "bad";
        const badgeText = r.status === "ok" ? "成功" : (r.error ? r.error.substring(0, 20) : "失败");
        return `<tr>
          <td>${r.seq}</td>
          <td>${r.input_tokens.toLocaleString()}</td>
          <td>${r.output_tokens.toLocaleString()}</td>
          <td>${typeof r.ttft === "number" ? r.ttft.toFixed(2) : r.ttft}</td>
          <td>${typeof r.decode_tps === "number" ? r.decode_tps.toFixed(1) : r.decode_tps}</td>
          <td>${typeof r.tpot_ms === "number" ? r.tpot_ms.toFixed(1) : r.tpot_ms}</td>
          <td>${typeof r.itl_max_ms === "number" ? r.itl_max_ms.toFixed(1) : r.itl_max_ms}</td>
          <td><span class="badge ${badgeCls}">${badgeText}</span></td>
        </tr>`;
      }
    );
  }

  // ── TC-10 朴素扫描明细表 ──
  if (R.tc10 && R.tc10.mode === "io_sweep" && R.tc10.sweep && R.tc10.sweep.length > 0) {
    const sweepTbody = document.getElementById("tc10tbody");
    // 切换表头为扫描模式字段
    const thead = document.getElementById("tc10thead");
    const repCol = (R.tc10.repetitions > 1) ? '<th>±1σ TTFT</th><th>±1σ Decode</th>' : '';
    if (thead) thead.innerHTML = '<tr><th>#</th><th>input_tokens</th><th>output_tokens</th><th>档位</th><th>ttft (s)</th>' + repCol + '<th>decode_tps</th><th>tpot_ms</th><th>总延迟(s)</th><th>rep</th><th>状态</th></tr>';
    const sub = document.getElementById("tc10tableSub");
    const modeText = (R.tc10.repetitions === 0 || R.tc10.repetitions > 1) ? "全并发" : "串行";
    const repText = R.tc10.repetitions > 1 ? "×" + R.tc10.repetitions : "";
    if (sub) sub.textContent = `朴素扫描(${modeText}${repText})：输入 10K→380K，步长 10K，输出按档位固定 · 共 ${R.tc10.total_requests || R.tc10.sweep.length} 条请求`;
    const tag = document.getElementById("tc10tableTag");
    if (tag) tag.textContent = "SWEEP · " + R.tc10.sweep.length + " STEPS" + repText;
    if (sweepTbody) {
      renderTableCommon(
        R.tc10.sweep,
        sweepTbody,
        document.getElementById("tc10page"),
        document.getElementById("tc10prev"),
        document.getElementById("tc10next"),
        r => {
          const badgeCls = r.status === "ok" ? "ok" : "bad";
          const badgeText = r.status === "ok" ? "成功" : "失败";
          const repInfo = (r.repetitions > 1) ? (r.ok_reps + "/" + r.repetitions + " ok") : "1";
          const stdCols = (R.tc10.repetitions > 1) ?
            `<td>${typeof r.ttft_std === "number" ? "±"+r.ttft_std.toFixed(2) : ""}</td>
             <td>${typeof r.decode_tps_std === "number" ? "±"+r.decode_tps_std.toFixed(1) : ""}</td>` : "";
          return `<tr>
            <td>${r.seq}</td>
            <td>${r.input_tokens.toLocaleString()}</td>
            <td>${r.output_tokens.toLocaleString()}</td>
            <td><span class="badge">${r.tier}</span></td>
            <td>${typeof r.ttft === "number" ? r.ttft.toFixed(2) : r.ttft}</td>
            ${stdCols}
            <td>${typeof r.decode_tps === "number" ? r.decode_tps.toFixed(1) : r.decode_tps}</td>
            <td>${typeof r.tpot_ms === "number" ? r.tpot_ms.toFixed(1) : r.tpot_ms}</td>
            <td>${typeof r.total_latency === "number" ? r.total_latency.toFixed(2) : r.total_latency}</td>
            <td>${repInfo}</td>
            <td><span class="badge ${badgeCls}">${badgeText}</span></td>
          </tr>`;
        }
      );
    }
  }'''


def write_pdf_report(html_path, pdf_path):
    """使用 Chrome/Edge 无头浏览器将 HTML 转为 PDF。"""
    from generate_report import find_browser
    import subprocess
    browser = find_browser()
    if not browser:
        print("[WARN] 未找到 Chrome/Edge 浏览器，跳过 PDF")
        return
    subprocess.run([browser, "--headless", "--disable-gpu",
                    f"--print-to-pdf={pdf_path}", "--no-pdf-header-footer",
                    f"file:///{html_path}"],
                   capture_output=True, timeout=30)
    if os.path.exists(pdf_path):
        print(f"[OK] PDF 报告已写入: {pdf_path}")
    else:
        print("[WARN] PDF 生成失败")


# ---------------------------------------------------------------------------
# 累积汇总 CSV — 每次运行追加一行，便于对比历史
# ---------------------------------------------------------------------------
SUMMARY_COLUMNS = [
    "模型名", "测试时间", "梯度并发范围", "梯度步长", "输入Tokens",
    "TPS (tok/s)", "TPM (tok/min)", "Decode TPS", "Decode TPM", "请求成功率",
    "Token 总数", "平均延迟 (s)",
    "TTFT-avg(s)", "TTFT-p99(s)", "TPOT-avg(ms)", "Decode (tok/s)",
    "缓存命中率",
    "TC-01连通性", "TC-02 TPS压测", "TC-03 TPM换算",
    "TC-04上下文", "TC-05鉴权", "TC-06限流",
    "TC-07工具调用", "TC-08流式性能", "TC-09缓存命中", "TC-10 IO档位性能",
    "总通过", "总失败", "测试结论",
]


def _build_summary_row(results: list, cfg: dict, test_time: str) -> list:
    """从测试结果构建一行汇总数据（与 SUMMARY_COLUMNS 对齐）。"""
    total_cases = len(results)
    passed_cases = sum(1 for r in results if r.get("passed"))
    failed_cases = total_cases - passed_cases
    total_tokens = 0
    total_latency = 0.0
    latency_count = 0

    for r in results:
        d = r.get("detail", {})
        if isinstance(d, dict):
            if "total_tokens" in d:
                total_tokens += int(d["total_tokens"])
            if "latency" in d:
                try:
                    total_latency += float(d["latency"])
                    latency_count += 1
                except (ValueError, TypeError):
                    pass

    avg_latency = round(total_latency / latency_count, 3) if latency_count else 0
    success_rate = f"{passed_cases / total_cases * 100:.1f}%" if total_cases else "N/A"
    conclusion = "通过" if failed_cases == 0 else f"{failed_cases}/{total_cases} 未通过"

    tps_val = "N/A"
    tpm_val = "N/A"
    decode_tps_val = "N/A"
    decode_tpm_val = "N/A"
    for r in results:
        if r.get("case_id") == "TC-02" and r.get("detail"):
            d = r["detail"]
            tps_val = d.get("tps_tokens", "N/A")
            tpm_val = d.get("tpm_tokens", "N/A")
            decode_tps_val = d.get("decode_tps", "N/A")
            decode_tpm_val = d.get("decode_tpm", "N/A")

    # ── 各用例通过状态（失败时附简短原因）──
    case_status = {}
    for r in results:
        cid = r.get("case_id", "?")
        if r.get("passed"):
            detail = r.get("detail", {})
            if cid == "TC-04" and detail.get("mode") == "multi_tier":
                case_status[cid] = f"通过({detail.get('passed_count', '?')})"
            elif cid == "TC-09":
                case_status[cid] = f"通过({detail.get('cache_hit_rate_pct', 'N/A')}%)"
            else:
                case_status[cid] = "通过"
        else:
            detail = r.get("detail", {})
            if cid == "TC-04" and detail.get("mode") == "multi_tier":
                reason = f"全部{len(detail.get('tiers', []))}档失败"
            elif cid == "TC-09":
                reason = f"未命中({detail.get('cache_hit_rate_pct', 0)}%)"
            else:
                reason = (
                    detail.get("failure_reason") or
                    detail.get("reason") or
                    detail.get("error") or
                    "失败"
                )
            reason = str(reason)[:80].replace("\n", " ").replace(",", "，")
            case_status[cid] = f"失败: {reason}"

    # ── 提取流式性能指标（精简版：TTFT-avg, TTFT-p99, TPOT-avg, Decode）──
    ttft_avg_s = "N/A"
    ttft_p99_s = "N/A"
    tpot_avg_s = "N/A"
    decode_tps_s = "N/A"
    cache_rate_s = "N/A"

    for r in results:
        if r.get("case_id") == "TC-08" and r.get("detail"):
            d = r["detail"]
            ttft_avg_s = str(d.get("ttft_avg", "N/A"))
            ttft_p99_s = str(d.get("ttft_p99", "N/A"))
            tpot_avg_s = str(d.get("tpot_avg_ms", "N/A"))
            decode_tps_s = str(d.get("decode_tps", "N/A"))
            cache_rate_s = str(d.get("cache_hit_rate", "N/A"))

    # ── 如果 TC-09 有缓存命中率，优先用 TC-09 的 ──
    for r in results:
        if r.get("case_id") == "TC-09" and r.get("detail"):
            d = r["detail"]
            cache_rate_s = f"{d.get('cache_hit_rate_pct', 'N/A')}%"

    return [
        cfg["api"].get("model", ""),
        test_time,
        f"{cfg.get('benchmark', {}).get('gradient', {}).get('start', '')}→{cfg.get('benchmark', {}).get('gradient', {}).get('max', '')}",
        str(cfg.get("benchmark", {}).get("gradient", {}).get("step", "")),
        str(cfg.get("benchmark", {}).get("input_tokens", "")),
        tps_val,
        tpm_val,
        decode_tps_val,
        decode_tpm_val,
        success_rate,
        str(total_tokens),
        str(avg_latency),
        ttft_avg_s,
        ttft_p99_s,
        tpot_avg_s,
        decode_tps_s,
        cache_rate_s,
        case_status.get("TC-01", ""),
        case_status.get("TC-02", ""),
        case_status.get("TC-03", ""),
        case_status.get("TC-04", ""),
        case_status.get("TC-05", ""),
        case_status.get("TC-06", ""),
        case_status.get("TC-07", ""),
        case_status.get("TC-08", ""),
        case_status.get("TC-09", ""),
        case_status.get("TC-10", ""),
        str(passed_cases),
        str(failed_cases),
        conclusion,
    ]


def _read_csv_rows(path: Path) -> tuple:
    """读取 CSV 文件，返回 (headers: list, rows: list[list[str]])。
    文件不存在或为空时返回 (None, [])。"""
    if not path.exists():
        return None, []
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            all_rows = list(reader)
    except Exception:
        return None, []
    if not all_rows:
        return None, []
    return all_rows[0], all_rows[1:]


def append_summary_csv(results: list, cfg: dict, output_path: Path, test_time: str) -> None:
    """
    在累积汇总 CSV 中追加一行。
    - 文件不存在 → 创建并写入表头 + 数据
    - 文件存在且表头一致 → 追加一行
    - 文件存在但表头不一致 → 读取旧数据，用新表头重建整个文件
    """
    new_row = _build_summary_row(results, cfg, test_time)
    existing_headers, existing_rows = _read_csv_rows(output_path)

    # ── 情况 1: 文件不存在，直接创建 ──
    if existing_headers is None:
        with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(SUMMARY_COLUMNS)
            writer.writerow(new_row)
        print(f"[OK] 累积汇总已创建: {output_path} (+1 行)")
        return

    # ── 情况 2: 表头一致，直接追加 ──
    if existing_headers == SUMMARY_COLUMNS:
        with open(output_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(new_row)
        print(f"[OK] 累积汇总已更新: {output_path} (+1 行)")
        return

    # ── 情况 3: 表头不一致（升级了列定义），重建文件 ──
    print(f"[INFO] 检测到表头变化，正在迁移旧数据...")
    print(f"       旧列数: {len(existing_headers)} → 新列数: {len(SUMMARY_COLUMNS)}")

    # 将旧行映射为新列（同名列保留值，新列留空，旧列丢弃）
    migrated_rows = []
    for old_row in existing_rows:
        if not old_row:
            continue
        new_row_mapped = []
        for col_name in SUMMARY_COLUMNS:
            try:
                idx = existing_headers.index(col_name)
                new_row_mapped.append(old_row[idx] if idx < len(old_row) else "")
            except ValueError:
                new_row_mapped.append("")  # 新列，旧数据没有
        migrated_rows.append(new_row_mapped)

    # 追加当前新行
    migrated_rows.append(new_row)

    # 重写整个文件
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(SUMMARY_COLUMNS)
        writer.writerows(migrated_rows)

    print(f"[OK] 累积汇总已重建: {output_path} ({len(migrated_rows)} 行数据，含本次)")

# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    args = parse_args()
    cfg = build_cfg(args)

    print("=" * 60)
    print("  Token 接口快速测试工具")
    print("=" * 60)
    print(f"  URL:      {cfg['api']['url']}")
    print(f"  Model:    {cfg['api']['model']}")
    print(f"  Timeout:  {cfg['api']['timeout']}s")
    grad = cfg["benchmark"].get("gradient", {})
    print(f"  梯度并发: {grad.get('start', '?')} → {grad.get('max', '?')} (步长 {grad.get('step', '?')})")
    print(f"  输入 tokens: {cfg['benchmark'].get('input_tokens', 1000)}")
    print(f"  上下文阈值: {cfg['context_test']['target_tokens']} tokens")
    if args.probe_context:
        print(f"  上下文模式: 渐进式探测")
    print()

    # ── 查询模型信息 ──
    api = cfg["api"]
    print("[INFO] 查询模型信息...")
    model_info = query_model_info(api["url"], api["key"], api["model"], api.get("timeout", 30))
    if model_info["ok"]:
        if model_info["context_length"]:
            print(f"       模型声明上下文长度: {model_info['context_length']} tokens")
        else:
            print(f"       模型列表可用，但未声明上下文长度")
    else:
        print(f"       无法查询模型信息: {model_info.get('error', '未知')}")

    # 加载用例
    cases_data = load_json(args.cases)
    cases = cases_data.get("cases", [])
    print(f"  用例文件:   {args.cases} ({len(cases)} 条)")

    # ── 独立开关：--io-benchmark 仅运行 TC-10 分档位性能测试 ──
    if args.io_benchmark:
        print("\n[INFO] --io-benchmark 已启用，仅运行 TC-10 分输入输出性能测试\n")
        # 混合分布模式：默认从 test_cases.json 的 TC-10 定义读取 anchors；
        # 若显式传了 --io-tiers 则退化为分档位模式（解析 in:out 列表）
        # 若传了 --naive-io-tier 则使用朴素四档测试模式
        src = next((c for c in cases if c["id"] == "TC-10"), None)
        src_params = (src or {}).get("params", {})

        if args.naive_io_tier:
            # 朴素扫描：默认串行, --io-concurrency N (N>1) 每步×N 全并发
            rep = args.io_concurrency or 1  # None/1=串行, N>1=并发×N
            sweep, sweep_label = _build_naive_sweep(args, max_context=args.io_max_context)
            params = {
                "sweep": sweep,
                "repetitions": rep,  # 1=串行, N>1=每步N条并发
                "max_output": args.io_max_context or 8192,
            }
            method = "io_sweep_benchmark"
            if rep > 1:
                sweep_label += f" [每步×{rep}, 按档依次并发]"
            else:
                sweep_label += " [串行]"
            print(f"         [朴素扫描] {sweep_label}")
        elif args.io_tiers:
            # 分档位模式
            tiers = parse_io_tiers(args.io_tiers)
            params = {
                "concurrency": args.io_concurrency or cfg["benchmark"].get("concurrency", 4),
                "total_requests": args.io_requests or 10,
                "tiers": tiers,
            }
            method = "io_tier_benchmark"
        else:
            # 混合分布模式（默认）：用 anchors 生成 N 条请求
            params = {
                "concurrency": args.io_concurrency or src_params.get("concurrency") or cfg["benchmark"].get("concurrency", 4),
                "total_requests": args.io_requests or src_params.get("total_requests") or 100,
                "max_output": src_params.get("max_output", 8192),
                "max_context": args.io_max_context,
                "in_anchors": src_params.get("in_anchors"),
                "out_anchors": src_params.get("out_anchors"),
            }
            method = "io_mix_benchmark"

        cases = [{
            "id": "TC-10",
            "name": "分输入输出档位性能",
            "type": "性能",
            "input": "命令行独立触发",
            "expected": "各请求成功并输出分位数",
            "method": method,
            "params": params,
        }]

    # ── 朴素扫描模式：替换完整测试套件中的 TC-10 ──
    elif args.naive_io_tier:
        rep = args.io_concurrency or 1  # None/1=串行, N>1=并发×N
        sweep, sweep_label = _build_naive_sweep(args, max_context=args.io_max_context)
        sweep_label += f" [{'每步×'+str(rep)+'按档并发' if rep > 1 else '串行'}]"
        for c in cases:
            if c["id"] == "TC-10":
                c["method"] = "io_sweep_benchmark"
                c["params"] = {
                    "sweep": sweep,
                    "repetitions": rep,
                    "max_output": args.io_max_context or 8192,
                }
                print(f"\n[INFO] --naive-io-tier 已启用，TC-10: {sweep_label}\n")
                break

    # 检查 Key
    key = cfg["api"].get("key", "")
    if not key or key == "sk-your-api-key-here":
        print("\n[WARN] !! API Key 未配置！")
        print("       请通过 -k 参数或 test_config.json 提供真实的 API Key。")
        print("       将使用占位 Key 继续执行（大部分用例会失败）。\n")

    # 注册测试方法
    methods = {
        "connectivity": test_connectivity,
        "tps_benchmark": test_tps_benchmark,
        "tpm_calc": test_tpm_calc,
        "context_limit": test_context_limit,
        "auth_failure": test_auth_failure,
        "rate_limit": test_rate_limit,
        "tool_calling": test_tool_calling,
        "streaming_benchmark": test_streaming_benchmark,
        "cache_hit": test_cache_hit,
        "io_sweep_benchmark": test_io_sweep_benchmark,
        "io_tier_benchmark": test_io_tier_benchmark,
        "io_mix_benchmark": test_io_mix_benchmark,
    }

    results = []
    prev_results = {}

    print("\n[RUN] 开始执行测试用例...\n")

    for i, case in enumerate(cases, 1):
        cid = case["id"]
        method_name = case.get("method", "")
        print(f"  [{i}/{len(cases)}] {cid} {case['name']} ... ", end="", flush=True)

        if method_name not in methods:
            print(f"SKIP (未知方法: {method_name})")
            continue

        try:
            if method_name == "tpm_calc":
                result = test_tpm_calc(cfg, case, prev_results.get("TC-02"))
            elif method_name == "context_limit":
                result = test_context_limit(cfg, case, probe_mode=args.probe_context,
                                           model_info=model_info)
            else:
                func = methods[method_name]
                result = func(cfg, case)

            results.append(result)
            prev_results[cid] = result
            status = "[PASS]" if result["passed"] else "[FAIL]"
            print(status)

            # 失败时立即打印原因
            if not result["passed"]:
                detail = result.get("detail", {})
                if detail.get("mode") == "multi_tier":
                    for t in detail.get("tiers", []):
                        if not t.get("ok"):
                            reason = t.get("failure_reason") or t.get("error") or ""
                            if reason:
                                print(f"         [{t['target']//1000}k] -> {reason[:180]}")
                            if t.get("is_timeout"):
                                print(f"         -> 建议缩短 context_test.timeout 或使用 --probe-context")
                else:
                    reason = detail.get("failure_reason") or detail.get("reason") or detail.get("error") or ""
                    if reason:
                        print(f"         -> {reason[:200]}")
                    if detail.get("is_timeout"):
                        print(f"         -> 建议使用 --probe-context 渐进探测实际上下文上限")
        except Exception as e:
            print(f"[ERROR] {e}")
            traceback.print_exc()
            error_result = {
                "case_id": cid,
                "passed": False,
                "detail": {"error": str(e), "traceback": traceback.format_exc()},
            }
            results.append(error_result)
            prev_results[cid] = error_result

    # 输出汇总
    print("\n" + "=" * 60)
    passed = sum(1 for r in results if r.get("passed"))
    failed = len(results) - passed
    print(f"  总计: {len(results)} | 通过: {passed} | 失败: {failed}")
    print("=" * 60 + "\n")

    # ── 生成时间戳文件名 ──
    test_time = datetime.now(TZ_BJ).strftime("%Y-%m-%d %H:%M:%S")
    timestamp = datetime.now(TZ_BJ).strftime("%Y%m%d_%H%M%S")

    # 确保输出目录存在
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    if args.output is None:
        detail_output = RESULTS_DIR / f"test_output_{timestamp}.csv"
    else:
        detail_output = args.output

    # ── 写入详情 CSV（每次独立文件）──
    write_single_csv(results, cases, cfg, detail_output)

    # ── 追加累积汇总 CSV（所有运行汇集到一个文件，一行一次）──
    append_summary_csv(results, cfg, args.summary, test_time)

    # ── 验收报告（MD + HTML + PDF）──
    report_dir = RESULTS_MD_DIR / timestamp
    report_dir.mkdir(parents=True, exist_ok=True)
    md_path = report_dir / f"test_report_{timestamp}.md"
    html_path = report_dir / f"test_report_{timestamp}.html"
    pdf_path = report_dir / f"test_report_{timestamp}.pdf"

    write_markdown_report(results, cases, cfg, args, md_path, test_time)
    write_html_report(results, cases, cfg, args, html_path, test_time)
    try:
        write_pdf_report(str(html_path.resolve()), str(pdf_path))
    except Exception as e:
        print(f"[WARN] PDF 生成失败: {e}")

    print(f"\n[DONE] 测试完成")
    print(f"       详情: {detail_output}")
    print(f"       报告: {md_path}")
    print(f"             {html_path}")
    if pdf_path.exists():
        print(f"             {pdf_path}")
    print(f"       汇总: {args.summary}  <- 累积对比，每次追加一行")
    print(f"       用例: {args.cases}  <- 持久化，可复现")
    print(f"\n  重新运行: python token_test.py -u {cfg['api']['url']} -k $KEY -m {cfg['api']['model']}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())