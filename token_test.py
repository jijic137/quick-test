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
        last_token_time = t_start
        completion_count = 0

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

            now = time.perf_counter()

            # 首 token 时间 — 精确判断 content/reasoning_content
            choices = chunk.get("choices", [])
            delta = choices[0].get("delta", {}) if choices else {}
            content_val = delta.get("content")
            reasoning_val = delta.get("reasoning_content")
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
            if first_token is None:
                first_token = t_start
            result["ttft"] = 0.001
            completion_count = uc
        elif first_token:
            result["ttft"] = round(first_token - t_start, 3)
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
        else:
            lines.append(f"```json\n{json.dumps(d, ensure_ascii=False, indent=2)}\n```")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Markdown 报告已写入: {output_path}")


def write_html_report(results, cases, cfg, args, output_path, test_time, model_info=None):
    """使用 template.html 从 Markdown 生成美观 HTML 报告。"""
    from generate_report import parse_markdown, generate_html
    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as tmp:
        md_tmp = tmp.name
    try:
        write_markdown_report(results, cases, cfg, args, Path(md_tmp), test_time)
        with open(md_tmp, 'r', encoding='utf-8') as f:
            md_text = f.read()
    finally:
        os.unlink(md_tmp)
    parsed = parse_markdown(md_text)
    html = generate_html(parsed)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] HTML 报告已写入: {output_path}")


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
    "TC-07工具调用", "TC-08流式性能", "TC-09缓存命中",
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