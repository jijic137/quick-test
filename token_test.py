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
SUMMARY_PATH  = BASE_DIR / "test_summary.csv"         # 累积汇总文件（每次追加一行）

# 北京时间
TZ_BJ = timezone(timedelta(hours=8))

# ── 线程安全存储 ────────────────────────────────────────────
_stats_lock = threading.Lock()
_req_counter = [0]  # 全局请求计数器（用 list 避免 nonlocal）

# ── 工具函数 ────────────────────────────────────────────────

def _gen_request_id() -> str:
    """生成唯一请求 ID，用于链路追踪。"""
    _req_counter[0] += 1
    ts = datetime.now(TZ_BJ).strftime("%Y%m%d%H%M%S%f")
    return f"req_{ts}_{_req_counter[0]:06d}"


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
        "--probe-context",
        action="store_true",
        default=False,
        help="启用渐进式上下文探测：从小量递增直到找到实际上限（较慢但精确）",
    )

    return parser.parse_args()


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
        "context_tokens": ctx.get("target_tokens", builtin["context_tokens"]),
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

            # 首 token 时间
            choices = chunk.get("choices", [])
            if choices and choices[0].get("delta", {}).get("content"):
                if first_token is None:
                    first_token = now
                last_token_time = now
                completion_count += 1
                result["token_times"].append(round(now - t_start, 4))

            # usage 通常在最后一块返回
            if "usage" in chunk and chunk["usage"]:
                u = chunk["usage"]
                result["prompt_tokens"] = int(u.get("prompt_tokens", 0))
                result["completion_tokens"] = int(u.get("completion_tokens", 0))
                result["total_tokens"] = int(u.get("total_tokens", 0))
                # 提取缓存命中 token 数
                ptd = u.get("prompt_tokens_details") or u.get("prompt_tokens_detail") or {}
                result["cached_tokens"] = int(ptd.get("cached_tokens", 0))

        result["total_latency"] = round(time.perf_counter() - t_start, 3)
        result["ok"] = first_token is not None

        if first_token:
            result["ttft"] = round(first_token - t_start, 3)
            # 如果 usage 没返回，用 token_times 数量
            if result["completion_tokens"] == 0:
                result["completion_tokens"] = completion_count
        else:
            result["error"] = "未收到任何 token"

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
    chars_per_token = cfg.get("context_test", {}).get("chars_per_token_estimate", 2.5)
    timeout = max(api.get("timeout", 60), params.get("timeout", 600))
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
        chars = int(test_size * chars_per_token)
        chunk = "测" * min(chars, 2_000_000)
        messages = [{"role": "user", "content": prompt_template.format(padding=chunk)}]

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
            chars = int(mid * chars_per_token)
            chunk = "测" * min(chars, 2_000_000)
            messages = [{"role": "user", "content": prompt_template.format(padding=chunk)}]

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
        },
        "context_test": {
            "target_tokens": args.context_tokens,
            "chars_per_token_estimate": 2.5,
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
    """TC-02: TPS 压测 — 并发请求，计算 TPS/TPM。"""
    api = cfg["api"]
    params = case.get("params", {})
    concurrency = cfg["benchmark"].get("concurrency", params.get("concurrency", 4))
    total = cfg["benchmark"].get("total_requests", params.get("total_requests", 20))
    max_tokens = params.get("max_tokens", 100)
    messages = params.get("messages", [{"role": "user", "content": "Hi"}])

    stats = {
        "total": total,
        "concurrency": concurrency,
        "ok": 0,
        "fail": 0,
        "total_latency": 0.0,
        "total_prompt_tokens": 0,
        "total_completion_tokens": 0,
        "total_tokens": 0,
        "errors": [],
    }

    t_start = time.perf_counter()

    def _worker(_idx):
        # 每个请求在末尾追加唯一标记，避免 KV-cache 命中导致 TPS 虚高
        varied_msgs = [dict(m) for m in messages]
        varied_msgs[-1]["content"] += f"\n\n[req:{_idx}]"
        r = api_request(
            url=api["url"], key=api["key"], model=api["model"],
            messages=varied_msgs, max_tokens=max_tokens,
            timeout=api.get("timeout", 60),
        )
        with _stats_lock:
            if r["ok"]:
                stats["ok"] += 1
                stats["total_latency"] += r["latency"]
                if r["usage"]:
                    stats["total_prompt_tokens"] += r["usage"]["prompt"]
                    stats["total_completion_tokens"] += r["usage"]["completion"]
                    stats["total_tokens"] += r["usage"]["total"]
            else:
                stats["fail"] += 1
                if r["error"] and len(stats["errors"]) < 20:
                    stats["errors"].append(r["error"])
        return r

    with ThreadPoolExecutor(max_workers=concurrency) as pool:
        futures = [pool.submit(_worker, i) for i in range(total)]
        for f in as_completed(futures):
            f.result()

    elapsed = time.perf_counter() - t_start
    success_rate = stats["ok"] / total * 100 if total else 0
    tps_tokens = stats["total_tokens"] / elapsed if elapsed > 0 else 0
    avg_latency = stats["total_latency"] / stats["ok"] if stats["ok"] else 0

    passed = stats["ok"] > 0 and tps_tokens > 0

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "elapsed": round(elapsed, 2),
            "ok": stats["ok"],
            "fail": stats["fail"],
            "success_rate": f"{success_rate:.1f}%",
            "tps_tokens": f"{tps_tokens:.1f}",
            "tpm_tokens": f"{tps_tokens * 60:.1f}",
            "avg_latency": round(avg_latency, 3),
            "total_tokens": stats["total_tokens"],
            "errors": stats["errors"][:5],
        },
    }


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
    """TC-04: 上下文窗口快速验收，优先查模型声明，支持渐进探测。"""
    api = cfg["api"]
    params = case.get("params", {})
    target_tokens = cfg.get("context_test", {}).get("target_tokens", params.get("target_tokens", 512000))
    chars_per_token = cfg.get("context_test", {}).get("chars_per_token_estimate", 2.5)
    timeout = max(api.get("timeout", 60), params.get("timeout", 600))

    # ── 渐进式探测模式 ──
    if probe_mode:
        probe_result = probe_context_limit(cfg, params)
        found = probe_result.get("found_limit", 0)
        return {
            "case_id": case["id"],
            "passed": probe_result["passed"],
            "detail": {
                "mode": "probe",
                "target": target_tokens,
                "found_limit": found,
                "threshold_reached": found is not None and found >= target_tokens,
                "attempts_count": len(probe_result.get("attempts", [])),
                "attempts": probe_result.get("attempts", []),
            },
        }

    # ── 快速模式：优先用模型声明判断 ──
    declared = None
    if model_info and model_info.get("ok") and model_info.get("context_length"):
        declared = model_info["context_length"]

    if declared is not None:
        passed = declared >= target_tokens
        return {
            "case_id": case["id"],
            "passed": passed,
            "detail": {
                "mode": "declared",
                "target": target_tokens,
                "declared_context_length": declared,
                "passed": passed,
                "reason": (
                    f"模型声明上下文长度 {declared} >= {target_tokens}，通过"
                    if passed else
                    f"模型声明上下文长度 {declared} < {target_tokens}，不满足要求"
                ),
            },
        }

    # ── 未声明：发一次请求验证 ──
    total_chars = int(target_tokens * chars_per_token)
    chunk = "测" * min(int(total_chars), 2_000_000)

    prompt_template = params.get(
        "prompt_template",
        "请总结以下文本的开头3个字和结尾3个字，用英文回复: {padding}"
    )
    messages = [
        {"role": "user", "content": prompt_template.format(padding=chunk)}
    ]

    t0 = time.perf_counter()
    result = api_request(
        url=api["url"], key=api["key"], model=api["model"],
        messages=messages,
        max_tokens=params.get("max_output_tokens", 50),
        timeout=timeout,
    )
    elapsed = round(time.perf_counter() - t0, 2)

    # ── 失败原因分析 ──
    failure_reason = None
    is_context_error = False
    is_timeout = False

    if not result["ok"]:
        status_code = result.get("status_code")
        error_msg = result.get("error") or ""

        if error_msg == "请求超时":
            is_timeout = True
            failure_reason = (
                f"请求超时 ({elapsed}s, timeout={timeout}s)。"
                "可能原因: 1) 模型处理长上下文耗时超过超时设置; "
                "2) 网关/代理层超时; 3) 模型不支持此长度导致无响应。"
                "建议: 使用 --probe-context 渐进探测实际支持的上限。"
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

        err_lower = error_msg.lower()
        is_context_error = is_context_error or any(
            kw in err_lower
            for kw in ["context", "token", "length", "limit", "maximum", "exceed", "too long", "truncat"]
        )

    passed = result["ok"]

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "mode": "live_test",
            "target": target_tokens,
            "declared": declared,
            "status_code": result["status_code"],
            "ok": result["ok"],
            "latency": elapsed,
            "is_context_error": is_context_error,
            "is_timeout": is_timeout,
            "failure_reason": failure_reason,
            "error": result.get("error"),
            "usage": result["usage"],
        },
    }


def test_auth_failure(cfg: dict, case: dict) -> dict:
    """TC-05: 鉴权失败 — 使用无效 Key，应返回 401/403。"""
    api = cfg["api"]
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
    max_tokens = params.get("max_tokens", 300)
    messages = params.get("messages", [{"role": "user", "content": "Hello"}])

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

    passed = stats["ok"] > 0

    return {
        "case_id": case["id"],
        "passed": passed,
        "detail": {
            "elapsed": round(elapsed, 2),
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
                context_passed = "是" if r.get("passed") else "否"
                if d.get("error") and not r.get("passed"):
                    context_error = str(d["error"])[:200]
                estimated_input = d.get("estimated_input_tokens", "N/A")

    avg_latency = round(total_latency / latency_count, 3) if latency_count else "N/A"
    success_rate = f"{passed_cases / total_cases * 100:.1f}%" if total_cases else "N/A"
    test_time = datetime.now(TZ_BJ).strftime("%Y-%m-%d %H:%M:%S")
    conclusion = "通过" if failed_cases == 0 else f"{failed_cases}/{total_cases} 未通过"

    tps_val = "N/A"
    tpm_val = "N/A"
    for r in results:
        if r.get("case_id") == "TC-02" and r.get("detail"):
            d = r["detail"]
            tps_val = d.get("tps_tokens", "N/A")
            tpm_val = d.get("tpm_tokens", "N/A")

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
        ["上下文验收门槛 (tokens)", str(cfg.get("context_test", {}).get("target_tokens", 512000))],
        ["实测输入 Token 估算", str(estimated_input)],
        ["512k 上下文是否通过", context_passed],
        ["上下文测试错误", context_error or "无"],
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
                actual = (
                    f"成功 {detail.get('ok')}/{detail.get('ok',0)+detail.get('fail',0)}, "
                    f"TPS={detail.get('tps_tokens')} tok/s, "
                    f"TPM={detail.get('tpm_tokens')} tok/min, "
                    f"平均延迟 {detail.get('avg_latency')}s"
                )
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
                        f"阈值 {detail.get('target')}, {detail.get('reason', '')}"
                    )
                elif mode == "probe":
                    actual = (
                        f"探测上限: {detail.get('found_limit')} tokens, "
                        f"阈值 {detail.get('target')}, "
                        f"{'达到' if detail.get('threshold_reached') else '未达到'}, "
                        f"共 {detail.get('attempts_count', 0)} 次尝试"
                    )
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
                if detail.get("verdict"):
                    actual += f" | {detail['verdict'][:150]}"
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
# 累积汇总 CSV — 每次运行追加一行，便于对比历史
# ---------------------------------------------------------------------------

SUMMARY_COLUMNS = [
    "模型名", "测试时间", "并发数", "总请求数",
    "TPS (tok/s)", "TPM (tok/min)", "请求成功率",
    "Token 总数", "平均延迟 (s)", "Prefill (tok/s)", "Decode (tok/s)",
    "TTFT (s)", "ITL (ms)", "TPOT (ms/tok)",
    "TC-01 连通性", "TC-02 TPS压测", "TC-03 TPM换算",
    "TC-04 上下文", "TC-05 鉴权", "TC-06 限流",
    "TC-07 工具调用", "TC-08 流式性能",
    "总通过", "总失败", "测试结论",
]


def append_summary_csv(results: list, cfg: dict, output_path: Path, test_time: str) -> None:
    """
    在累积汇总 CSV 中追加一行。
    如果文件不存在则先写入表头。
    """
    # ── 汇总数据 ──
    total_cases = len(results)
    passed_cases = sum(1 for r in results if r.get("passed"))
    failed_cases = total_cases - passed_cases
    total_tokens = 0
    total_latency = 0.0
    latency_count = 0
    context_passed = "否"

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
                context_passed = "是" if r.get("passed") else "否"

    avg_latency = round(total_latency / latency_count, 3) if latency_count else 0
    success_rate = f"{passed_cases / total_cases * 100:.1f}%" if total_cases else "N/A"
    conclusion = "通过" if failed_cases == 0 else f"{failed_cases}/{total_cases} 未通过"

    tps_val = "N/A"
    tpm_val = "N/A"
    for r in results:
        if r.get("case_id") == "TC-02" and r.get("detail"):
            d = r["detail"]
            tps_val = d.get("tps_tokens", "N/A")
            tpm_val = d.get("tpm_tokens", "N/A")

    # ── 各用例通过状态（失败时附简短原因）──
    case_status = {}
    for r in results:
        cid = r.get("case_id", "?")
        if r.get("passed"):
            case_status[cid] = "通过"
        else:
            detail = r.get("detail", {})
            # 取最精炼的失败原因
            reason = (
                detail.get("failure_reason") or
                detail.get("reason") or
                detail.get("error") or
                "失败"
            )
            # 截断到 80 字符，适合 CSV 单元格
            reason = str(reason)[:80].replace("\n", " ").replace(",", "，")
            case_status[cid] = f"失败: {reason}"

    # ── 提取流式性能指标 ──
    prefill_tps_s = "N/A"
    decode_tps_s = "N/A"
    ttft_s = "N/A"
    itl_s = "N/A"
    tpot_s = "N/A"
    for r in results:
        if r.get("case_id") == "TC-08" and r.get("detail"):
            d = r["detail"]
            prefill_tps_s = str(d.get("prefill_tps", "N/A"))
            decode_tps_s = str(d.get("decode_tps", "N/A"))
            ttft_s = f"{d.get('ttft_avg', 'N/A')} (p50:{d.get('ttft_p50','')} p99:{d.get('ttft_p99','')})"
            itl_s = f"{d.get('itl_avg_ms', 'N/A')} (max p99:{d.get('itl_max_p99_ms','')})"
            tpot_s = f"{d.get('tpot_avg_ms', 'N/A')} (p50:{d.get('tpot_p50','')} p99:{d.get('tpot_p99','')})"

    # ── 构建行数据 ──
    row = [
        cfg["api"].get("model", ""),
        test_time,
        str(cfg.get("benchmark", {}).get("concurrency", "")),
        str(cfg.get("benchmark", {}).get("total_requests", "")),
        tps_val,
        tpm_val,
        success_rate,
        str(total_tokens),
        str(avg_latency),
        prefill_tps_s,
        decode_tps_s,
        ttft_s,
        itl_s,
        tpot_s,
        case_status.get("TC-01", ""),
        case_status.get("TC-02", ""),
        case_status.get("TC-03", ""),
        case_status.get("TC-04", ""),
        case_status.get("TC-05", ""),
        case_status.get("TC-06", ""),
        case_status.get("TC-07", ""),
        case_status.get("TC-08", ""),
        str(passed_cases),
        str(failed_cases),
        conclusion,
    ]

    # ── 写入（带表头判断）──
    file_exists = output_path.exists()
    with open(output_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(SUMMARY_COLUMNS)
        writer.writerow(row)

    print(f"[OK] 累积汇总已更新: {output_path} (+1 行)")

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
    print(f"  并发/请求: {cfg['benchmark']['concurrency']}/{cfg['benchmark']['total_requests']}")
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
                reason = detail.get("failure_reason") or detail.get("reason") or detail.get("error") or ""
                if reason:
                    print(f"         -> {reason[:200]}")
                if detail.get("mode") == "live_test" and detail.get("is_timeout"):
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

    print(f"\n[DONE] 测试完成")
    print(f"       详情: {detail_output}")
    print(f"       汇总: {args.summary}  <- 累积对比，每次追加一行")
    print(f"       用例: {args.cases}  <- 持久化，可复现")
    print(f"\n  重新运行: python token_test.py -u {cfg['api']['url']} -k $KEY -m {cfg['api']['model']}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
