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
        type=int,
        default=defaults["context_tokens"],
        help="上下文验收阈值 tokens (默认: 512000)",
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
    "Token 总数", "平均延迟 (s)",
    "TC-01 连通性", "TC-02 TPS压测", "TC-03 TPM换算",
    "TC-04 上下文", "TC-05 鉴权", "TC-06 限流",
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
        case_status.get("TC-01", ""),
        case_status.get("TC-02", ""),
        case_status.get("TC-03", ""),
        case_status.get("TC-04", ""),
        case_status.get("TC-05", ""),
        case_status.get("TC-06", ""),
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
