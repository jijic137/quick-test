# Token 接口快速测试工具

针对 OpenAI 兼容 Chat Completions API 的自动化测试套件。覆盖功能验证、性能压测、上下文容量、异常处理、流式性能、缓存命中、IO 档位扫描等 10 项用例。每次运行生成 CSV 详情 + MD 报告 + **深色主题 Chart.js 仪表盘 HTML**，并追加累积汇总表方便横向对比。

## 快速开始

```bash
# 1. 安装依赖
pip install requests markdown

# 2. 命令行运行
python token_test.py \
  -u https://your-api.example.com/v1/chat/completions \
  -k sk-your-api-key \
  -m gpt-4o

# 3. 查看结果
#    仪表盘: reports/20260723_150000/test_report_20260723_150000.html  ← 浏览器打开
#    详情:   results/test_output_20260723_150000.csv
#    MD:     reports/20260723_150000/test_report_20260723_150000.md
#    汇总:   test_summary.csv  （累积对比，每次追加一行）
```

也可以编辑 `test_config.json` 设置默认值后直接运行：

```bash
python token_test.py
```

## 命令行参数

### 基础配置

| 参数 | 简写 | 说明 | 默认值来源 |
|---|---|---|---|
| `--url` | `-u` | API 接口地址 | test_config.json |
| `--key` | `-k` | API Key | test_config.json |
| `--model` | `-m` | 模型名称 | test_config.json |
| `--timeout` | `-t` | 请求超时（秒） | test_config.json (60) |
| `--platform` | | 平台信息备注（如 `GPU: A100×8`），写入报告 | — |

### 通用压测参数

| 参数 | 简写 | 说明 | 默认值 |
|---|---|---|---|
| `--concurrency` | `-c` | 并发数 | 4 |
| `--requests` | `-n` | 总请求数 | 20 |
| `--input-tokens` | | 输入 token 数（支持 k/m 后缀，如 `1k`, `50k`） | 1k |
| `--max-workers` | | 线程池上限 | 50 |

### 上下文测试

| 参数 | 说明 | 默认值 |
|---|---|---|
| `--context-tokens` | 上下文验收阈值（支持 k/m 后缀） | 512k |
| `--probe-context` | 启用渐进式探测：从小量递增直到找到实际上限 | 关闭 |

### 梯度压测

| 参数 | 说明 | 默认值 |
|---|---|---|
| `--g-start` | 起始并发数 | 100 |
| `--g-step` | 并发步长 | 200 |
| `--g-max` | 最大并发数 | 取 `-c` 的值 |

### 输出路径

| 参数 | 简写 | 说明 | 默认值 |
|---|---|---|---|
| `--output` | `-o` | 详情 CSV 路径 | 自动生成时间戳文件名 |
| `--summary` | | 累积汇总 CSV 路径 | test_summary.csv |
| `--cases` | | 用例 JSON 路径 | test_cases.json |

### TC-10 分输入输出档位性能（专用）

| 参数 | 说明 |
|---|---|
| `--io-benchmark` | 独立开关：仅运行 TC-10，忽略其他用例 |
| `--io-tiers` | 自定义档位，格式 `in:out` 逗号分隔。例: `1k:300,8k:1000,32k:2000` |
| `--io-concurrency` | 并发数（混合分布模式统一并发；朴素模式默认全并发，=1 串行，>1 每步×N） |
| `--io-requests` | 总请求数（混合分布默认 100；朴素模式默认 38 步） |
| `--io-max-context` | 最大上下文上限（超出的输入跳过不测），支持 k/m 后缀 |
| `--naive-io-tier` | 朴素扫描模式：10K→380K 步长 10K，输出按档位固定。默认全并发 38 步同时发出 |

### TC-10 朴素扫描模式详解

```
输入:  10K  20K  30K  ...  380K  （步长 10K，共 38 步）
输出:  0.2K (≤50K) | 0.6K (60~80K) | 1.3K (90~160K) | 7K (>160K)
```

三种执行方式：

```bash
# 全并发（默认）：38 步同时发出，测并发争抢下的性能衰减曲线
python token_test.py --io-benchmark --naive-io-tier

# 串行：逐条发出，测纯输入→性能曲线（无并发干扰）
python token_test.py --io-benchmark --naive-io-tier --io-concurrency 1

# 并发×N：每步重复 N 条并发取均值 + 标准差（38×N 条同时发出）
python token_test.py --io-benchmark --naive-io-tier --io-concurrency 5
```

## 测试用例

所有用例定义在 [test_cases.json](test_cases.json)，可编辑扩展。

| 编号 | 名称 | 类型 | 方法 | 说明 |
|---|---|---|---|---|
| TC-01 | 接口连通性 | 功能 | `connectivity` | 单次请求，验证 HTTP 200 + usage 返回 |
| TC-02 | TPS 压测 | 性能 | `tps_benchmark` | 多并发请求，计算 TPS / 每并发 TPS |
| TC-03 | TPM 换算 | 性能 | `tpm_calc` | 验证 TPM = TPS × 60 |
| TC-04 | 上下文验收 | 容量 | `context_limit` | 发送指定长度上下文，验证不超限 |
| TC-05 | 鉴权失败 | 异常 | `auth_failure` | 无效 Key 应返回 401/403 |
| TC-06 | 限流/超时 | 异常 | `rate_limit` | 高频并发，观察 429 / 超时处理 |
| TC-07 | 工具调用 | 功能 | `tool_calling` | 验证 function calling 是否可用 |
| TC-08 | 流式性能 | 性能 | `streaming_benchmark` | 流式请求，测 TTFT / Decode TPS / TPOT / ITL |
| TC-09 | 缓存命中 | 性能 | `cache_hit` | 重复相同请求，检测前缀缓存命中率 |
| TC-10 | IO 档位性能 | 性能 | `io_mix_benchmark` / `io_tier_benchmark` / `io_sweep_benchmark` | 分输入输出档位压测，支持混合分布 / 分档位 / 朴素扫描三种模式 |

### TC-10 三种模式

| 模式 | 方法 | 触发方式 | 说明 |
|---|---|---|---|
| 混合分布 | `io_mix_benchmark` | `--io-benchmark`（默认） | 按 P50/P90/P99 锚点线性插值生成 N 条 (in, out) 请求，并发执行 |
| 分档位 | `io_tier_benchmark` | `--io-benchmark --io-tiers 1k:300,8k:1000` | 自定义多个固定 (in, out) 档位，每档独立并发 |
| 朴素扫描 | `io_sweep_benchmark` | `--io-benchmark --naive-io-tier` | 输入 10K→380K 步长 10K，输出按档位固定，全并发/串行可选 |

## 输出文件

每次运行生成以下文件：

```
reports/20260723_150000/
├── test_report_20260723_150000.md     ← Markdown 验收报告
├── test_report_20260723_150000.html   ← Chart.js 深色仪表盘（浏览器打开）
└── test_report_20260723_150000.pdf    ← PDF 导出（自动，需 Chrome/Edge）

results/
└── test_output_20260723_150000.csv    ← 详情 CSV（每次独立文件）

test_summary.csv                       ← 累积汇总（每次追加一行）
```

### HTML 仪表盘

深色主题，Chart.js 渲染，包含：

- **KPI 卡片**（7 个）：通过率、失败数、TC-10 请求数/成功率/TTFT/TPOT/Decode TPS
- **TC-10 混合分布图表**：Input/Output Tokens 分位数柱状图、TTFT / TPOT 分位数 vs SLO
- **TC-10 朴素扫描图表**：TTFT / Decode TPS / TPOT 随输入增长曲线，按档位着色，并发模式显示 ±1σ 误差
- **TC-02 并发梯度图表**：TPS + Decode TPS 双轴对比
- **全部用例汇总表**：ID / 名称 / 类型 / 状态 / 关键指标
- **TC-10 请求明细表**：分页（每页 20 条），扫描模式显示档位 + 标准差

### CSV 详情文件

单文件分上下两部分：

```
┌──────────────────────────────────────────────┐
│  Token 接口测试报告                           │  ← 报告区（2 列：指标 | 值）
│  被测接口 URL, https://...                    │
│  TPS (tokens/秒), 86.9                       │
│  ...                                         │
├──────────────────────────────────────────────┤
│  (空行)                                      │
├──────────────────────────────────────────────┤
│  编号, 用例名称, 测试类型, ..., 是否通过       │  ← 用例表（7 列）
│  TC-01, 接口连通性, 功能, ..., 通过           │
└──────────────────────────────────────────────┘
```

### 累积汇总 CSV

每次运行追加一行，字段包括：模型名、测试时间、并发数、总请求数、TPS、TPM、成功率、Token 总数、平均延迟、各 TC 状态、总通过/失败、测试结论。

## 流式性能指标说明

流式请求（TC-08、TC-10）使用 `api_request_stream()` 采集逐 token 时间戳，计算以下指标：

| 指标 | 公式 | 单位 | 说明 |
|---|---|---|---|
| **TTFT** | t₁ − t_start | 秒 (s) | Time To First Token，prefill 阶段耗时 |
| **Decode Time** | total_latency − TTFT | 秒 (s) | 纯 token 生成时间 |
| **Decode TPS** | completion_tokens ÷ Decode Time | tok/s | 单请求解码吞吐 |
| **Prefill TPS** | prompt_tokens ÷ TTFT | tok/s | prefill 阶段吞吐 |
| **TPOT** | Decode Time ÷ completion_tokens × 1000 | 毫秒 (ms) | Time Per Output Token，每 token 平均生成延迟 |
| **ITL** | tᵢ − tᵢ₋₁ | 秒 (s) | Inter-Token Latency，相邻 token 间隔 |

```
t_start                                                     t_end
  │                                                           │
  ├────────── Prefill ──────────┬────────── Decode ────────────┤
  │                             │                              │
  │<───────── TTFT ────────────>│                              │
  │                             │<────── Decode Time ─────────>│
  │                             t₁   t₂   t₃   ...   tₙ       │
  │                             │<ITL>│                        │
```

## 项目结构

```
quick-test/
├── README.md                          ← 本文件
├── token_test.py                      ← 主测试脚本
├── generate_report.py                 ← 报告生成模块（MD→HTML/PDF）
├── template.html                      ← HTML 模板（generate_report.py CLI 用）
├── test_config.json                   ← 默认配置（URL/Key/模型/参数）
├── test_cases.json                    ← 测试用例持久化定义
├── test_summary.csv                   ← 累积汇总（每次追加一行）
├── results/
│   └── test_output_*.csv              ← 各次运行详情 CSV
└── reports/
    └── YYYYMMDD_HHMMSS/
        ├── test_report_*.md           ← Markdown 报告
        ├── test_report_*.html         ← Chart.js 仪表盘
        └── test_report_*.pdf          ← PDF 导出
```

## 可复现性

- 用例定义在 `test_cases.json`，参数固定，可纳入 Git
- 配置通过 `test_config.json` 或命令行传入，相同输入 → 相同流程
- 每次运行生成独立时间戳目录，历史结果不丢失
- 汇总 CSV 累积所有运行记录，支持跨版本对比

## 自定义用例

编辑 `test_cases.json`，按以下结构添加用例：

```json
{
  "id": "TC-11",
  "name": "我的用例",
  "type": "功能",
  "input": "输入描述，支持 {{变量}}",
  "expected": "预期结果",
  "method": "connectivity",
  "params": {
    "max_tokens": 100,
    "messages": [{"role": "user", "content": "Hello"}]
  }
}
```

可用的 `method` 值：

| method | 对应测试函数 |
|---|---|
| `connectivity` | 基本连通性 |
| `tps_benchmark` | TPS 并发压测 |
| `tpm_calc` | TPM 换算验证 |
| `context_limit` | 上下文上限 |
| `auth_failure` | 鉴权失败 |
| `rate_limit` | 限流/超时 |
| `tool_calling` | 工具调用 |
| `streaming_benchmark` | 流式性能 |
| `cache_hit` | 缓存命中 |
| `io_tier_benchmark` | TC-10 分档位模式 |
| `io_mix_benchmark` | TC-10 混合分布模式 |
| `io_sweep_benchmark` | TC-10 朴素扫描模式 |

## License

MIT
