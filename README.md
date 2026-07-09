# Token 接口快速测试工具

针对 OpenAI 兼容 Chat Completions API 的自动化测试套件。测试用例持久化、可复现，每次运行生成独立结果文件并追加到累积汇总表，方便横向对比不同模型/配置的表现。

## 快速开始

```bash
# 1. 安装依赖
pip install requests

# 2. 命令行运行（推荐）
python token_test.py \
  -u https://your-api.example.com/v1/chat/completions \
  -k sk-your-api-key \
  -m gpt-4o

# 3. 查看结果
#    详情: results/test_output_20260709_180000.csv  （每次运行独立文件）
#    汇总: test_summary.csv                  （累积对比，每次追加一行）
```

也可以编辑 `test_config.json` 设置默认值后直接运行：

```bash
python token_test.py
```

## 命令行参数

| 参数 | 简写 | 说明 | 默认值来源 |
|---|---|---|---|
| `--url` | `-u` | API 接口地址 | test_config.json |
| `--key` | `-k` | API Key | test_config.json |
| `--model` | `-m` | 模型名称 | test_config.json |
| `--concurrency` | `-c` | 压测并发数 | test_config.json |
| `--requests` | `-n` | 压测总请求数 | test_config.json |
| `--timeout` | `-t` | 请求超时（秒） | test_config.json |
| `--output` | `-o` | 详情输出路径 | 自动生成时间戳文件名 |
| `--summary` | | 累积汇总路径 | test_summary.csv |
| `--cases` | | 用例 JSON 路径 | test_cases.json |
| `--context-tokens` | | 上下文验收阈值 | test_config.json |

## 测试用例

所有用例定义在 [test_cases.json](test_cases.json)，可编辑、扩展，纳入版本管理。

| 编号 | 名称 | 类型 | 说明 |
|---|---|---|---|
| TC-01 | 接口连通性 | 功能 | 单次请求，验证 HTTP 200 + usage |
| TC-02 | TPS 压测 | 性能 | 并发请求，计算 TPS / TPM |
| TC-03 | TPM 换算 | 性能 | 验证 TPM = TPS × 60 |
| TC-04 | 512k 上下文验收 | 容量 | 发送超长上下文，验证不超限 |
| TC-05 | 鉴权失败 | 异常 | 无效 Key 应返回 401/403 |
| TC-06 | 限流/超时 | 异常 | 高频并发，观察 429 / 超时处理 |

## 输出文件

### 详情文件 `test_output_YYYYMMDD_HHMMSS.csv`

每次运行生成一个带时间戳的独立文件，内容分上下两部分：

```
┌──────────────────────────────────────┐
│  Token 接口测试报告                    │  ← 报告区（2列：指标 | 值）
│  被测接口 URL, https://...            │
│  TPS (tokens/秒), 86.9               │
│  ...                                 │
│  测试结论, 通过                       │
├──────────────────────────────────────┤
│  (空行)                              │
├──────────────────────────────────────┤
│  编号, 用例名称, 测试类型, ..., 是否通过│  ← 用例表（7列）
│  TC-01, 接口连通性, 功能, ..., 通过   │
│  ...                                 │
└──────────────────────────────────────┘
```

### 汇总文件 `test_summary.csv`

所有运行汇集到一个文件，每次追加一行，便于横向对比：

| 模型名 | 测试时间 | 并发数 | 总请求数 | TPS | TPM | 成功率 | Token总数 | 平均延迟 | TC-01~06 | 总通过 | 总失败 | 测试结论 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| gpt-4o | 2026-07-09 18:00 | 4 | 20 | 86.9 | 5214.9 | 83.3% | 1532 | 3.2s | ... | 5 | 1 | 1/6 未通过 |

## 项目结构

```
quick-test/
├── README.md              ← 本文件
├── token_test.py           ← 主测试脚本
├── test_config.json        ← 默认配置（URL/Key/模型/参数）
├── test_cases.json         ← 测试用例持久化定义
├── template.xlsx           ← 参考模板
├── results/
│   └── test_output_*.csv   ← 各次运行详情（自动生成，按时间戳命名）
└── test_summary.csv        ← 累积汇总（自动生成，每次追加一行）
```

## 可复现性

- 用例定义在 `test_cases.json`，参数固定，可纳入 Git
- 配置通过 `test_config.json` 或命令行传入，相同输入 → 相同流程
- 每次运行生成独立时间戳文件，历史结果不丢失
- 汇总文件积累所有运行记录，支持跨版本对比

## 自定义用例

编辑 `test_cases.json`，按以下结构添加用例：

```json
{
  "id": "TC-07",
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

可用的 `method` 值：`connectivity`、`tps_benchmark`、`tpm_calc`、`context_limit`、`auth_failure`、`rate_limit`。

## License

MIT
