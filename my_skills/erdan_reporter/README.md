# 二单元 Report Generator

处理二单元 Excel 数据，筛选 7 天内 120s 通话数为空/未完成的记录，按 SS 小组和销售分类。

## 功能

- ✅ 读取 Excel 文件（`.xlsx`）
- ✅ 时间范围：过去 7 天（包含当天）
- ✅ 筛选条件：`120s通话数` 为空、`-` 或 `0`（未完成）
- ✅ 按 `SS小组` 和 `SS`（销售名称）分文件夹
- ✅ 标注警戒级别：RED（≥3天未完成）、YELLOW（1-2天未完成）
- ✅ 生成详细处理报告

## 快速使用

```python
from erdan_reporter import handle_erdan_command

# 使用默认输出路径和报告日期（今天）
result = handle_erdan_command('/path/to/file.xlsx')

# 指定报告日期
result = handle_erdan_command('/path/to/file.xlsx', report_date='2025-12-07')

# 指定自定义输出路径
result = handle_erdan_command('/path/to/file.xlsx', 
                             output_base='/your/custom/path/exports/organized')

print(result)
```

## 输出目录结构

```
exports/organized/
└── 二单元/
    └── 2025-12-07/
        ├── SH-SS01小组/
        │   ├── 销售1.csv
        │   ├── 销售2.csv
        │   └── ...
        ├── SH-SS09小组/
        │   └── ...
        └── ...
```

## CSV 输出列

原始 Excel 列 + 以下三列：
- `age_days`: 距报告日期的天数
- `alert`: 警戒级别（RED ≥3天 / YELLOW 1-2天 / 空=0天）
- `complete_date`: 完课时间（YYYY-MM-DD）

## 处理规则

- **时间范围**：报告日期及前 6 天（共 7 天）
- **筛选条件**：`120s通话数` 为空、`-` 或 `0`（表示未完成）
- **分组方式**：
  - 一级：`SS小组`（如 `SH-SS01小组`）
  - 二级：`SS`（销售名称）
- **警戒标记**：
  - RED：距完课时间 ≥3 天（严重延期）
  - YELLOW：距完课时间 1-2 天（需要注意）
  - （空）：距完课时间 0 天（最近）

## 输入数据要求

Excel 文件必须包含以下列：
- `大区`（可选）
- `SS小组` - 小组名称
- `SS` - 销售名称/学员代码
- `学员ID`（可选）
- `二单元第八节课完课时间` - 完课时间
- `20s通话数`（可选，处理器会忽略）
- `120s通话数` - 关键列，用于判断是否完成

## API 参考

### `ErDanProcessor` 类

#### `__init__(file_path, report_date=None, output_base=None)`
初始化处理器
- `file_path` (str): Excel 文件路径
- `report_date` (str, optional): 报告日期，格式 'YYYY-MM-DD'，默认为今天
- `output_base` (str, optional): 输出基础路径，默认为 `my_skills/exports/organized`

#### `process() -> (int, int, Path)`
执行处理
- 返回: (总行数, 匹配行数, 输出目录)

#### `report() -> str`
生成处理报告
- 返回: 格式化的报告字符串

### 命令处理函数

#### `handle_erdan_command(file_path, report_date=None, output_base=None) -> str`
处理 `/二单元` 命令，返回报告文本

## 示例输出

```
📊 二单元数据处理报告
==================================================
报告日期: 2025-12-07
时间范围: 2025-12-01 ~ 2025-12-07（7天）
筛选条件: 120s通话数 为空/'-'/0（未完成）

处理结果:
  • 总行数: 164
  • 匹配行数（7天未完成）: 89
  • 生成 CSV 文件数: 33
  • 输出目录: /Users/mac/Desktop/.../exports/organized/二单元

📁 按SS小组分类（2025-12-07）:
  • 【SH-SS01小组】- 7 个销售
  • 【SH-SS09小组】- 7 个销售
  • 【SH-SS11小组】- 7 个销售
  • 【SH-SS13小组】- 7 个销售
  • 【SH-SS15小组】- 5 个销售
```

---

**版本**: 1.0.0  
**最后更新**: 2025-12-07
