# SOP & 二单元 Report Generator

快速数据处理工具，支持两个命令：

- `/sop <文件路径>` — 处理 SOP 数据（SS首通、SS首单元，排除 CC）
- `/二单元 <文件路径>` — 处理二单元数据

## 功能

- ✅ 自动按报告日期、生成日期（3日窗口）、销售名称组织文件夹
- ✅ 计算任务年龄（`age_hours`），标注超过24小时为 `YELLOW` 警戒
- ✅ 支持自定义报告日期（默认为今天）
- ✅ 自动检测 CSV 编码
- ✅ 生成详细处理报告

## 快速使用

```python
from sop_reporter import handle_sop_command, handle_erdan_command

# 处理 SOP 数据（使用默认报告日期 = 今天）
print(handle_sop_command('/path/to/file.csv'))

# 处理 SOP 数据（指定报告日期）
print(handle_sop_command('/path/to/file.csv', '2025-06-21'))

# 处理二单元数据
print(handle_erdan_command('/path/to/file.csv', '2025-06-21'))
```

## 输出目录结构

```
exports/
├── organized_sop/
│   └── SOP/
│       └── 2025-06-21/
│           ├── 销售1.csv
│           ├── 销售2.csv
│           └── ...
└── organized_erdan/
    └── 二单元/
        └── 2025-06-21/
            ├── 销售1.csv
            ├── 销售2.csv
            └── ...
```

## CSV 输出列

原始列 + 以下三列：
- `age_hours`: 任务距生成时间的小时数
- `alert`: 超过 24 小时时标记为 `YELLOW`
- `gen_date`: 生成日期（YYYY-MM-DD）

## 处理规则

### SOP 模式 (`/sop`)
- 仅包含：`sop类型` 为 `SS首通` 或 `SS首单元`
- 排除：包含 `CC` 的类型
- 状态：必须为"未完成"
- 日期范围：报告日期及前两日（3日窗口）

### 二单元模式 (`/二单元`)
- 仅包含：`sop类型` 包含 `二单元`
- 状态：必须为"未完成"
- 日期范围：报告日期及前两日（3日窗口）

## API 参考

### `DataProcessor` 类

#### `__init__(file_path, report_date=None)`
初始化处理器
- `file_path` (str): CSV 文件路径
- `report_date` (str, optional): 报告日期，格式 'YYYY-MM-DD'，默认为今天

#### `process(mode) -> (int, int, Path)`
执行处理
- `mode` (str): 'sop' 或 'erdan'
- 返回: (总行数, 匹配行数, 输出目录)

#### `report(mode) -> str`
生成处理报告
- `mode` (str): 'sop' 或 'erdan'
- 返回: 格式化的报告字符串

### 命令处理函数

#### `handle_sop_command(file_path, report_date=None) -> str`
处理 `/sop` 命令，返回报告文本

#### `handle_erdan_command(file_path, report_date=None) -> str`
处理 `/二单元` 命令，返回报告文本

## 常见问题

**Q: 为什么没有生成任何文件？**
A: 检查以下几点：
1. 文件路径是否正确（建议用绝对路径）
2. `sop生成日期` 的数据是否在报告日期前三日内
3. `sop类型` 和 `sop状态` 的值是否匹配规则
4. CSV 编码是否为常见格式（脚本会自动尝试检测）

**Q: 如何使用历史数据？**
A: 指定对应的 `report_date`，例如数据是 2025-06-21 的，就传入 `report_date='2025-06-21'`。

---

**版本**: 1.0.0  
**最后更新**: 2025-12-07
