---
name: retail-data-migration
description: 零售汇总分析数据迁移技能。将 Excel 文件（零售汇总分析-YYYY-MM-DD格式）增量导入 MySQL 数据库。适用于定时任务自动执行或手动触发未导入数据的导入操作。
---

# 零售汇总分析数据迁移

将 `D:\数据专用\基础全量数据` 目录下的**日期格式** Excel 文件（`零售汇总分析-YYYY-MM-DD.xlsx`）增量导入到 `YMSS_DING.ERP_Retail_Summary` 表。

## 核心逻辑

脚本 `D:\数据专用\基础全量数据\零售汇总分析\import_retail_summary.py` 实现了完整的数据迁移流程：

```
Excel → CSV → MySQL
```

### 增量导入机制

- 通过 `sourceFile` 字段判断文件是否已导入（`get_imported_files`）
- 每次运行只处理**未导入的新文件**，不会重复导入
- 支持**定时任务**（Cron）自动执行，也支持**手动触发**

## 数据库配置

- 主机：`112.124.40.48`
- 端口：`3306`
- 用户名：`root`
- 密码：`881206YanYan`
- 数据库：`YMSS_DING`
- 表名：`ERP_Retail_Summary`

## 源文件目录

| 目录 | 用途 | 文件名格式 |
|------|------|-----------|
| `D:\数据专用\基础全量数据` | 每日增量文件（定时任务扫描） | `零售汇总分析-YYYY-MM-DD.xlsx` |
| `D:\数据专用\基础全量数据\零售汇总分析` | 历史归档文件 | `零售汇总分析YY年M月-M月.xlsx` |

本技能**只处理** `D:\数据专用\基础全量数据` 目录下的日期格式文件。

## 执行方式

### 方式一：定时任务自动执行

```bash
python "D:\数据专用\基础全量数据\零售汇总分析\import_retail_summary.py"
```

推荐 Cron 表达式（每日早8点执行）：

```
0 8 * * *
```

### 方式二：手动触发

当定时任务未执行或需要立即导入新数据时，对 Claude 说：

> "执行零售数据迁移" 或 "运行 import_retail_summary.py"

## 数据处理规则

### Excel 表头处理

Excel 文件使用**合并单元格表头**：

- **Row 1**：分类行（如"含税"/"除税"）
- **Row 2**：具体字段名（如"实收金额"/"成本金额"）
- **合并后表头**：`{Row2字段名}（{Row1分类}）`，例：`实收金额（含税）`

Row 1 中无合并的列，直接使用 Row 1 值作为字段名。

### 数据清洗

- 跳过最后一行"合计"数据行
- 空单元格 → `NULL`
- `NaN` 值 → `NULL`
- 日期格式 → `YYYY-MM-DD`
- 时间格式 → `HH:MM:SS`

### 字段映射（39列）

| 英文字段名 | 中文注释 | 类型 |
|-----------|---------|------|
| documentNo | 单据编号 | VARCHAR(50) |
| date | 日期 | DATE |
| time | 时间 | TIME |
| storeNo | 门店编号 | VARCHAR(20) |
| storeName | 门店名称 | VARCHAR(100) |
| productNo | 商品编号 | VARCHAR(50) |
| productName | 商品名称 | VARCHAR(200) |
| genericName | 通用名 | VARCHAR(100) |
| spec | 商品规格 | VARCHAR(100) |
| manufacturer | 生产厂家 | VARCHAR(200) |
| origin | 商品产地 | VARCHAR(100) |
| unit | 单位 | VARCHAR(20) |
| salesperson | 销售员 | VARCHAR(50) |
| cashier | 收款员 | VARCHAR(50) |
| quantity | 数量 | DECIMAL(18,4) |
| receivableAmount | 应收金额 | DECIMAL(18,4) |
| receivedAmountWithTax | 含税-实收金额 | DECIMAL(18,4) |
| costAmountWithTax | 含税-成本金额 | DECIMAL(18,4) |
| grossProfitWithTax | 含税-毛利 | DECIMAL(18,4) |
| revenueExTax | 除税-收入 | DECIMAL(18,4) |
| costAmountExTax | 除税-成本金额 | DECIMAL(18,4) |
| grossProfitExTax | 除税-毛利 | DECIMAL(18,4) |
| grossMarginRate | 毛利率 | DECIMAL(10,4) |
| taxRate | 税率 | DECIMAL(10,4) |
| batchNo | 批号 | VARCHAR(50) |
| expiryDate | 有效期至 | VARCHAR(20) |
| warehouseName | 库房名称 | VARCHAR(100) |
| categoryLv1 | 一级分类 | VARCHAR(100) |
| prescriptionCategory | 处方分类 | VARCHAR(50) |
| ecommercePrice | 电商价 | DECIMAL(18,4) |
| jushuitanOrderNo | 聚水潭内部单号 | VARCHAR(100) |
| shopName | 电商店铺名称 | VARCHAR(200) |
| platformOrderNo | 平台订单号 | VARCHAR(500) |
| salesCategory | 销售类别 | VARCHAR(50) |
| cEndCategory | c端分类 | VARCHAR(100) |
| cEndProcurement | c端采购 | VARCHAR(100) |
| cEndFiling | c端备案 | VARCHAR(100) |
| trackingNo | 快递单号 | VARCHAR(100) |

附加字段：

| 字段名 | 说明 |
|--------|------|
| sourceFile | 源文件名（用于增量判断） |
| importTime | 导入时间（自动记录） |

### 日志输出

运行日志保存在：`D:\数据专用\基础全量数据\零售汇总分析\import_log.txt`

日志内容包括：

- 每步操作的 timestamp
- 文件扫描结果
- Excel → CSV 转换行数
- MySQL 插入成功/失败数
- 最终统计（Excel总数、成功插入数、错误数、表记录总数）

## 异常处理

| 异常情况 | 处理方式 |
|---------|---------|
| 数据库连接失败 | 重试3次，间隔5秒 |
| 批量插入错误 | 跳过并记录日志 |
| 文件未找到 | 跳过并输出警告 |
| 字段类型错误 | 该行跳过，错误计数+1 |

## 验证方法

运行完成后检查：

1. 日志中 `成功插入` 数是否与预期行数一致
2. `MySQL 表记录数` 是否为历史记录 + 新导入记录之和
3. 抽样验证字段映射是否正确
