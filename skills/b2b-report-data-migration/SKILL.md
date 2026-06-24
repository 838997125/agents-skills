---
name: b2b-report-data-migration
description: |
  将 `D:\数据专用\基础全量数据` 目录下的电商B端报表Excel文件增量导入 MySQL 数据库。适用于：定时任务自动导入新文件、手动触发导入未处理的Excel文件、或者当定时任务漏跑时补导入数据。

  When to trigger: 用户提到"电商B端报表"、"B2B报表"、"import_all_reports"、"数据导入"、"增量导入"、"定时任务漏跑"、"补充导入"、"导入未处理的Excel"等场景时使用此技能。
  Also triggers when user mentions database `YMSS_DING` table `B2B_Report` or Excel files in `D:\数据专用\基础全量数据` that need to be imported.
---

# B2B 报表数据迁移技能

将 `D:\数据专用\基础全量数据` 目录下的电商B端报表 Excel 文件增量导入 MySQL 数据库 `YMSS_DING.B2B_Report` 表。

## 数据库连接

| 参数 | 值 |
|------|-----|
| Host | 112.124.40.48 |
| Port | 3306 |
| User | root |
| Password | 881206YanYan |
| Database | YMSS_DING |
| local_infile | true |
| charset | utf8mb4 |

## 源文件类型

### 类型1：周期性文件（增量导入）

文件命名格式：`电商B端报表-YYYY-MM-DD.xlsx`（如 `电商B端报表-2026-05-16.xlsx`）

使用脚本：`scripts/import_all_reports.py`

处理逻辑：
- 使用正则 `^电商B端报表-\d{4}-\d{2}-\d{2}\.xlsx$` 匹配文件
- 通过 `imported_files` 表跟踪已处理文件，仅处理新文件
- 文件处理成功后才标记为已处理
- 适合定时任务场景

### 类型2：一次性大批量文件

文件：`电商B端报表25年9-26年5.xlsx`（或其他不带日期格式的汇总文件）

使用脚本：`scripts/import_single_file.py`

处理逻辑：直接导入，不做文件名模式匹配

## 表结构 (B2B_Report)

```sql
CREATE TABLE IF NOT EXISTS B2B_Report (
    id INT AUTO_INCREMENT PRIMARY KEY COMMENT '自增ID',
    documentNo VARCHAR(32) COMMENT '单据编号',
    `date` DATE COMMENT '日期',
    `time` TIME COMMENT '时间',
    storeNo VARCHAR(20) COMMENT '门店编号',
    storeName VARCHAR(128) COMMENT '门店名称',
    productNo VARCHAR(32) COMMENT '商品编号',
    productName VARCHAR(128) COMMENT '商品名称',
    commonName VARCHAR(128) COMMENT '通用名',
    specs VARCHAR(64) COMMENT '商品规格',
    manufacturer VARCHAR(128) COMMENT '生产厂家',
    origin VARCHAR(128) COMMENT '商品产地',
    unit VARCHAR(16) COMMENT '单位',
    salesperson VARCHAR(32) COMMENT '销售员',
    quantity DECIMAL(18,4) COMMENT '数量',
    taxRate DECIMAL(5,2) COMMENT '税率',
    taxInclusiveAmount DECIMAL(18,4) COMMENT '含税金额',
    amount DECIMAL(18,4) COMMENT '金额',
    costAmount DECIMAL(18,4) COMMENT '成本金额',
    profit DECIMAL(18,4) COMMENT '毛利',
    taxInclusiveCost DECIMAL(18,4) COMMENT '含税成本金额',
    taxInclusiveProfit DECIMAL(18,4) COMMENT '含税毛利',
    batchNo VARCHAR(32) COMMENT '批号',
    expiryDate DATE COMMENT '有效期至',
    warehouseName VARCHAR(64) COMMENT '库房名称',
    ecommerceOrderNo VARCHAR(64) COMMENT '电商单号',
    platformName VARCHAR(32) COMMENT '平台名称',
    category VARCHAR(64) COMMENT '一级分类',
    prescriptionType VARCHAR(32) COMMENT '处方分类',
    trackingNo VARCHAR(256) COMMENT '快递单号',
    cEndPurchase VARCHAR(8) COMMENT 'c端采购'
) ENGINE=InnoDB DEFAULT CHARSET=utfmb4 COMMENT='电商B端报表';
```

**注意**：`id` 为自增主键，`documentNo` 可以重复（不作为唯一键）。

## 字段映射（中文字段名 → 英文字段名）

| 中文列名 | 英文字段名 |
|---------|-----------|
| 单据编号 | documentNo |
| 日期 | date |
| 时间 | time |
| 门店编号 | storeNo |
| 门店名称 | storeName |
| 商品编号 | productNo |
| 商品名称 | productName |
| 通用名 | commonName |
| 商品规格 | specs |
| 生产厂家 | manufacturer |
| 商品产地 | origin |
| 单位 | unit |
| 销售员 | salesperson |
| 数量 | quantity |
| 税率 | taxRate |
| 含税金额 | taxInclusiveAmount |
| 金额 | amount |
| 成本金额 | costAmount |
| 毛利 | profit |
| 含税成本金额 | taxInclusiveCost |
| 含税毛利 | taxInclusiveProfit |
| 批号 | batchNo |
| 有效期至 | expiryDate |
| 库房名称 | warehouseName |
| 电商单号 | ecommerceOrderNo |
| 平台名称 | platformName |
| 一级分类 | category |
| 处方分类 | prescriptionType |
| 快递单号 | trackingNo |
| c端采购 | cEndPurchase |

## 数据清洗规则

1. **删除合计行**：如果最后一行的第一列值是"合计"，删除该行
2. **空值处理**：`NaN`、`nan`、空白字符串、`'NULL'` → Python `None`（对应 SQL `NULL`）
3. **有效期(expiryDate)解析**：
   - 完整格式 `YYYY-MM-DD`：直接使用
   - 只有年月 `YYYY-MM`：补齐为 `YYYY-MM-01`
   - 其他格式：截取前10字符
   - 无效值：→ `None`

## 导入执行

### 增量导入（新文件，定时任务场景）

```bash
python3 scripts/import_all_reports.py
```

**流程**：
1. 查找 `电商B端报表-YYYY-MM-DD.xlsx` 格式且未在 `imported_files` 表中的文件
2. 确保 `B2B_Report` 表存在
3. 禁用索引（`unique_checks=0`, `foreign_key_checks=0`）
4. 逐行 INSERT，每 5000 行提交一次
5. 成功后标记文件为已处理
6. 重新启用索引
7. 输出验证信息

### 一次性大批量导入

```bash
python3 scripts/import_single_file.py
```

指定其他文件路径时，修改脚本中 `SOURCE_FILE` 变量。

## 验证步骤

导入完成后执行以下验证：

```sql
-- 1. 总行数
SELECT COUNT(*) FROM B2B_Report;

-- 2. 单据编号重复检查
SELECT documentNo, COUNT(*) as cnt 
FROM B2B_Report 
GROUP BY documentNo 
HAVING cnt > 1 
LIMIT 10;

-- 3. expiryDate 修复检查（当月1日表示原日期不完整）
SELECT COUNT(*) 
FROM B2B_Report 
WHERE DAY(expiryDate) = 1 AND LENGTH(expiryDate) = 10;

-- 4. 数据样本
SELECT id, documentNo, date, expiryDate, storeName, productName 
FROM B2B_Report 
LIMIT 5;
```

## 跟踪表 (imported_files)

增量导入使用 `imported_files` 表避免重复导入：

```sql
CREATE TABLE IF NOT EXISTS imported_files (
    id INT AUTO_INCREMENT PRIMARY KEY,
    filename VARCHAR(255) NOT NULL UNIQUE,
    imported_at DATETIME NOT NULL
);
```

## 常见错误处理

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| `secure_file_priv` 拒绝 | MySQL 安全限制 | 使用 INSERT 而非 `LOAD DATA INFILE`（脚本已默认使用 INSERT） |
| `not all arguments converted` | SQL 占位符与值数量不匹配 | 检查 INSERT 语句中 `%s` 数量是否等于字段数（应为30个） |
| 重复数据 | 文件被重复处理 | 确认 `imported_files` 表中已记录，或先手动清理 |

## 定时任务配置

建议通过操作系统的定时任务（Windows Task Scheduler / cron）定期执行：

```bash
# Windows Task Scheduler 示例
python.exe D:\数据专用\基础全量数据\import_all_reports.py

# cron 示例（每天早上9点执行）
0 9 * * * /c/Python314/python.exe /d/数据专用/基础全量数据/import_all_reports.py
```

**注意**：定时任务漏跑时，直接手动触发 `python3 scripts/import_all_reports.py` 即可补导入未处理的Excel文件。