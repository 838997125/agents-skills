#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量Excel转MySQL导入脚本
处理目录下所有"电商B端报表"Excel文件，导入到B2B_Report表
"""

import pandas as pd
import numpy as np
import pymysql
import os
import time
import re

DB_CONFIG = {
    'host': '112.124.40.48',
    'port': 3306,
    'user': 'root',
    'password': '881206YanYan',
    'database': 'YMSS_DING',
    'local_infile': True,
    'charset': 'utf8mb4'
}

SOURCE_DIR = r'D:\数据专用\基础全量数据'

COLUMN_MAPPING = {
    '单据编号': 'documentNo',
    '日期': 'date',
    '时间': 'time',
    '门店编号': 'storeNo',
    '门店名称': 'storeName',
    '商品编号': 'productNo',
    '商品名称': 'productName',
    '通用名': 'commonName',
    '商品规格': 'specs',
    '生产厂家': 'manufacturer',
    '商品产地': 'origin',
    '单位': 'unit',
    '销售员': 'salesperson',
    '数量': 'quantity',
    '税率': 'taxRate',
    '含税金额': 'taxInclusiveAmount',
    '金额': 'amount',
    '成本金额': 'costAmount',
    '毛利': 'profit',
    '含税成本金额': 'taxInclusiveCost',
    '含税毛利': 'taxInclusiveProfit',
    '批号': 'batchNo',
    '有效期至': 'expiryDate',
    '库房名称': 'warehouseName',
    '电商单号': 'ecommerceOrderNo',
    '平台名称': 'platformName',
    '一级分类': 'category',
    '处方分类': 'prescriptionType',
    '快递单号': 'trackingNo',
    'c端采购': 'cEndPurchase'
}

CREATE_TABLE_SQL = """
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='电商B端报表';
"""

INSERT_SQL = """
INSERT INTO B2B_Report (
    documentNo, `date`, `time`, storeNo, storeName, productNo, productName,
    commonName, specs, manufacturer, origin, unit, salesperson, quantity,
    taxRate, taxInclusiveAmount, amount, costAmount, profit, taxInclusiveCost,
    taxInclusiveProfit, batchNo, expiryDate, warehouseName, ecommerceOrderNo,
    platformName, category, prescriptionType, trackingNo, cEndPurchase
) VALUES (
    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
)
"""


def find_new_excel_files():
    """查找目录下未处理的新Excel文件（电商B端报表-YYYY-MM-DD格式）"""
    pattern = re.compile(r'^电商B端报表-\d{4}-\d{2}-\d{2}\.xlsx$')

    processed_files = set()
    try:
        conn = pymysql.connect(**DB_CONFIG)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS imported_files (
                id INT AUTO_INCREMENT PRIMARY KEY,
                filename VARCHAR(255) NOT NULL UNIQUE,
                imported_at DATETIME NOT NULL
            )
        """)
        cursor.execute("SELECT filename FROM imported_files")
        processed_files = set(row[0] for row in cursor.fetchall())
        cursor.close()
        conn.close()
    except Exception as e:
        print(f'  获取已处理文件列表失败: {e}')

    new_files = []
    for f in os.listdir(SOURCE_DIR):
        if pattern.match(f) and f not in processed_files:
            new_files.append(os.path.join(SOURCE_DIR, f))

    return sorted(new_files)


def is_file_processed(cursor, filename):
    """检查文件是否已处理"""
    cursor.execute("SELECT 1 FROM imported_files WHERE filename = %s", (filename,))
    return cursor.fetchone() is not None


def mark_file_processed(cursor, filename):
    """标记文件为已处理"""
    cursor.execute("""
        INSERT INTO imported_files (filename, imported_at)
        VALUES (%s, NOW())
    """, (filename,))


def connect_with_retry(max_retries=3, retry_interval=5):
    for attempt in range(1, max_retries + 1):
        try:
            conn = pymysql.connect(**DB_CONFIG)
            print(f'数据库连接成功 (尝试 {attempt}/{max_retries})')
            return conn
        except pymysql.Error as e:
            print(f'数据库连接失败 (尝试 {attempt}/{max_retries}): {e}')
            if attempt < max_retries:
                print(f'等待 {retry_interval} 秒后重试...')
                time.sleep(retry_interval)
            else:
                raise


def clean_value(val):
    if pd.isna(val) or val is None or str(val).strip() == '' or str(val).strip().upper() == 'NULL':
        return None
    return str(val).strip()


def parse_expiry_date(val):
    if val is None or str(val).strip() == '' or str(val).strip().upper() == 'NULL':
        return None

    val = str(val).strip()

    if len(val) == 10 and val[4] == '-' and val[7] == '-':
        return val

    if len(val) == 7 and val[4] == '-':
        return f'{val}-01'

    if len(val) >= 10:
        return val[:10]

    return None


def process_excel_to_df(excel_path):
    print(f'正在读取: {excel_path}')

    df = pd.read_excel(excel_path, dtype=str, engine='openpyxl')

    if len(df) > 0 and str(df.iloc[-1, 0]).strip() == '合计':
        df = df.iloc[:-1]
        print(f'  已删除合计行')

    df = df.rename(columns=COLUMN_MAPPING)
    df = df.replace({np.nan: None, 'nan': None, 'NaN': None, '': None})

    print(f'  数据行数: {len(df)}')
    return df


def import_file(cursor, excel_path, conn):
    """导入单个文件"""
    filename = os.path.basename(excel_path)

    df = process_excel_to_df(excel_path)

    success_count = 0
    error_count = 0

    for i, row in df.iterrows():
        try:
            expiry_date = parse_expiry_date(row.get('expiryDate'))

            values = (
                clean_value(row.get('documentNo')),
                clean_value(row.get('date')),
                clean_value(row.get('time')),
                clean_value(row.get('storeNo')),
                clean_value(row.get('storeName')),
                clean_value(row.get('productNo')),
                clean_value(row.get('productName')),
                clean_value(row.get('commonName')),
                clean_value(row.get('specs')),
                clean_value(row.get('manufacturer')),
                clean_value(row.get('origin')),
                clean_value(row.get('unit')),
                clean_value(row.get('salesperson')),
                clean_value(row.get('quantity')),
                clean_value(row.get('taxRate')),
                clean_value(row.get('taxInclusiveAmount')),
                clean_value(row.get('amount')),
                clean_value(row.get('costAmount')),
                clean_value(row.get('profit')),
                clean_value(row.get('taxInclusiveCost')),
                clean_value(row.get('taxInclusiveProfit')),
                clean_value(row.get('batchNo')),
                expiry_date,
                clean_value(row.get('warehouseName')),
                clean_value(row.get('ecommerceOrderNo')),
                clean_value(row.get('platformName')),
                clean_value(row.get('category')),
                clean_value(row.get('prescriptionType')),
                clean_value(row.get('trackingNo')),
                clean_value(row.get('cEndPurchase')),
            )
            cursor.execute(INSERT_SQL, values)
            success_count += 1
        except Exception as e:
            error_count += 1
            if error_count <= 5:
                print(f'  错误 Row {i}: {str(e)}')

    conn.commit()
    print(f'  成功: {success_count} 行, 失败: {error_count} 行')
    return success_count, error_count


def main():
    print('=' * 60)
    print('批量Excel转MySQL导入程序（增量模式）')
    print('=' * 60)

    new_files = find_new_excel_files()
    print(f'找到 {len(new_files)} 个未处理的新文件:')
    for f in new_files:
        print(f'  - {os.path.basename(f)}')

    if not new_files:
        print('没有新文件需要处理')
        return

    conn = connect_with_retry()
    cursor = conn.cursor()

    try:
        print('\n[1] 确保表结构存在...')
        cursor.execute(CREATE_TABLE_SQL)
        print('表 B2B_Report 检查/创建完成')

        print('\n[2] 确保跟踪表存在...')
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS imported_files (
                id INT AUTO_INCREMENT PRIMARY KEY,
                filename VARCHAR(255) NOT NULL UNIQUE,
                imported_at DATETIME NOT NULL
            )
        """)
        print('跟踪表检查完成')

        cursor.execute("SELECT COUNT(*) FROM B2B_Report")
        existing_count = cursor.fetchone()[0]
        print(f'现有数据行数: {existing_count}')

        print('\n[3] 禁用索引...')
        cursor.execute("SET unique_checks=0")
        cursor.execute("SET foreign_key_checks=0")

        total_success = 0
        total_error = 0

        for excel_path in new_files:
            filename = os.path.basename(excel_path)

            # 再次检查是否已处理（防止并发情况）
            if is_file_processed(cursor, filename):
                print(f'  文件 {filename} 已处理，跳过')
                continue

            print(f'\n--- 处理文件: {filename} ---')
            try:
                s, e = import_file(cursor, excel_path, conn)
                total_success += s
                total_error += e
                # 数据导入成功后标记文件为已处理
                mark_file_processed(cursor, filename)
                conn.commit()
            except Exception as ex:
                print(f'  文件处理失败: {ex}')
                continue

        print('\n[4] 重新启用索引...')
        cursor.execute("SET unique_checks=1")
        cursor.execute("SET foreign_key_checks=1")

        print('\n[5] 验证导入结果...')
        cursor.execute("SELECT COUNT(*) FROM B2B_Report")
        db_count = cursor.fetchone()[0]
        print(f'总导入行数: {db_count}')
        print(f'本次新增: {total_success} 行')

        print('\n[6] 单据编号重复检查:')
        cursor.execute('SELECT documentNo, COUNT(*) as cnt FROM B2B_Report GROUP BY documentNo HAVING cnt > 1 LIMIT 10')
        dups = cursor.fetchall()
        if dups:
            print(f'  存在重复单据编号（显示前10个）:')
            for dup in dups:
                print(f'    documentNo={dup[0]}, 出现次数={dup[1]}')

        print('\n[7] expiryDate修复检查:')
        cursor.execute("SELECT COUNT(*) FROM B2B_Report WHERE DAY(expiryDate) = 1 AND LENGTH(expiryDate) = 10")
        fixed_count = cursor.fetchone()[0]
        print(f'  expiryDate为当月1日的记录数: {fixed_count}')

        print('\n' + '=' * 60)
        print(f'导入完成! 总成功: {total_success} 行, 总失败: {total_error} 行')
        print('=' * 60)

    finally:
        cursor.close()
        conn.close()


if __name__ == '__main__':
    main()