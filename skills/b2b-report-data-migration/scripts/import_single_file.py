#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导入 电商B端报表25年9-26年5.xlsx 到 MySQL 数据库
使用与 import_all_reports.py 相同的表结构，但独立处理这个一次性大量数据文件
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

SOURCE_FILE = r'D:\数据专用\基础全量数据\电商B端报表25年9-26年5.xlsx'

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


def main():
    print('=' * 60)
    print('导入 电商B端报表25年9-26年5.xlsx')
    print('=' * 60)

    if not os.path.exists(SOURCE_FILE):
        print(f'文件不存在: {SOURCE_FILE}')
        return

    print(f'\n[1] 读取Excel文件...')
    df = pd.read_excel(SOURCE_FILE, dtype=str, engine='openpyxl')
    print(f'原始行数: {len(df)}')

    if len(df) > 0 and str(df.iloc[-1, 0]).strip() == '合计':
        df = df.iloc[:-1]
        print('已删除合计行')

    df = df.rename(columns=COLUMN_MAPPING)
    df = df.replace({np.nan: None, 'nan': None, 'NaN': None, '': None})
    print(f'处理后行数: {len(df)}')

    conn = connect_with_retry()
    cursor = conn.cursor()

    try:
        print('\n[2] 确保表结构存在...')
        cursor.execute("""
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
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='电商B端报表'
        """)

        cursor.execute("SELECT COUNT(*) FROM B2B_Report")
        existing = cursor.fetchone()[0]
        print(f'现有数据行数: {existing}')

        print('\n[3] 禁用索引...')
        cursor.execute("SET unique_checks=0")
        cursor.execute("SET foreign_key_checks=0")

        print('\n[4] 导入数据...')
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

            if (i + 1) % 5000 == 0:
                conn.commit()
                print(f'  已处理 {i + 1} 行...')

        conn.commit()
        print(f'\n导入统计: 成功 {success_count} 行, 失败 {error_count} 行')

        print('\n[5] 重新启用索引...')
        cursor.execute("SET unique_checks=1")
        cursor.execute("SET foreign_key_checks=1")

        print('\n[6] 验证结果...')
        cursor.execute("SELECT COUNT(*) FROM B2B_Report")
        final_count = cursor.fetchone()[0]
        print(f'导入前行数: {existing}')
        print(f'本次导入: {success_count} 行')
        print(f'导入后总行数: {final_count}')

        print('\n' + '=' * 60)
        print('导入完成!')
        print('=' * 60)

    finally:
        cursor.close()
        conn.close()


if __name__ == '__main__':
    main()