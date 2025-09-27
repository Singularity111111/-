import pandas as pd
import numpy as np
from datetime import datetime

# ==============================================================================
#                           *** 日期设置区域 ***
#
#  - 如果要计算特定日期，请在这里修改日期字符串，例如 '2025-08-25'。
#  - 如果想让脚本自动计算【最新】的日期，请将值设置为 None，像这样：target_date = None
#
target_date_str = '2025-08-25'  # <--- 请在这里修改您想要的日期
#
# ==============================================================================


def generate_product_daily_report(target_date=None):
    """
    根据分解的逻辑，专门生成【产品日报】。
    V3: 增加了指定日期的功能。
    """
    try:
        # --- 0. 文件路径定义 ---
        f_book1 = 'Book1.xlsx'
        f_book2 = 'Book2.xlsx'
        f_book4 = 'Book4.xlsx'
        f_pakistan_cost = '巴基斯坦消耗.xlsx'

        # --- 1. 加载数据 ---
        print("开始加载 Excel 数据文件...")
        df_book1 = pd.read_excel(f_book1)
        df_book2 = pd.read_excel(f_book2)
        df_book4 = pd.read_excel(f_book4)
        df_cost = pd.read_excel(f_pakistan_cost)
        print("数据加载完毕。")

        # --- 2. 统一并清洗日期格式 ---
        print("正在统一并清洗日期格式...")
        for df in [df_book1, df_book2, df_book4, df_cost]:
            if '日期' in df.columns:
                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                df.dropna(subset=['日期'], inplace=True)
        print("日期格式处理完毕。")

        # --- 3. 确定并筛选目标日期的数据 ---
        if df_book2.empty:
            print("错误：核心文件 (Book2.xlsx) 中没有有效的日期数据。")
            return
            
        if target_date:
            report_date = pd.to_datetime(target_date)
            print(f"已指定日期为: {report_date.strftime('%Y-%m-%d')}，将为此日期生成日报...")
        else:
            report_date = df_book2['日期'].max()
            print(f"未指定日期，将自动检测最新日期: {report_date.strftime('%Y-%m-%d')}...")

        report_date_str = report_date.strftime('%Y-%m-%d')
        
        # 筛选各表当天的数据
        report_data = df_book2[df_book2['日期'] == report_date].copy()
        if report_data.empty:
            print(f"错误：Book2.xlsx 中找不到日期为 {report_date_str} 的数据。")
            return
            
        df_book1_today = df_book1[df_book1['日期'] == report_date]
        df_book4_today = df_book4[df_book4['日期'] == report_date]
        df_cost_today = df_cost[df_cost['日期'] == report_date]
        
        # --- 4. 提取和计算各个字段 ---
        print("正在计算各项指标...")
        
        final_report = pd.DataFrame()
        final_report['日期'] = [report_date_str]
        final_report['产品'] = 'aa'
        
        total_cost = 0
        if not df_cost_today.empty and '和' in df_cost_today.columns:
            total_cost = df_cost_today['和'].iloc[0]
        final_report['消耗'] = total_cost

        base_metrics = [
            '新增用户数', '充值人数', '充值金额', '提现金额', '首存人数', '首存充值金额', 
            '首存付费率(%)', '首存盈余率(%)', '首存ARPPU', '新增付费人数', '新增充值金额', 
            '新增付费率(%)', '老用户充值人数', '老用户充值金额', '老用户付费率(%)', 
            '老用户ARPPU', '老用户盈余率(%)', 'ARPPU'
        ]
        for metric in base_metrics:
            if metric in report_data.columns:
                final_report[metric] = report_data[metric].iloc[0]
            else:
                final_report[metric] = 0

        final_report['注册成本'] = total_cost / final_report['新增用户数'].replace(0, np.nan)
        final_report['首充成本'] = total_cost / final_report['首存人数'].replace(0, np.nan)
        final_report['充减提'] = final_report['充值金额'] - final_report['提现金额']
        final_report['盈余率(%)'] = (final_report['充减提'] / final_report['充值金额'].replace(0, np.nan)) * 100

        retention_data = df_book4_today[df_book4_today['来源渠道'] == '账号来源汇总']
        retention_cols = ['次日复充率(%)', '3日复充率(%)', '7日复充率(%)', '15日复充率(%)', '30日复充率(%)']
        if not retention_data.empty:
            for col in retention_cols:
                final_report[col] = retention_data[col].iloc[0]
        else:
            for col in retention_cols:
                final_report[col] = 0

        fission_channels = ['fission', 'agent', 'wheel']
        fission_data = df_book1_today[df_book1_today['渠道来源'].str.contains('|'.join(fission_channels), na=False)]
        
        if not fission_data.empty and '充值人数' in fission_data.columns and '首存人数' in fission_data.columns and fission_data['充值人数'].sum() > 0:
            fission_rate = fission_data['首存人数'].sum() / fission_data['充值人数'].sum()
        else:
            fission_rate = 0
        final_report['裂变率'] = fission_rate * 100

        for col in ['LTV-7天', 'LTV-15天', 'LTV-30天', '历史消耗', '历史充提差']:
            final_report[col] = 0
            
        final_report.fillna(0, inplace=True)

        # --- 5. 最终格式整理 ---
        print("正在按最终格式整理...")
        final_columns_order = [
            '日期', '产品', '消耗', '注册成本', '首充成本', '新增用户数', '充值人数', '充值金额', '提现金额',
            '充减提', '盈余率(%)', '首存人数', '首存充值金额', '首存付费率(%)', '首存盈余率(%)', '首存ARPPU',
            '新增付费人数', '新增充值金额', '新增付费率(%)', '老用户充值人数', '老用户充值金额', '老用户付费率(%)',
            '老用户ARPPU', '老用户盈余率(%)', 'ARPPU', '次日复充率(%)', '3日复充率(%)', '7日复充率(%)',
            '15日复充率(%)', '30日复充率(%)', 'LTV-7天', 'LTV-15天', 'LTV-30天', '裂变率',
            '历史消耗', '历史充提差'
        ]
        
        final_report = final_report.reindex(columns=final_columns_order).fillna(0)

        # --- 6. 保存文件 ---
        output_filename = f'产品日报_{report_date_str}.csv'
        print(f"正在保存结果到文件: {output_filename}...")
        final_report.to_csv(output_filename, index=False, encoding='utf-8-sig', float_format='%.2f')

        print("\n处理完成！")
        print(f"已为您生成【产品日报】。请在文件夹中查找名为 '{output_filename}' 的文件。")

    except FileNotFoundError as e:
        print(f"\n错误：文件未找到 - {e}")
    except Exception as e:
        print(f"\n处理数据时发生了一个意料之外的错误: {e}")

if __name__ == "__main__":
    generate_product_daily_report(target_date=target_date_str)
