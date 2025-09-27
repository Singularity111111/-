import pandas as pd
import numpy as np
from datetime import datetime

# ==============================================================================
#                           *** 配置区域 ***
# ------------------------------------------------------------------------------
# 1. 日期设置:
#    - 如需计算特定日期，请修改日期字符串，例如 '2025-08-25'。
#    - 如需自动计算最新日期，请将值设置为 None, 即 target_date_str = None
# ------------------------------------------------------------------------------
target_date_str = '2025-08-25'  # <--- 请在这里修改您想要的日期

# ------------------------------------------------------------------------------
# 2. 文件名配置:
#    - 这里定义了脚本需要读取的所有数据源文件名。
#    - 如果未来您的源文件名发生变化，请在这里进行修改。
# ------------------------------------------------------------------------------
FILE_CONFIG = {
    "main_data": "Book1.xlsx",              # 总代数据原 (用于裂变率计算)
    "product_data": "Book2.xlsx",           # 产品数据源 (核心数据)
    "retention_data": "Book4.xlsx",         # 复充率表
    "cost_data": "巴基斯坦消耗.xlsx"      # 每日总消耗表
}
# ==============================================================================


def generate_product_daily_report(target_date=None):
    """
    根据配置区域的设置，生成【产品日报】。
    """
    try:
        # --- 0. 从配置中读取文件名 ---
        f_book1 = FILE_CONFIG["main_data"]
        f_book2 = FILE_CONFIG["product_data"]
        f_book4 = FILE_CONFIG["retention_data"]
        f_pakistan_cost = FILE_CONFIG["cost_data"]

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
            print(f"错误：核心文件 ({f_book2}) 中没有有效的日期数据。")
            return
            
        if target_date:
            report_date = pd.to_datetime(target_date)
            print(f"已指定日期为: {report_date.strftime('%Y-%m-%d')}，将为此日期生成日报...")
        else:
            report_date = df_book2['日期'].max()
            print(f"未指定日期，将自动检测最新日期: {report_date.strftime('%Y-%m-%d')}...")

        report_date_str = report_date.strftime('%Y-%m-%d')
        
        report_data = df_book2[df_book2['日期'] == report_date].copy()
        if report_data.empty:
            print(f"错误：{f_book2} 中找不到日期为 {report_date_str} 的数据。")
            return
            
        df_book1_today = df_book1[df_book1['日期'] == report_date]
        df_book4_today = df_book4[df_book4['日期'] == report_date]
        df_cost_today = df_cost[df_cost['日期'] == report_date]
        
        # --- 4. 提取和计算各个字段 ---
        print("正在计算各项指标...")
        final_report = pd.DataFrame([{'日期': report_date_str, '产品': 'aa'}])
        
        total_cost = df_cost_today['和'].iloc[0] if not df_cost_today.empty and '和' in df_cost_today.columns else 0
        final_report['消耗'] = total_cost

        base_metrics = [
            '新增用户数', '充值人数', '充值金额', '提现金额', '首存人数', '首存充值金额', 
            '首存付费率(%)', '首存盈余率(%)', '首存ARPPU', '新增付费人数', '新增充值金额', 
            '新增付费率(%)', '老用户充值人数', '老用户充值金额', '老用户付费率(%)', 
            '老用户ARPPU', '老用户盈余率(%)', 'ARPPU'
        ]
        for metric in base_metrics:
            final_report[metric] = report_data[metric].iloc[0] if metric in report_data.columns else 0

        final_report['注册成本'] = total_cost / final_report['新增用户数'].replace(0, np.nan)
        final_report['首充成本'] = total_cost / final_report['首存人数'].replace(0, np.nan)
        final_report['充减提'] = final_report['充值金额'] - final_report['提现金额']
        final_report['盈余率(%)'] = (final_report['充减提'] / final_report['充值金额'].replace(0, np.nan)) * 100

        retention_data = df_book4_today[df_book4_today['来源渠道'] == '账号来源汇总']
        retention_cols = ['次日复充率(%)', '3日复充率(%)', '7日复充率(%)', '15日复充率(%)', '30日复充率(%)']
        for col in retention_cols:
            final_report[col] = retention_data[col].iloc[0] if not retention_data.empty and col in retention_data else 0

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
        print("请检查【文件名配置】区域的设置是否正确，并确保所有必需的 .xlsx 文件都在脚本的同一个文件夹里。")
    except Exception as e:
        print(f"\n处理数据时发生了一个意料之外的错误: {e}")

if __name__ == "__main__":
    generate_product_daily_report(target_date=target_date_str)
