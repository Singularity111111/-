"""
集团分公司业务投入效能自动化分析系统
该脚本自动读取各分公司的月度财务和业务数据，计算核心KPI，
生成健康度评分，并最终输出一个包含数据汇总和图表仪表盘的Excel报告。
"""
import pandas as pd
import os
import datetime
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows

# 确保matplotlib能够显示中文
plt.rcParams['font.sans-serif'] = ['SimHei'] # 指定默认字体
plt.rcParams['axes.unicode_minus'] = False # 解决保存图像是负号'-'显示为方块的问题

class PerformanceAnalyzer:
    """
    业务投入效能分析核心类。
    处理数据加载、KPI计算、评分和报告生成。
    """
    def __init__(self, data_path, report_path, score_rules):
        """
        初始化分析器。
        :param data_path: 存放分公司Excel数据的文件夹路径
        :param report_path: 存放生成报告的文件夹路径
        :param score_rules: 健康度评分规则
        """
        self.data_path = data_path
        self.report_path = report_path
        self.score_rules = score_rules
        self.all_data = pd.DataFrame()
        self.kpi_data = pd.DataFrame()

    def create_dummy_data(self):
        """
        为方便测试，自动生成模拟的财务和业务数据Excel文件。
        """
        print("正在生成模拟数据...")
        
        # 模拟数据文件夹
        if not os.path.exists(self.data_path):
            os.makedirs(self.data_path)

        companies = ['公司A', '公司B', '公司C']
        
        # 生成2023年7月、8月、9月的数据
        for company in companies:
            for month in range(7, 10):
                date_str = f'2023{month:02}'
                
                # 财务数据
                financial_df = pd.DataFrame({
                    '科目': ['营业收入', '营业成本', '市场费用', '研发费用', '管理费用'],
                    '金额': [
                        100000 * (1 + (month-7)*0.2) + (10000* (month-7)),  # 模拟增长
                        40000 * (1 + (month-7)*0.2) + (5000* (month-7)),
                        20000 * (1 + (month-7)*0.1),
                        15000,
                        5000
                    ]
                })
                financial_df.to_excel(os.path.join(self.data_path, f'{company}_财务_{date_str}.xlsx'), index=False)
                
                # 业务数据
                business_df = pd.DataFrame({
                    '指标': ['月初用户数', '月末用户数', '新增用户数', '总订单数', '付费用户数'],
                    '数值': [
                        50000 + (month-7)*10000,
                        60000 + (month-7)*10000,
                        10000 + (month-7)*1000,
                        8000 + (month-7)*500,
                        1500 + (month-7)*100
                    ]
                })
                business_df.to_excel(os.path.join(self.data_path, f'{company}_业务_{date_str}.xlsx'), index=False)

        print("模拟数据生成完毕。")

    def load_and_integrate_data(self):
        """
        从指定文件夹读取并整合所有分公司的数据。
        """
        print("正在读取和整合数据...")
        data_list = []
        try:
            for filename in os.listdir(self.data_path):
                if filename.endswith('.xlsx'):
                    parts = os.path.splitext(filename)[0].split('_')
                    if len(parts) == 3:
                        company = parts[0]
                        data_type = parts[1]
                        date_str = parts[2]
                        
                        file_path = os.path.join(self.data_path, filename)
                        df = pd.read_excel(file_path, engine='openpyxl')
                        df.columns = ['指标', '数值']
                        
                        # 将数据转为一行
                        df_transposed = df.set_index('指标').transpose()
                        df_transposed['公司'] = company
                        df_transposed['日期'] = date_str
                        df_transposed['数据类型'] = data_type
                        data_list.append(df_transposed)
            
            self.all_data = pd.concat(data_list, ignore_index=True)
            print("数据整合完毕。")
            return True

        except Exception as e:
            print(f"数据读取或整合过程中发生错误: {e}")
            return False

    def calculate_kpis(self):
        """
        计算各分公司的核心KPI。
        """
        print("正在计算关键绩效指标...")
        
        # 将数据按公司和日期进行整合
        financial_df = self.all_data[self.all_data['数据类型'] == '财务'].drop(columns='数据类型').reset_index(drop=True)
        business_df = self.all_data[self.all_data['数据类型'] == '业务'].drop(columns='数据类型').reset_index(drop=True)
        
        merged_df = pd.merge(financial_df, business_df, on=['公司', '日期'], how='outer')
        merged_df['日期'] = pd.to_datetime(merged_df['日期'], format='%Y%m')
        merged_df = merged_df.sort_values(by=['公司', '日期']).reset_index(drop=True)
        
        # 计算KPI
        self.kpi_data = merged_df.copy()
        
        # 营收月度增长率
        self.kpi_data['上月营收'] = self.kpi_data.groupby('公司')['营业收入'].shift(1)
        self.kpi_data['营收月度增长率'] = (self.kpi_data['营业收入'] - self.kpi_data['上月营收']) / self.kpi_data['上月营收']
        
        # 获客成本(CAC)
        self.kpi_data['获客成本(CAC)'] = self.kpi_data['市场费用'] / self.kpi_data['新增用户数']
        
        # 用户留存率（这里需要连续两个月数据，用月末用户数模拟）
        self.kpi_data['上月月末用户数'] = self.kpi_data.groupby('公司')['月末用户数'].shift(1)
        # 简化计算，假设本月活跃用户中，来自上月用户的比例为（本月月末用户数 - 本月新增用户数）/ 上月月末用户数
        self.kpi_data['用户留存率'] = (self.kpi_data['月末用户数'] - self.kpi_data['新增用户数']) / self.kpi_data['上月月末用户数']

        # LTV（估算）
        # 毛利率 = (营业收入 - 营业成本) / 营业收入
        self.kpi_data['毛利率'] = (self.kpi_data['营业收入'] - self.kpi_data['营业成本']) / self.kpi_data['营业收入']
        # LTV = (月度营收 / 付费用户数) * 毛利率 * (1 / (1 - 月度留存率))
        self.kpi_data['LTV（估算）'] = (self.kpi_data['营业收入'] / self.kpi_data['付费用户数']) * self.kpi_data['毛利率'] * (1 / (1 - self.kpi_data['用户留存率']))
        
        # LTV/CAC 比率
        self.kpi_data['LTV/CAC 比率'] = self.kpi_data['LTV（估算）'] / self.kpi_data['获客成本(CAC)']

        print("KPI计算完毕。")

    def calculate_scores(self):
        """
        基于KPI计算健康度评分和评级。
        """
        print("正在计算健康度评分...")
        
        def get_score_and_rating(row):
            total_score = 0
            
            # 营收增长率评分
            if row['营收月度增长率'] > self.score_rules['营收月度增长率']['高']: total_score += 5
            elif row['营收月度增长率'] > self.score_rules['营收月度增长率']['中']: total_score += 3
            else: total_score += 0
            
            # LTV/CAC 比率评分
            if row['LTV/CAC 比率'] > self.score_rules['LTV/CAC 比率']['高']: total_score += 5
            elif row['LTV/CAC 比率'] > self.score_rules['LTV/CAC 比率']['中']: total_score += 3
            else: total_score += 0

            # 评级
            if total_score >= 8: rating = 'A[优秀]'
            elif total_score >= 5: rating = 'B[良好]'
            elif total_score >= 2: rating = 'C[及格]'
            else: rating = 'D[危险]'
            
            return pd.Series([total_score, rating])

        # 仅对非空值行进行评分
        self.kpi_data[['综合得分', '健康度评级']] = self.kpi_data.apply(
            lambda row: get_score_and_rating(row) if pd.notna(row['营收月度增长率']) and pd.notna(row['LTV/CAC 比率']) else [None, None], axis=1
        )
        print("健康度评分计算完毕。")

    def generate_report(self):
        """
        生成最终的Excel报告。
        """
        print("正在生成报告...")
        
        if self.kpi_data.empty:
            print("KPI数据为空，无法生成报告。")
            return False

        today = datetime.date.today().strftime('%Y%m%d')
        report_filename = os.path.join(self.report_path, f'分公司效能分析报告_{today}.xlsx')
        
        if not os.path.exists(self.report_path):
            os.makedirs(self.report_path)

        try:
            # 创建Excel工作簿
            wb = Workbook()
            
            # 1. 汇总仪表盘 Sheet
            ws_dashboard = wb.active
            ws_dashboard.title = "汇总仪表盘"
            
            # 准备图表数据
            latest_kpi = self.kpi_data.dropna(subset=['健康度评级']).sort_values('日期').drop_duplicates(subset=['公司'], keep='last')
            
            # 图表1: LTV/CAC比率对比
            plt.figure(figsize=(10, 6))
            latest_kpi.plot(x='公司', y='LTV/CAC 比率', kind='bar', rot=0)
            plt.title('各分公司LTV/CAC比率对比')
            plt.xlabel('分公司')
            plt.ylabel('比率')
            plt.grid(axis='y', linestyle='--')
            ltv_cac_chart_path = os.path.join(self.report_path, 'ltv_cac_chart.png')
            plt.savefig(ltv_cac_chart_path)
            
            # 图表2: 健康度综合得分对比
            plt.figure(figsize=(10, 6))
            latest_kpi.plot(x='公司', y='综合得分', kind='bar', rot=0, color='orange')
            plt.title('各分公司健康度综合得分对比')
            plt.xlabel('分公司')
            plt.ylabel('综合得分')
            plt.grid(axis='y', linestyle='--')
            score_chart_path = os.path.join(self.report_path, 'score_chart.png')
            plt.savefig(score_chart_path)

            # 将图表嵌入到Excel
            img1 = OpenpyxlImage(ltv_cac_chart_path)
            img2 = OpenpyxlImage(score_chart_path)
            img1.anchor = 'A1'
            img2.anchor = 'A20'
            ws_dashboard.add_image(img1)
            ws_dashboard.add_image(img2)
            
            # 2. 明细数据 Sheet
            ws_details = wb.create_sheet(title="明细数据")
            for r in dataframe_to_rows(self.kpi_data, index=False, header=True):
                ws_details.append(r)
            
            wb.save(report_filename)
            print(f"报告生成完毕，文件保存在：{report_filename}")
            return True

        except Exception as e:
            print(f"报告生成过程中发生错误: {e}")
            return False
        
def main():
    """
    主程序入口。
    """
    # ----- 脚本配置区，您可以修改以下参数 -----
    config = {
        # 数据文件夹路径
        'data_folder': r"D:\分公司月度数据",
        # 报告输出文件夹路径
        'report_folder': r"D:\效能分析报告",
        # 健康度评分规则
        'score_rules': {
            '营收月度增长率': {'高': 0.15, '中': 0.05}, # > 15% 得5分, > 5% 得3分
            'LTV/CAC 比率': {'高': 3, '中': 1}, # > 3 得5分, > 1 得3分
        }
    }
    
    # ----------------------------------------
    
    analyzer = PerformanceAnalyzer(
        data_path=config['data_folder'],
        report_path=config['report_folder'],
        score_rules=config['score_rules']
    )

    # 1. (可选) 生成模拟数据，如果您已经准备好真实数据，请注释掉此行
    analyzer.create_dummy_data()

    # 2. 数据获取与整合
    if not analyzer.load_and_integrate_data():
        return

    # 3. 核心指标计算
    analyzer.calculate_kpis()
