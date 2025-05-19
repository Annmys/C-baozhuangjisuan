import win32com.client
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from openai import OpenAI
import json
import os

class WPSExcelAnalyzer:
    def __init__(self, api_key=None):
        """初始化WPS Excel分析器"""
        self.wps = win32com.client.Dispatch("ket.Application")
        self.ai_client = OpenAI(api_key=api_key) if api_key else None
        
    def open_workbook(self, file_path):
        """打开WPS工作簿"""
        try:
            self.workbook = self.wps.Workbooks.Open(file_path)
            return True
        except Exception as e:
            print(f"打开文件失败: {e}")
            return False
            
    def read_range(self, range_str, sheet_name=None):
        """读取指定范围的数据"""
        try:
            if sheet_name:
                sheet = self.workbook.Worksheets(sheet_name)
            else:
                sheet = self.workbook.ActiveSheet
            
            range_data = sheet.Range(range_str)
            data = []
            for row in range_data:
                data.append([cell.Value for cell in row])
            
            # 转换为DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])
            return df
        except Exception as e:
            print(f"读取数据失败: {e}")
            return None
            
    def analyze_with_ai(self, data_description):
        """使用AI分析数据"""
        if not self.ai_client:
            print("未配置AI API密钥")
            return None
            
        try:
            response = self.ai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "你是一个数据分析专家。"},
                    {"role": "user", "content": f"请分析以下数据并给出见解：\n{data_description}"}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"AI分析失败: {e}")
            return None
            
    def create_visualization(self, df, x_col, y_col, hue_col=None, col_col=None):
        """创建可视化图表"""
        try:
            plt.figure(figsize=(10, 6))
            if col_col:
                g = sns.lmplot(data=df, x=x_col, y=y_col, col=col_col, hue=hue_col)
            else:
                g = sns.lmplot(data=df, x=x_col, y=y_col, hue=hue_col)
            plt.tight_layout()
            return g
        except Exception as e:
            print(f"创建可视化失败: {e}")
            return None
            
    def save_results(self, output_path, analysis_results, fig=None):
        """保存分析结果"""
        try:
            # 保存AI分析结果
            with open(f"{output_path}_analysis.txt", "w", encoding="utf-8") as f:
                f.write(analysis_results)
            
            # 保存图表
            if fig:
                fig.savefig(f"{output_path}_plot.png")
            
            return True
        except Exception as e:
            print(f"保存结果失败: {e}")
            return False
            
    def close(self):
        """关闭WPS"""
        try:
            self.workbook.Close()
            self.wps.Quit()
        except:
            pass

def main():
    # 配置参数
    EXCEL_FILE = "your_excel_file.xlsx"
    SHEET_NAME = "Sheet1"
    RANGE_STR = "A2:G246"
    OUTPUT_PATH = "analysis_results"
    API_KEY = "your-openai-api-key"  # 替换为你的OpenAI API密钥
    
    # 创建分析器实例
    analyzer = WPSExcelAnalyzer(api_key=API_KEY)
    
    try:
        # 打开工作簿
        if not analyzer.open_workbook(EXCEL_FILE):
            return
        
        # 读取数据
        df = analyzer.read_range(RANGE_STR, SHEET_NAME)
        if df is None:
            return
            
        # 使用AI分析数据
        data_description = df.describe().to_string()
        analysis_results = analyzer.analyze_with_ai(data_description)
        
        # 创建可视化
        fig = analyzer.create_visualization(
            df,
            x_col="总账单",
            y_col="小费",
            col_col="时间",
            hue_col="是否吸烟"
        )
        
        # 保存结果
        analyzer.save_results(OUTPUT_PATH, analysis_results, fig)
        
    finally:
        # 关闭WPS
        analyzer.close()

if __name__ == "__main__":
    main() 