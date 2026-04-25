"""专门测试散点图创建"""
import sys
import os
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_to_chart import ExcelChartGenerator

def test_scatter_chart():
    print("=" * 60)
    print("测试散点图创建")
    print("=" * 60)
    
    generator = ExcelChartGenerator()
    
    sample_file = os.path.join(os.path.dirname(__file__), "示例销售数据.xlsx")
    
    print(f"\n加载文件: {sample_file}")
    generator.load_excel_file(sample_file)
    print("文件加载成功")
    
    print("\n获取数据信息:")
    data_info = generator.get_data_info()
    print(f"  列名: {data_info['columns']}")
    print(f"  数值列: {data_info['numeric_columns']}")
    
    print("\n配置散点图:")
    chart_configs = [
        {
            "type": "scatter",
            "title": "客户数量与满意度评分散点图",
            "category_col": None,
            "data_col": None,
            "x_col": "客户数量",
            "y_col": "满意度评分"
        }
    ]
    print(f"  配置: {chart_configs[0]}")
    
    print("\n尝试创建图表工作表...")
    try:
        generator.create_chart_sheet(chart_configs)
        print("  ✓ 图表工作表创建成功")
        
        output_path = os.path.join(os.path.dirname(__file__), "散点图测试输出.xlsx")
        result = generator.save_file(output_path)
        print(f"  ✓ 文件保存成功: {result}")
        print(f"  ✓ 文件大小: {os.path.getsize(result) / 1024:.1f} KB")
        
    except Exception as e:
        print(f"  ✗ 创建失败: {str(e)}")
        print("\n详细错误信息:")
        traceback.print_exc()
        return False
    
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)
    return True

if __name__ == "__main__":
    success = test_scatter_chart()
    sys.exit(0 if success else 1)
