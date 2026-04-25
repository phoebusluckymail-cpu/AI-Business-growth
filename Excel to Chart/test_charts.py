"""自动化测试Excel图表生成功能"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_to_chart import ExcelChartGenerator

def test_chart_generation():
    print("=" * 60)
    print("开始测试Excel图表生成功能")
    print("=" * 60)
    
    generator = ExcelChartGenerator()
    
    sample_file = os.path.join(os.path.dirname(__file__), "示例销售数据.xlsx")
    
    if not os.path.exists(sample_file):
        print(f"[ERROR] 找不到示例文件: {sample_file}")
        return False
    
    print(f"\n[1/5] 加载Excel文件: {sample_file}")
    try:
        generator.load_excel_file(sample_file)
        print("  ✓ 文件加载成功")
    except Exception as e:
        print(f"  ✗ 文件加载失败: {str(e)}")
        return False
    
    print(f"\n[2/5] 获取数据信息")
    try:
        data_info = generator.get_data_info()
        print(f"  ✓ 数据形状: {data_info['shape'][0]} 行 × {data_info['shape'][1]} 列")
        print(f"  ✓ 列名: {', '.join(data_info['columns'])}")
        print(f"  ✓ 数值列: {', '.join(data_info['numeric_columns'])}")
        print(f"  ✓ 类别列: {', '.join(data_info['categorical_columns'])}")
    except Exception as e:
        print(f"  ✗ 获取数据信息失败: {str(e)}")
        return False
    
    print(f"\n[3/5] 配置图表")
    chart_configs = [
        {
            "type": "bar",
            "title": "各月份产品A销售额柱状图",
            "category_col": "月份",
            "data_col": "产品A销售额",
            "x_col": None,
            "y_col": None
        },
        {
            "type": "line",
            "title": "客户数量趋势折线图",
            "category_col": "月份",
            "data_col": "客户数量",
            "x_col": None,
            "y_col": None
        },
        {
            "type": "pie",
            "title": "12月各产品销售额占比饼图",
            "category_col": "月份",
            "data_col": "产品B销售额",
            "x_col": None,
            "y_col": None
        },
        {
            "type": "scatter",
            "title": "客户数量与满意度评分散点图",
            "category_col": None,
            "data_col": None,
            "x_col": "客户数量",
            "y_col": "满意度评分"
        }
    ]
    print(f"  ✓ 已配置 {len(chart_configs)} 个图表:")
    for idx, config in enumerate(chart_configs, 1):
        print(f"    {idx}. {config['type']}: {config['title']}")
    
    print(f"\n[4/5] 生成图表工作表")
    try:
        generator.create_chart_sheet(chart_configs)
        print("  ✓ 图表工作表创建成功")
        print(f"  ✓ 工作表列表: {', '.join(generator.wb.sheetnames)}")
    except Exception as e:
        print(f"  ✗ 生成图表失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    print(f"\n[5/5] 保存文件")
    try:
        output_path = os.path.join(os.path.dirname(__file__), "测试输出_示例销售数据.xlsx")
        result = generator.save_file(output_path)
        print(f"  ✓ 文件保存成功: {result}")
        print(f"  ✓ 文件大小: {os.path.getsize(result) / 1024:.1f} KB")
    except Exception as e:
        print(f"  ✗ 保存文件失败: {str(e)}")
        return False
    
    print("\n" + "=" * 60)
    print("✓ 所有测试通过！")
    print("=" * 60)
    return True

if __name__ == "__main__":
    success = test_chart_generation()
    if success:
        print("\n测试完成，程序运行正常！")
        sys.exit(0)
    else:
        print("\n测试失败，请检查错误信息！")
        sys.exit(1)
