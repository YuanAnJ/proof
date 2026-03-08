"""
快速测试批量导入解析逻辑
运行方式：D:/Code/proof/.venv/Scripts/python.exe test_docx_parse.py
"""
import sys
import os

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Django 环境设置
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'mysite.settings')
import django
django.setup()

from forms.utils.docx_import import (
    _extract_date_from_context,
    _extract_title_from_context,
    _extract_authors_from_context,
    _extract_accept_level_from_context,
    _extract_instruction_level_from_context,
)


def test_parsing():
    """测试各个字段的提取逻辑"""
    
    # 模拟docx正文内容
    sample_text = """
    深圳大学
    
    信息采用证明
    
    根据上级党政部门2017年02月信息情况反馈，兹证明深圳大学小纪、厦门大学小李、复旦大学小赵撰写的咨询报告《有关于咨询决策的报告》获中办内参采用，并获中央主要领导肯定性批示，为有关部门和领导决策提供积极参考。
    
    此证明仅供职称评审、课题结项、评奖评优、绩效考核使用。
    
    特此证明。
    """
    
    print("=" * 60)
    print("测试 DOCX 解析逻辑")
    print("=" * 60)
    
    # 测试日期提取
    print("\n1. 反馈日期提取:")
    date_info = _extract_date_from_context(sample_text)
    print(f"   原始值: {date_info['feedback_date_raw']}")
    print(f"   是否仅到月: {date_info['is_partial_date']}")
    print(f"   → 预期: 2017-02 (仅到月)")
    
    # 测试题目提取
    print("\n2. 报告题目提取:")
    title = _extract_title_from_context(sample_text)
    print(f"   提取结果: {title}")
    print(f"   → 预期: 有关于咨询决策的报告")
    
    # 测试作者和单位提取
    print("\n3. 作者和单位提取:")
    authors = _extract_authors_from_context(sample_text)
    for idx, (name, unit) in enumerate(authors, 1):
        print(f"   作者{idx}: {name}, 单位: {unit}")
    print(f"   → 预期:")
    print(f"      作者1: 小纪, 单位: 深圳大学")
    print(f"      作者2: 小李, 单位: 厦门大学")
    print(f"      作者3: 小赵, 单位: 复旦大学")
    
    # 测试采纳级别提取
    print("\n4. 采纳级别提取:")
    accept_level = _extract_accept_level_from_context(sample_text)
    print(f"   提取结果: {accept_level}")
    print(f"   → 预期: 中办")
    
    # 测试批示级别提取
    print("\n5. 批示级别提取:")
    instruction_level = _extract_instruction_level_from_context(sample_text)
    print(f"   提取结果: {instruction_level}")
    print(f"   → 预期: 中央主要领导")
    
    print("\n" + "=" * 60)
    print("测试完成！请对比实际提取结果与预期值。")
    print("如果有差异，可能需要调整正则表达式模式。")
    print("=" * 60)


if __name__ == "__main__":
    test_parsing()
