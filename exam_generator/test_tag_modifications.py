import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_reader import read_questions_from_excel
from word_generator import generate_exam_document
from question import Question

def test_tag_modifications():
    """测试TYPE.TAG标签修改功能"""
    print("开始测试TYPE.TAG标签修改...")
    
    # 创建测试数据
    test_questions = [
        Question(
            question_id='001',
            title='这是一道阅读理解题',
            question_type='阅读理解',
            option_a='选项A',
            option_b='选项B',
            option_c='选项C',
            option_d='选项D',
            score=2,
            correct_option='A',
            analysis='这是解析',
            reading_passage='这是一篇测试文章，用于验证TYPE.TAG标签的修改是否正确。'
        ),
        Question(
            question_id='002',
            title='这是第二道阅读理解题',
            question_type='阅读理解',
            option_a='选项A',
            option_b='选项B',
            option_c='选项C',
            option_d='选项D',
            score=2,
            correct_option='B',
            analysis='这是解析',
            reading_passage='这是一篇测试文章，用于验证TYPE.TAG标签的修改是否正确。'
        ),
        Question(
            question_id='003',
            title='这是一道普通选择题',
            question_type='单选题',
            option_a='选项A',
            option_b='选项B',
            option_c='选项C',
            option_d='选项D',
            score=2,
            correct_option='C',
            analysis='这是解析'
        )
    ]
    
    # 测试生成Word文档
    output_file = 'test_tags_new.docx'
    success = generate_exam_document(test_questions, output_file)
    
    if success:
        print(f"成功生成测试文档: {output_file}")
        print("\n检查生成的文档，确认以下内容：")
        print("1. 文章部分是否包含 <TYPE.TAG>文本行 标签")
        print("2. 每道题目是否包含 <TYPE.TAG>选择题 标签")
        print("3. 移除了 <TYPE.TAG>阅读理解 和 【阅读理解文章】 的标签")
        print("4. 普通题型是否使用 <TYPE.TAG>{q_type} 标签")
    else:
        print("生成文档失败")
    
    print("\n测试完成！")

if __name__ == "__main__":
    test_tag_modifications()