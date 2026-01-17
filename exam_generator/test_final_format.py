import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_reader import read_questions_from_excel
from word_generator import generate_exam_document

def test_final_format():
    """测试最终排版格式是否符合要求"""
    print("开始测试最终排版格式...")
    
    # 读取实际的阅读理解Excel文件
    excel_file = "D:\code\AI_Code\exam_data\exam_data\exam_generator\阅读理解.xlsx"
    if not os.path.exists(excel_file):
        print(f"错误：找不到文件 {excel_file}")
        return False
    
    try:
        # 读取前10道题用于测试
        questions = read_questions_from_excel(excel_file)
        print(f"成功读取 {len(questions)} 道题")
        
        # 生成测试文档
        output_file = "final_format_test.docx"
        success = generate_exam_document(questions[:10], output_file, title="考试试卷")
        
        if success:
            print(f"成功生成测试文档: {output_file}")
            print("\n检查文档是否符合以下要求：")
            print("1. 试卷名称：考试试卷（1-10题）")
            print("2. 阅读理解,每小题2.0分。前后有<TYPE.TAG>文本行标签")
            print("3. 文章内容直接输出，没有额外标签")
            print("4. 每道题前有<TYPE.TAG>选择题标签和空行")
            print("5. 题目序号连续（1-10）")
            print("6. 每道题后有答案、分数和解析")
            return True
        else:
            print("生成文档失败")
            return False
            
    except Exception as e:
        print(f"测试过程中出错: {str(e)}")
        return False

if __name__ == "__main__":
    test_final_format()