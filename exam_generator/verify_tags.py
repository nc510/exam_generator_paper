from docx import Document

def verify_tags_in_document(file_path):
    """验证文档中的TYPE.TAG标签是否符合最新要求"""
    print(f"验证文档: {file_path}")
    
    # 打开文档
    doc = Document(file_path)
    
    # 读取所有段落
    paragraphs = [para.text for para in doc.paragraphs]
    
    # 检查标签
    text_line_tag_count = 0
    choice_question_tag_count = 0
    other_type_tag_count = 0
    reading_tag_count = 0
    article_header_count = 0
    
    print("\n文档内容片段（包含TYPE.TAG的段落）:")
    for i, para in enumerate(paragraphs):
        if "<TYPE.TAG>" in para:
            print(f"段落 {i+1}: {para}")
            
            if "<TYPE.TAG>文本行" in para:
                text_line_tag_count += 1
            elif "<TYPE.TAG>选择题" in para:
                choice_question_tag_count += 1
            elif "<TYPE.TAG>阅读理解" in para:
                reading_tag_count += 1
            else:
                other_type_tag_count += 1
        
        if "【阅读理解文章" in para:
            article_header_count += 1
    
    print(f"\n验证结果:")
    print(f"- <TYPE.TAG>文本行 标签数: {text_line_tag_count}")
    print(f"- <TYPE.TAG>选择题 标签数: {choice_question_tag_count}")
    print(f"- <TYPE.TAG>阅读理解 标签数: {reading_tag_count}")
    print(f"- 其他题型标签数: {other_type_tag_count}")
    print(f"- 【阅读理解文章】标题数: {article_header_count}")
    
    # 检查是否满足最新要求
    issues = []
    
    if text_line_tag_count == 0:
        issues.append("文章部分缺少 <TYPE.TAG>文本行 标签")
    
    if choice_question_tag_count == 0:
        issues.append("题目部分缺少 <TYPE.TAG>选择题 标签")
    
    if reading_tag_count > 0:
        issues.append("仍存在 <TYPE.TAG>阅读理解 标签")
    
    if article_header_count > 0:
        issues.append("仍存在 【阅读理解文章】标题")
    
    if issues:
        print("\n❌ 验证失败！发现以下问题:")
        for issue in issues:
            print(f"- {issue}")
        return False
    else:
        print("\n✅ 验证通过！所有标签符合要求:")
        print("- 文章部分使用 <TYPE.TAG>文本行 标签")
        print("- 每道题目使用 <TYPE.TAG>选择题 标签")
        print("- 移除了 <TYPE.TAG>阅读理解 和 【阅读理解文章】 的标签")
        return True

if __name__ == "__main__":
    verify_tags_in_document("test_tags_new.docx")