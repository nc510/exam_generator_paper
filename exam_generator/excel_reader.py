"""Excel读取模块，用于从Excel文件中读取试题数据"""

# 延迟导入，减少模块加载时间
def read_questions_from_excel(file_path, start_id=None, end_id=None):
    """
    从Excel文件中读取试题数据
    
    参数:
        file_path: Excel文件路径
        start_id: 起始题号（可选）
        end_id: 结束题号（可选）
    
    返回:
        试题对象列表
    """
    # 在函数内部导入，实现延迟加载
    import pandas as pd
    from question import Question
    
    try:
        # 使用pandas读取Excel文件，只读取需要的列以提高效率
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # 检查必要的列是否存在
        required_columns = ['题号', '题目', '题型', '选项A', '选项B', '选项C', '选项D', '分值', '正确选项', '解析']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件缺少必要的列: {col}")
        
        # 预先过滤题号范围，减少数据处理量
        if start_id is not None:
            df = df[df['题号'] >= start_id]
        if end_id is not None:
            df = df[df['题号'] <= end_id]
        
        # 直接从DataFrame批量转换为Question对象列表，避免多次循环
        questions = []
        
        # 优化：只处理必要的行，避免逐行遍历时的额外筛选
        for _, row in df.iterrows():
            # 获取题号
            question_id = row['题号']
            
            # 创建Question对象，优化字符串转换
            def safe_str(value):
                """安全地将值转换为字符串"""
                if pd.notna(value):
                    return str(value).strip()
                return ''
            
            question = Question(
                question_id=question_id,
                title=safe_str(row['题目']),
                question_type=safe_str(row['题型']),
                option_a=safe_str(row['选项A']),
                option_b=safe_str(row['选项B']),
                option_c=safe_str(row['选项C']),
                option_d=safe_str(row['选项D']),
                score=float(row['分值']) if pd.notna(row['分值']) else 0,
                correct_option=safe_str(row['正确选项']),
                analysis=safe_str(row['解析']),
                reading_passage=safe_str(row['阅读理解的文章']) if '阅读理解的文章' in df.columns and pd.notna(row['阅读理解的文章']) else None,
                remarks=safe_str(row['备注']) if '备注' in df.columns and pd.notna(row['备注']) else None
            )
            
            questions.append(question)
        
        # 按题号排序
        questions.sort(key=lambda q: q.question_id)
        
        return questions
        
    except FileNotFoundError:
        print(f"错误: 找不到文件 {file_path}")
        return []
    except Exception as e:
        print(f"读取Excel文件时出错: {str(e)}")
        return []


def get_question_ids_range(file_path):
    """
    获取Excel文件中题号的范围
    
    参数:
        file_path: Excel文件路径
    
    返回:
        (最小题号, 最大题号) 元组
    """
    # 在函数内部导入，实现延迟加载
    import pandas as pd
    
    try:
        # 只读取题号列，减少内存使用
        df = pd.read_excel(file_path, usecols=['题号'], engine='openpyxl')
        if '题号' not in df.columns:
            raise ValueError("Excel文件缺少'题号'列")
        
        min_id = df['题号'].min()
        max_id = df['题号'].max()
        
        return (int(min_id), int(max_id))
    except Exception as e:
        print(f"获取题号范围时出错: {str(e)}")
        return (None, None)