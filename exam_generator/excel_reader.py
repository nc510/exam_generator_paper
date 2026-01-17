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
    import os
    from question import Question
    
    try:
        print(f"正在尝试读取Excel文件: {file_path}")
        print(f"文件是否存在: {os.path.exists(file_path)}")
        
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
        
        # 保存当前阅读理解文章，用于处理后续的题目
        current_reading_passage = None
        
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
            
            # 检查是否是阅读理解题型
            question_type = safe_str(row['题型'])
            
            # 处理阅读理解文章
            reading_passage = safe_str(row['阅读理解的文章']) if '阅读理解的文章' in df.columns and pd.notna(row['阅读理解的文章']) else ''
            
            if question_type == '阅读理解':
                if reading_passage:
                    # 如果有文章内容，更新当前文章并使用
                    current_reading_passage = reading_passage
                else:
                    # 如果没有文章内容，使用之前保存的文章
                    reading_passage = current_reading_passage
            else:
                # 非阅读理解题型，重置当前文章
                current_reading_passage = None
            
            question = Question(
                question_id=question_id,
                title=safe_str(row['题目']),
                question_type=question_type,
                option_a=safe_str(row['选项A']),
                option_b=safe_str(row['选项B']),
                option_c=safe_str(row['选项C']),
                option_d=safe_str(row['选项D']),
                score=float(row['分值']) if pd.notna(row['分值']) else 0,
                correct_option=safe_str(row['正确选项']),
                analysis=safe_str(row['解析']),
                reading_passage=reading_passage if question_type == '阅读理解' else None,
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
    import os
    import sys
    
    try:
        print(f"正在尝试获取Excel文件题号范围: {file_path}")
        print(f"文件是否存在: {os.path.exists(file_path)}")
        
        if not os.path.exists(file_path):
            print("文件不存在")
            return (None, None)
            
        print(f"文件大小: {os.path.getsize(file_path)} bytes")
        
        # 只读取题号列，减少内存使用
        # 使用更简单的参数配置，避免可能的阻塞问题
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"成功读取Excel文件，行数: {len(df)}, 列数: {len(df.columns)}")
        print(f"所有列名: {list(df.columns)}")
        
        if '题号' not in df.columns:
            print("Excel文件缺少'题号'列")
            return (None, None)
        
        # 获取题号列并过滤无效值
        question_ids = df['题号'].dropna()
        print(f"有效题号数量: {len(question_ids)}")
        
        if not question_ids.empty:
            min_id = int(question_ids.min())
            max_id = int(question_ids.max())
            print(f"最小题号: {min_id}, 最大题号: {max_id}")
            return (min_id, max_id)
        else:
            print("没有找到有效的题号")
            return (None, None)
    except Exception as e:
        print(f"获取题号范围时出错: {str(e)}")
        print(f"异常类型: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        return (None, None)