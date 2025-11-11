"""试题类模块，定义存储试题信息的数据结构"""


class Question:
    """试题类，用于存储每道题的完整信息"""
    
    def __init__(self, question_id, title, question_type, option_a, option_b, option_c, option_d, 
                 score, correct_option, analysis, reading_passage=None, remarks=None):
        """
        初始化试题对象
        
        参数:
            question_id: 题号
            title: 题目内容
            question_type: 题型
            option_a: 选项A
            option_b: 选项B
            option_c: 选项C
            option_d: 选项D
            score: 分值
            correct_option: 正确选项
            analysis: 解析
            reading_passage: 阅读理解的文章（可选）
            remarks: 备注（可选）
        """
        self.question_id = question_id
        self.title = title
        self.question_type = question_type
        self.option_a = option_a
        self.option_b = option_b
        self.option_c = option_c
        self.option_d = option_d
        self.score = score
        self.correct_option = correct_option
        self.analysis = analysis
        self.reading_passage = reading_passage
        self.remarks = remarks
        
    def shuffle_options(self):
        """
        打乱选项顺序，同时更新正确选项
        
        返回:
            打乱后的选项映射字典
        """
        # 在方法内部延迟导入random模块
        import random
        
        # 优化：使用字典存储选项标签和值，提高查找效率
        option_dict = {}
        
        # 安全的字符串处理函数
        def safe_get_option(option):
            return str(option).strip() if option else ""
        
        # 收集非空选项
        if safe_get_option(self.option_a):
            option_dict['A'] = safe_get_option(self.option_a)
        if safe_get_option(self.option_b):
            option_dict['B'] = safe_get_option(self.option_b)
        if safe_get_option(self.option_c):
            option_dict['C'] = safe_get_option(self.option_c)
        if safe_get_option(self.option_d):
            option_dict['D'] = safe_get_option(self.option_d)
        
        # 如果选项太少，无法打乱
        if len(option_dict) <= 1:
            return {}
        
        # 保存原始正确选项的值
        original_correct_value = option_dict.get(self.correct_option)
        
        # 获取所有选项标签和值
        option_list = list(option_dict.items())  # [(标签, 值), ...]
        
        # 打乱选项顺序
        random.shuffle(option_list)
        
        # 创建新的选项映射
        new_options = {}
        option_labels = ['A', 'B', 'C', 'D']
        
        # 将打乱后的选项重新分配到A、B、C、D
        for i, (_, value) in enumerate(option_list):
            if i < len(option_labels):
                new_options[option_labels[i]] = value
        
        # 更新选项
        self.option_a = new_options.get('A', "")
        self.option_b = new_options.get('B', "")
        self.option_c = new_options.get('C', "")
        self.option_d = new_options.get('D', "")
        
        # 更新正确选项
        if original_correct_value:
            # 查找原始正确选项值在新位置的标签
            for label, value in new_options.items():
                if value == original_correct_value:
                    self.correct_option = label
                    break
        
        # 返回映射字典（如果需要的话）
        return {label: label for label in new_options}
    
    def __str__(self):
        """返回试题的字符串表示"""
        return f"题号: {self.question_id}, 题型: {self.question_type}, 分值: {self.score}"