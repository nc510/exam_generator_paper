"""主程序模块，实现图形界面和程序入口"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# 延迟导入大型库，提高启动速度
# 这些函数将在需要时动态导入相应模块
def read_questions_from_excel_lazy(file_path, start_id=None, end_id=None):
    """延迟导入并调用read_questions_from_excel函数"""
    from excel_reader import read_questions_from_excel
    return read_questions_from_excel(file_path, start_id, end_id)

def get_question_ids_range_lazy(file_path):
    """延迟导入并调用get_question_ids_range函数"""
    from excel_reader import get_question_ids_range
    return get_question_ids_range(file_path)

def generate_exam_document_lazy(questions, output_file, title="考试试卷", shuffle_options=False):
    """延迟导入并调用generate_exam_document函数"""
    from word_generator import generate_exam_document
    return generate_exam_document(questions, output_file, title, shuffle_options)


class ExamGeneratorApp:
    """考试试卷生成器应用程序类"""
    
    def __init__(self, root):
        """
        初始化应用程序
        
        参数:
            root: tkinter根窗口
        """
        self.root = root
        self.root.title("考试试卷生成器")
        self.root.geometry("800x600")  # 扩大窗口尺寸以容纳预览区域
        
        # 设置窗口图标
        try:
            self.root.iconbitmap("logo2.ico")
        except Exception:
            # 如果图标设置失败，忽略错误，程序继续运行
            pass
        
        # 创建菜单栏
        self.create_menu()
        
        # 设置中文字体
        self.font_config = {"font": ("SimHei", 10)}
        
        # 文件路径变量
        self.excel_file_path = tk.StringVar()
        
        # 题号范围变量
        self.start_id_var = tk.StringVar()
        self.end_id_var = tk.StringVar()
        self.min_id = None
        self.max_id = None
        
        # 选项乱序开关
        self.shuffle_var = tk.BooleanVar(value=False)
        
        # 抽题方式（0:固定顺序，1:随机抽题）
        self.selection_method_var = tk.IntVar(value=0)
        
        # 随机抽题次数
        self.random_count_var = tk.StringVar(value="10")
        
        # 试题数据
        self.questions = []
        
        # 选中的试题ID集合
        self.selected_questions = set()
        
        # 创建界面
        self.create_widgets()
    
    def create_menu(self):
        """创建菜单栏和关于菜单"""
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        
        # 创建关于菜单
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="关于程序", command=self.show_about)
        
        # 添加关于菜单到菜单栏
        menubar.add_cascade(label="帮助", menu=about_menu)
        
        # 设置菜单栏
        self.root.config(menu=menubar)
    
    def show_about(self):
        """显示关于对话框，包含作者信息"""
        about_text = (
            "考试试卷生成器 v1.0\n\n"
            "作者：倾尽温柔 VS 一世无尘\n"
            "机构：咸宁市理工中等职业技术学校\n"
            "QQ：121666880\n\n"
            "本程序用于快速生成考试试卷，支持从Excel文件读取试题，\n"
            "可自定义题号范围、选项乱序和抽题方式。"
        )
        messagebox.showinfo("关于", about_text)
    
    def create_widgets(self):
        """创建图形界面组件"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧控制面板
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # Excel文件选择
        file_frame = ttk.LabelFrame(control_frame, text="Excel文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(file_frame, textvariable=self.excel_file_path, width=30).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="浏览...", command=self.browse_excel_file).pack(side=tk.RIGHT, padx=5)
        
        # 选项设置
        options_frame = ttk.LabelFrame(control_frame, text="选项设置", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        ttk.Checkbutton(options_frame, text="选项乱序", variable=self.shuffle_var).pack(anchor=tk.W, padx=5)
        
        # 抽题方式设置
        selection_frame = ttk.LabelFrame(control_frame, text="抽题方式设置", padding="10")
        selection_frame.pack(fill=tk.X, pady=5)
        
        # 固定顺序抽题
        ttk.Radiobutton(selection_frame, text="固定顺序抽题", variable=self.selection_method_var, value=0).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 随机抽题
        ttk.Radiobutton(selection_frame, text="随机抽题", variable=self.selection_method_var, value=1).grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 随机抽题数量
        ttk.Label(selection_frame, text="抽题数量:").grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(selection_frame, textvariable=self.random_count_var, width=10).grid(row=1, column=2, padx=5, pady=5)
        
        # 加载试题按钮
        load_button_frame = ttk.Frame(control_frame, padding="10")
        load_button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(load_button_frame, text="加载试题预览", command=self.load_questions_for_preview, style="TButton").pack(fill=tk.X)
        
        # 题号范围设置 - 移动到加载试题预览按钮下方
        range_frame = ttk.LabelFrame(control_frame, text="题号范围设置", padding="10")
        range_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(range_frame, text="起始题号:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(range_frame, textvariable=self.start_id_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(range_frame, text="结束题号:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(range_frame, textvariable=self.end_id_var, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Button(range_frame, text="获取题号范围", command=self.update_id_range).grid(row=1, column=0, columnspan=4, padx=5, pady=5)
        
        # 生成按钮
        button_frame = ttk.Frame(control_frame, padding="10")
        button_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(button_frame, text="生成试卷", command=self.generate_exam, style="Accent.TButton").pack(fill=tk.X)
        
        # 右侧试题预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="试题预览", padding="10")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 创建带有滑块的文本框
        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD, font=("SimHei", 10))
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # 确保滚动条正确绑定，支持滑块翻页
        self.preview_text.config(yscrollcommand=scrollbar.set)
        # 设置文本框可以编辑但不可直接修改内容，只允许通过程序更新
        self.preview_text.config(state=tk.NORMAL)
        
        # 绑定点击事件
        self.preview_text.bind("<Button-1>", self.on_text_click)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 设置样式
        style = ttk.Style()
        style.configure("Accent.TButton", font=("SimHei", 11, "bold"))
    
    def browse_excel_file(self):
        """浏览并选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_file_path.set(file_path)
            # 自动更新题号范围
            self.update_id_range()
    
    def update_id_range(self):
        """更新题号范围"""
        file_path = self.excel_file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("错误", "请先选择有效的Excel文件")
            return
        
        self.status_var.set("正在读取题号范围...")
        self.root.update()
        
        # 获取题号范围（延迟导入）
        min_id, max_id = get_question_ids_range_lazy(file_path)
        
        if min_id is not None and max_id is not None:
            self.min_id = min_id
            self.max_id = max_id
            self.start_id_var.set(str(min_id))
            self.end_id_var.set(str(max_id))
            self.status_var.set(f"已获取题号范围: {min_id} - {max_id}")
        else:
            self.status_var.set("获取题号范围失败")
            messagebox.showerror("错误", "无法获取题号范围，请检查Excel文件格式")
    
    def load_questions_for_preview(self):
        """加载试题用于预览"""
        file_path = self.excel_file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("错误", "请先选择有效的Excel文件")
            return
        
        try:
            start_id = int(self.start_id_var.get())
            end_id = int(self.end_id_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的题号数字")
            return
        
        self.status_var.set("正在加载试题预览...")
        self.root.update()
        
        # 读取试题数据（延迟导入）
        self.questions = read_questions_from_excel_lazy(file_path, start_id, end_id)
        
        if not self.questions:
            messagebox.showerror("错误", "未读取到试题数据")
            self.status_var.set("就绪")
            return
        
        # 清空预览区域和选中状态
        self.preview_text.delete(1.0, tk.END)
        self.selected_questions.clear()
        
        # 准备预览文本，减少多次插入操作
        preview_content = []
        shuffle_options = self.shuffle_var.get()
        
        for idx, question in enumerate(self.questions, 1):
            # 如果启用了选项乱序，则对预览的试题进行乱序处理
            if shuffle_options:
                question.shuffle_options()
            
            # 准备题目行
            question_header = f"{idx}. {question.title} (题型: {question.question_type}, 分值: {question.score}, 原题号: {idx})\n"
            preview_content.append(question_header)
            
            # 准备选项
            options = []
            if question.option_a and str(question.option_a).strip():
                options.append(f"  A. {question.option_a}")
            if question.option_b and str(question.option_b).strip():
                options.append(f"  B. {question.option_b}")
            if question.option_c and str(question.option_c).strip():
                options.append(f"  C. {question.option_c}")
            if question.option_d and str(question.option_d).strip():
                options.append(f"  D. {question.option_d}")
            
            if options:
                preview_content.append("\n".join(options) + "\n")
            
            # 准备答案和解析
            if question.correct_option:
                preview_content.append(f"  参考答案: {question.correct_option}\n")
            
            if question.analysis and str(question.analysis).strip():
                preview_content.append(f"  解析: {question.analysis}\n")
            
            # 添加空行分隔不同试题
            preview_content.append("\n")
        
        # 一次性插入所有内容，减少GUI更新次数
        self.preview_text.insert(tk.END, "".join(preview_content))
        self.status_var.set(f"已加载 {len(self.questions)} 道试题")
    
    def on_text_click(self, event):
        """处理文本点击事件"""
        # 获取点击的行号
        line_start = self.preview_text.index(f"@{event.x},{event.y} linestart")
        line_end = self.preview_text.index(f"@{event.x},{event.y} lineend")
        
        # 获取点击行的文本
        line_text = self.preview_text.get(line_start, line_end)
        
        # 尝试提取显示序号
        try:
            # 检查是否是题目行（以数字开头后接小数点）
            if line_text.strip() and line_text[0].isdigit() and '.' in line_text:
                # 提取序号
                display_idx_str = line_text.split('.')[0].strip()
                display_idx = int(display_idx_str)
                
                # 根据显示序号获取对应的试题
                if 1 <= display_idx <= len(self.questions):
                    question = self.questions[display_idx - 1]
                    
                    # 切换选中状态
                    # 注意：这里仍然使用question.question_id作为选中标识，因为这是Excel中的原始题号
                    if question.question_id in self.selected_questions:
                        self.selected_questions.remove(question.question_id)
                        # 恢复默认背景色
                        self.preview_text.tag_remove("selected", line_start, line_end)
                    else:
                        self.selected_questions.add(question.question_id)
                        # 设置选中背景色
                        self.preview_text.tag_add("selected", line_start, line_end)
        except (ValueError, IndexError):
            pass
        
        # 配置选中标签样式
        self.preview_text.tag_config("selected", background="#cce5ff")
    
    def generate_exam(self):
        """生成考试试卷"""
        # 验证输入
        file_path = self.excel_file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("错误", "请先选择有效的Excel文件")
            return
        
        try:
            start_id = int(self.start_id_var.get())
            end_id = int(self.end_id_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的题号数字")
            return
        
        if start_id > end_id:
            messagebox.showerror("错误", "起始题号不能大于结束题号")
            return
        
        # 验证题号范围
        if self.min_id is not None and self.max_id is not None:
            if start_id < self.min_id or end_id > self.max_id:
                if not messagebox.askyesno("警告", f"题号范围超出文件中的实际范围({self.min_id}-{self.max_id})，是否继续？"):
                    return
        
        # 选择输出文件
        output_file = filedialog.asksaveasfilename(
            title="保存试卷文件",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")]
        )
        
        if not output_file:
            return
        
        # 重用已加载的试题数据，避免重复读取
        questions = []
        if self.questions and start_id == int(self.start_id_var.get()) and end_id == int(self.end_id_var.get()):
            # 如果预览中的试题范围与当前选择的范围相同，则直接使用预览中的试题
            questions = self.questions.copy()
            self.status_var.set("正在准备试题数据...")
        else:
            self.status_var.set("正在读取试题数据...")
            # 读取试题数据（延迟导入）
            questions = read_questions_from_excel_lazy(file_path, start_id, end_id)
        
        # 应用筛选逻辑
        if self.selected_questions:
            filtered_questions = [q for q in questions if q.question_id in self.selected_questions]
            if not filtered_questions:
                messagebox.showinfo("提示", "没有找到选中的试题，将使用所有试题")
            else:
                questions = filtered_questions
        elif self.selection_method_var.get() == 1:  # 随机抽题
            try:
                import random
                random_count = int(self.random_count_var.get())
                if random_count <= 0:
                    messagebox.showerror("错误", "抽题数量必须大于0")
                    self.status_var.set("就绪")
                    return
                
                # 如果抽题数量超过可用试题数量，则使用所有试题
                if random_count > len(questions):
                    messagebox.showinfo("提示", f"抽题数量({random_count})超过可用试题数量({len(questions)})，将使用所有试题")
                else:
                    # 随机抽取指定数量的试题
                    questions = random.sample(questions, random_count)
            except ValueError:
                messagebox.showerror("错误", "请输入有效的抽题数量")
                self.status_var.set("就绪")
                return
        
        if not questions:
            messagebox.showerror("错误", "未读取到试题数据")
            self.status_var.set("就绪")
            return
        
        self.status_var.set("正在生成试卷...")
        self.root.update()
        
        shuffle_options = self.shuffle_var.get()
        
        # 根据抽题方式生成不同的标题
        if self.selection_method_var.get() == 0:  # 固定顺序抽题
            title = f"考试试卷（{start_id}-{end_id}题）"
        else:  # 随机抽题
            try:
                random_count = int(self.random_count_var.get())
                title = f"考试试卷（随机抽取{min(random_count, len(questions))}题）"
            except ValueError:
                title = f"考试试卷（随机抽取试题）"
        
        # 生成Word文档（延迟导入）
        success = generate_exam_document_lazy(
            questions=questions,
            output_file=output_file,
            title=title,
            shuffle_options=shuffle_options
        )
        
        if success:
            self.status_var.set("试卷生成成功")
            if messagebox.askyesno("成功", f"试卷已成功生成！\n文件路径: {output_file}\n\n是否打开文件？"):
                # 尝试打开文件
                try:
                    os.startfile(output_file)
                except Exception:
                    pass
        else:
            self.status_var.set("试卷生成失败")
            messagebox.showerror("错误", "生成试卷时发生错误")


def main():
    """程序入口"""
    root = tk.Tk()
    app = ExamGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()