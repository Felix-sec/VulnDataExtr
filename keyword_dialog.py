import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

class KeywordDialog:
    def __init__(self, parent, keyword_manager):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("关键词管理")
        self.dialog.geometry("800x600")
        self.keyword_manager = keyword_manager
        
        # 创建选项卡
        self.notebook = ttk.Notebook(self.dialog)
        self.exact_frame = ttk.Frame(self.notebook)
        self.fuzzy_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.exact_frame, text="精准匹配")
        self.notebook.add(self.fuzzy_frame, text="模糊匹配")
        self.notebook.pack(expand=True, fill="both")
        
        # 创建精准匹配界面
        self.create_keyword_frame(self.exact_frame, False)
        # 创建模糊匹配界面
        self.create_keyword_frame(self.fuzzy_frame, True)
        
        # 在每个标签页添加批量导入按钮
        self.create_import_buttons(self.exact_frame, False)
        self.create_import_buttons(self.fuzzy_frame, True)
    
    def create_keyword_frame(self, frame, is_fuzzy):
        # 输入区域
        input_frame = ttk.LabelFrame(frame, text="添加关键词", padding=5)
        input_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(input_frame, text="关键词:").grid(row=0, column=0, padx=5)
        keyword_entry = ttk.Entry(input_frame, width=30)
        keyword_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(input_frame, text="类型:").grid(row=0, column=2, padx=5)
        type_entry = ttk.Entry(input_frame, width=20)
        type_entry.grid(row=0, column=3, padx=5)
        
        def add_keyword():
            keyword = keyword_entry.get().strip()
            type_name = type_entry.get().strip()
            if keyword and type_name:
                self.keyword_manager.add_keyword(keyword, type_name, is_fuzzy)
                keyword_entry.delete(0, tk.END)
                type_entry.delete(0, tk.END)
                refresh_list()
            else:
                messagebox.showwarning("警告", "关键词和类型不能为空！")
        
        ttk.Button(input_frame, text="添加", command=add_keyword).grid(row=0, column=4, padx=5)
        
        # 列表区域
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        columns = ("关键词", "类型")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def refresh_list():
            tree.delete(*tree.get_children())
            keywords = self.keyword_manager.fuzzy_keywords if is_fuzzy else self.keyword_manager.exact_keywords
            for keyword, type_name in keywords.items():
                tree.insert("", "end", values=(keyword, type_name))
        
        def remove_selected():
            selected = tree.selection()
            if selected:
                item = tree.item(selected[0])
                keyword = item['values'][0]
                self.keyword_manager.remove_keyword(keyword, is_fuzzy)
                refresh_list()
        
        ttk.Button(frame, text="删除选中", command=remove_selected).pack(pady=5)
        
        refresh_list() 
    
    def create_import_buttons(self, frame, is_fuzzy):
        # 创建按钮框架并居中
        import_frame = ttk.Frame(frame)
        import_frame.pack(fill="x", padx=5, pady=10)
        
        # 创建一个子框架来容纳按钮，并使其居中
        button_frame = ttk.Frame(import_frame)
        button_frame.pack(anchor="center")
        
        # 导入按钮 - 设置宽度和高度
        ttk.Button(button_frame, text="导入Excel", width=15, padding=(5, 8),
                  command=lambda: self.import_from_excel(is_fuzzy)).pack(side="left", padx=10)
        
        # 导出按钮 - 设置宽度和高度
        ttk.Button(button_frame, text="导出Excel", width=15, padding=(5, 8),
                  command=lambda: self.export_to_excel(is_fuzzy)).pack(side="left", padx=10)
    
    def import_from_excel(self, is_fuzzy):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if not filename:
            return
            
        try:
            df = pd.read_excel(filename)
            
            # 检查必要的列
            required_columns = ['关键词', '类型']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("错误", "Excel文件必须包含'关键词'和'类型'列！")
                return
            
            # 准备导入数据
            keywords_data = list(zip(df['关键词'], df['类型']))
            
            # 创建导入选项对话框
            option_dialog = tk.Toplevel(self.dialog)
            option_dialog.title("导入选项")
            option_dialog.geometry("300x150")
            
            overwrite_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(option_dialog, text="覆盖已存在的关键词", 
                          variable=overwrite_var).pack(pady=10)
            
            def do_import():
                success_count, skip_count, errors = self.keyword_manager.batch_import(
                    keywords_data, is_fuzzy, overwrite_var.get()
                )
                
                result_msg = f"成功导入: {success_count}\n跳过: {skip_count}"
                if errors:
                    result_msg += f"\n\n错误信息:\n" + "\n".join(errors)
                
                messagebox.showinfo("导入结果", result_msg)
                option_dialog.destroy()
                
                # 刷新显示
                for frame in [self.exact_frame, self.fuzzy_frame]:
                    for child in frame.winfo_children():
                        if isinstance(child, ttk.Frame):
                            for widget in child.winfo_children():
                                if isinstance(widget, ttk.Treeview):
                                    self.refresh_list(widget, frame == self.fuzzy_frame)
            
            ttk.Button(option_dialog, text="开始导入", 
                      command=do_import).pack(pady=10)
            ttk.Button(option_dialog, text="取消", 
                      command=option_dialog.destroy).pack(pady=5)
            
            # 使对话框模态
            option_dialog.transient(self.dialog)
            option_dialog.grab_set()
            self.dialog.wait_window(option_dialog)
            
        except Exception as e:
            messagebox.showerror("错误", f"导入过程中出错：\n{str(e)}")
    
    def export_to_excel(self, is_fuzzy):
        try:
            # 获取要导出的关键词库
            keywords = self.keyword_manager.fuzzy_keywords if is_fuzzy else self.keyword_manager.exact_keywords
            
            # 检查关键词库是否为空
            if not keywords:
                messagebox.showwarning("警告", "关键词库为空，无法导出！")
                return
            
            # 创建DataFrame
            df = pd.DataFrame([
                {'关键词': keyword, '类型': type_name}
                for keyword, type_name in keywords.items()
            ])
            
            # 让用户选择保存位置
            filename = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel文件", "*.xlsx"),
                    ("所有文件", "*.*")
                ]
            )
            
            if filename:
                # 导出到Excel
                df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"关键词库已导出到：\n{filename}")
        
        except Exception as e:
            messagebox.showerror("错误", f"导出过程中出错：\n{str(e)}")
    
    def refresh_list(self, tree, is_fuzzy):
        """刷新指定树形视图的显示"""
        tree.delete(*tree.get_children())
        keywords = self.keyword_manager.fuzzy_keywords if is_fuzzy else self.keyword_manager.exact_keywords
        for keyword, type_name in keywords.items():
            tree.insert("", "end", values=(keyword, type_name)) 