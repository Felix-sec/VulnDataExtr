import tkinter as tk
from tkinter import filedialog, messagebox
import json
import re
from pathlib import Path
import pandas as pd
from keyword_manager import KeywordManager
from keyword_dialog import KeywordDialog

class JsonExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("HTML JSON提取器")
        self.root.geometry("800x600")
        
        # 输入文件框
        self.input_frame = tk.LabelFrame(root, text="输入HTML文件", padx=10, pady=5)
        self.input_frame.pack(fill="x", padx=10, pady=5)
        
        self.input_path = tk.StringVar()
        self.input_entry = tk.Entry(self.input_frame, textvariable=self.input_path, width=80)
        self.input_entry.pack(side="left", padx=5)
        
        self.browse_btn = tk.Button(self.input_frame, text="浏览", command=self.browse_input)
        self.browse_btn.pack(side="left", padx=5)
        
        # 拖放提示移到输入框右边
        self.drop_label = tk.Label(self.input_frame, text="(支持拖放HTML文件)", pady=5)
        self.drop_label.pack(side="left", padx=5)
        
        # 输出文件框
        self.output_frame = tk.LabelFrame(root, text="输出JSON文件", padx=10, pady=5)
        self.output_frame.pack(fill="x", padx=10, pady=5)
        
        self.output_path = tk.StringVar()
        self.output_entry = tk.Entry(self.output_frame, textvariable=self.output_path, width=80)
        self.output_entry.pack(side="left", padx=5)
        
        self.save_btn = tk.Button(self.output_frame, text="浏览", command=self.browse_output)
        self.save_btn.pack(side="left", padx=5)
        
        # 将提取JSON按钮移到输出文件框中
        self.extract_btn = tk.Button(self.output_frame, text="提取JSON", 
                                   command=self.extract_json)
        self.extract_btn.pack(side="left", padx=5)
        
        # 添加关键词管理器
        self.keyword_manager = KeywordManager()
        
        # 创建漏洞分类框架
        vuln_frame = tk.LabelFrame(root, text="漏洞类型分类", padx=5, pady=5)
        vuln_frame.pack(fill="x", padx=10, pady=5)
        
        # 添加匹配模式选择
        self.match_mode = tk.StringVar(value="both")
        match_label = tk.Label(vuln_frame, text="匹配模式:")
        match_label.pack(side="left", padx=5)
        
        modes = [
            ("精准匹配", "exact"),
            ("模糊匹配", "fuzzy"),
            ("精准+模糊", "both")
        ]
        
        for text, mode in modes:
            tk.Radiobutton(vuln_frame, text=text, variable=self.match_mode, 
                          value=mode).pack(side="left", padx=5)
        
        # 添加关键词管理按钮
        self.keyword_btn = tk.Button(vuln_frame, text="关键词管理", 
                                   command=self.show_keyword_dialog)
        self.keyword_btn.pack(side="left", padx=15)
        
        # 添加导出按钮
        self.vuln_btn = tk.Button(vuln_frame, text="导出分类", 
                                 command=self.export_vuln_types)
        self.vuln_btn.pack(side="left", padx=5)
        
        # 日志框
        self.log_frame = tk.LabelFrame(root, text="运行日志", padx=10, pady=5)
        self.log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建文本框和滚动条
        self.log_text = tk.Text(self.log_frame, height=15, width=80)
        self.scrollbar = tk.Scrollbar(self.log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=self.scrollbar.set)
        
        # 放置文本框和滚动条
        self.scrollbar.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        
        # 状态显示
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(root, textvariable=self.status_var)
        self.status_label.pack(pady=5)
        
        # 绑定拖放事件
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def log_message(self, message, level="INFO"):
        """添加日志消息到日志框"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] [{level}] {message}\n")
        self.log_text.see("end")  # 自动滚动到最新消息
        
    def browse_input(self):
        filename = filedialog.askopenfilename(
            filetypes=[("HTML文件", "*.html"), ("所有文件", "*.*")]
        )
        if filename:
            self.input_path.set(filename)
            # 自动设置输出文件名
            output_path = Path(filename).with_suffix('.json')
            self.output_path.set(str(output_path))

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if filename:
            self.output_path.set(filename)

    def handle_drop(self, event):
        file_path = event.data
        self.log_message(f"收到拖放文件: {file_path}")
        
        # 移除花括号并转换为Path对象
        file_path = Path(file_path.strip('{}'))
        
        # 检查文件扩展名（不区分大小写）
        if file_path.suffix.lower() in ('.html', '.htm'):
            self.input_path.set(str(file_path))
            output_path = file_path.with_suffix('.json')
            self.output_path.set(str(output_path))
            self.log_message(f"已设置输入文件: {file_path}")
            self.log_message(f"已设置输出文件: {output_path}")
        else:
            self.log_message(f"无效的文件类型: {file_path}", "ERROR")

    def extract_json(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()
        
        if not input_file or not output_file:
            self.log_message("请选择输入和输出文件！", "ERROR")
            return
            
        try:
            # 尝试多种编码格式
            encodings = ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']
            file_content = None
            
            for encoding in encodings:
                try:
                    with open(input_file, encoding=encoding) as f:
                        file_content = f.read()
                    self.log_message(f"成功使用 {encoding} 编码读取文件")
                    break
                except UnicodeDecodeError:
                    continue
            
            if file_content is None:
                self.log_message(f"无法读取文件，已尝试以下编码: {', '.join(encodings)}", "ERROR")
                return
            
            self.log_message(f"正在处理文件: {input_file}")
            pat_list = re.findall(r'<script>window.data = (.*?);</script>', file_content)
            
            if not pat_list:
                self.log_message("未在HTML文件中找到匹配的JSON数据！", "ERROR")
                return
                
            data_json = json.loads(pat_list[0])
            
            with open(output_file, 'w', encoding='utf8') as f:
                json.dump(data_json, f, ensure_ascii=False, indent=4)
                
            success_msg = f"JSON数据已成功保存到: {output_file}"
            self.log_message(success_msg, "SUCCESS")
            self.status_var.set("JSON提取成功！")
            
        except Exception as e:
            error_msg = f"处理过程中出现错误：{str(e)}"
            self.log_message(error_msg, "ERROR")
            self.status_var.set(f"错误: {str(e)}")

    def show_keyword_dialog(self):
        KeywordDialog(self.root, self.keyword_manager)
    
    def export_vuln_types(self):
        try:
            # 先让用户选择保存位置
            output_excel = filedialog.asksaveasfilename(
                title="保存漏洞分类Excel",
                defaultextension=".xlsx",
                initialfile="漏洞类型分类.xlsx",
                filetypes=[
                    ("Excel文件", "*.xlsx"),
                    ("所有文件", "*.*")
                ]
            )
            
            if not output_excel:  # 用户取消选择
                return
            
            # 读取JSON文件
            with open(self.output_path.get(), 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 获取漏洞列表
            vuln_list = data['categories'][3]['children'][0]['data']['vulns_info']['vuln_distribution']['vuln_list']
            
            # 准备Excel数据
            excel_data = []
            for i, vuln in enumerate(vuln_list, 1):
                # 处理描述和解决方案
                description = '\n'.join(filter(None, vuln.get('i18n_description', [])))
                solution = '\n'.join(filter(None, vuln.get('i18n_solution', [])))
                
                # 获取漏洞等级中文
                level_map = {'high': '高危', 'middle': '中危', 'low': '低危'}
                level = level_map.get(vuln.get('vuln_level', ''), '未知')
                
                # 获取漏洞类型
                vuln_type = self.keyword_manager.get_type(
                    vuln.get('i18n_name', ''), 
                    self.match_mode.get()
                )
                
                excel_data.append({
                    '序号': i,
                    '漏洞名称': vuln.get('i18n_name', ''),
                    '类型': vuln_type,
                    '漏洞等级': level,
                    '影响主机个数': vuln.get('vuln_count', 0),
                    '受影响主机': vuln.get('target', ''),
                    '详细描述': description,
                    '解决办法': solution
                })
            
            # 创建DataFrame并导出到Excel
            df = pd.DataFrame(excel_data)
            df.to_excel(output_excel, index=False)
            
            self.log_message(f"漏洞类型分类已导出到: {output_excel}", "SUCCESS")
            
        except Exception as e:
            self.log_message(f"导出漏洞类型分类时出错: {str(e)}", "ERROR")

if __name__ == "__main__":
    # 需要安装tkinterdnd2包来支持拖放功能
    from tkinterdnd2 import *
    
    root = TkinterDnD.Tk()
    app = JsonExtractorGUI(root)
    root.mainloop() 