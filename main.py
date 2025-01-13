# -*- coding:utf-8 -*-
import json
import re
from pathlib import Path

def extract_json_from_html(input_file, output_file):
    """
    从HTML文件中提取JSON数据并保存
    
    Args:
        input_file (str): 输入HTML文件路径
        output_file (str): 输出JSON文件路径
        
    Returns:
        bool: 是否成功提取
    """
    try:
        with open(input_file, encoding='utf8') as f:
            file_content = f.read()

        pat_list = re.findall(r'<script>window.data = (.*?);</script>', file_content)
        
        if not pat_list:
            return False
            
        data_json = json.loads(pat_list[0])
        
        with open(output_file, 'w', encoding='utf8') as f:
            json.dump(data_json, f, ensure_ascii=False, indent=4)
            
        return True
        
    except Exception as e:
        raise Exception(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    # 命令行方式运行时的代码
    input_file = "index.html"
    output_file = "data.json"
    
    if extract_json_from_html(input_file, output_file):
        print(f"JSON数据已保存到: {output_file}")
    else:
        print("未找到匹配的JSON数据！")
