import json
from pathlib import Path
import re

class KeywordManager:
    def __init__(self):
        self.exact_keywords = {}  # 精准匹配关键词库
        self.fuzzy_keywords = {}  # 模糊匹配关键词库
        self.load_keywords()
    
    def load_keywords(self):
        """从文件加载关键词库"""
        try:
            if Path('exact_keywords.json').exists():
                with open('exact_keywords.json', 'r', encoding='utf-8') as f:
                    self.exact_keywords = json.load(f)
            if Path('fuzzy_keywords.json').exists():
                with open('fuzzy_keywords.json', 'r', encoding='utf-8') as f:
                    self.fuzzy_keywords = json.load(f)
        except Exception as e:
            print(f"加载关键词库失败: {e}")
    
    def save_keywords(self):
        """保存关键词库到文件"""
        try:
            with open('exact_keywords.json', 'w', encoding='utf-8') as f:
                json.dump(self.exact_keywords, f, ensure_ascii=False, indent=4)
            with open('fuzzy_keywords.json', 'w', encoding='utf-8') as f:
                json.dump(self.fuzzy_keywords, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存关键词库失败: {e}")
    
    def add_keyword(self, keyword, type_name, is_fuzzy=False):
        """添加关键词"""
        if is_fuzzy:
            self.fuzzy_keywords[keyword] = type_name
        else:
            self.exact_keywords[keyword] = type_name
        self.save_keywords()
    
    def remove_keyword(self, keyword, is_fuzzy=False):
        """删除关键词"""
        if is_fuzzy:
            self.fuzzy_keywords.pop(keyword, None)
        else:
            self.exact_keywords.pop(keyword, None)
        self.save_keywords()
    
    def get_type(self, vuln_name, match_mode='both'):
        """
        根据漏洞名称获取类型
        match_mode: 'exact'(仅精准匹配), 'fuzzy'(仅模糊匹配), 'both'(先精准后模糊)
        """
        if match_mode in ['exact', 'both']:
            # 精准匹配
            if vuln_name in self.exact_keywords:
                return self.exact_keywords[vuln_name]
        
        if match_mode in ['fuzzy', 'both']:
            # 模糊匹配
            for keyword, type_name in self.fuzzy_keywords.items():
                if keyword in vuln_name:
                    return type_name
        
        return "未知"  # 修改默认返回值
    
    def batch_import(self, keywords_data, is_fuzzy=False, overwrite=True):
        """
        批量导入关键词
        
        Args:
            keywords_data: list of tuples [(keyword, type_name), ...]
            is_fuzzy: 是否为模糊匹配关键词
            overwrite: 是否覆盖已存在的关键词
        
        Returns:
            tuple: (成功数量, 跳过数量, 错误信息列表)
        """
        success_count = 0
        skip_count = 0
        errors = []
        
        target_dict = self.fuzzy_keywords if is_fuzzy else self.exact_keywords
        
        for keyword, type_name in keywords_data:
            try:
                keyword = str(keyword).strip()
                type_name = str(type_name).strip()
                
                if not keyword or not type_name:
                    errors.append(f"无效的数据: {keyword} -> {type_name}")
                    continue
                
                if keyword in target_dict and not overwrite:
                    skip_count += 1
                    continue
                
                target_dict[keyword] = type_name
                success_count += 1
                
            except Exception as e:
                errors.append(f"处理 {keyword} 时出错: {str(e)}")
        
        if success_count > 0:
            self.save_keywords()
            
        return success_count, skip_count, errors 