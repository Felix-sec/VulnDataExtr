# -*- coding:utf-8 -*-
import json
import re


def htmlToJson():
    with open('index.html', encoding='utf8') as f:
        fileContent = f.read()

    patList = re.findall(r'<script>window.data = (.*?);</script><title></title>', fileContent)
    data_Json = json.loads(patList[0])
    print(type(data_Json))
    print(data_Json)
    return data_Json

if __name__ == "__main__":
    data = htmlToJson()
    with open('data.json', 'w', encoding='utf8') as file:
        # 使用json.dump()方法将字典转换为json并写入文件
        # ensure_ascii=False 保证非ASCII字符以原始形式保存
        # indent=4 设置缩进为4个空格，使输出更加易读
        json.dump(data, file, ensure_ascii=False, indent=4)
