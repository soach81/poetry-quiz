#!/usr/bin/env python3
"""解析古诗词docx文件，生成JSON数据"""
import json
import re
import sys
from docx import Document

def parse_poem(text, poem_type):
    """解析单首诗词文本"""
    # 匹配 作者《标题》内容 格式
    match = re.match(r'^(.+?)《(.+?)》(.+)$', text.strip())
    if not match:
        return None

    author = match.group(1).strip()
    title = match.group(2).strip()
    content = match.group(3).strip()

    # 按所有标点拆分：逗号、句号、问号、感叹号
    # 这样每个"句"就是两个标点之间的内容
    lines = re.split(r'[，。！？、]', content)

    # 清理空行
    lines = [l.strip() for l in lines if l.strip()]

    # 如果拆分后行数和标点数不匹配，可能是拆分有问题
    # 直接按原始内容返回
    if len(lines) == 0:
        return None

    return {
        "author": author,
        "title": title,
        "type": poem_type,
        "lines": lines,
        "raw": content
    }

def main():
    docx_path = r"C:\Users\99114\Desktop\鑫禹姐的文件\古诗背诵\文本.docx"
    doc = Document(docx_path)

    poems = []
    current_type = "诗"

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 检测类型标记
        if text == "一、诗":
            current_type = "诗"
            continue
        elif text == "二、词":
            current_type = "词"
            continue

        # 解析诗词
        poem = parse_poem(text, current_type)
        if poem:
            poems.append(poem)

    # 输出JSON
    output_path = r"C:\Users\99114\Desktop\古诗默写题库\poems.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(poems, f, ensure_ascii=False, indent=2)

    print(f"成功解析 {len(poems)} 首诗词")
    print(f"诗: {sum(1 for p in poems if p['type'] == '诗')} 首")
    print(f"词: {sum(1 for p in poems if p['type'] == '词')} 首")
    print(f"输出文件: {output_path}")

if __name__ == "__main__":
    main()
