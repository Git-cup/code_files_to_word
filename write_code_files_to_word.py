import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
from collections import defaultdict

def write_code_files_to_word(word_file):
    # 获取当前脚本所在的目录和文件名
    directory = os.path.dirname(os.path.abspath(__file__))
    current_script = os.path.basename(__file__)
    
    # 创建一个新的Word文档
    doc = Document()
    total_lines = 0  # 初始化总行数
    file_count = 0  # 初始化文件总数
    file_type_counts = defaultdict(int)  # 用于统计每种文件类型的数量
    
    # 常见的代码文件扩展名
    code_extensions = {
        '.java', '.py', '.js', '.jsx', '.ts', '.tsx', '.cpp', '.c', '.h', '.hpp',
        '.cs', '.rb', '.go', '.swift', '.php', '.html', '.htm', '.css', '.scss', 
        '.sass', '.less', '.xml', '.json', '.yaml', '.yml', '.sql', '.sh', '.bat',
        '.ps1', '.r', '.pl', '.pm', '.lua', '.vb', '.vbs', '.asp', '.aspx', '.jsp',
        '.kt', '.m', '.mm', '.erl', '.ex', '.exs', '.groovy', '.scala', '.dart', 
        '.rs', '.clj', '.cljs', '.coffee', '.f90', '.f95', '.for', '.ada', '.adb',
        '.awk', '.pas', '.pp', '.inc', '.d', '.nim', '.lisp', '.scm', '.rkt'
    }
    
    # 遍历指定目录下的所有代码文件
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file == current_script:
                continue  # 跳过当前脚本文件
            file_ext = os.path.splitext(file)[1]
            if file_ext in code_extensions:
                file_path = os.path.join(root, file)
                
                # 统计文件类型
                file_count += 1
                file_type_counts[file_ext] += 1
                
                # 打开并读取代码文件的内容
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    
                # 将非空行逐行写入Word文档并设置格式
                for line in lines:
                    if line.strip():  # 去掉空行
                        para = doc.add_paragraph(line.rstrip())
                        # 设置字体大小
                        run = para.runs[0]
                        run.font.size = Pt(10)
                        # 设置段落格式
                        para_format = para.paragraph_format
                        para_format.space_before = Pt(0)  # 段前0行
                        para_format.space_after = Pt(0)   # 段后0行
                        para_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE  # 多倍行距
                        para_format.line_spacing = 1.0  # 行距为1.0倍
                        total_lines += 1  # 增加总行数

    # 保存Word文档
    doc.save(word_file)
    
    # 打印汇总信息
    print(f"Total files: {file_count}")
    for file_type, count in file_type_counts.items():
        print(f"{file_type} files: {count}")
    print(f"Total lines of code: {total_lines}")

# 使用示例
word_file = 'output.docx'
write_code_files_to_word(word_file)
