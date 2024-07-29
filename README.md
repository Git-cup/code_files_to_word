### 引言
1. 遍历代码文件里的代码，写入到word里。在申请软著申请时比较方便一些。
2. 写入word后，每页有平均54~55行代码，满足软著申请的每页代码行数要求。
3. 如果一行代码过长，在word里会是两行显示，所以写入word里后，展示的代码行数，会比较脚本统计的代码行数要多，请以脚本统计出来的为准。

### 环境要求：

1. **操作系统**：支持Python的操作系统，如Windows、macOS、Linux。
2. **Python版本**：Python 3.6或更高版本。

### 必备的Python库：

1. **`python-docx`**：用于创建和操作Word文档。

### 安装步骤：

1. **安装Python**：
   确保已安装Python 3.6或更高版本。可以在[Python官方网站](https://www.python.org/downloads/)下载并安装最新版本的Python。

2. **安装`python-docx`库**：
   使用pip安装`python-docx`库。可以通过以下命令安装：
   ```bash
   pip install python-docx
   ```

### 检查和安装步骤：

1. **检查Python安装**：
   在命令行或终端中输入以下命令以检查Python版本：
   ```bash
   python --version
   ```
   或者：
   ```bash
   python3 --version
   ```

2. **安装`python-docx`库**：
   在命令行或终端中输入以下命令以安装`python-docx`库：
   ```bash
   pip install python-docx
   ```
   或者：
   ```bash
   pip3 install python-docx
   ```

### 运行代码：

将Python脚本（`write_code_files_to_word.py`）保存到跟代码文件通过一个文件夹中，然后在命令行或终端中导航到该文件所在的目录并运行以下命令：
```bash
python write_code_files_to_word.py
```
或
```bash
python3 write_code_files_to_word.py
```

### 确保代码文件与脚本放在同一目录中：

将你希望脚本处理的代码文件与Python脚本放在同一个目录中。运行脚本时，它会自动在当前目录中查找并处理所有支持的代码文件类型，当查找到自己会跳过。示例如下：

![image](https://github.com/user-attachments/assets/b9e917a7-746f-4135-a803-a47579debc73)


### 依赖总结：

- Python 3.6或更高版本
- `python-docx`库

通过安装和设置上述环境和依赖项，你就可以运行该Python脚本，将当前目录中的代码文件内容写入Word文档，并生成汇总信息。

汇总示例
```bash
Total files: 40
.html files: 13
.py files: 1
.java files: 26
Total lines of code: 27573
```
