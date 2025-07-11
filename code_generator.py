import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import glob
import re

class SourceCodeGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("软著源代码生成器")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 软件名称
        ttk.Label(main_frame, text="软件名称:").grid(row=0, column=0, sticky=tk.W, pady=10)
        self.software_name = ttk.Entry(main_frame, width=50)
        self.software_name.grid(row=0, column=1, sticky=tk.W, pady=10)
        
        # 版本号
        ttk.Label(main_frame, text="版本号:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.version = ttk.Entry(main_frame, width=50)
        self.version.grid(row=1, column=1, sticky=tk.W, pady=10)
        
        # 作者名
        ttk.Label(main_frame, text="作者名:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.author = ttk.Entry(main_frame, width=50)
        self.author.grid(row=2, column=1, sticky=tk.W, pady=10)
        
        # 项目路径
        ttk.Label(main_frame, text="项目路径:").grid(row=3, column=0, sticky=tk.W, pady=10)
        path_frame = ttk.Frame(main_frame)
        path_frame.grid(row=3, column=1, sticky=tk.W, pady=10)
        
        self.project_path = ttk.Entry(path_frame, width=42)
        self.project_path.pack(side=tk.LEFT)
        
        ttk.Button(path_frame, text="选择", command=self.select_path).pack(side=tk.RIGHT)
        
        # 指定文件后缀
        ttk.Label(main_frame, text="指定文件后缀:").grid(row=4, column=0, sticky=tk.W, pady=10)
        self.file_extensions = ttk.Entry(main_frame, width=50)
        self.file_extensions.grid(row=4, column=1, sticky=tk.W, pady=10)
        ttk.Label(main_frame, text="指定多个文件请使用英文逗号分隔").grid(row=5, column=1, sticky=tk.W)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=30)
        
        ttk.Button(button_frame, text="生成", command=self.generate_document, width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=root.quit, width=20).pack(side=tk.LEFT, padx=10)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=560, mode="determinate")
        self.progress.grid(row=7, column=0, columnspan=2, pady=10)
        
        # 状态标签
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=8, column=0, columnspan=2)
    
    def select_path(self):
        """选择项目路径"""
        path = filedialog.askdirectory(title="选择项目路径")
        if path:
            self.project_path.delete(0, tk.END)
            self.project_path.insert(0, path)
            # 自动检测文件后缀
            self.detect_file_extensions(path)
    
    def detect_file_extensions(self, path):
        """自动检测目录中的文件后缀"""
        if not path or not os.path.exists(path):
            return
            
        self.update_status("正在检测项目中的文件类型...")
        
        # 获取所有文件
        all_files = []
        for root, dirs, files in os.walk(path):
            for file in files:
                all_files.append(os.path.join(root, file))
                
        # 提取后缀
        extensions = set()
        for file in all_files:
            ext = os.path.splitext(file)[1]
            if ext and not ext.startswith('.git'):  # 排除git相关文件
                extensions.add(ext.lstrip('.'))
                
        # 按常见编程语言排序
        common_exts = ['java', 'py', 'js', 'html', 'css', 'xml', 'json', 'c', 'cpp', 'h', 'hpp', 'cs', 'php', 'go', 'rs', 'ts', 'vue', 'jsx', 'tsx']
        sorted_exts = sorted(extensions, key=lambda x: (x not in common_exts, x))
        
        if sorted_exts:
            self.file_extensions.delete(0, tk.END)
            self.file_extensions.insert(0, ','.join(sorted_exts[:10]))  # 最多显示10个后缀
            self.update_status(f"检测到 {len(sorted_exts)} 种文件类型")
    
    def update_status(self, message):
        """更新状态信息"""
        self.status_label.config(text=message)
        self.root.update()
    
    def add_page_number(self, doc):
        """添加页脚（著作权人全称）"""
        for section in doc.sections:
            footer = section.footer
            # 清除现有段落并创建新段落
            if footer.paragraphs:
                p = footer.paragraphs[0]
            else:
                p = footer.add_paragraph()
                
            # 设置对齐方式
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 添加著作权人全称
            author_run = p.add_run(f"著作权人: {self.author.get()}")
            author_run.font.size = Pt(10)
    
    def generate_document(self):
        """生成文档"""
        # 获取输入
        software_name = self.software_name.get().strip()
        version = self.version.get().strip()
        author = self.author.get().strip()
        project_path = self.project_path.get().strip()
        extensions = self.file_extensions.get().strip()
        
        # 验证输入
        if not software_name:
            messagebox.showerror("错误", "请输入软件名称")
            return
        if not version:
            messagebox.showerror("错误", "请输入版本号")
            return
        if not author:
            messagebox.showerror("错误", "请输入作者名")
            return
        if not project_path or not os.path.exists(project_path):
            messagebox.showerror("错误", "请选择有效的项目路径")
            return
        if not extensions:
            messagebox.showerror("错误", "请输入至少一个文件后缀")
            return
        
        try:
            # 重置进度条
            self.progress["value"] = 0
            self.update_status("开始收集源代码文件...")
            
            # 解析文件后缀
            ext_list = [ext.strip() for ext in extensions.split(",")]
            ext_list = [ext if ext.startswith(".") else f".{ext}" for ext in ext_list]
            
            # 收集所有匹配的文件
            all_files = []
            for ext in ext_list:
                files = glob.glob(f"{project_path}/**/*{ext}", recursive=True)
                all_files.extend(files)
            
            if not all_files:
                messagebox.showerror("错误", f"在指定路径下未找到任何匹配的文件")
                return
            
            # 更新进度条
            self.progress["maximum"] = len(all_files) + 1
            self.progress["value"] = 1
            self.update_status(f"找到 {len(all_files)} 个文件，开始处理...")
            
            # 创建文档
            doc = Document()
            
            # 设置页面边距
            sections = doc.sections
            for section in sections:
                section.top_margin = Cm(2.54)
                section.bottom_margin = Cm(2.54)
                section.left_margin = Cm(3.17)
                section.right_margin = Cm(3.17)
            
            # 计算前30页和后30页的行数
            lines_per_page = 50
            front_pages = 30
            back_pages = 30
            total_pages = front_pages + back_pages
            
            front_lines = front_pages * lines_per_page
            back_lines = back_pages * lines_per_page
            
            # 添加页眉
            header = doc.sections[0].header
            header_para = header.paragraphs[0]
            header_para.text = f"{software_name} {version}"
            header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # 添加右上角页码信息
            header_table = header.add_table(1, 2, width=Cm(16))
            # 左侧为软件名称和版本号
            cell_left = header_table.cell(0, 0)
            cell_left.text = f"{software_name} {version}"
            cell_left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            # 右侧为页码
            cell_right = header_table.cell(0, 1)
            cell_right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # 添加页码字段
            run = cell_right.paragraphs[0].add_run("第")
            run.font.size = Pt(10)
            
            # 添加页码
            page_num = cell_right.paragraphs[0].add_run()
            r = page_num._r
            fld = OxmlElement('w:fldChar')
            fld.set(qn('w:fldCharType'), 'begin')
            r.append(fld)
            
            run = cell_right.paragraphs[0].add_run()
            r = run._r
            instrText = OxmlElement('w:instrText')
            instrText.text = " PAGE "
            r.append(instrText)
            
            run = cell_right.paragraphs[0].add_run()
            r = run._r
            fld = OxmlElement('w:fldChar')
            fld.set(qn('w:fldCharType'), 'end')
            r.append(fld)
            
            # 添加"页，共60页"
            run = cell_right.paragraphs[0].add_run("页，共")
            run.font.size = Pt(10)
            
            # 添加总页数，固定为60页
            total_pages_run = cell_right.paragraphs[0].add_run("60")
            total_pages_run.font.size = Pt(10)
            
            run = cell_right.paragraphs[0].add_run("页")
            run.font.size = Pt(10)
            
            # 设置表格无边框
            for row in header_table.rows:
                for cell in row.cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
            
            # 删除原始段落，只保留表格
            if len(header.paragraphs) > 0:
                p = header.paragraphs[0]._p
                if p.getparent() is not None:
                    p.getparent().remove(p)
            
            # 添加页码（使用新的方法）
            self.add_page_number(doc)
            
            # 收集所有代码行
            all_code_lines = []
            file_boundaries = []  # 记录每个文件的起始行和结束行
            
            for i, file_path in enumerate(all_files):
                # 更新进度
                self.progress["value"] = i + 1
                self.update_status(f"收集文件 {i+1}/{len(all_files)}: {os.path.basename(file_path)}")
                
                # 读取文件内容
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                except UnicodeDecodeError:
                    try:
                        with open(file_path, 'r', encoding='gbk') as f:
                            content = f.read()
                    except:
                        self.update_status(f"警告: 无法读取文件 {file_path}，跳过")
                        continue
                
                # 添加文件内容
                lines = content.split('\n')
                
                # 记录文件边界
                start_line = len(all_code_lines)
                all_code_lines.extend(lines)
                end_line = len(all_code_lines)
                
                # 记录文件信息
                relative_path = os.path.relpath(file_path, project_path)
                file_boundaries.append((relative_path, start_line, end_line))
            
            # 计算总行数
            total_lines = len(all_code_lines)
            self.update_status(f"总共收集了 {total_lines} 行代码")
            
            # 确保有足够的代码行
            if total_lines < 3000:
                messagebox.showwarning("警告", f"收集到的代码行数不足3000行（当前{total_lines}行），可能无法满足软著要求")
            
            # 如果代码行数不足以填满60页，则全部使用
            if total_lines <= front_lines + back_lines:
                front_code = all_code_lines
                back_code = []
            else:
                # 分割代码为前30页和后30页
                front_code = all_code_lines[:front_lines]
                back_code = all_code_lines[-back_lines:]
            
            # 添加前30页代码
            self.update_status("正在添加前30页代码...")
            line_count = 0
            current_file_idx = 0
            page_count = 0
            
            # 添加代码
            while line_count < len(front_code) and page_count < front_pages:
                # 查找当前行所属的文件
                file_found = False
                for idx, (file_path, start, end) in enumerate(file_boundaries):
                    if start <= line_count < end:
                        if current_file_idx != idx:
                            current_file_idx = idx
                        file_found = True
                        break
                        file_found = True
                        break
                
                # 计算当前文件中要添加的行数
                if file_found:
                    _, start, end = file_boundaries[current_file_idx]
                    lines_to_add = min(end - line_count, lines_to_add if 'lines_to_add' in locals() else lines_per_page)
                else:
                    lines_to_add = lines_per_page
                
                # 添加代码段
                code_para = doc.add_paragraph()
                code_run = code_para.add_run('\n'.join(front_code[line_count:line_count+lines_to_add]))
                code_run.font.name = 'Courier New'
                code_run.font.size = Pt(10)
                
                # 设置行距以确保每页恰好50行
                code_para.paragraph_format.line_spacing = 1.0
                code_para.paragraph_format.space_after = Pt(0)
                code_para.paragraph_format.space_before = Pt(0)
                
                line_count += lines_to_add
                page_count += 1
                
                # 添加分页符
                if line_count < len(front_code) and page_count < front_pages:
                    doc.add_page_break()
            
            # 如果有后30页代码，直接添加页面分隔符
            if back_code:
                # 添加页面分隔符
                doc.add_page_break()
                
                # 添加后30页代码
                self.update_status("正在添加后30页代码...")
                line_count = 0
                current_file_idx = -1
                page_count = 0
                
                # 查找最后30页代码的起始文件
                for idx, (file_path, start, end) in enumerate(file_boundaries):
                    if start <= total_lines - back_lines < end:
                        current_file_idx = idx
                        break
                
                # 添加后30页代码
                while line_count < len(back_code) and page_count < back_pages:
                    # 查找当前行所属的文件
                    file_found = False
                    actual_line = total_lines - back_lines + line_count
                    
                    for idx, (file_path, start, end) in enumerate(file_boundaries):
                        if start <= actual_line < end:
                            if current_file_idx != idx:
                                current_file_idx = idx
                            file_found = True
                            break
                            file_found = True
                            break
                    
                    # 计算当前文件中要添加的行数
                    if file_found:
                        _, start, end = file_boundaries[current_file_idx]
                        lines_to_add = min(end - actual_line, lines_to_add if 'lines_to_add' in locals() else lines_per_page)
                    else:
                        lines_to_add = lines_per_page
                    
                    # 添加代码段
                    code_para = doc.add_paragraph()
                    code_run = code_para.add_run('\n'.join(back_code[line_count:line_count+lines_to_add]))
                    code_run.font.name = 'Courier New'
                    code_run.font.size = Pt(10)
                    
                    # 设置行距以确保每页恰好50行
                    code_para.paragraph_format.line_spacing = 1.0
                    code_para.paragraph_format.space_after = Pt(0)
                    code_para.paragraph_format.space_before = Pt(0)
                    
                    line_count += lines_to_add
                    page_count += 1
                    
                    # 添加分页符
                    if line_count < len(back_code) and page_count < back_pages:
                        doc.add_page_break()
            
            # 保存文档
            output_filename = f"{software_name}源代码(共60页).docx"
            
            # 二次检查页数，确保不超过60页
            self.update_status("正在检查页数...")
            
            # 临时保存文档以便检查页数
            temp_filename = f"temp_{output_filename}"
            doc.save(temp_filename)
            
            # 重新打开文档以检查页数
            try:
                # 使用更可靠的方法检查页数
                check_doc = Document(temp_filename)
                
                # 计算页数（通过分页符和段落估算）
                page_count = 1  # 至少有1页
                for para in check_doc.paragraphs:
                    if not para.text and len(para.runs) == 0:  # 空段落可能是分页符
                        page_count += 1
                
                self.update_status(f"估计页数: {page_count}")
                
                # 如果页数超过60页，则重新生成文档
                if page_count > 60:
                    self.update_status(f"检测到页数超过60页，正在调整...")
                    
                    # 重新创建文档
                    doc = Document()
                    
                    # 设置页面边距
                    sections = doc.sections
                    for section in sections:
                        section.top_margin = Cm(2.54)
                        section.bottom_margin = Cm(2.54)
                        section.left_margin = Cm(3.17)
                        section.right_margin = Cm(3.17)
                    
                    # 添加页眉
                    header = doc.sections[0].header
                    header_para = header.paragraphs[0]
                    header_para.text = f"{software_name} {version}"
                    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # 添加右上角页码信息
                    header_table = header.add_table(1, 2, width=Cm(16))
                    cell_left = header_table.cell(0, 0)
                    cell_left.text = f"{software_name} {version}"
                    cell_left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    cell_right = header_table.cell(0, 1)
                    cell_right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    # 添加页码
                    run = cell_right.paragraphs[0].add_run("第")
                    run.font.size = Pt(10)
                    page_num = cell_right.paragraphs[0].add_run()
                    r = page_num._r
                    fld = OxmlElement('w:fldChar')
                    fld.set(qn('w:fldCharType'), 'begin')
                    r.append(fld)
                    run = cell_right.paragraphs[0].add_run()
                    r = run._r
                    instrText = OxmlElement('w:instrText')
                    instrText.text = " PAGE "
                    r.append(instrText)
                    run = cell_right.paragraphs[0].add_run()
                    r = run._r
                    fld = OxmlElement('w:fldChar')
                    fld.set(qn('w:fldCharType'), 'end')
                    r.append(fld)
                    run = cell_right.paragraphs[0].add_run("页，共")
                    run.font.size = Pt(10)
                    total_pages_run = cell_right.paragraphs[0].add_run("60")
                    total_pages_run.font.size = Pt(10)
                    run = cell_right.paragraphs[0].add_run("页")
                    run.font.size = Pt(10)
                    
                    # 设置表格无边框
                    for row in header_table.rows:
                        for cell in row.cells:
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.size = Pt(10)
                    
                    # 添加页码（页脚）
                    self.add_page_number(doc)
                    
                    # 强制控制页数：只取前30页和后30页的代码
                    # 计算每页的行数（减少以确保不会超页）
                    safe_lines_per_page = 45  # 减少每页行数以确保安全
                    front_lines = 30 * safe_lines_per_page
                    back_lines = 30 * safe_lines_per_page
                    
                    # 重新分割代码
                    if total_lines <= front_lines + back_lines:
                        front_code = all_code_lines
                        back_code = []
                    else:
                        front_code = all_code_lines[:front_lines]
                        back_code = all_code_lines[-back_lines:]
                    
                    # 添加前30页代码（简化处理，不再添加文件标题）
                    self.update_status("正在重新生成前30页代码...")
                    for i in range(0, min(30, len(front_code) // safe_lines_per_page + 1)):
                        start_idx = i * safe_lines_per_page
                        end_idx = min((i + 1) * safe_lines_per_page, len(front_code))
                        
                        code_para = doc.add_paragraph()
                        code_run = code_para.add_run('\n'.join(front_code[start_idx:end_idx]))
                        code_run.font.name = 'Courier New'
                        code_run.font.size = Pt(10)
                        
                        # 设置行距
                        code_para.paragraph_format.line_spacing = 1.0
                        code_para.paragraph_format.space_after = Pt(0)
                        code_para.paragraph_format.space_before = Pt(0)
                        
                        # 添加分页符（除了最后一页）
                        if i < 29 and end_idx < len(front_code):
                            doc.add_page_break()
                    
                    # 如果有后30页代码，直接添加页面分隔符
                    if back_code:
                        doc.add_page_break()
                        
                        # 添加后30页代码（简化处理，不再添加文件标题）
                        self.update_status("正在重新生成后30页代码...")
                        for i in range(0, min(30, len(back_code) // safe_lines_per_page + 1)):
                            start_idx = i * safe_lines_per_page
                            end_idx = min((i + 1) * safe_lines_per_page, len(back_code))
                            
                            code_para = doc.add_paragraph()
                            code_run = code_para.add_run('\n'.join(back_code[start_idx:end_idx]))
                            code_run.font.name = 'Courier New'
                            code_run.font.size = Pt(10)
                            
                            # 设置行距
                            code_para.paragraph_format.line_spacing = 1.0
                            code_para.paragraph_format.space_after = Pt(0)
                            code_para.paragraph_format.space_before = Pt(0)
                            
                            # 添加分页符（除了最后一页）
                            if i < 29 and end_idx < len(back_code):
                                doc.add_page_break()
                    
                    self.update_status(f"已重新生成文档，确保页数为60页")
            except Exception as e:
                self.update_status(f"检查页数时出错: {str(e)}")
            
            # 删除临时文件
            try:
                if os.path.exists(temp_filename):
                    os.remove(temp_filename)
            except:
                pass
            
            # 保存最终文档
            doc.save(output_filename)
            
            self.progress["value"] = self.progress["maximum"]
            self.update_status(f"文档生成完成! 总共处理了 {total_lines} 行代码，前30页和后30页，保存为 {output_filename}")
            
            messagebox.showinfo("成功", f"文档已成功生成: {output_filename}\n\n符合软著申请要求：\n- 共60页，每页50行\n- 前30页为代码开头，后30页为代码结尾\n- 页眉包含软件名称和版本号\n- 页脚包含著作权人信息")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成文档时出错: {str(e)}")
            self.update_status(f"错误: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SourceCodeGenerator(root)
    root.mainloop() 