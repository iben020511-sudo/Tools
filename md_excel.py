"""
Markdown与Excel相互转换工具
支持多种输入编码，输出统一为UTF-8
"""

import pandas as pd
import os
import re
import argparse
from pathlib import Path

class MarkdownExcelConverter:
    # 支持的编码列表
    SUPPORTED_ENCODINGS = [
        'utf-8', 'utf-8-sig',
        'gbk', 'gb2312', 'gb18030',  # 中文编码
        'big5',  # 繁体中文
        'shift_jis', 'euc-jp',  # 日文
        'euc-kr', 'cp949',  # 韩文
        'latin-1', 'iso-8859-1',
        'cp1251', 'cp1252', 'koi8-r',  # 欧洲语言
        'ascii'
    ]
    
    def __init__(self):
        pass
    
    def detect_encoding(self, file_path, sample_size=1024):
        """
        自动检测文件编码
        """
        encodings_to_try = self.SUPPORTED_ENCODINGS.copy()
        
        for encoding in encodings_to_try:
            try:
                with open(file_path, 'rb') as f:
                    raw_data = f.read(sample_size)
                    # 尝试解码
                    raw_data.decode(encoding, errors='strict')
                    # 如果成功，尝试读取整个文件的一小部分验证
                    with open(file_path, 'r', encoding=encoding) as test_file:
                        test_file.read(100)
                    return encoding
            except (UnicodeDecodeError, LookupError):
                continue
        
        # 如果都不行，使用chardet（如果可用）或返回utf-8
        try:
            import chardet
            with open(file_path, 'rb') as f:
                raw_data = f.read(sample_size * 4)
                result = chardet.detect(raw_data)
                if result['confidence'] > 0.7:
                    return result['encoding']
        except ImportError:
            pass
        
        return 'utf-8'  # 默认编码
    
    def md_to_excel(self, md_file, excel_file=None, sheet_name='Sheet1'):
        """
        将Markdown表格转换为Excel文件
        """
        # 如果未指定输出文件名，使用相同的文件名
        if excel_file is None:
            excel_file = Path(md_file).with_suffix('.xlsx')
        
        # 检测编码并读取Markdown文件
        encoding = self.detect_encoding(md_file)
        print(f"检测到编码: {encoding}")
        
        try:
            with open(md_file, 'r', encoding=encoding) as f:
                content = f.read()
        except Exception as e:
            print(f"读取文件时出错: {e}")
            return False
        
        # 提取Markdown表格
        tables = self._extract_markdown_tables(content)
        
        if not tables:
            print("未找到Markdown表格")
            return False
        
        # 写入Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            for i, table in enumerate(tables):
                sheet_name_i = f"{sheet_name}_{i+1}" if i > 0 else sheet_name
                table.to_excel(writer, sheet_name=sheet_name_i, index=False)
        
        print(f"转换完成！已保存到: {excel_file}")
        return True
    
    def _extract_markdown_tables(self, content):
        """
        从Markdown内容中提取表格
        """
        tables = []
        
        # 使用正则表达式匹配Markdown表格
        # 格式: | Header1 | Header2 |
        #       |---------|---------|
        #       | Cell1   | Cell2   |
        table_pattern = r'(\|.*\|[ \t]*\n)((?:\|.*\|[ \t]*\n)*)'
        
        # 改进的正则表达式，匹配完整的表格
        lines = content.split('\n')
        table_lines = []
        in_table = False
        current_table = []
        
        for line in lines:
            line = line.rstrip()
            # 检查是否是表格行（包含 | 字符）
            if '|' in line and line.strip().startswith('|'):
                if not in_table:
                    in_table = True
                current_table.append(line)
            else:
                if in_table and current_table:
                    # 处理当前表格
                    table_df = self._parse_table_lines(current_table)
                    if table_df is not None:
                        tables.append(table_df)
                    current_table = []
                    in_table = False
        
        # 处理最后一个表格
        if in_table and current_table:
            table_df = self._parse_table_lines(current_table)
            if table_df is not None:
                tables.append(table_df)
        
        return tables
    
    def _parse_table_lines(self, lines):
        """
        解析表格行并转换为DataFrame
        """
        if len(lines) < 2:
            return None
        
        # 处理表头
        header_line = lines[0]
        headers = [cell.strip() for cell in header_line.split('|')[1:-1]]
        
        # 检查是否有分隔行（第二行应该是分隔行）
        separator_line = lines[1]
        if re.match(r'^\|[\s:-]+\|[\s:-]+\|', separator_line):
            data_lines = lines[2:]
        else:
            data_lines = lines[1:]
        
        # 解析数据行
        data = []
        for line in data_lines:
            if line.strip() and '|' in line:
                cells = [cell.strip() for cell in line.split('|')[1:-1]]
                if len(cells) == len(headers):
                    data.append(cells)
        
        if not headers:
            return None
        
        return pd.DataFrame(data, columns=headers)
    
    def excel_to_md(self, excel_file, md_file=None, sheet_name=None):
        """
        将Excel文件转换为Markdown表格
        """
        # 如果未指定输出文件名，使用相同的文件名
        if md_file is None:
            md_file = Path(excel_file).with_suffix('.md')
        
        try:
            # 读取Excel文件
            if sheet_name:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                tables = [(sheet_name, df)]
            else:
                # 读取所有工作表
                xls = pd.ExcelFile(excel_file)
                tables = []
                for sheet in xls.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet)
                    tables.append((sheet, df))
        except Exception as e:
            print(f"读取Excel文件时出错: {e}")
            return False
        
        # 转换为Markdown
        md_content = ""
        
        for sheet_name, df in tables:
            if len(tables) > 1:
                md_content += f"## {sheet_name}\n\n"
            
            # 将NaN替换为空字符串
            df = df.fillna('')
            
            # 生成Markdown表格
            if not df.empty:
                # 表头
                headers = df.columns.tolist()
                md_content += "| " + " | ".join(str(h) for h in headers) + " |\n"
                
                # 分隔线
                md_content += "| " + " | ".join(["---"] * len(headers)) + " |\n"
                
                # 数据行
                for _, row in df.iterrows():
                    row_data = [str(cell) if pd.notna(cell) else "" for cell in row]
                    md_content += "| " + " | ".join(row_data) + " |\n"
            
            md_content += "\n"
        
        # 以UTF-8编码写入文件
        try:
            with open(md_file, 'w', encoding='utf-8') as f:
                f.write(md_content)
            print(f"转换完成！已保存到: {md_file}")
            return True
        except Exception as e:
            print(f"写入文件时出错: {e}")
            return False
    
    def convert_csv_to_md(self, csv_file, md_file=None, delimiter=','):
        """
        将CSV文件转换为Markdown表格
        """
        if md_file is None:
            md_file = Path(csv_file).with_suffix('.md')
        
        # 尝试检测编码
        encoding = self.detect_encoding(csv_file)
        print(f"检测到编码: {encoding}")
        
        try:
            # 尝试读取CSV
            df = pd.read_csv(csv_file, encoding=encoding, delimiter=delimiter)
        except Exception as e:
            print(f"读取CSV文件时出错: {e}")
            # 尝试其他常见的分隔符
            for sep in [';', '\t', '|']:
                try:
                    df = pd.read_csv(csv_file, encoding=encoding, delimiter=sep)
                    print(f"使用分隔符: {repr(sep)}")
                    break
                except:
                    continue
            else:
                return False
        
        # 转换为Markdown
        return self._dataframe_to_md(df, md_file)
    
    def _dataframe_to_md(self, df, md_file):
        """
        将DataFrame转换为Markdown并保存
        """
        if df.empty:
            print("数据框为空")
            return False
        
        # 将NaN替换为空字符串
        df = df.fillna('')
        
        # 生成Markdown表格
        md_content = ""
        headers = df.columns.tolist()
        md_content += "| " + " | ".join(str(h) for h in headers) + " |\n"
        md_content += "| " + " | ".join(["---"] * len(headers)) + " |\n"
        
        for _, row in df.iterrows():
            row_data = [str(cell) if pd.notna(cell) else "" for cell in row]
            md_content += "| " + " | ".join(row_data) + " |\n"
        
        # 以UTF-8编码写入文件
        try:
            with open(md_file, 'w', encoding='utf-8') as f:
                f.write(md_content)
            print(f"转换完成！已保存到: {md_file}")
            return True
        except Exception as e:
            print(f"写入文件时出错: {e}")
            return False

def main():
    parser = argparse.ArgumentParser(description='Markdown与Excel相互转换工具')
    parser.add_argument('input', help='输入文件路径')
    parser.add_argument('-o', '--output', help='输出文件路径')
    parser.add_argument('-t', '--to', choices=['excel', 'md', 'auto'], default='auto',
                       help='转换目标类型: excel, md, 或auto(根据扩展名自动判断)')
    parser.add_argument('-s', '--sheet', help='Excel工作表名称(仅当转换到md时有效)')
    parser.add_argument('-e', '--encoding', help='指定输入文件编码(自动检测)')
    
    args = parser.parse_args()
    
    converter = MarkdownExcelConverter()
    
    # 确定转换方向
    input_path = Path(args.input)
    input_ext = input_path.suffix.lower()
    
    if args.to == 'auto':
        if input_ext in ['.md', '.markdown']:
            to_type = 'excel'
        elif input_ext in ['.xlsx', '.xls']:
            to_type = 'md'
        elif input_ext == '.csv':
            to_type = 'md'
        else:
            print("无法自动判断转换方向，请使用-t参数指定")
            return
    else:
        to_type = args.to
    
    # 执行转换
    if to_type == 'excel':
        converter.md_to_excel(args.input, args.output)
    elif to_type == 'md':
        if input_ext == '.csv':
            converter.convert_csv_to_md(args.input, args.output)
        else:
            converter.excel_to_md(args.input, args.output, args.sheet)

if __name__ == "__main__":
    main()