import os
import shutil
import pandas as pd
from PyQt5 import Qt
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QCheckBox, QHBoxLayout, QWidget


def generate_unique_filename(output_folder, base_filename):
    """
    生成唯一的文件名，避免重名
    
    :param output_folder: 输出文件夹路径
    :param base_filename: 原始文件名
    :return: 唯一的文件名
    """
    # 分离文件名和扩展名
    name, ext = os.path.splitext(base_filename)
    
    # 初始计数器
    counter = 1
    unique_filename = base_filename
    
    # 如果文件已存在，则添加数字后缀
    while os.path.exists(os.path.join(output_folder, unique_filename)):
        unique_filename = f"{name}_{counter}{ext}"
        counter += 1
    
    return unique_filename

def rename_files(excel_data, source_files, output_folder, col1_index, col2_index=None):
    """
    批量重命名文件
    
    :param excel_data: pandas DataFrame
    :param source_files: 源文件完整路径列表
    :param output_folder: 输出文件夹路径
    :param col1_index: 第一列索引（从0开始）
    :param col2_index: 第二列索引（可选，从0开始）
    :return: 日志列表
    """
    logs = []
    
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 遍历源文件
    for file_path in source_files:
        try:
            # 获取原始文件名和扩展名
            original_filename = os.path.basename(file_path)
            file_ext = os.path.splitext(original_filename)[1]
            
            # 读取当前文件的Excel数据
            current_excel_data = pd.read_excel(file_path)
            
            # 将列名转换为字符串
            current_excel_data.columns = [str(col) for col in current_excel_data.columns]
            
            # 生成新文件名的基础部分
            new_filename = ''
            
            # 第一列
            if col1_index is not None and col1_index < len(current_excel_data.columns):
                # 获取第一列的第一行值
                new_filename = str(current_excel_data.iloc[:, col1_index].iloc[0])
            
            # 第二列（如果选择了）
            if col2_index is not None and col2_index < len(current_excel_data.columns):
                new_filename += f'_{current_excel_data.iloc[:, col2_index].iloc[0]}'
            
            # 添加原始文件扩展名
            new_filename += file_ext
            
            # 生成唯一文件名
            unique_new_filename = generate_unique_filename(output_folder, new_filename)
            
            # 构建完整路径
            output_path = os.path.join(output_folder, unique_new_filename)
            
            # 复制文件
            shutil.copy2(file_path, output_path)
            
            # 记录日志
            logs.append(f'重命名：{original_filename} -> {unique_new_filename}')
        
        except Exception as e:
            logs.append(f'处理文件 {file_path} 失败：{str(e)}')
    
    return logs

def select_files(self):
    # 支持多种Excel文件格式
    file_paths, _ = QFileDialog.getOpenFileNames(
        self, 
        '选择Excel文件', 
        '', 
        'Excel Files (*.xlsx *.xls *.csv *.xlsm *.xlsb)'
    )
    
    if not file_paths:
        return
    
    try:
        # 读取第一个Excel文件
        self.excel_data = pd.read_excel(file_paths[0])
        
        # 详细调试信息
        print("Excel文件内容：")
        print(self.excel_data)
        print("Excel列信息：", list(self.excel_data.columns))
        print("Excel数据类型：", self.excel_data.dtypes)
        
        # 将列名转换为字符串
        self.excel_data.columns = [str(col) for col in self.excel_data.columns]
        
        # 将所有列转换为字符串
        for col in self.excel_data.columns:
            self.excel_data[col] = self.excel_data[col].astype(str)
        
        # 更新文件标签
        self.file_label.setText(f'已选择 {len(file_paths)} 个Excel文件')
        
        # 填充列选择下拉框
        columns = [''] + list(self.excel_data.columns)
        self.column_combo1.clear()
        self.column_combo2.clear()
        self.column_combo1.addItems(columns)
        self.column_combo2.addItems(columns)
        
        # 清空之前的文件列表
        self.file_table.setRowCount(0)
        self.selected_files = []
        
        # 使用Excel的行数生成文件列表
        for row, filename in enumerate(self.excel_data.index):
            row_count = self.file_table.rowCount()
            self.file_table.insertRow(row_count)
            
            # 文件名
            filename_item = QTableWidgetItem(f'file_{row+1}')
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)
            self.file_table.setItem(row_count, 0, QTableWidgetItem(''))  # 路径
            self.file_table.setItem(row_count, 1, filename_item)  # 文件名
            
            # 复选框
            checkbox = QCheckBox()
            checkbox.setChecked(True)  # 默认全选
            checkbox.stateChanged.connect(self.update_selected_files)
            
            checkbox_widget = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(Qt.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.file_table.setCellWidget(row_count, 2, checkbox_widget)
            
            # 默认添加到选中文件列表
            self.selected_files.append(f'file_{row+1}')
        
        # 更新全选复选框
        self.select_all_checkbox.setChecked(True)
        
        self.log_text.append(f'已加载 {len(self.selected_files)} 个文件')
        self.log_text.append(f'列信息：{", ".join(columns[1:])}')
        
    except Exception as e:
        self.log_text.append(f'读取Excel失败：{str(e)}')
        import traceback
        traceback.print_exc() 