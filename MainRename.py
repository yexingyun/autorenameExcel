import os
import sys
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QHBoxLayout, QWidget, QLabel, QFileDialog, QComboBox, 
                             QTextEdit, QTableWidget, QTableWidgetItem, QCheckBox)
from PyQt5.QtCore import Qt
from rename_utils import rename_files

class FileRenameApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.excel_data = None
        self.output_folder = None
        
        # 配置文件路径
        self.config_path = os.path.join(
            os.path.expanduser('~'), 
            '.rename_tool_config.json'
        )
        
        # 初始化配置
        self.init_config()
        
        self.initUI()
    
    def init_config(self):
        """初始化配置文件"""
        try:
            # 如果配置文件不存在，创建默认配置
            if not os.path.exists(self.config_path):
                default_config = {
                    'output_folder': '',
                    'col1_index': -1,
                    'col2_index': -1
                }
                with open(self.config_path, 'w') as f:
                    json.dump(default_config, f)
            
            # 读取配置
            with open(self.config_path, 'r') as f:
                config = json.load(f)
            
            # 恢复配置
            self.output_folder = config.get('output_folder', '')
            self.last_col1_index = config.get('col1_index', -1)
            self.last_col2_index = config.get('col2_index', -1)
        
        except Exception as e:
            print(f"初始化配置失败：{e}")
            # 如果出错，使用默认值
            self.output_folder = ''
            self.last_col1_index = -1
            self.last_col2_index = -1
    
    def save_config(self, col1_index=None, col2_index=None):
        """保存当前配置"""
        try:
            # 读取现有配置
            with open(self.config_path, 'r') as f:
                config = json.load(f)
            
            # 更新配置
            if self.output_folder:
                config['output_folder'] = self.output_folder
            
            if col1_index is not None:
                config['col1_index'] = col1_index
            
            if col2_index is not None:
                config['col2_index'] = col2_index
            
            # 写入配置文件
            with open(self.config_path, 'w') as f:
                json.dump(config, f)
        
        except Exception as e:
            print(f"保存配置失败：{e}")
    
    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        if folder_path:
            self.output_folder = folder_path
            self.output_label.setText(f'已选择：{folder_path}')
            
            # 自动保存输出文件夹配置
            self.save_config(
                col1_index=self.column_combo1.currentIndex() - 1, 
                col2_index=self.column_combo2.currentIndex() - 1
            )
    
    def select_files(self):
        # 支持多种文件格式
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, 
            '选择文件', 
            '', 
            'All Files (*.*)'
        )
        
        if not file_paths:
            return
        
        # 清空之前的文件列表
        self.file_table.setRowCount(0)
        self.selected_files = []
        
        for file_path in file_paths:
            row_count = self.file_table.rowCount()
            self.file_table.insertRow(row_count)
            
            # 文件路径
            path_item = QTableWidgetItem(os.path.dirname(file_path))
            path_item.setFlags(path_item.flags() & ~Qt.ItemIsEditable)
            self.file_table.setItem(row_count, 0, path_item)
            
            # 文件名
            filename_item = QTableWidgetItem(os.path.basename(file_path))
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)
            self.file_table.setItem(row_count, 1, filename_item)
            
            # 复选框
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            checkbox.stateChanged.connect(self.update_selected_files)
            
            checkbox_widget = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(Qt.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.file_table.setCellWidget(row_count, 2, checkbox_widget)
            
            # 添加到选中文件列表
            self.selected_files.append(file_path)
        
        # 更新全选复选框状态
        self.select_all_checkbox.setChecked(True)
        
        self.log_text.append(f'已选择 {len(self.selected_files)} 个文件')
        
        # 尝试加载第一个Excel文件的列信息
        excel_extensions = ['.xlsx', '.xls', '.csv', '.xlsm', '.xlsb']
        excel_files = [f for f in file_paths if os.path.splitext(f)[1].lower() in excel_extensions]
        
        if excel_files:
            try:
                # 读取第一个Excel文件
                self.excel_data = pd.read_excel(excel_files[0])
                
                # 将列名转换为字符串
                self.excel_data.columns = [str(col) for col in self.excel_data.columns]
                
                # 将所有列转换为字符串
                for col in self.excel_data.columns:
                    self.excel_data[col] = self.excel_data[col].astype(str)
                
                # 填充列选择下拉框（使用列索引）
                columns = [''] + [f'第 {i+1} 列' for i in range(len(self.excel_data.columns))]
                self.column_combo1.clear()
                self.column_combo2.clear()
                self.column_combo1.addItems(columns)
                self.column_combo2.addItems(columns)
                
                self.log_text.append(f'已加载Excel列信息：{", ".join(self.excel_data.columns)}')
                
                # 恢复上次选择的列
                if self.last_col1_index >= 0:
                    self.column_combo1.setCurrentIndex(self.last_col1_index + 1)
                if self.last_col2_index >= 0:
                    self.column_combo2.setCurrentIndex(self.last_col2_index + 1)
                
                # 如果有上次的输出文件夹，自动设置
                if self.output_folder:
                    self.output_label.setText(f'已选择：{self.output_folder}')
            
            except Exception as e:
                self.log_text.append(f'读取Excel列信息失败：{str(e)}')
                import traceback
                traceback.print_exc()
    
    def column_changed(self, index):
        """列选择发生变化时保存配置"""
        sender = self.sender()
        
        if sender == self.column_combo1:
            # 保存第一列索引
            self.save_config(col1_index=index - 1)
        elif sender == self.column_combo2:
            # 保存第二列索引
            self.save_config(col2_index=index - 1)
    
    def initUI(self):
        self.setWindowTitle('文件批量重命名工具')
        self.setGeometry(100, 100, 1000, 700)
        
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # 文件选择区域
        file_select_layout = QHBoxLayout()
        self.file_label = QLabel('未选择文件')
        select_file_btn = QPushButton('选择文件')
        select_file_btn.clicked.connect(self.select_files)
        file_select_layout.addWidget(self.file_label)
        file_select_layout.addWidget(select_file_btn)
        
        # 文件列表表格
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(3)  # 增加一列
        self.file_table.setHorizontalHeaderLabels(['文件路径', '文件名', '选择'])
        
        # 全选复选框
        self.select_all_checkbox = QCheckBox('全选')
        self.select_all_checkbox.setChecked(True)
        self.select_all_checkbox.stateChanged.connect(self.toggle_all_files)
        
        # 输出文件夹选择区域
        output_layout = QHBoxLayout()
        self.output_label = QLabel('未选择输出文件夹')
        select_output_btn = QPushButton('选择输出文件夹')
        select_output_btn.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(select_output_btn)
        
        # Excel列选择
        column_layout = QHBoxLayout()
        self.column_combo1 = QComboBox()
        self.column_combo2 = QComboBox()
        column_layout.addWidget(QLabel('选择第一列：'))
        column_layout.addWidget(self.column_combo1)
        column_layout.addWidget(QLabel('选择第二列（可选）：'))
        column_layout.addWidget(self.column_combo2)
        
        # 为列选择下拉框添加变化监听
        self.column_combo1.currentIndexChanged.connect(self.column_changed)
        self.column_combo2.currentIndexChanged.connect(self.column_changed)
        
        # 预览表格
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(4)
        self.preview_table.setHorizontalHeaderLabels(['原文件名', '列1值', '列2值', '新文件名'])
        
        # 重命名和输出按钮
        action_layout = QHBoxLayout()
        preview_btn = QPushButton('预览重命名')
        preview_btn.clicked.connect(self.preview_rename)
        rename_btn = QPushButton('开始重命名')
        rename_btn.clicked.connect(self.start_rename)
        action_layout.addWidget(preview_btn)
        action_layout.addWidget(rename_btn)
        
        # 日志区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        
        # 添加所有布局
        main_layout.addLayout(file_select_layout)
        main_layout.addWidget(self.file_table)
        main_layout.addLayout(output_layout)  # 新增输出文件夹选择
        main_layout.addLayout(column_layout)
        main_layout.addWidget(self.preview_table)
        main_layout.addLayout(action_layout)
        main_layout.addWidget(QLabel('操作日志：'))
        main_layout.addWidget(self.log_text)
        
        # 添加全选复选框到布局
        select_all_layout = QHBoxLayout()
        select_all_layout.addWidget(self.select_all_checkbox)
        select_all_layout.addStretch()
        main_layout.insertLayout(1, select_all_layout)  # 在文件列表上方插入
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 如果有上次的输出文件夹，自动设置
        if self.output_folder:
            self.output_label.setText(f'已选择：{self.output_folder}')
    
    def preview_rename(self):
        if self.excel_data is None:
            self.log_text.append('请先选择Excel文件')
            return
        
        if not self.output_folder:
            self.log_text.append('请先选择输出文件夹')
            return
        
        # 获取选择的列索引
        col1_index = self.column_combo1.currentIndex() - 1
        col2_index = self.column_combo2.currentIndex() - 1
        
        # 保存当前配置
        self.save_config(col1_index, col2_index)
        
        # 清空预览表格
        self.preview_table.setRowCount(0)
        
        # 遍历选择的文件
        for index, file_path in enumerate(self.selected_files):
            # 检查是否有足够的Excel行数据
            if index >= len(self.excel_data):
                break
            
            row_count = self.preview_table.rowCount()
            self.preview_table.insertRow(row_count)
            
            # 原文件名
            self.preview_table.setItem(row_count, 0, QTableWidgetItem(os.path.basename(file_path)))
            
            # 获取对应的Excel行
            row = self.excel_data.iloc[index]
            
            # 列1值
            if col1_index >= 0:
                self.preview_table.setItem(row_count, 1, QTableWidgetItem(str(row.iloc[col1_index])))
            
            # 列2值
            if col2_index >= 0:
                self.preview_table.setItem(row_count, 2, QTableWidgetItem(str(row.iloc[col2_index])))
            
            # 新文件名
            new_name = str(row.iloc[col1_index]) if col1_index >= 0 else ''
            if col2_index >= 0:
                new_name += f'_{row.iloc[col2_index]}'
            new_name += os.path.splitext(file_path)[1]
            
            self.preview_table.setItem(row_count, 3, QTableWidgetItem(new_name))
    
    def start_rename(self):
        # 检查必要条件
        if self.excel_data is None:
            self.log_text.append('请先选择Excel文件')
            return
        
        if not self.output_folder:
            self.log_text.append('请先选择输出文件夹')
            return
        
        # 检查是否选择了文件
        if not self.selected_files:
            self.log_text.append('请先选择要重命名的文件')
            return
        
        # 获取选择的列索引
        col1_index = self.column_combo1.currentIndex() - 1
        col2_index = self.column_combo2.currentIndex() - 1
        
        # 保存当前配置
        self.save_config(col1_index, col2_index)
        
        try:
            # 使用实际选择的文件的完整路径
            logs = rename_files(
                excel_data=self.excel_data, 
                source_files=self.selected_files, 
                output_folder=self.output_folder, 
                col1_index=col1_index, 
                col2_index=col2_index
            )
            
            # 显示日志
            for log in logs:
                self.log_text.append(log)
            
            self.log_text.append('重命名完成')
        
        except Exception as e:
            self.log_text.append(f'重命名失败：{str(e)}')

    def toggle_all_files(self, state):
        # 全选/取消全选
        for row in range(self.file_table.rowCount()):
            checkbox = self.file_table.cellWidget(row, 2).findChild(QCheckBox)
            checkbox.setChecked(state == Qt.Checked)

    def update_selected_files(self):
        # 更新选中的文件列表
        self.selected_files = []
        for row in range(self.file_table.rowCount()):
            checkbox = self.file_table.cellWidget(row, 2).findChild(QCheckBox)
            if checkbox.isChecked():
                file_path = os.path.join(
                    self.file_table.item(row, 0).text(), 
                    self.file_table.item(row, 1).text()
                )
                self.selected_files.append(file_path)
        
        # 同步更新全选复选框状态
        all_checked = all(
            self.file_table.cellWidget(row, 2).findChild(QCheckBox).isChecked() 
            for row in range(self.file_table.rowCount())
        )
        self.select_all_checkbox.setChecked(all_checked)
        
        self.log_text.append(f'已选择 {len(self.selected_files)} 个文件')

def main():
    app = QApplication(sys.argv)
    ex = FileRenameApp()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
