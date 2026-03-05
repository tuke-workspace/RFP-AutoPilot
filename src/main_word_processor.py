#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
RFP智能引擎
"""

import sys
import os
from pathlib import Path

# PyQt5 imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QPushButton, QLabel, QFileDialog, QTextEdit,
                           QProgressBar, QMessageBox, QHBoxLayout, QCheckBox,
                           QFrame, QGridLayout, QGroupBox, QAbstractButton,
                           QDialogButtonBox, QComboBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon

# Document processing imports
from document_processor import DocumentProcessor, load_response_templates, get_template_folder_path, PRODUCT_LIST


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("RFP智能引擎")
        self.setGeometry(200, 50, 1350, 1000)

        # 设置窗口图标（如果图标文件存在）
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # 创建内容部件
        content_widget = QWidget()
        content_widget.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                font-family: 'Microsoft YaHei UI';
            }
        """)
        self.setCentralWidget(content_widget)

        # 创建垂直布局
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 15, 20, 15)

        # 标题部分
        self.create_header_section(layout)

        # 功能选择部分
        self.create_function_selection_section(layout)

        # 文件选择部分
        self.create_file_selection_section(layout)

        # 进度条
        self.create_progress_section(layout)

        # 日志显示区域
        self.create_log_section(layout)

        # 按钮区域
        self.create_button_section(layout)

        # 状态栏
        status_bar = self.statusBar()
        if status_bar:
            status_bar.showMessage("就绪")

        # 初始化变量
        self.input_file = None
        self.output_file = None
        self.processor = None

        # 启动时加载回应条款模板
        self.load_templates_on_startup()

    def load_templates_on_startup(self):
        """启动时加载回应条款模板"""
        template_dir = get_template_folder_path()
        self.log_text.append(f"正在加载回应条款模板...")
        self.log_text.append(f"模板目录: {template_dir}")

        if not os.path.exists(template_dir):
            self.log_text.append(f"⚠️ 回应条款文件夹不存在，将使用默认应答")
            self.log_text.append(f"请在程序同级目录创建 '回应条款' 文件夹并放入模板文件")
        else:
            templates = load_response_templates()
            if templates:
                self.log_text.append(f"✅ 已加载 {len(templates)} 个回应模板:")
                for keyword in templates.keys():
                    self.log_text.append(f"   • {keyword}")
            else:
                self.log_text.append(f"⚠️ 未找到有效的模板文件，将使用默认应答")

        self.log_text.append("-" * 50)
        self.log_text.append("就绪，请选择要处理的文档")

    def create_header_section(self, layout):
        """创建头部标题部分"""
        # 标题标签
        title_label = QLabel("RFP智能引擎")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont("Microsoft YaHei UI", 16, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("""
            QLabel { 
                color: #2c3e50; 
                margin: 10px;
                padding: 8px;
            }
        """)
        layout.addWidget(title_label)
        
        # 作者信息
        author_label = QLabel("开发者：Tuke | 版本：V3.0")
        author_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        author_label.setFont(QFont("Microsoft YaHei UI", 9))
        author_label.setStyleSheet("QLabel { color: #7f8c8d; margin: 5px; }")
        layout.addWidget(author_label)
        
    def create_function_selection_section(self, layout):
        """创建功能选择部分"""
        # 功能选择组
        function_group = QGroupBox("🎯 处理功能选择")
        function_group.setFont(QFont("Microsoft YaHei UI", 11, QFont.Bold))
        function_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 10px;
                margin: 8px;
                padding-top: 15px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                           stop:0 #ffffff, stop:1 #f8f9fa);
                color: #2c3e50;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                background: white;
                border-radius: 4px;
                color: #3498db;
            }
        """)
        layout.addWidget(function_group)

        function_layout = QVBoxLayout(function_group)
        function_layout.setSpacing(12)
        function_layout.setContentsMargins(20, 15, 20, 15)

        # 格式调整选项 - 主选项
        self.format_checkbox = QCheckBox("🔧 格式调整（全选所有细分功能）")
        self.format_checkbox.setFont(QFont("Microsoft YaHei UI", 10))
        self.format_checkbox.setChecked(True)
        self.format_checkbox.setStyleSheet("""
            QCheckBox {
                color: #2c3e50;
                spacing: 8px;
                padding: 10px;
                border-radius: 6px;
                border: 1px solid #bbdefb;
                background-color: #f8fffe;
            }
            QCheckBox:hover {
                border: 1px solid #2196f3;
                background-color: #f3f8ff;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        self.format_checkbox.stateChanged.connect(self.on_format_checkbox_changed)
        function_layout.addWidget(self.format_checkbox)

        # 格式调整细分选项区域
        self.format_sub_frame = QFrame()
        self.format_sub_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f7ff;
                border-radius: 6px;
                border: 1px solid #bbdefb;
                margin-left: 20px;
                padding: 5px;
            }
        """)
        function_layout.addWidget(self.format_sub_frame)

        sub_layout = QGridLayout(self.format_sub_frame)
        sub_layout.setSpacing(8)
        sub_layout.setContentsMargins(15, 10, 15, 10)

        # 细分功能复选框样式
        sub_checkbox_style = """
            QCheckBox {
                color: #2c3e50;
                spacing: 5px;
                padding: 5px;
                font-size: 9pt;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
            }
        """

        # 创建细分功能选项
        self.sub_outline_checkbox = QCheckBox("① 大纲生成")
        self.sub_outline_checkbox.setChecked(True)
        self.sub_outline_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_outline_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_outline_checkbox, 0, 0)

        self.sub_numbering_checkbox = QCheckBox("② 自动序号转文本")
        self.sub_numbering_checkbox.setChecked(True)
        self.sub_numbering_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_numbering_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_numbering_checkbox, 0, 1)

        self.sub_image_checkbox = QCheckBox("③ 图片调整")
        self.sub_image_checkbox.setChecked(True)
        self.sub_image_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_image_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_image_checkbox, 1, 0)

        self.sub_table_checkbox = QCheckBox("④ 表格调整")
        self.sub_table_checkbox.setChecked(True)
        self.sub_table_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_table_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_table_checkbox, 1, 1)

        self.sub_keyword_checkbox = QCheckBox("⑤ 关键词标蓝")
        self.sub_keyword_checkbox.setChecked(True)
        self.sub_keyword_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_keyword_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_keyword_checkbox, 2, 0)

        self.sub_symbol_checkbox = QCheckBox("⑥ 星号标红")
        self.sub_symbol_checkbox.setChecked(True)
        self.sub_symbol_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_symbol_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_symbol_checkbox, 2, 1)

        self.sub_header_footer_checkbox = QCheckBox("⑦ 删除页眉页脚")
        self.sub_header_footer_checkbox.setChecked(True)
        self.sub_header_footer_checkbox.setStyleSheet(sub_checkbox_style)
        self.sub_header_footer_checkbox.stateChanged.connect(self.on_sub_checkbox_changed)
        sub_layout.addWidget(self.sub_header_footer_checkbox, 3, 0)
        
        # 智能应答选项
        self.response_checkbox = QCheckBox("🧠 智能应答（自动识别条款、生成标准化应答内容）")
        self.response_checkbox.setFont(QFont("Microsoft YaHei UI", 10))
        self.response_checkbox.setChecked(True)
        self.response_checkbox.setStyleSheet("""
            QCheckBox {
                color: #2c3e50;
                spacing: 8px;
                padding: 10px;
                border-radius: 6px;
                border: 1px solid #c8e6c9;
                background-color: #f8fff8;
            }
            QCheckBox:hover {
                border: 1px solid #4caf50;
                background-color: #f1f8e9;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        function_layout.addWidget(self.response_checkbox)
        self.response_checkbox.stateChanged.connect(self.on_response_checkbox_changed)

        # 产品选择区域（智能应答的子选项）
        self.product_sub_frame = QFrame()
        self.product_sub_frame.setStyleSheet("""
            QFrame {
                background-color: #f0fff0;
                border-radius: 6px;
                border: 1px solid #c8e6c9;
                margin-left: 20px;
                padding: 5px;
            }
        """)
        function_layout.addWidget(self.product_sub_frame)

        product_layout = QGridLayout(self.product_sub_frame)
        product_layout.setSpacing(10)
        product_layout.setContentsMargins(15, 10, 15, 10)

        # 产品选择（合并所有领域的产品到一个列表）
        product_label = QLabel("选择产品：")
        product_label.setFont(QFont("Microsoft YaHei UI", 9))
        product_layout.addWidget(product_label, 0, 0)

        self.product_combo = QComboBox()
        self.product_combo.setFont(QFont("Microsoft YaHei UI", 9))
        # 常规应答为默认选项
        self.product_combo.addItem("常规应答", "")
        # 添加所有产品选项
        for product in PRODUCT_LIST:
            self.product_combo.addItem(product, product)
        self.product_combo.setStyleSheet("""
            QComboBox {
                padding: 5px 10px;
                border: 1px solid #c8e6c9;
                border-radius: 4px;
                background: white;
                min-width: 200px;
            }
            QComboBox:hover {
                border: 1px solid #4caf50;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
        """)
        product_layout.addWidget(self.product_combo, 0, 1)

        # 提示信息
        product_hint = QLabel("💡 选择具体产品将使用对应模板进行精准匹配，常规应答使用通用模板")
        product_hint.setFont(QFont("Microsoft YaHei UI", 8))
        product_hint.setStyleSheet("color: #666; padding-top: 5px;")
        product_layout.addWidget(product_hint, 1, 0, 1, 2)

        
    def create_file_selection_section(self, layout):
        """创建文件选择部分"""
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("未选择文件")
        self.file_label.setFont(QFont("Microsoft YaHei UI", 10))
        self.file_label.setStyleSheet("""
            QLabel { 
                background-color: white;
                padding: 12px;
                border-radius: 6px;
                border: 1px solid #ddd;
                color: #6c757d;
            }
        """)
        file_layout.addWidget(self.file_label)
        
        self.select_btn = QPushButton("📁 选择文件")
        self.select_btn.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
        self.select_btn.setFixedSize(150, 40)
        self.select_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 #4fc3f7, stop:1 #29b6f6);
                color: white;
                border: none;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 #29b6f6, stop:1 #0288d1);
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background: #0277bd;
                transform: translateY(0px);
            }
        """)
        self.select_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.select_btn)
        
        layout.addLayout(file_layout)
        
    def create_progress_section(self, layout):
        """创建进度条部分"""
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(25)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #ddd;
                border-radius: 4px;
                text-align: center;
                font-weight: bold;
                color: #2c3e50;
                background: white;
            }
            QProgressBar::chunk {
                background-color: #28a745;
                border-radius: 2px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
    def create_log_section(self, layout):
        """创建日志显示部分"""
        # 日志组
        log_group = QGroupBox("📋 处理日志")
        log_group.setFont(QFont("Microsoft YaHei UI", 11, QFont.Bold))
        log_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #27ae60;
                border-radius: 10px;
                margin: 8px;
                padding-top: 15px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 #ffffff, stop:1 #f8f9fa);
                color: #2c3e50;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                background: white;
                border-radius: 4px;
                color: #27ae60;
            }
        """)
        layout.addWidget(log_group)
        
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(20, 15, 20, 15)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setMinimumHeight(180)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 #263238, stop:1 #37474f);
                color: #4caf50;
                border: 2px solid #546e7a;
                border-radius: 8px;
                padding: 12px;
                font-family: 'Consolas', monospace;
                selection-background-color: #3498db;
            }
            QScrollBar:vertical {
                background: #455a64;
                width: 12px;
                border-radius: 6px;
                margin: 2px;
            }
            QScrollBar::handle:vertical {
                background: #78909c;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #90a4ae;
            }
        """)
        log_layout.addWidget(self.log_text)
        
    def create_button_section(self, layout):
        """创建按钮部分"""
        button_layout = QHBoxLayout()
        button_layout.setSpacing(30)
        
        # 添加弹性空间
        button_layout.addStretch()
        
        self.process_btn = QPushButton("⚡ 开始处理")
        self.process_btn.setEnabled(False)
        self.process_btn.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        self.process_btn.setFixedSize(160, 50)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                font-size: 16px;
            }
            QPushButton:hover:enabled {
                background-color: #c0392b;
                box-shadow: 0 2px 8px rgba(231, 76, 60, 0.4);
            }
            QPushButton:pressed:enabled {
                background-color: #a93226;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        self.process_btn.clicked.connect(self.process_document)
        button_layout.addWidget(self.process_btn)
        
        self.open_btn = QPushButton("📖 打开结果")
        self.open_btn.setEnabled(False)
        self.open_btn.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        self.open_btn.setFixedSize(160, 50)
        self.open_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                font-size: 16px;
            }
            QPushButton:hover:enabled {
                background-color: #2980b9;
                box-shadow: 0 2px 8px rgba(52, 152, 219, 0.4);
            }
            QPushButton:pressed:enabled {
                background-color: #1f4e79;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)
        self.open_btn.clicked.connect(self.open_result)
        button_layout.addWidget(self.open_btn)
        
        # 添加弹性空间
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
    def on_format_checkbox_changed(self, state):
        """格式调整主选项状态改变时，同步更新所有子选项"""
        is_checked = state == 2  # Qt.Checked 的值为 2
        # 阻止子选项的信号，避免循环触发
        self.sub_outline_checkbox.blockSignals(True)
        self.sub_numbering_checkbox.blockSignals(True)
        self.sub_image_checkbox.blockSignals(True)
        self.sub_table_checkbox.blockSignals(True)
        self.sub_keyword_checkbox.blockSignals(True)
        self.sub_symbol_checkbox.blockSignals(True)
        self.sub_header_footer_checkbox.blockSignals(True)

        self.sub_outline_checkbox.setChecked(is_checked)
        self.sub_numbering_checkbox.setChecked(is_checked)
        self.sub_image_checkbox.setChecked(is_checked)
        self.sub_table_checkbox.setChecked(is_checked)
        self.sub_keyword_checkbox.setChecked(is_checked)
        self.sub_symbol_checkbox.setChecked(is_checked)
        self.sub_header_footer_checkbox.setChecked(is_checked)

        self.sub_outline_checkbox.blockSignals(False)
        self.sub_numbering_checkbox.blockSignals(False)
        self.sub_image_checkbox.blockSignals(False)
        self.sub_table_checkbox.blockSignals(False)
        self.sub_keyword_checkbox.blockSignals(False)
        self.sub_symbol_checkbox.blockSignals(False)
        self.sub_header_footer_checkbox.blockSignals(False)

    def on_sub_checkbox_changed(self):
        """子选项状态改变时，更新主选项状态"""
        all_checked = (
            self.sub_outline_checkbox.isChecked() and
            self.sub_numbering_checkbox.isChecked() and
            self.sub_image_checkbox.isChecked() and
            self.sub_table_checkbox.isChecked() and
            self.sub_keyword_checkbox.isChecked() and
            self.sub_symbol_checkbox.isChecked() and
            self.sub_header_footer_checkbox.isChecked()
        )
        any_checked = (
            self.sub_outline_checkbox.isChecked() or
            self.sub_numbering_checkbox.isChecked() or
            self.sub_image_checkbox.isChecked() or
            self.sub_table_checkbox.isChecked() or
            self.sub_keyword_checkbox.isChecked() or
            self.sub_symbol_checkbox.isChecked() or
            self.sub_header_footer_checkbox.isChecked()
        )

        # 阻止主选项的信号，避免循环触发
        self.format_checkbox.blockSignals(True)
        if all_checked:
            self.format_checkbox.setChecked(True)
        elif not any_checked:
            self.format_checkbox.setChecked(False)
        else:
            # 部分选中时，主选项保持不变或显示为半选状态
            self.format_checkbox.setChecked(True)
        self.format_checkbox.blockSignals(False)

    def on_response_checkbox_changed(self, state):
        """智能应答复选框状态改变时，控制产品选择区域的启用/禁用"""
        is_checked = state == 2  # Qt.Checked 的值为 2
        self.product_sub_frame.setEnabled(is_checked)
        if not is_checked:
            # 取消选择时重置产品选择
            self.product_combo.setCurrentIndex(0)

    def get_selected_product(self):
        """获取当前选择的产品名称，如果没有选择返回None"""
        if not self.response_checkbox.isChecked():
            return None
        product = self.product_combo.currentData()
        return product if product else None

    def get_format_options(self):
        """获取格式调整的细分选项"""
        return {
            'outline': self.sub_outline_checkbox.isChecked(),
            'numbering': self.sub_numbering_checkbox.isChecked(),
            'image': self.sub_image_checkbox.isChecked(),
            'table': self.sub_table_checkbox.isChecked(),
            'keyword': self.sub_keyword_checkbox.isChecked(),
            'symbol': self.sub_symbol_checkbox.isChecked(),
            'header_footer': self.sub_header_footer_checkbox.isChecked()
        }

    def is_any_format_option_selected(self):
        """检查是否选择了任何格式调整选项"""
        options = self.get_format_options()
        return any(options.values())

    def select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx *.doc);;所有文件 (*.*)"
        )
        
        if file_path:
            self.input_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.setText(f"已选择: {filename}")
            self.file_label.setStyleSheet("""
                QLabel { 
                    background-color: #e8f5e8;
                    padding: 12px;
                    border-radius: 6px;
                    border: 1px solid #4caf50;
                    color: #2e7d32;
                    font-weight: bold;
                }
            """)
            self.process_btn.setEnabled(True)
            self.log_text.clear()
            self.log_text.append(f"已选择文件: {file_path}")
            status_bar = self.statusBar()
            if status_bar:
                status_bar.showMessage("已选择文件，可以开始处理")
    
    def get_file_size(self, file_path):
        """获取文件大小的友好显示"""
        try:
            size_bytes = os.path.getsize(file_path)
            if size_bytes < 1024:
                return f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                return f"{size_bytes / 1024:.1f} KB"
            else:
                return f"{size_bytes / (1024 * 1024):.1f} MB"
        except:
            return "未知大小"
            
    def process_document(self):
        """处理文档"""
        if not self.input_file:
            self.show_message("警告", "请先选择文件！", "warning")
            return

        # 检查功能选择
        format_enabled = self.is_any_format_option_selected()
        response_enabled = self.response_checkbox.isChecked()

        if not format_enabled and not response_enabled:
            self.show_message("警告", "请至少选择一个处理功能！", "warning")
            return

        # 获取格式调整细分选项
        format_options = self.get_format_options() if format_enabled else None

        # 确定处理模式
        if format_enabled and response_enabled:
            process_mode = "both"
            mode_text = "格式调整+智能应答"
        elif format_enabled:
            process_mode = "format"
            mode_text = "格式调整"
        else:
            process_mode = "response"
            mode_text = "智能应答"

        # 获取选中的产品
        selected_product = self.get_selected_product()

        # 禁用按钮和选项
        self.process_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.open_btn.setEnabled(False)
        self.format_checkbox.setEnabled(False)
        self.response_checkbox.setEnabled(False)
        self.format_sub_frame.setEnabled(False)
        self.product_sub_frame.setEnabled(False)

        # 重置进度条
        self.progress_bar.setValue(0)

        # 创建处理线程（传递格式选项和选中的产品）
        self.processor = DocumentProcessor(self.input_file, process_mode, format_options, selected_product)
        self.processor.progress.connect(self.update_progress)
        self.processor.log.connect(self.append_log)
        self.processor.finished.connect(self.processing_finished)
        self.processor.error.connect(self.processing_error)

        # 开始处理
        self.processor.start()
        status_bar = self.statusBar()
        if status_bar:
            status_bar.showMessage(f"正在进行{mode_text}处理...")
        
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
        
    def append_log(self, message):
        """添加日志"""
        # 添加时间戳和美化格式
        from datetime import datetime
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        formatted_message = f"{timestamp} {message}"
        self.log_text.append(formatted_message)
        
    def show_message(self, title, message, msg_type="info"):
        """显示自定义样式的消息框"""
        msg = QMessageBox(self)
        
        if msg_type == "warning":
            icon_text = "⚠️"
            button_color = "#f39c12"
            border_color = "#f39c12"
        elif msg_type == "error":
            icon_text = "❌"
            button_color = "#e74c3c"
            border_color = "#e74c3c"
        elif msg_type == "success":
            icon_text = "🌟"
            button_color = "#27ae60"
            border_color = "#27ae60"
        else:
            icon_text = "💡"
            button_color = "#3498db"
            border_color = "#3498db"
        
        # 不设置系统图标，避免显示左侧图标
        msg.setIcon(QMessageBox.NoIcon)
        msg.setWindowTitle(f"{icon_text} {title}")
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.Ok)
        
        # 设置消息框样式
        msg.setStyleSheet(f"""
            QMessageBox {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 #ffffff, stop:1 #f8f9fa);
                font-family: 'Microsoft YaHei UI';
                border: 2px solid {border_color};
                border-radius: 15px;
                min-width: 450px;
                max-width: 500px;
                min-height: 250px;
            }}
            QMessageBox QLabel {{
                color: #2c3e50;
                font-size: 17px;
                font-family: 'Microsoft YaHei UI';
                font-weight: 500;
                padding: 30px 40px 25px 40px;
                margin: 0px;
                min-width: 370px;
                max-width: 420px;
                line-height: 1.8;
                background: transparent;
                border: none;
                qproperty-alignment: 'AlignCenter';
                qproperty-wordWrap: true;
            }}
            QMessageBox QDialogButtonBox {{
                padding: 0px 0px 20px 0px;
                margin: 0px;
                background: transparent;
            }}
            QMessageBox QDialogButtonBox QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 {button_color}, stop:1 {self.darken_color(button_color)});
                color: white;
                border: none;
                border-radius: 10px;
                padding: 12px 30px;
                font-weight: bold;
                font-size: 14px;
                font-family: 'Microsoft YaHei UI';
                min-width: 90px;
                min-height: 40px;
                margin: 0px auto;
            }}
            QMessageBox QDialogButtonBox QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                           stop:0 {self.darken_color(button_color)}, stop:1 {self.darken_color(self.darken_color(button_color))});
                transform: translateY(-2px);
                box-shadow: 0 6px 12px rgba(0,0,0,0.2);
            }}
            QMessageBox QDialogButtonBox QPushButton:pressed {{
                background: {self.darken_color(self.darken_color(button_color))};
                transform: translateY(0px);
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
        """)
        
        # 确保按钮居中
        msg.show()
        
        # 获取按钮并手动居中
        buttons = msg.findChildren(QAbstractButton)
        if buttons:
            button = buttons[0]
            button_box = msg.findChild(QDialogButtonBox)
            if button_box:
                button_box.setCenterButtons(True)
        
        msg.exec_()
    
    def darken_color(self, color):
        """加深颜色"""
        color_map = {
            "#f39c12": "#e67e22",
            "#e74c3c": "#c0392b", 
            "#27ae60": "#229954",
            "#3498db": "#2980b9",
            "#e67e22": "#d35400",
            "#c0392b": "#a93226",
            "#229954": "#1e8449",
            "#2980b9": "#1f4e79"
        }
        return color_map.get(color, color)

    def processing_finished(self, output_file):
        """处理完成"""
        self.output_file = output_file
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        self.format_checkbox.setEnabled(True)
        self.response_checkbox.setEnabled(True)
        self.format_sub_frame.setEnabled(True)
        self.product_sub_frame.setEnabled(self.response_checkbox.isChecked())
        status_bar = self.statusBar()
        if status_bar:
            status_bar.showMessage("处理完成！")
        
        filename = os.path.basename(output_file) if output_file else "未知文件"
        success_message = f"""✨ 文档处理完成！

📄 输出文件：{filename}

🎉 文件已保存至原目录
您可以点击"打开结果"按钮查看处理后的文档

💝 感谢您的使用！"""
        
        self.show_message("处理完成", success_message, "success")
        
    def processing_error(self, error_msg):
        """处理错误"""
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.format_checkbox.setEnabled(True)
        self.response_checkbox.setEnabled(True)
        self.format_sub_frame.setEnabled(True)
        self.product_sub_frame.setEnabled(self.response_checkbox.isChecked())

        status_bar = self.statusBar()
        if status_bar:
            status_bar.showMessage("❌ 处理失败")
        
        # 创建自定义错误消息框
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle("❌ 处理失败")
        msg.setText("💥 文档处理过程中出现错误")
        msg.setInformativeText(f"🔍 错误详情：\n{error_msg}")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setStyleSheet("""
            QMessageBox {
                background-color: white;
                font-family: 'Microsoft YaHei UI';
            }
            QMessageBox QLabel {
                color: #2c3e50;
                font-size: 12px;
            }
            QMessageBox QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QMessageBox QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        msg.exec_()
        
    def open_result(self):
        """打开结果文件"""
        if self.output_file and os.path.exists(self.output_file):
            try:
                os.startfile(self.output_file)
                status_bar = self.statusBar()
                if status_bar:
                    status_bar.showMessage(f"📖 已打开文件: {os.path.basename(self.output_file)}")
            except Exception as e:
                QMessageBox.warning(self, "⚠️ 警告", f"无法打开文件：\n{str(e)}")
        else:
            # 创建自定义警告消息框
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("⚠️ 警告")
            msg.setText("📄 结果文件不存在")
            msg.setInformativeText("请先完成文档处理，然后再尝试打开结果文件。")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                    font-family: 'Microsoft YaHei UI';
                }
                QMessageBox QLabel {
                    color: #2c3e50;
                    font-size: 12px;
                }
                QMessageBox QPushButton {
                    background-color: #f39c12;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QMessageBox QPushButton:hover {
                    background-color: #e67e22;
                }
            """)
            msg.exec_()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格，更现代
    
    # 设置应用程序信息
    app.setApplicationName("RFP智能引擎")
    app.setOrganizationName("Tuke")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
