#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2023/6/27 20:02
# @Author  : xuhui
# @FileName: main.py
# @Software: PyCharm
import codecs
import datetime
import os
import re
import sys

import qtawesome as qta
from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication, QLineEdit, QPushButton, QFileDialog, QTextEdit, QVBoxLayout, \
    QMessageBox, QTabWidget, QMainWindow, QAction, QDialog
from PyQt5.QtWidgets import QLabel, QWidget


from ico import icon
from pcie.WriteToExcel import write_to_excell as wte
from pcie.compare import compare_data as cd
from preset.WriteToExcelPre import write_to_excel_pre as wtep
from preset.compare_data import compare_data_pre as cdp
from qss.qss import qss
from xhmi import ReportGenerate as rg
from xhmi.compare_xhmi import compare_data_xhmi as cdx
from xhmi.write_to_excel import write_to_excel_xhmi as wtex


def create_button_with_icon(icon_name, active_icon_name, color='white', active_color='blue'):
    button = QPushButton()
    icon = qta.icon(icon_name, active=active_icon_name, color=color, color_active=active_color)
    button.setIcon(icon)
    button.setIconSize(QSize(26, 26))
    return button


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.last_path = ''  # 添加记录上次使用路径的属性
        self.initUI()
        # 使用资源文件中的ICON
        self.setWindowIcon(QtGui.QIcon(':/icon.ico'))
        self.show()

    def initUI(self):
        self.setWindowTitle('Margin Tool')
        self.setGeometry(555, 215, 900, 700)

        self.tab_widget = QTabWidget()
        self.create_pcie_tab()
        self.create_xhmi_tab()
        self.create_preset_tab()

        hbox = QVBoxLayout()
        hbox.addWidget(self.tab_widget)
        centralWidget = QWidget()
        centralWidget.setLayout(hbox)
        self.setCentralWidget(centralWidget)

        image_label = QLabel(self)
        pixmap = QPixmap('pic/Suma.png')
        pixmap = pixmap.scaled(125, 30)
        image_label.setPixmap(pixmap)

        vbox = QVBoxLayout()
        vbox.addWidget(image_label, alignment=Qt.AlignBottom | Qt.AlignHCenter)

        corner_widget = QWidget(self)
        corner_widget.setLayout(vbox)
        hbox.addWidget(corner_widget)

        menubar = self.menuBar()
        self.setMenuBar(menubar)

        help_menu = menubar.addMenu('Help')
        open_readme_action = QAction('Open README', self)
        open_readme_action.triggered.connect(self.open_readme)
        help_menu.addAction(open_readme_action)

        self.readme_dialog = QDialog()  # 创建对话框，默认隐藏
        self.readme_dialog.setWindowTitle('README')
        self.readme_dialog.setGeometry(150, 150, 800, 800)
        self.readme_text = QTextEdit(self.readme_dialog)
        self.readme_text.setReadOnly(True)
        self.readme_text.setGeometry(0, 0, 800, 800)

    def open_readme(self):
        try:
            with codecs.open('README.html', 'r', 'utf-8') as readme_file:
                readme_text = readme_file.read()
                self.readme_text.setHtml(readme_text)
        except FileNotFoundError:
            print("README file not found")

        self.readme_dialog.show()

    def create_pcie_tab(self):
        # 创建 PCIe 标签页并添加控件
        self.tab_pci = QWidget()
        path_label = QLabel('请选择PCIe Margin文件夹路径', self.tab_pci)
        self.path_edit = QLineEdit(self.tab_pci)
        choose_button = create_button_with_icon('fa5s.folder-open', 'fa5s.folder-open', color='white',
                                                active_color='blue')
        choose_button.setText('选择文件夹')
        choose_button.clicked.connect(self.choose_folder)
        self.log_edit = QTextEdit(self.tab_pci)
        self.log_edit.setReadOnly(True)
        process_button = create_button_with_icon('fa.hourglass-start', 'fa.hourglass-start', color='white',
                                                 active_color='blue')
        process_button.setText('处理')
        process_button.clicked.connect(self.process_files)

        vbox = QVBoxLayout()
        vbox.addWidget(path_label)
        vbox.addWidget(self.path_edit)
        vbox.addWidget(choose_button)
        vbox.addWidget(self.log_edit)
        vbox.addWidget(process_button)

        self.tab_pci.setLayout(vbox)

        # 设置标签页水平排列
        self.tab_widget.setTabShape(QTabWidget.Triangular)
        # 添加到主窗口中
        self.tab_widget.addTab(self.tab_pci, "PCIe Margin")

        self.tab_widget.setTabIcon(self.tab_widget.indexOf(self.tab_pci),
                                   qta.icon('mdi.expansion-card', color='white'))

    def create_preset_tab(self):
        # 创建 Preset 标签页并添加控件
        self.tab_pre = QWidget()
        path_label = QLabel('请选择PCIe Preset文件夹路径', self.tab_pre)
        self.path_edit_pre = QLineEdit(self.tab_pre)
        choose_button = create_button_with_icon('fa5s.folder-open', 'fa5s.folder-open', color='white',
                                                active_color='blue')
        choose_button.setText('选择文件夹')
        choose_button.clicked.connect(self.choose_folder_pre)
        self.log_edit_pre = QTextEdit(self.tab_pre)
        self.log_edit_pre.setReadOnly(True)
        process_button = create_button_with_icon('fa.hourglass-start', 'fa.hourglass-start', color='white',
                                                 active_color='blue')
        process_button.setText('写入所有Preset到Excel')
        process_button.clicked.connect(self.process_files_pre)

        vbox = QVBoxLayout()
        vbox.addWidget(path_label)
        vbox.addWidget(self.path_edit_pre)
        vbox.addWidget(choose_button)
        vbox.addWidget(self.log_edit_pre)
        vbox.addWidget(process_button)

        self.tab_pre.setLayout(vbox)

        # 添加到主窗口中
        self.tab_widget.addTab(self.tab_pre, "Preset Scan")
        self.tab_widget.setTabIcon(self.tab_widget.indexOf(self.tab_pre),
                                   qta.icon('mdi.table-settings', color='white'))

    def create_xhmi_tab(self):
        # 创建 XHMI 标签页并添加控件
        # xhmi_tab = QWidget()
        self.xhmi_tab = QWidget()
        path_label = QLabel('请选择xHMI文件夹路径', self.xhmi_tab)
        self.path_edit_xhmi = QLineEdit(self.xhmi_tab)
        self.path_edit_xhmi.setReadOnly(True)
        choose_button = create_button_with_icon('fa5s.folder-open', 'fa5s.folder-open', color='white',
                                                active_color='blue')
        choose_button.setText('选择文件夹')
        choose_button.clicked.connect(lambda: self.choose_folder_xhmi(self.last_path))  # 修改 4：传入初始路径参数
        self.log_edit_xhmi = QTextEdit(self.xhmi_tab)
        self.log_edit_xhmi.setReadOnly(True)
        process_button = create_button_with_icon('fa.hourglass-start', 'fa.hourglass-start', color='white',
                                                 active_color='blue')
        process_button.setText('处理')
        process_button.clicked.connect(self.process_files_xhmi)
        report_button = create_button_with_icon('fa5s.file-excel', 'fa5s.file-excel', color='white',
                                                active_color='blue')
        report_button.setText('生成报告')
        report_button.clicked.connect(self.generate_report)

        vbox = QVBoxLayout()
        vbox.addWidget(path_label)
        vbox.addWidget(self.path_edit_xhmi)
        vbox.addWidget(choose_button)
        vbox.addWidget(self.log_edit_xhmi)
        vbox.addWidget(process_button)
        vbox.addWidget(report_button)

        self.xhmi_tab.setLayout(vbox)

        # 添加到主窗口中
        self.tab_widget.addTab(self.xhmi_tab, "xHMI Margin")
        self.tab_widget.setTabIcon(self.tab_widget.indexOf(self.xhmi_tab),
                                   qta.icon('mdi.cpu-64-bit', color='white'))
        self.tab_widget.setIconSize(QtCore.QSize(35, 35))

    def choose_folder(self):
        # 使用上次打开的路径作为默认路径
        folder_path = QFileDialog.getExistingDirectory(self, '选择文件夹', self.last_path)
        if folder_path:
            self.path_edit.setText(folder_path)
            self.last_path = folder_path  # 记录本次使用的路径

    def choose_folder_pre(self):
        # 使用上次打开的路径作为默认路径
        folder_path = QFileDialog.getExistingDirectory(self, '选择文件夹', self.last_path)
        if folder_path:
            self.path_edit_pre.setText(folder_path)
            self.last_path = folder_path  # 记录本次使用的路径

    def log(self, text):
        self.log_edit.append(text)
        self.log_edit_xhmi.append(text)
        self.log_edit_pre.append(text)

    def process_files(self):
        # 清空日志框
        self.log_edit.clear()

        # 获取文件夹路径
        folder_path = self.path_edit.text().strip()
        if not folder_path:
            self.log('文件夹路径不能为空！')
            return

        # 获取符合条件的txt文件名
        txt_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.txt'):
                    if re.match(r'^D[01234567]T[01234567]P[01234567]L[01234567]\.txt$', file) or re.match(
                            r'^S[01]D0T[01234567]P[01234567]L[01234567]\.txt$', file):
                        txt_files.append(os.path.join(root, file))
        if len(txt_files) == 0:
            QMessageBox.warning(self, '警告', f'未找到符合条件的txt文件！')
            return
        else:
            # 获取当前日期时间
            now = datetime.datetime.now()
            date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")

            # 创建文件路径和名称
            txt_file_path = os.path.join(folder_path, f"result_{date_time_str}.txt")
            excel_file_path = os.path.join(folder_path, f"result_{date_time_str}.xlsx")

            # 创建一个txt文件，将符合条件的文件名写进去，方便后面提取
            file_list = []
            with open(txt_file_path, 'w') as file:
                for item in txt_files[:]:
                    match = re.search(r'^(.*\\1\\)(S[01]D0T[01234567]P[01234567]L[01234567])(.*)$', item) or re.search(
                        r'^(.*\\1\\)(D[01234567]T[01234567]P[01234567]L[01234567])(.*)$', item)
                    if match:
                        file_list.append(match.group(2).strip())

            # 比较并打印数据
            cd(txt_files, txt_file_path, file_list, self.log)

            # 读取txt文件中的数据，并将其写入Excel文件
            data = []
            with open(txt_file_path, 'r') as f:
                for line in f.readlines():
                    data.append(line.strip().split(','))
            wte(data, excel_file_path)
            self.log(f'已导入到Excel文件：{excel_file_path}')

    def process_files_pre(self):
        # 清空日志框
        self.log_edit_pre.clear()

        # 获取文件夹路径
        folder_path = self.path_edit_pre.text().strip()
        if not folder_path:
            self.log_edit_pre.append('文件夹路径不能为空！')
            return

        # 获取符合条件的txt文件名
        txt_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.txt'):
                    if re.match(r'^D[01234567]T[01234567]P[01234567]L[01234567]\.txt$', file) or re.match(
                            r'^S[01]D0T[01234567]P[01234567]L[01234567]\.txt$', file):
                        txt_files.append(os.path.join(root, file))
        if len(txt_files) == 0:
            QMessageBox.warning(self, '警告', f'未找到符合条件的txt文件！')
        else:
            # 获取当前日期时间
            now = datetime.datetime.now()
            date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")

            # 创建文件路径和名称
            file_path = os.path.join(folder_path, f"Preset_Resut_{date_time_str}.txt")
            excel_path = os.path.join(folder_path, f"Preset_Result_{date_time_str}.xlsx")

            file_list = []
            with open(file_path, 'w') as file:
                for item in txt_files[:]:
                    match = re.search(r'^(.*\\1\\)(S[01]D0T[01234567]P[01234567]L[01234567])(.*)$', item) or re.search(
                        r'^(.*\\1\\)(D[01234567]T[01234567]P[01234567]L[01234567])(.*)$', item)
                    if match:
                        file_list.append(match.group(2).strip())

            # 调用 compare_data 函数，并将 file_list 作为参数传递进去
            cdp(txt_files, file_path, self.log)
            # 读取txt文件中的数据，并将其写入Excel文件
            data = []
            with open(file_path, 'r') as f:
                for line in f.readlines():
                    data.append(line.strip().split(','))
            # 将数据写入 Excel 文件
            wtep(data, excel_path)
            self.log_edit_pre.append(f'已导入到Excel文件：{excel_path}')

    def choose_folder_xhmi(self, init_path=''):
        # 使用上次打开的路径作为默认路径
        folder_path = QFileDialog.getExistingDirectory(self, '选择文件夹', init_path)
        if folder_path:
            self.path_edit_xhmi.setText(folder_path)
            self.last_path_xhmi = folder_path  # 记录本次使用的路径
            # 获取该目录下的所有子目录，作为处理的文件夹列表
            self.folder_paths_xhmi = [os.path.join(folder_path, name) for name in os.listdir(folder_path)
                                      if os.path.isdir(os.path.join(folder_path, name))]

    def process_files_xhmi(self):
        # 清空日志框
        self.log_edit_xhmi.clear()

        # 获取文件夹路径
        folder_path = self.path_edit_xhmi.text().strip()
        if not folder_path:
            self.log('文件夹路径不能为空！')
            return

        # 如果self.folder_paths为空，则说明用户没有手动指定需要处理的文件夹，此时将默认将整个文件夹进行处理
        if len(self.folder_paths_xhmi) == 0:
            self.folder_paths_xhmi.append(folder_path)

        # 定义保存文件路径和名称
        file_paths = []
        excel_file_paths = []

        # 遍历每个文件夹，对其中的文件进行处理
        for path in self.folder_paths_xhmi:

            # 获取符合条件的txt文件名
            txt_files = []
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.endswith('.txt'):
                        if re.match(r'^S[01]D0T[0123]P[0123]L[0123]\.txt$', file):
                            txt_files.append(os.path.join(root, file))
            if len(txt_files) == 0:
                QMessageBox.warning(self, '警告', f'未在文件夹 {path} 中找到符合条件的txt文件！')
            else:
                # 获取当前日期时间
                now = datetime.datetime.now()
                date_time_str = now.strftime("%Y-%m-%d_%H-%M-%S")

                # 创建文件路径和名称
                file_path = os.path.join(path, f"Min_result_{date_time_str}.txt")
                excel_file_path = os.path.join(path, f"Min_result_{date_time_str}.xlsx")

                # 创建一个txt文件，将符合条件的文件名写进去，方便后面提取
                file_list = []

                with open(file_path, 'w') as file:
                    for item in txt_files[:16]:
                        match = re.search(r'^(.*)(S[01]D0T[0123]P[0123]L[0123])(.*)$', item)
                        if match:
                            file_list.append(match.group(2).strip())

                # 比较并打印数据
                cdx(txt_files, file_path, file_list, self.log)

                # 保存文件路径和名称到列表
                file_paths.append(file_path)
                excel_file_paths.append(excel_file_path)

        # 循环完成后读取所有符合条件的文件名数据，并将其写入Excel文件
        for i in range(len(file_paths)):
            data = []
            with open(file_paths[i], 'r') as f:
                for line in f.readlines():
                    data.append(line.strip().split(','))
            wtex(data, excel_file_paths[i])
            self.log(f'已导入到Excel文件：{excel_file_paths[i]}')

        # 处理完毕弹出提示框
        QMessageBox.information(self, '处理完毕', '已经处理完毕！')

    def generate_report(self):
        # 获取文件夹路径和处理结果信息
        folder_path = self.path_edit_xhmi.text()

        if not folder_path:
            QMessageBox.warning(self, '警告', '请先选择文件夹路径！')
            return

        # 调用 ReportGenerate 方法生成报告
        rg.Report(folder_path)

        self.log(f'已生成Report：{folder_path}')
        # 处理完毕弹出提示框
        QMessageBox.information(self, '处理完毕', '报告已生成！')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    icon.qInitResources()
    Qss = qss
    app.setStyleSheet(Qss)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
