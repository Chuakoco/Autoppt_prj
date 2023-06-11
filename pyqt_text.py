# file: review.py
#!/usr/bin/python

import sys
from PyQt6.QtWidgets import (QWidget, QLabel, QLineEdit, QSystemTrayIcon, QMessageBox, QMainWindow,
        QTextEdit, QGridLayout, QApplication, QPushButton, QFileDialog)
from PyQt6 import QtGui
from PyQt6.QtGui import QPixmap, QClipboard, QIcon
from PyQt6.QtCore import Qt, QTimer


class ContentPage(QWidget):

    def __init__(self):
        super().__init__()
        self.page_element_dict = {
            "Label_ChapterTitle": QLabel('标题'),
            "Label_ImageTitle": QLabel("图片"),
            "Label_TextTitle": QLabel("文本"),
            "Label_GenerateCode": QLabel("生成代码"),
            "Label_OutputCode": QLabel("输出代码"),
            "Block_ChapterTitle": QLineEdit(self),
            "Block_Image": QLabel(self),
            "Block_Text": QTextEdit(self),
            "Block_OutputCode": QTextEdit(self),
            "Btn_OpenImage": QPushButton(self),
            "Btn_GenerateCode": QPushButton(self),
            "Btn_CopyCode": QPushButton(self),
        }
        self.initUI()

    def upload_image(self):
        # 打开文件对话框以选择图片文件
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择图片")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
        file_dialog.setNameFilter("Images (*.png *.xpm *.jpg *.jpeg *.bmp)")

        if file_dialog.exec() == QFileDialog.DialogCode.Accepted:
            selected_files = file_dialog.selectedFiles()
            file_path = selected_files[0]

            # 显示选中的图片
            pixmap = QPixmap(file_path)
            self.page_element_dict["Block_Image"].setPixmap(pixmap.scaledToHeight(200, mode=Qt.TransformationMode.SmoothTransformation))

    def copy_text(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.page_element_dict["Block_OutputCode"].toPlainText())

        # Show notification
        QMessageBox.information(self, "Copy Succeeded", "Text copied to clipboard", QMessageBox.StandardButton.Ok,
                                QMessageBox.StandardButton.Ok)

    def generate_code_method(self):
        # plain_text = self.content_block.toPlainText()
        # self.output_code_block.setPlainText(plain_text)

        text = self.page_element_dict["Block_Text"].toPlainText()
        self.page_element_dict["Block_OutputCode"].setPlainText(text)
        # self.text_edit2.setPlainText(text)

    def manage_page_element(self):
        # open image button
        image_upload_btn = self.page_element_dict["Btn_OpenImage"]
        image_upload_btn.setText("打开图片")
        image_upload_btn.clicked.connect(self.upload_image)

        # generate code button
        generate_code_btn = self.page_element_dict["Btn_GenerateCode"]
        generate_code_btn.setText("生成代码")
        generate_code_btn.clicked.connect(self.generate_code_method)

        # copy generated code to clipboard
        copy_button = self.page_element_dict["Btn_CopyCode"]
        copy_button.setText("复制代码")
        copy_button.clicked.connect(self.copy_text)

    def manage_page_layout(self):
        # Layout
        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(self.page_element_dict["Label_ChapterTitle"], 1, 0)
        grid.addWidget(self.page_element_dict["Block_ChapterTitle"], 1, 1)

        grid.addWidget(self.page_element_dict["Label_ImageTitle"], 2, 0)
        grid.addWidget(self.page_element_dict["Block_Image"], 2, 1)
        grid.addWidget(self.page_element_dict["Btn_OpenImage"], 3, 1)

        grid.addWidget(self.page_element_dict["Label_TextTitle"], 4, 0)
        grid.addWidget(self.page_element_dict["Block_Text"], 4, 1)

        grid.addWidget(self.page_element_dict["Label_GenerateCode"], 5, 0)
        grid.addWidget(self.page_element_dict["Btn_GenerateCode"], 5, 1)

        grid.addWidget(self.page_element_dict["Label_OutputCode"], 6, 0)
        grid.addWidget(self.page_element_dict["Btn_CopyCode"], 6, 1)
        grid.addWidget(self.page_element_dict["Block_OutputCode"], 7, 1)

        self.setLayout(grid)
        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('Review')
        self.show()

    def initUI(self):
        self.manage_page_element()
        self.manage_page_layout()

def main():
    app = QApplication(sys.argv)
    ex = ContentPage()
    ex.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()