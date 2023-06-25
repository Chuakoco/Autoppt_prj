import sys
import os
import shutil
from PyQt6.QtWidgets import (QWidget, QLabel, QLineEdit, QMessageBox,
                             QTextEdit, QGridLayout, QApplication, QPushButton,
                             QFileDialog, QHBoxLayout, QListWidget, QListWidgetItem)
from PyQt6.QtGui import QPixmap, QGuiApplication
from PyQt6.QtCore import Qt

from ppt_template import VBACodeGenerator


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.create_cache()
        self.setWindowTitle("Auto PPT")
        self.setWindowTitle("Autoppt")
        self.page_elements = {
            "L_Label_Navi": QLabel("当前页数"),
            "L_Block_Navi": QListWidget(self),
            "L_Btn_addPage": QPushButton("新建页面"),
            "L_Btn_removePage": QPushButton("删除页面"),
            "M_Label_ChapterTitle": QLabel('标题'),
            "M_Label_ImageTitle": QLabel("图片"),
            "M_Label_TextTitle": QLabel("文本"),
            "M_Label_OutputCode": QLabel("输出代码"),
            "M_Block_ChapterTitle": QLineEdit(self),
            "M_Block_Image": QLabel(self),
            "M_Block_Text": QTextEdit(self),
            "R_Block_OutputCode": QTextEdit(self),
            "M_Btn_OpenImage": QPushButton("上传图片"),
            "R_Btn_GenerateCode": QPushButton("生成代码"),
            "R_Btn_CopyCode": QPushButton("复制代码"),
            "M_Btn_save": QPushButton("保存"),
            "Path_image": "",
        }
        self.current_page = 0
        self.pages = {} #1: {'Chapter_Name': '', 'Image_Path': '', 'Content_Text': ''}
        self.initUI()

    def create_cache(self):
        if not os.path.exists('./temp/image'):
            os.mkdir('./temp/image')

    def clear_cache(self):
        for img in os.listdir('./temp/image'):
            if img.endswith('.jpg'):
                os.remove(img)

    def remove_cache(self):
        if os.path.exists('./temp/image'):
            shutil.rmtree('./temp/image')

    def closeEvent(self, event):
        # Show a confirmation dialog
        reply = QMessageBox.question(
            self, "Confirmation", "Are you sure you want to exit?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        # If the user confirms, close the window
        if reply == QMessageBox.StandardButton.Yes:
            self.remove_cache()
            event.accept()
        else:
            event.ignore()

    def manage_page_element(self):
        # connect change page action
        self.page_elements["L_Block_Navi"].addItem(QListWidgetItem(f"Page {1}. "))
        self.pages[0] = {'Chapter_Name': '', 'Image_Path': '', 'Content_Text': ''}
        self.page_elements["L_Block_Navi"].currentRowChanged.connect(self.action_changePage)

        # connect action to add page button
        self.page_elements["L_Btn_addPage"].clicked.connect(self.action_addPage)

        # connect action to remove page button
        self.page_elements["L_Btn_removePage"].clicked.connect(self.action_delPage)

        # connect save action to save button
        self.page_elements["M_Btn_save"].clicked.connect(self.action_savePage)

        # connect upload image button
        self.page_elements["M_Btn_OpenImage"].clicked.connect(self.upload_image)

        # connect to generate code
        self.page_elements["R_Btn_GenerateCode"].clicked.connect(self.action_generateCode)

        # connect to copy code
        self.page_elements["R_Btn_CopyCode"].clicked.connect(self.action_copyText)

    def manage_page_layout(self):
        # Layout
        grid = QGridLayout()
        grid.setSpacing(10)

        # left layout
        left_layout = QGridLayout()
        left_layout.addWidget(self.page_elements["L_Label_Navi"], 1, 0, 1, -1)
        left_layout.addWidget(self.page_elements["L_Block_Navi"], 2, 0, 1, -1)
        left_layout.addWidget(self.page_elements["L_Btn_addPage"], 3, 0, 1, 1)
        left_layout.addWidget(self.page_elements["L_Btn_removePage"], 3, 1, 1, 1)

        # middle layout
        mid_layout = QGridLayout()
        mid_layout.setSpacing(10)
        mid_layout.addWidget(self.page_elements["M_Label_ChapterTitle"], 1, 0)
        mid_layout.addWidget(self.page_elements["M_Block_ChapterTitle"], 1, 1)
        mid_layout.addWidget(self.page_elements["M_Label_ImageTitle"], 2, 0)
        mid_layout.addWidget(self.page_elements["M_Block_Image"], 2, 1)
        mid_layout.addWidget(self.page_elements["M_Btn_OpenImage"], 3, 1)
        mid_layout.addWidget(self.page_elements["M_Label_TextTitle"], 4, 0)
        mid_layout.addWidget(self.page_elements["M_Block_Text"], 4, 1)
        mid_layout.addWidget(self.page_elements["M_Btn_save"], 5, 0, 1, -1)

        # right layout
        right_layout = QGridLayout()
        right_layout.addWidget(self.page_elements["R_Btn_GenerateCode"], 1, 0)
        right_layout.addWidget(self.page_elements["R_Btn_CopyCode"], 1, 1)
        right_layout.addWidget(self.page_elements["R_Block_OutputCode"], 2, 0, -1, 2)

        layout = QHBoxLayout()
        layout.addLayout(left_layout)
        layout.addLayout(mid_layout)
        layout.addLayout(right_layout)


        self.setLayout(layout)
        self.setGeometry(200, 200, 1200, 500)
        # Center the window on the screen
        screen_geometry = QGuiApplication.screens()[0].geometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

        self.setWindowTitle("Notepad App")
        self.show()

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
            self.page_elements["Path_image"] = file_path

            # 显示选中的图片
            pixmap = QPixmap(file_path)
            self.page_elements["M_Block_Image"].setPixmap(pixmap.scaledToHeight(
                200, mode=Qt.TransformationMode.SmoothTransformation
            ))

    def action_addPage(self):
        # detect current page
        current_row = self.page_elements["L_Block_Navi"].currentRow()

        # reorder page elements
        for page_idx in range(len(self.pages)-1, current_row, -1):
            self.pages[page_idx+1] = self.pages[page_idx]
        self.pages[current_row+1] = {'Chapter_Name': '', 'Image_Path': '', 'Content_Text': ''}

        # insert item to navigation list
        page_item = QListWidgetItem(f"Page {current_row + 2}. ")
        self.page_elements["L_Block_Navi"].insertItem(current_row+1, page_item)

        # change item name in navigation list name
        for page_item in range(current_row+1, self.page_elements["L_Block_Navi"].count()):
            self.page_elements["L_Block_Navi"].item(page_item).setText(
                f"Page {page_item + 1}. {self.pages[page_item]['Chapter_Name']}"
            )

        # select to the new added item
        new_item = self.page_elements["L_Block_Navi"].item(current_row+1)
        if new_item:
            self.page_elements["L_Block_Navi"].setCurrentItem(new_item)



    def action_delPage(self):
        if self.page_elements["L_Block_Navi"].count() <= 1:
            return
        current_row = self.page_elements["L_Block_Navi"].currentRow()

        # rename page index
        for page_item in range(current_row, self.page_elements["L_Block_Navi"].count()):
            self.page_elements["L_Block_Navi"].item(page_item).setText(
                f"Page {page_item}. {self.pages[page_item]['Chapter_Name']}"
            )

        # relink page contents
        for page_idx in range(current_row, self.page_elements["L_Block_Navi"].count()-1):
            self.pages[page_idx] = self.pages[page_idx + 1]
        # self.pages.pop(self.page_elements["L_Block_Navi"].count()-1)

        self.page_elements["L_Block_Navi"].takeItem(current_row)  # element is dropped from pages

        # select to the item above the deleted item
        new_item = self.page_elements["L_Block_Navi"].item(current_row)
        if new_item:
            self.page_elements["L_Block_Navi"].setCurrentItem(new_item)

    def action_savePage(self):
        current_row = self.page_elements["L_Block_Navi"].currentRow()
        self.page_elements["L_Block_Navi"].item(current_row).setText(
            f"Page {current_row + 1}. {self.page_elements['M_Block_ChapterTitle'].text()}"
        )
        self.pages[current_row] = {
            'Chapter_Name': self.page_elements["M_Block_ChapterTitle"].text(),
            'Image_Path': self.page_elements["Path_image"],
            'Content_Text': self.page_elements["M_Block_Text"].toPlainText()
        }

    def action_changePage(self, index):
        self.page_elements["M_Block_ChapterTitle"].setText(self.pages[index]["Chapter_Name"])
        image_path = self.pages[index]["Image_Path"]
        if image_path:
            self.page_elements["M_Block_Image"].setPixmap(
                QPixmap(image_path).scaledToHeight(200, mode=Qt.TransformationMode.SmoothTransformation)
            )
        else:
            self.page_elements["M_Block_Image"].setPixmap(QPixmap())
        self.page_elements["M_Block_Text"].setPlainText(self.pages[index]["Content_Text"])

    def action_generateCode(self):
        # save all pages
        self.action_savePage()

        # use ppt template to generate code
        output_code = VBACodeGenerator().concat_code(content_dict=self.pages)
        self.page_elements["R_Block_OutputCode"].setPlainText(output_code)

    def action_copyText(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.page_elements["R_Block_OutputCode"].toPlainText())

        # Show notification
        QMessageBox.information(self, "Copy Succeeded", "Text copied to clipboard", QMessageBox.StandardButton.Ok,
                                QMessageBox.StandardButton.Ok)


    def initUI(self):
        self.manage_page_element()
        self.manage_page_layout()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


