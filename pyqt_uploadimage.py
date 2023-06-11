import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QFileDialog
from PyQt6.QtGui import QPixmap


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("图片上传应用")

        # 创建一个标签用于显示图片
        self.image_label = QLabel(self)
        self.image_label.setFixedSize(400, 400)

        # 创建一个按钮
        self.upload_button = QPushButton("上传图片", self)
        self.upload_button.clicked.connect(self.upload_image)

        # 创建一个垂直布局，并将标签和按钮添加到其中
        layout = QVBoxLayout()
        layout.addWidget(self.image_label)
        layout.addWidget(self.upload_button)

        # 创建一个主部件，将布局设置为主部件的布局
        main_widget = QWidget()
        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

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
            self.image_label.setPixmap(pixmap.scaledToHeight(1000, mode=Qt.TransformationMode.SmoothTransformation))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
