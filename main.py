import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTextEdit, QPushButton


class NotebookWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Notebook")

        # Create text edit widgets
        self.text_edit1 = QTextEdit(self)
        self.text_edit2 = QTextEdit(self)

        # Create a button
        self.button = QPushButton("Copy Text", self)
        self.button.clicked.connect(self.copy_text)

        # Create a layout and add the widgets
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit1)
        layout.addWidget(self.button)
        layout.addWidget(self.text_edit2)

        # Create a main widget and set the layout
        main_widget = QWidget()
        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

    def copy_text(self):
        # Get the text from the first text edit and set it in the second text edit
        text = self.text_edit1.toPlainText()
        self.text_edit2.setPlainText(text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = NotebookWindow()
    window.show()
    sys.exit(app.exec())
