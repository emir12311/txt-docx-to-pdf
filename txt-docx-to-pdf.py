import sys, os, subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QComboBox, QLineEdit,
    QPushButton, QFileDialog, QHBoxLayout, QVBoxLayout, QMessageBox
)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Converter")
        self.setGeometry(960, 540, 400, 200)
        self.file_name = ""
        self.setup_ui()

    def setup_ui(self):
        self.label = QLabel("Please pick your file type to convert.")

        self.combobox = QComboBox()
        self.combobox.addItems(["Docx", "Txt"])
        self.combobox.setCurrentIndex(-1)

        self.label2 = QLabel("Pick your file:")
        self.line_edit = QLineEdit()
        self.browse = QPushButton("Browse")

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.line_edit)
        file_layout.addWidget(self.browse)

        self.convert = QPushButton("Convert To PDF")

        self.browse.clicked.connect(self.filepicker)
        self.convert.clicked.connect(self.convertfile)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.combobox)
        layout.addWidget(self.label2)
        layout.addLayout(file_layout)
        layout.addWidget(self.convert)

        self.setLayout(layout)

    def filepicker(self):
        current_type = self.combobox.currentText()
        if current_type == "Txt":
            file_name, _ = QFileDialog.getOpenFileName(self, "Select a File", "", "Text Files (*.txt)")
        elif current_type == "Docx":
            file_name, _ = QFileDialog.getOpenFileName(self, "Select a File", "", "Document Files (*.docx)")
        else:
            self.warningcb()
            return

        if file_name:
            self.file_name = file_name
            self.line_edit.setText(file_name)

    def warningcb(self):
        mb = QMessageBox()
        mb.setWindowTitle("Warning")
        mb.setText("Please select a valid file type.")
        mb.setStandardButtons(QMessageBox.Ok)
        mb.exec_()

    def convertfile(self):
        if not self.file_name:
            self.warningcb()
            return

        name, ext = os.path.splitext(self.file_name)
        ext = ext.lower().replace(".", "")
        output_dir = os.path.dirname(self.file_name)

        if ext in ["txt", "docx"]:
            try:
                subprocess.run([
                    "soffice", "--headless", "--convert-to", "pdf",
                    self.file_name, "--outdir", output_dir
                ], check=True)
                self.show_success()
            except FileNotFoundError:
                self.show_error("LibreOffice not found. Make sure 'soffice' is in your PATH.")
            except subprocess.CalledProcessError as e:
                self.show_error(f"{e}")
        else:
            self.warningcb()

    def show_error(self, message):
        mb = QMessageBox()
        mb.setWindowTitle("Error")
        mb.setText(f"Conversion failed:\n{message}")
        mb.setStandardButtons(QMessageBox.Ok)
        mb.exec_()

    def show_success(self, message="Conversion finished successfully!"):
        mb = QMessageBox()
        mb.setWindowTitle("Success")
        mb.setText(message)
        mb.setStandardButtons(QMessageBox.Ok)
        mb.exec_()


app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())
