import sys
from PyQt6.QtWidgets import QApplication, QLabel, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QSizePolicy, QCheckBox, QComboBox, QSpacerItem, QMessageBox
from PyQt6.QtCore import Qt
from functools import partial
from modules.msg_reader import MsgReader
from modules.xls_manipulator import XlsProcessor
import os
from datetime import date
import time
from typing import Literal

class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        
        # configuring window
        self.setWindowTitle("PyQt6 Example")
        self.setGeometry(100, 100, 415, 500)
        
        # creating modul objects
        self.file_dir = os.path.expanduser(r'~/Desktop/emails')

        if not os.path.exists(self.file_dir):
            os.mkdir(self.file_dir)
            QMessageBox.information(self, 'Info!', 'A levelek számára létrehoztam az "emails" könyvtárat az asztalon.'
                                                   '\nIde másolhatod a feldogozni kívánt leveleket.')
        elif not os.listdir(self.file_dir):
            QMessageBox.information(self, 'Info!', 'A leveleket tartalmazó "emails" könyvtár üres.')
        
        xls = r'path/to/file'
        self.reader = MsgReader(file_dir=self.file_dir)
        self.xls_processor = XlsProcessor(file_path=xls)

        
        # creating layout
        main_layout = QVBoxLayout()
        content_layout = QHBoxLayout()
        self.left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()
        
        # left side's content
        #label = QLabel("Main Label")
        #self.left_layout.addWidget(label)
        
        #stock_text = QLabel("Stock Text")
        #self.left_layout.addWidget(stock_text)

        # adding spacer to push content to bottom
        spacer = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.left_layout.addItem(spacer)

        # creating inputs' field
        self.create_input_fields(
            fields=[
                {"label": "Approver", "type": "text", "name": "approver_input"},
                {"label": "Accepted", "type": "checkbox", "name": "accepted_checkbox"},
            ]
        )
        
        # buttons
        btn_layout = QHBoxLayout()
        button_1 = QPushButton("Submit")
        button_2 = QPushButton("Exit")
        button_1.clicked.connect(self.on_submit)  # connecting "Submit" btn with its method
        button_2.clicked.connect(self.close)  # connecting "Cancel" btn with its method
        btn_layout.addWidget(button_1)
        btn_layout.addWidget(button_2)
        self.left_layout.addLayout(btn_layout)
        
        # right side's content
        self.text_box = QTextEdit()
        self.text_box.setReadOnly(True)
        right_layout.addWidget(self.text_box)
        
        # adding left and right layouts to the content layout
        content_layout.addLayout(self.left_layout)
        content_layout.addLayout(right_layout)
        
        # adding content layout to the main layout
        main_layout.addLayout(content_layout)
        
        # layout for footer
        footer_layout = QHBoxLayout()
        
        # ME-EC footer
        stock_text_footer = QLabel("ME-EC/ENG45-HU")
        stock_text_footer.setStyleSheet("font-size: 10px; margin: 0px; padding: 5px;")
        stock_text_footer.setAlignment(Qt.AlignmentFlag.AlignLeft)
        footer_layout.addWidget(stock_text_footer)
        
        # footer signature "@Powered by G0g1"
        powered_by_label = QLabel("@Powered by G0g1")
        powered_by_label.setStyleSheet("font-size: 10px; margin: 0px; padding: 5px;")
        powered_by_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        footer_layout.addWidget(powered_by_label)
        
        main_layout.addLayout(footer_layout)
        
        # setting main layout to the window
        self.setLayout(main_layout)
    

    def create_input_fields(self, fields: list[dict]):
        for field in fields:

            h_layout = QHBoxLayout()
            text_label = QLabel(f"{field['label']}:")
            h_layout.addWidget(text_label)

            if field['type'] == 'text':
                text_input = QLineEdit()
                setattr(self, field['name'], text_input)
                h_layout.addWidget(text_input)

            elif field['type'] == 'checkbox':
                checkbox = QCheckBox()
                setattr(self, field['name'], checkbox)
                h_layout.addWidget(checkbox)

            elif field['type'] == 'dropdown':
                dropdown = QComboBox()
                dropdown.addItems(field['options'])
                setattr(self, field['name'], dropdown)
                h_layout.addWidget(dropdown)
            
            self.left_layout.addLayout(h_layout)


    def on_submit(self):

        files = self.reader.list_of_files()
        self.read_table_from_msg(files)
    

    def log_message(self, file, status: Literal['processing', 'failed', 'processed'], error=None):
        if status == 'processing':
            self.text_box.append(f'Processing "{file}"...\n')
        elif status == 'failed':
            self.text_box.append(f'[-] Processing "{file}" has failed.\n[!] {error}\n{"#" * 20} \n')
        elif status == 'processed':
            self.text_box.append(f'[+] "{file}" has been processed.\n{"#" * 20} \n')
        QApplication.processEvents()


    def read_table_from_msg(self, file_list):

        for file in file_list:

            col_n = self.approver_input.text()
            col_o = date.today()
            col_s = 'x' if self.accepted_checkbox.isChecked() else None

            if file.endswith('.msg'):
                
                self.log_message(
                    file=file, 
                    status='processing'
                    )

                html_str = self.reader.converting_msg_to_html(
                    file_path=self.reader.file_dir,
                    file_name=file
                )

                try:
                    df = self.reader.converting_html_to_df(html=html_str)

                except ValueError as e:
                    self.log_message(
                        file=file,
                        status='failed',
                        error=e
                        )
                        
                    time.sleep(0.5)
                    continue

                else:
                    formated_df = self.reader.processing_dataframe(dataframe=df)
                    additional_datas = self.xls_processor.additional_datas(
                        col_n=col_n,
                        col_o=col_o,
                        col_s=col_s
                    )

                    self.xls_processor.data_to_excel(
                        sheet_name='Components',
                        data=formated_df,
                        add_datas=additional_datas
                    )

                    self.log_message(
                        file=file,
                        status='processed'
                    )

                    os.remove(f'{self.file_dir}\{file}')
                    time.sleep(0.5)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
