import pandas as pd
import sys
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, 
    QPushButton, QFileDialog, QVBoxLayout,
    QWidget, QApplication, QHBoxLayout,
    QLineEdit, QGridLayout, QScrollArea
)

def searchData(file_path, search_string):
    data = pd.read_excel(file_path, skiprows=8)
    result = []
   
    for i in range(len(data)):
        if search_string.lower() in data.iloc[i]['CÂU HỎI'].lower():
            question = data.iloc[i]['CÂU HỎI']
            answerIndex = data.iloc[i]['ĐÁP ÁN ĐÚNG'] + 1
            answer = data.iloc[i, answerIndex]
            result.append([question, str(answer), str(i + 1)])
    return result

class MainWindow(QMainWindow):
    file_path = ""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Crawler - Python")
        self.resize(800, 600)
        self.center()
      
        self.file_label = QLabel("Select a file to start", self)
        self.file_button = QPushButton("Browse file", self)
        self.input_field = QLineEdit(self)
        
        self.input_field.setPlaceholderText("Enter search string")
        self.search_button = QPushButton("Search", self)
        self.initUI()
        
    def initUI(self):        
        central_widget = QWidget()
        self.setStyleSheet("background-color: #323232;")
        self.setCentralWidget(central_widget)
        
        file_sellection_layout = QHBoxLayout()
        self.file_button.clicked.connect(self.pick_file)
        
        self.file_label.setStyleSheet("background-color: #151E29; border-radius: 8px; padding: 10px; color: white;")
        self.file_button.setStyleSheet("width: 100px; border-radius: 8px; padding: 10px; color: white; background-color:#656565")
        
        file_sellection_layout.addWidget(self.file_label, 7)
        file_sellection_layout.addWidget(self.file_button, 3)
        
        search_layout = QHBoxLayout()
        self.search_button.clicked.connect(self.search)
        
        self.input_field.setStyleSheet("background-color: #141D27; border-radius: 8px; padding: 10px; color: white;")
        self.search_button.setStyleSheet("width: 100px; border-radius: 8px; padding: 10px; color: white; background-color:#656565")
        
        search_layout.addWidget(self.input_field, 7)
        search_layout.addWidget(self.search_button, 3)
        
        self.table_layout = QGridLayout()
        self.table_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.table_layout.setColumnStretch(0, 1)
        self.table_layout.setColumnStretch(1, 15)
        self.table_layout.setColumnStretch(2, 15)
        
        scroll_content = QWidget()
        scroll_content.setLayout(self.table_layout)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(scroll_content)
        scroll_area.setStyleSheet("""
            QScrollBar:vertical {
                background: #2e2e2e;
                width: 6px;
                margin: 0px 0px 0px 0px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #606060;
                min-height: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #909090;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        layout.addLayout(file_sellection_layout)
        layout.addLayout(search_layout)
        layout.addWidget(scroll_area)

        central_widget.setLayout(layout)
    
    def pick_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Excel (*.xls *.xlsx)")
        if file_path:
            self.file_label.setText(f"{file_path}")
            self.file_path = file_path
            
    def clear_table(self):
        while self.table_layout.count() :
            item = self.table_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.setParent(None)
                widget.deleteLater()
            elif item.layout() is not None:
                self.clear_table(item.layout())
            
    def search(self):
        self.clear_table()
        if (self.file_path == ""): return
        result = searchData(self.file_path, self.input_field.text())    
        
        questionTile = QLabel("Question")
        answerTile = QLabel("Answer")
        indexTile = QLabel("Index")
        
        questionTile.setStyleSheet("font-weight: bold; font-size: 16px;")
        answerTile.setStyleSheet("font-weight: bold; font-size: 16px;")
        indexTile.setStyleSheet("font-weight: bold; font-size: 16px;")
        
        self.table_layout.addWidget(indexTile, 0, 0, alignment=Qt.AlignCenter)
        self.table_layout.addWidget(questionTile, 0, 1, alignment=Qt.AlignCenter)
        self.table_layout.addWidget(answerTile, 0, 2, alignment=Qt.AlignCenter)
        
        for i in range(len(result)):
            question = QLabel(result[i][0])
            answer = QLabel(result[i][1])
            index = QLabel(result[i][2])
            
            question.setStyleSheet("background-color: #2F2F2F; border-radius: 8px; padding: 4px; color: white;")
            answer.setStyleSheet("background-color: #607d29; border-radius: 8px; padding: 4px; color: white;")
            
            question.setWordWrap(True)
            answer.setWordWrap(True)
            
            self.table_layout.addWidget(index, i+1, 0)
            self.table_layout.addWidget(question, i+1, 1)
            self.table_layout.addWidget(answer, i+1, 2)
        
    def center(self):
        screen_geometry = QApplication.desktop().screenGeometry()
        window_geometry = self.frameGeometry()
        center_point = screen_geometry.center()
        window_geometry.moveCenter(center_point)
        self.move(window_geometry.topLeft())    
        
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
