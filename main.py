import sys
import win32com.client as win32
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                             QListWidget, QFileDialog, QMessageBox, QHBoxLayout)

class HwpMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.files = []

    def initUI(self):
        self.setWindowTitle('HWP/HWPX 파일 합치기')
        self.setGeometry(300, 300, 500, 400)

        layout = QVBoxLayout()

        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton('파일 추가')
        self.btn_add.clicked.connect(self.add_files)
        
        self.btn_clear = QPushButton('목록 초기화')
        self.btn_clear.clicked.connect(self.clear_list)

        self.btn_merge = QPushButton('하나로 합치기')
        self.btn_merge.clicked.connect(self.merge_hwp)
        self.btn_merge.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_clear)
        layout.addLayout(btn_layout)
        layout.addWidget(self.btn_merge)

        self.setLayout(layout)

    def add_files(self):
        fnames, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'HWP Files (*.hwp *.hwpx)')
        if fnames:
            self.files.extend(fnames)
            self.file_list.addItems(fnames)

    def clear_list(self):
        self.files.clear()
        self.file_list.clear()

    def merge_hwp(self):
        if not self.files:
            QMessageBox.warning(self, "경고", "합칠 파일을 먼저 추가하세요.")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, '저장 위치 선택', '', 'HWP Files (*.hwp *.hwpx)')
        if not save_path:
            return

        try:
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") # 보안 모듈 승인 가정
            hwp.XHwpWindows.Item(0).Visible = False # 백그라운드 실행

            # 첫 번째 파일 열기
            hwp.Open(self.files[0])
            
            # 두 번째 파일부터 이어 붙이기
            for file in self.files[1:]:
                hwp.MovePos(3) # 문서 끝으로 이동
                hwp.InsertFile(file)

            hwp.SaveAs(save_path)
            hwp.Quit()
            QMessageBox.information(self, "완료", "파일이 성공적으로 합쳐졌습니다!")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"작업 중 오류 발생: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = HwpMergerApp()
    ex.show()
    sys.exit(app.exec_())
