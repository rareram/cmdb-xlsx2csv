import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QComboBox, QVBoxLayout, QHBoxLayout, QWidget, QTextEdit, QFrame)
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QFont, QPainter, QColor, QPixmap
import pandas as pd
import os
import re

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CMDB 'xlsx to csv' Converter")
        self.setFixedSize(400, 400)
        self.setAcceptDrops(True)

        # main widget layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(2)
        layout.setContentsMargins(10, 0, 10, 0)

        top_container = QWidget()
        top_container.setFixedHeight(44)    # margin 10 + image height 34 = 44
        container_layout = QHBoxLayout(top_container)
        container_layout.setContentsMargins(0, 10, 0, 0)
        container_layout.setSpacing(0)

        encoding_widget = QWidget()
        encoding_layout = QHBoxLayout(encoding_widget)
        encoding_layout.setContentsMargins(0, 0, 0, 0)
        encoding_layout.setSpacing(5)

        # encoding select combobox
        encoding_label = QLabel("① encoding 선택:")
        self.encoding_combo = QComboBox()
        self.encoding_combo.addItems(['euc-kr', 'utf-8'])
        self.encoding_combo.setFixedWidth(100)

        encoding_layout.addWidget(encoding_label)
        encoding_layout.addWidget(self.encoding_combo)
        encoding_layout.addStretch()

        # logo
        try:
            logo_label = QLabel()
            logo_pixmap = QPixmap("logo.png")
            if not logo_pixmap.isNull():
                logo_label.setPixmap(logo_pixmap)
                logo_label.setFixedSize(70, 34)
            else:
                print("로고 이미지를 불러올 수 없습니다.")
        except Exception as e:
            print(f"로고 로딩 오류: {str(e)}")
            logo_label = QLabel()

        container_layout.addWidget(encoding_widget)
        container_layout.addWidget(logo_label, alignment=Qt.AlignRight | Qt.AlignTop)

        layout.addWidget(top_container)

        # comment
        instruction_label = QLabel("② 엑셀 파일을 아래에 드래그&드롭 하세요.")
        instruction_label.setAlignment(Qt.AlignLeft)
        instruction_label.setFixedHeight(20)
        layout.addWidget(instruction_label)

        # drag & drop area
        self.drop_frame = QFrame()
        self.drop_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        self.drop_frame.setFixedSize(380, 220)
        layout.addWidget(self.drop_frame)
        layout.setAlignment(self.drop_frame, Qt.AlignHCenter)

        # log display area
        self.log_text = QTextEdit()
        self.log_text.setFixedSize(380, 90)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        layout.setAlignment(self.log_text, Qt.AlignHCenter)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        font = QFont("맑은 고딕", 25, QFont.Bold)
        painter.setFont(font)
        painter.setPen(QColor(200, 200, 200, 100))

        texts = ["구성관리조회", "xlsx → csv", "변환 유틸리티"]
        fm = painter.fontMetrics()
        total_height = fm.height() * len(texts)
        frame_rect = self.drop_frame.geometry()

        for i, text in enumerate(texts):
            text_width = fm.width(text)
            x = frame_rect.x() + (frame_rect.width() - text_width) / 2
            y = frame_rect.y() + (frame_rect.height() - total_height) / 2 + fm.height() * i
            painter.drawText(QPoint(int(x), int(y + fm.height())), text)

    def clean_text(self, text):
        if pd.isna(text):                                   # None , NaN 처리
            return ''

        text = str(text)                                    # 모든 데이터를 문자열로 변환
        text = text.replace('\xa0', ' ')                    # non-breaking space 를 일반 공백으로 변환
        text = ' '.join(text.split())                       # 모든 종류의 공백문자를 일반 공백으로 변환
        text = text.replace('\n', ' ').replace('\r', ' ')   # 줄바꿈 제거
        text = re.sub(r'[^\x00-\x7F가-힇\s]', '', text)     # 특수문자 제거 또는 변환

        return text.strip()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file_path in files:
            if file_path.endswith('.xlsx'):
                self.convert_excel_to_csv(file_path)

    def log_message(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def convert_excel_to_csv(self, excel_path):
        try:
            # Sheet1만 읽기
            self.log_message(f"변환 시작: {os.path.basename(excel_path)}")
            df = pd.read_excel(excel_path, sheet_name="Sheet1")
            self.log_message("Excel 파일 로드 완료")

            # 2행까지 삭제
            df = df.iloc[2:]
            self.log_message("상단 2행 삭제 완료")

            # 모든셀 데이터 정제
            for column in df.columns:
                df[column] = df[column].apply(self.clean_text)
            self.log_message("데이터 정제 완료")

            # csv 경로 설정
            csv_path = os.path.splitext(excel_path)[0] + '.csv'
            encoding = self.encoding_combo.currentText()

            try:
                df.to_csv(csv_path, index=False, encoding=encoding)
            except UnicodeEncodeError:
                if encoding == 'euc-kr':
                    self.log_message("euc-kr 인코딩 실패, utf-8로 전환하여 저장을 시도합니다.")
                    self.encoding_combo.setCurrentText('utf-8')
                    df.to_csv(csv_path, index=False, encoding='utf-8')

            self.log_message(f"변환 완료: {os.path.basename(csv_path)}")

        except Exception as e:
            self.log_message(f"오류 발생: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())