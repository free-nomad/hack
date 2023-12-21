import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QMessageBox
from PyQt5.QtGui import QPixmap
import pandas as pd
import os
import sys

class ExcelFileSelector(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("파일 선택")
        self.setGeometry(100, 100, 500, 250)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        layout = QVBoxLayout()

        self.file_labels = []  # 파일 경로를 표시할 라벨 리스트

        # 각각의 파일 선택 버튼을 생성하고 연결
        file_types = ['카페24 파일', '노스노스 파일', '기준 파일']
        for i, file_type in enumerate(file_types):
            btn_select_file = QPushButton(f"{file_type}", self)
            btn_select_file.clicked.connect(lambda checked, num=i: self.select_file(num))
            layout.addWidget(btn_select_file)

            file_label = QLabel("선택된 파일 없음", self)
            self.file_labels.append(file_label)
            layout.addWidget(file_label)

        # 실행 버튼 추가
        btn_execute = QPushButton("실행", self)
        btn_execute.clicked.connect(self.execute_function)
        layout.addWidget(btn_execute)

        self.central_widget.setLayout(layout)

        # 파일 경로를 저장할 변수
        self.file_paths = ["", "", ""]

        # 로고 이미지 추가
        self.add_logo()

    def add_logo(self):
        label = QLabel(self)
        pixmap = QPixmap(os.path.join(os.path.dirname(__file__), 'logo.png'))
        label.setPixmap(pixmap)
        label.resize(pixmap.width(), pixmap.height())
        label.move(50, 10)

    def select_file(self, index):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setViewMode(QFileDialog.List)
        file_dialog.setOptions(options)

        file_name, _ = file_dialog.getOpenFileName(self, f"{self.file_labels[index].text()} 선택", "", "모든 파일 (*);;")  # 모든 파일 선택

        if file_name:
            # 선택된 파일의 이름을 라벨에 표시
            self.file_labels[index].setText(file_name.split('/')[-1])  # 파일 이름만 표시
            self.file_paths[index] = file_name

    def execute_function(self):
        # 파일 경로를 변수로 전달하여 data_col 함수 실행
        cafe24, nosnos, matching = self.file_paths
        cafe24_data, nosnos_data, matching_data = data_col(cafe24, nosnos, matching)

        # collect 함수 실행
        result = collect(cafe24_data, nosnos_data, matching_data)
        QMessageBox.information(self, "완료", "데이터 처리가 완료되었습니다.")
        # 결과를 엑셀 파일로 저장
        result.to_excel('result.xlsx', index=False)  # index=False로 설정하면 인덱스를 엑셀에 저장하지 않습니다.

def data_col(cafe, nos, mat):
    cafe24 = pd.read_csv(cafe)
    nosnos = pd.read_excel(nos)
    matching = pd.read_excel(mat)

    return cafe24, nosnos, matching

def collect(cafe24, nosnos, matching):
    cafe24_item_code = []
    nosnos_product_code = []
    supplier_name = []
    cafe24_product_name = []
    nosnos_product_name = []
    cafe24_stock_quantity = []
    cafe24_safety_stock = []
    cafe24_all_quantity = []
    nosnos_real_time_available_stock = []

    for i in range(len(matching)):

    

        if type(matching['cafe_code'][i]) == float:
            cafe24_item_code.append(0)
            nosnos_product_code.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['상품코드'].values[0])
            supplier_name.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['공급사명'].values[0])
            cafe24_product_name.append(0)
            nosnos_product_name.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['출고상품명'].values[0])
            cafe24_stock_quantity.append(0)
            cafe24_safety_stock.append(0)
            nosnos_real_time_available_stock.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['실시간 가용재고'].values[0])
        else:
            try:
                cafe24_item_code.append(cafe24[cafe24['품목코드'].str.contains(matching.loc[i, 'cafe_code'])]['품목코드'].values[0])
                nosnos_product_code.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['상품코드'].values[0])
                supplier_name.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['공급사명'].values[0])
                cafe24_product_name.append(cafe24[cafe24['품목코드'].str.contains(matching.loc[i, 'cafe_code'])]['상품명'].values[0])
                nosnos_product_name.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['출고상품명'].values[0])
                cafe24_stock_quantity.append(cafe24[cafe24['품목코드'].str.contains(matching.loc[i, 'cafe_code'])]['재고수량'].values[0])
                cafe24_safety_stock.append(cafe24[cafe24['품목코드'].str.contains(matching.loc[i, 'cafe_code'])]['안전재고'].values[0])
                cafe24_all_quantity
                nosnos_real_time_available_stock.append(nosnos[nosnos['상품코드'].str.contains(matching.loc[i, 'nosnos_code'])]['실시간 가용재고'].values[0])
            except:
                pass

    data = {
        '카페24 품목코드': cafe24_item_code,
        'nosnos 상품코드': nosnos_product_code,
        '공급사명': supplier_name,
        '카페24 상품명': cafe24_product_name,
        'nosnos 상품명': nosnos_product_name,
        '카페24 재고수량': cafe24_stock_quantity,
        '카페24 안전재고': cafe24_safety_stock,
        'nosnos 실시간가용재고': nosnos_real_time_available_stock
    }


    result = pd.DataFrame(data)

    # 데이터프레임을 엑셀 파일로 저장
    result.to_excel('result.xlsx', index=False)  # index=False로 설정하면 인덱스를 엑셀에 저장하지 않습니다.

    return result

def main():
    app = QApplication(sys.argv)
    window = ExcelFileSelector()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
