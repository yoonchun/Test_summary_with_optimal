import sys
import time
import re
import os.path
import pandas as pd
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
import TestSummary


class MyForm(QMainWindow, TestSummary.Ui_MainWindow):

    def __init__(self):
        super().__init__()
        # self.setWindowIcon(QIcon('main.png'))
        self.setupUi(self)

    def open_file(self):

        self.label.setText("Open Files...")

        file_names, _ = QFileDialog.getOpenFileNames(self, "Open file", './', "Excel(*.xlsx)")

        for file in file_names:
            self.listWidget.addItem(file)
        self.label.setText("File lists are updated")

    def get_summary(self):

        self.label.setText("Working with files...")

        files = []
        for x in range(self.listWidget.count()):
            files.append(self.listWidget.item(x).text())

        try:
            for file in files:

                # Read excel파일, 첫번째 시트를 Read... sheet_name=0 옵션 이용)
                df_original = pd.read_excel(file, sheet_name=0)

                # Test Summary 에 필요한 컬럼만 추출
                df_table = df_original[
                    ['Test House', 'Project', 'Test Program', 'Lot', 'Total Units', 'First Pass Good', 'Total Good']]

                # Grouping(Index 없이)후, 각 컬럼의 갯수 및 합
                df_new = df_table.groupby(by=['Test House', 'Project', 'Test Program'], as_index=False) \
                    .aggregate({'Lot': 'count', 'Total Units': 'sum', 'First Pass Good': 'sum', 'Total Good': 'sum'})

                # FPY / FTY 의 컬럼 생성 (.map함수 이용하여 % & 소숫점 이하 2자리로 포맷팅하였음)
                df_new['FPY'] = (df_new['First Pass Good'] / df_new['Total Units']).map(lambda n: '{:,.2%}'.format(n))
                df_new['FTY'] = (df_new['Total Good'] / df_new['Total Units']).map(lambda n: '{:,.2%}'.format(n))

                # Test Program Revision 을 추출하기 위해 사용될 새로운 컬럼(Temp) 추가
                df_new['Temp'] = df_new['Test Program']

                idx = 0
                for value in df_new['Temp']:
                    # RevXXX.prog의 경우 Rev정보 추출 (대소문자 상관없이... re.IGNORECASE 옵션이용)
                    if re.findall(r'(rev.*?)(?:.prog)', value, re.IGNORECASE):
                        x = re.findall(r'(rev.*?)(?:.prog)', value, re.IGNORECASE)
                        df_new.at[idx, 'Temp'] = x[0]

                    # RevXXX.tp의 경우 Rev정보 추출 (대소문자 상관없이... re.IGNORECASE 옵션이용)
                    elif re.findall(r'(rev.*?)(?:.tp)', value, re.IGNORECASE):
                        x = re.findall(r'(rev.*?)(?:.tp)', value, re.IGNORECASE)
                        df_new.at[idx, 'Temp'] = x[0]

                    # Flex & Pax의 Test Program형식에 맞는 경우 Rev정보 추출 (ex: FT8U-PM811-3-12)
                    elif re.findall(r'^[a-zA-Z0-9]+-[a-zA-Z0-9]+-[a-zA-Z0-9]+-[a-zA-Z0-9]+$', value):
                        x = value.split('-')
                        df_new.at[idx, 'Temp'] = "REV" + x[3]

                    # 모든 형식에 부합하지 않을 경우 NA로 반환
                    else:
                        df_new.at[idx, 'Temp'] = "NA"

                    idx += 1

                # Comment 컬럼에 Weekly summary 포맷에 맞게 필요 내용을 조합.
                # 문자열의 조합이기 때문에 Dataframe형식이 string이 아닌 경우에는, astype(str)로 형변환 필요.
                df_new['Comment'] = '@' + df_new['Test House'] + ' : ' + df_new['Lot'].astype(str) + ' Lots (' \
                                    + round((df_new['Total Units'] / 1000), 1).astype(str) + 'k) test. FPY is ' \
                                    + df_new['FPY'] + '. FTY is ' + df_new['FTY'] + ' with ' \
                                    + df_new['Temp'] + '.'

                # 임시적으로 Rev추출을 위해 만들었던 Temp컬럼은 삭제
                df_new.pop("Temp")

                # 가독성을 위하여 숫자에 thousands separator를 추가 (.map함수이용)
                df_new['Total Units'] = df_new['Total Units'].map('{:,}'.format)
                df_new['First Pass Good'] = df_new['First Pass Good'].map('{:,}'.format)
                df_new['Total Good'] = df_new['Total Good'].map('{:,}'.format)

                # 파일 이름에 들어갈 Time stamp
                current_time = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))

                # 파일 name 과 extension 을 분리
                file_name, file_ext = os.path.splitext(file)

                # 최종 table 을 csv로 저장 (Dataframe의 Index를 제거... index=False 이용)
                df_new.to_csv(file_name + "_summary_" + current_time + ".csv", index=False)

            self.label.setText("Done!")

            if len(files) == 0:
                QMessageBox.about(self, 'Message Box', 'No file loaded!')
            else:
                QMessageBox.about(self, 'Message Box', 'Complete to generate summary!')

        except Exception as e:
            QMessageBox.about(self, 'Message Box', str(e))

    def clear(self):
        self.listWidget.clear()
        self.label.setText("Cleared!")

    def exit(self):
        sys.exit()


app = QApplication(sys.argv)
w = MyForm()
w.show()
app.exec_()
# sys.exit(app.exec())
