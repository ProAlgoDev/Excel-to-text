import pandas as pd
import os
import re
from datetime import datetime, timedelta
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QMessageBox
from PyQt5.QtGui import QFont
import threading

class Covert:
    def __init__(self,name):
        try:
            self.inputExcelFile =name
            self.inputName = str(self.inputExcelFile)
            if self.inputName[-5:].lower() == ".xlsx":
                self.inputName = self.inputName[:-5]
            elif self.inputName[-4:].lower() == ".xls":
                self.inputName = self.inputName[:-4]
            excelFile = pd.read_excel(self.inputExcelFile,header=None)
            excelFile = excelFile.dropna(how='all')
            # excelFile = excelFile.iloc[:, :4]
            excelFile.to_csv("out.csv", index = None,  quoting=1)
            self.dataframeObject = pd.read_csv("out.csv")
        except:pass
    def second(self):
        print("second")
        try:
            with pd.read_csv("out.csv", header=None, chunksize=1) as reader:
                file = open(f"{self.inputName}.txt",'w',encoding='utf-8')
                date = ''
                patern = r'\d+'
                temp = ''
                for row_num, chunk in enumerate(reader, start=0):
                        values_in_second_column = str(chunk.iloc[0, 1])
                        value_in_first_column = str(chunk.iloc[0, 0])
                        try:
                            if '2023 оны' in values_in_second_column or '2023 ОНЫ' in values_in_second_column:
                                if row_num != 0 and date != '':
                                    file.write("\n")
                                h = re.findall(patern, values_in_second_column)
                                if h[0] == '2023':
                                    month = self.add(h[1])
                                    day = self.add(h[2])
                                if len(h) <=2 and len(h[0]) <= 2:
                                    month = self.add(h[0])
                                    day = self.add(h[1])
                                date = f"2023/{month}/{day}"
                                if temp != date:
                                    file.write(date+"\n")
                                    temp = date
                            if 'САРЫН' in values_in_second_column or 'сарын' in values_in_second_column:
                                if row_num != 0 and date != '' and temp != date:
                                    file.write("\n")
                                h = re.findall(patern, values_in_second_column)
                                if h[0] == '2023':
                                    month = self.add(h[1])
                                    day = self.add(h[2])
                                if len(h) <=2 and len(h[0]) <= 2:
                                    month = self.add(h[0])
                                    day = self.add(h[1])
                                date = f"2023/{month}/{day}"
                                if temp != date:
                                    file.write(date+"\n")
                                    temp = date
                            if 'САРЫН' in value_in_first_column or 'сарын' in value_in_first_column:
                                print("3333333333333")
                                if row_num != 0 and date != '':
                                    file.write("\n")
                                h = re.findall(patern, value_in_first_column)
                                if h[0] == '2023':
                                    month = self.add(h[1])
                                    day = self.add(h[2])
                                if len(h) <=2 and len(h[0]) <= 2:
                                    month = self.add(h[0])
                                    day = self.add(h[1])
                                date = f"2023/{month}/{day}"

                                print(date)
                                if temp != date:
                                    file.write(date+"\n")
                                    temp = date
                            try:
                            # Convert the time to a datetime object
                                value_in_first_column = datetime.strptime(value_in_first_column, "%H:%M:%S")
                                # Format the datetime object to "12:00" (without seconds)
                                value_in_first_column = value_in_first_column.strftime("%H:%M")
                                tempContent = values_in_second_column
                                tempContent = self.remove_space(tempContent)
                                file.write(value_in_first_column+"\t"+values_in_second_column+"\n")
                            except:pass
                        except: 
                            pass
                file.close()
            global complete
            complete.append(self.inputExcelFile)
            
        except:
            global failed
            failed.append(self.inputExcelFile)
            pass 

    def first(self):
        print("first")
        try:
            with pd.read_csv("out.csv", header=None, chunksize=1) as reader:
                file = open(f"{self.inputName}.txt",'w',encoding='utf-8')
                date = ''
                patern = r'\d+'
                temp = ''
                flag = ''
                for row_num, chunk in enumerate(reader, start=0):
                        values_in_second_column = str(chunk.iloc[0, 1])
                        value_in_first_column = str(chunk.iloc[0, 0])
                        value_in_third_column = str(chunk.iloc[0, 2])
                        g = re.findall(patern, value_in_first_column[:10])
                        try:
                            if '2023 ОНЫ' in value_in_third_column:
                                if row_num != 0 and date != '':
                                    file.write("\n")
                                value_in_third_column = value_in_third_column.replace(" ","")
                                year = value_in_third_column[:4]
                                month = value_in_third_column.split("ОНЫ")[1].split("-РСАРЫН")[0]
                                day = value_in_third_column.split("ОНЫ")[1].split("-РСАРЫН")[1].split("-НЫ")[0]
                                day = day[:2]
                                day = self.add(day)
                                month = self.add(month)
                                date = f"{year}/{month}/{day}"
                                file.write(date+"\n")
                                temp = date
                            if 'САРЫН' in value_in_third_column:
                                if row_num != 0 and date != '' and temp != date:
                                    file.write("\n")
                                h = re.findall(patern, value_in_third_column)
                                test = len(h)
                                if len(h) <=2 and test <=2:
                                    month = self.add(h[0])
                                    day = self.add(h[1])
                                    date = f"2023/{month}/{day}"
                                    file.write(date+"\n")
                                    flag = date
                            try:
                                g = re.findall(patern, value_in_first_column[:10])
                                h = re.findall(patern, value_in_third_column)
                                if h and "сарын" in value_in_third_column:
                                    if len(h) == 3 and h[0] == '2023':
                                            month = self.add(h[1])
                                            day = self.add(h[2])
                                            date = h[0]+'/'+month+'/'+day
                                            if temp == '' and temp != date:
                                                file.write(date+"\n") 
                                                temp = date 
                                            elif temp != date:
                                                file.write("\n"+date+"\n")
                                                temp = date
                                    if len(h) == 2 and len(h[0]) <= 2:
                                        year = "2023"
                                        if len(h[0]) <= 2:
                                            month = self.add(h[0])
                                            day = self.add(h[1])
                                            date = year +'/' + month +'/' + day
                                            if temp =='' and temp != date:
                                                file.write(date+"\n")  
                                                temp = date
                                            elif temp != date:
                                                file.write("\n"+date+"\n")
                                                temp = date
                                    
                            except: pass

                            flag = date
                            if flag == '':
                                g = re.findall(patern, value_in_first_column[:10])
                                if len(g) == 3 and g[0] == '2023':
                                        month = self.add(g[1])
                                        day = self.add(g[2])
                                        date = g[0]+'/'+month+'/'+day
                                        if temp == '' and temp != date:
                                            file.write(date+"\n")  
                                            temp = date
                                        elif temp != date:
                                            temp = date
                                            file.write("\n"+date+"\n")
                                if len(g) == 2 and len(g[0]) <= 2:
                                        year = "2023"
                                        month = self.add(g[0])
                                        day = self.add(g[1])
                                        date = year +'/' + month +'/' + day
                                        print(temp)
                                        if temp =='' and temp != date:
                                            file.write(date+"\n")
                                            temp = date

                                        elif temp != date:
                                            file.write("\n"+date+"\n")
                                            temp = date
                                flag= ''
                                date=''
                        except: pass
                        try:
                            value_in_first_column = datetime.strptime(value_in_first_column, "%H:%M:%S")
                            value_in_first_column = value_in_first_column.strftime("%H:%M")
                            tempContent = self.remove_space(value_in_third_column)
                            
                            file.write(value_in_first_column+"\t"+tempContent+"\n")
                        except:pass
                file.close()
            global complete
            complete.append(self.inputExcelFile)
            
        except: 
            global failed
            failed.append(self.inputExcelFile)
            pass
    def third(self):
        print("third")
        global failed
        global complete
        try:
            with pd.read_csv("out.csv", header=None, chunksize=1) as reader:
                file = open(f"{self.inputName}.txt",'w',encoding='utf-8')
                temp = ''
                date = ''
                error = False
                patern = r'\d+'
                text = ''
                check = False
                checked = 0
                for row_num, chunk in enumerate(reader, start=0):
                        value_in_first_column = str(chunk.iloc[0, 0])
                        values_in_second_column = str(chunk.iloc[0, 1])
                        values_in_third_column = str(chunk.iloc[0, 2])
                        try:
                            g = re.findall(patern, value_in_first_column[:10])
                            h = re.findall(patern, values_in_second_column)
                            if len(g) == 3 and g[0] == '2023':
                                if len(g[1]) <=2 and len(g[2]) <=2:
                                    month = self.add(g[1])
                                    day = self.add(g[2])
                                    date = g[0]+'/'+month+'/'+day
                                    if temp == '' and temp != date:
                                        text+=date + "\n"
                                        temp = date
                                    elif temp != date:
                                        text +="\n"+date+"\n"
                                        print("first",temp)
                                        print("second",date)
                                        check = True
                                        temp = date
                                    else: 
                                        check = False
                            if len(g) == 2:
                                continue
                            if len(h) > 3:
                                continue
                            if len(h[0]) <=2 and len(h[1]) <=2:
                                check_time = int(h[0])
                                if check == True:
                                    if check_time >= 0 and check_time <= 3:
                                        tempDate = date
                                        input_date = datetime.strptime(tempDate, "%Y/%m/%d")
                                        result_date = input_date - timedelta(days=1)
                                        yeart = str(result_date.year)
                                        montht = str(result_date.month)
                                        dayt = str(result_date.day)
                                        montht = self.add(montht)
                                        dayt = self.add(dayt)
                                        lines = text.splitlines()
                                        if lines:
                                            lines.pop()
                                            lines.pop()
                                        text = "\n".join(lines)
                                        text +="\n"
                                        
                                        tempResult = yeart+"/"+montht + "/" + dayt
                                        temp = tempResult
                                    if check_time > 8:
                                        lines = text.splitlines()
                                        if lines:
                                            lines.pop()
                                            lines.pop()
                                        text = "\n".join(lines)
                                        text +="\n"
                                time = h[0]+":"+h[1]
                                if len(values_in_third_column) < 4:
                                    error = True
                                tempContent = self.remove_space(values_in_third_column)
                                
                                text += time+"\t"+tempContent+"\n"
                        except:
                            pass
                file.write(text)
                file.close()
            if not error:
                complete.append(self.inputExcelFile)
            elif error:
                failed.append(self.inputExcelFile)
        except: 
            failed.append(self.inputExcelFile)
            pass
    def remove_space(self,text):
        check = False
        temp = ''
        temptext = text
        for i in temptext:
            if i == ' ':
                if check:
                    temp += ''
                else: temp+=i
                check = True
            else:
                check = False
                temp +=i 
        return temp
             
    def main(self):
        patern = r'\d+'
        try:
            try:
                test2 = str(self.dataframeObject.iloc[8,2])
            except:
                test2 = ''
                pass
            try:
                test = str(self.dataframeObject.iloc[8,0])
                g = re.search(patern, test)
                number = int(g.group())
                print(number)
            except:
                try:
                    test = str(self.dataframeObject.iloc[9,0])
                except:
                    test = ''
                    pass
            try:
                test1 = str(self.dataframeObject.iloc[8,1])
                g = re.search(patern, test1)
                number = int(g.group())
                print(number)
            except:
                try:
                    test1 = str(self.dataframeObject.iloc[9,1])
                except:
                    test1 = ''
                    pass
                pass
            print(len(test),len(test1))
            
            if len(test) >= 9 and len(test) < 12:
                self.third()
            if len(test) <=8 and len(test) > 6 and len(test1) < 12:
                self.first()
            if len(test) >12:
                self.third()
            elif len(test) >12:
                self.third()
            if len(test) <=8 and len(test) > 6 and len(test1) > 12:
                self.second()
            if len(test) < 4 and test != '':
                self.dataframeObject = self.dataframeObject.drop(self.dataframeObject.columns[0],axis=1)
                self.dataframeObject.to_csv("out.csv", index = None,  quoting=1)
                print(self.dataframeObject)
                self.dataframeObject = pd.read_csv("out.csv")
                self.main()

                
        except:
            global failed
            failed.append(self.inputExcelFile)
            pass
    def add(self,date):
        if len(date) == 1:
            date = "0" + date
        
        return date


path = ''
complete = []
failed = []
state = ''


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.setWindowTitle("Converter")
        self.setGeometry(100, 100, 500, 220)

        font = QFont("Arial", 14)  # Create a QFont object with Arial font and size 14

        # Create a button to open the file dialog
        self.open = QPushButton("Select Folder", self)
        self.open.setGeometry(70, 50, 150,50)
        self.open.setFont(font)
        self.open.clicked.connect(self.open_folder_dialog)

        self.path = QLabel("Folder Path: ", self)
        self.path.move(20,110)
        self.path.setFont(font)

        self.label = QLabel("", self)
        self.label.move(130,110)
        self.label.setFixedSize(350,30)
        self.label.setFont(font)
        
        self.state = QLabel("", self)
        self.state.move(200,140)
        self.state.setFixedSize(300,30)
        self.state.setFont(font)

        self.convert = QPushButton("Convert", self)
        self.convert.setGeometry(270, 50, 150,50)
        self.convert.setFont(font)
        self.convert.clicked.connect(self.start_convert)

    def get_excelfile(self,path):
        excel_files = []
        if path == '':
            self.show_alert("Select folder")
        for root, _, files in os.walk(path):
            for file in files:
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    excel_files.append(os.path.join(root, file))
        return excel_files
    
    def show_alert(self,text):
        # Create a QMessageBox
        alert_text = text
        alert = QMessageBox(self)

        # Set the icon, title, and text of the QMessageBox
        alert.setIcon(QMessageBox.Information)
        alert.setWindowTitle("Alert")
        alert.setText(alert_text)

        font = alert.font()
        font.setPointSize(14)
        alert.setFont(font)

        # Add an "OK" button to the QMessageBox
        alert.setStandardButtons(QMessageBox.Ok)

        # Show the QMessageBox and handle the result
        result = alert.exec_()
        if result == QMessageBox.Ok:
            print("OK")
    def start(self, files):
        fileList = files
        for row_num, i in enumerate(fileList, start=0):
            instance = Covert(i)
            instance.main()
            self.state.setText("Processing...")
        # instance = Covert("25 ТВ -1-р суваг хөтөлбөр 2023-7.31-2023.8.6 .xlsx")
        # instance.main()
        file = open("report.txt",'w',encoding='utf-8')
        file.write("Completed\n")
        for i in complete:
            file.write("\t"+i+"\n")
        file.write("Failed\n")
        for j in failed:
            file.write("\t"+ j + "\n")
        file.write("\nResult\n")
        file.write(f"{len(complete)}  Completed\n")
        file.write(f"{len(failed)}  Failed\n")
        file.close()
        # if os.path.exists('out.csv'):
        #     os.remove('out.csv')
        self.state.setText("Finished")
    def start_convert(self):
        global path
        global complete
        global failed
        files = []
        files = self.get_excelfile(path)
        if len(files) == 0:
            self.show_alert("Excel file does not exist.")
        print(path)

        startConvert = threading.Thread(target=self.start,args=(files,))
        startConvert.daemon = True
        startConvert.start()

        
    def open_folder_dialog(self):

        folder_dialog = QFileDialog(self)
        folder_dialog.setFileMode(QFileDialog.Directory)
        folder_dialog.setOption(QFileDialog.ShowDirsOnly)

        self.folder_path = folder_dialog.getExistingDirectory(self, "Select Folder")
        global path
        path = self.folder_path
        self.label.setText(self.folder_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())