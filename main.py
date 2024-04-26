from openpyxl import load_workbook
import sys
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import *
from PySide6.QtCore import QFile, QIODevice

FisieruExcel = None
TabelExcel = None
window = None

def iesire():
    print("Am iesit")
    sys.exit(0)

def DespreNebun():
    print("Afisam")
    window.findChild(QPushButton, "pushButton_Despre").setText("Nu stiu sa fac popup sry")

def filtruCrazy():
    print("Filtram o nebunie mare")

def IncarcaTabel():
    print("Incarc tabelul")
    global FisieruExcel
    global TabelExcel
    global window
    FisieruExcel = load_workbook(filename = 'Fish_table.xlsx')
    WorksheetExcel = FisieruExcel.active
    Textul = window.findChild(QPlainTextEdit, "plainTextEdit")
    print( WorksheetExcel.title)
    #Textul.setPlainText(WorksheetExcel.title + "\n" + WorksheetExcel.cell(1,1).value)
    textulTot = ""
    coloane = 1
    linii = 1
    while WorksheetExcel.cell(1,coloane).value:
         print("valuare: "+str(WorksheetExcel.cell(1,coloane).value)+"'")
         coloane += 1
    while WorksheetExcel.cell(linii,1).value:
         linii += 1
    TabelExcel = [[""]*(coloane-1)]*(linii-1)
    
    Tabelul = window.findChild(QTableView, "tableWidget")
    Tabelul.setRowCount(linii-2)
    Tabelul.setColumnCount(coloane-1)

    for i in range(1, linii):
        for j in range(1, coloane):
            TabelExcel[i-1][j-1] = str(WorksheetExcel.cell(i, j).value)
            if i-1 > 0:
                Tabelul.setItem(i-2, j-1, QTableWidgetItem(TabelExcel[i-1][j-1]))
            else:
                Tabelul.setHorizontalHeaderItem(j-1, QTableWidgetItem(TabelExcel[i-1][j-1]))
            textulTot += str(WorksheetExcel.cell(i, j).value) + " "
        textulTot += "\n"
    
    Textul.setPlainText(textulTot)



if __name__ == "__main__":
    app = QApplication(sys.argv)

    ui_file_name = "Interfata.ui"
    ui_file = QFile(ui_file_name)
    if not ui_file.open(QIODevice.ReadOnly):
        print(f"Cannot open {ui_file_name}: {ui_file.errorString()}")
        sys.exit(-1)
    loader = QUiLoader()
    window = loader.load(ui_file)
    ui_file.close()

    ButonIesire = window.findChild(QPushButton, "pushButton_Iesire")
    ButonIesire.clicked.connect(iesire)
    ButonIncarca = window.findChild(QPushButton, "pushButton_Incarca")
    ButonIncarca.clicked.connect(IncarcaTabel)
    ButonDespre = window.findChild(QPushButton, "pushButton_Despre")
    ButonDespre.clicked.connect(DespreNebun)

    if not window:
        print(loader.errorString())
        sys.exit(-1)
    window.show()

    sys.exit(app.exec())