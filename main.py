from openpyxl import load_workbook
import sys
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import *
from PySide6.QtCore import QFile, QIODevice

FisieruExcel = 0
def iesire():
    print("Am iesit")
    sys.exit(0)

def IncarcaTabel():
    print("Incarc tabelul")
    global FisieruExcel
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
    while WorksheetExcel.cell(1,linii).value:
         linii += 1
    for i in range(1, linii):
        for j in range(1, coloane):
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

    if not window:
        print(loader.errorString())
        sys.exit(-1)
    window.show()

    sys.exit(app.exec())