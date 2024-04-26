from openpyxl import load_workbook
import sys
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import *
from PySide6.QtCore import QFile, QIODevice

TabelExcel = None
window = None
linii = 0
coloane = 0

def iesire():
    print("Am iesit")
    sys.exit(0)

def DespreNebun():
    print("Afisam")
    window.findChild(QPushButton, "pushButton_Despre").setText("Nu stiu sa fac popup sry")

def filtruCrazy():
    print("Filtram o nebunie mare")
    global window, TabelExcel
    Tabelul = window.findChild(QTableView, "tableWidget")
    Tabelul.setRowCount(0)
    #Tabelul.setColumnCount(0)
    #filtrul pentru anotimp
    elementeTabel = 0
    elementCautat = window.findChild(QComboBox, "comboBoxSezon").currentText()
    for i in range(1, linii):
        if elementCautat in TabelExcel[i-1][1] or elementCautat == "Any":
            Tabelul.insertRow(elementeTabel)
            elementeTabel += 1
            for j in range(1, coloane):
                Tabelul.setItem(elementeTabel-1, j-1, QTableWidgetItem(TabelExcel[i-1][j-1]))


def IncarcaTabel():
    print("Incarc tabelul")
    global TabelExcel
    global window
    global coloane, linii
    FisieruExcel = load_workbook(filename = 'Fish_table.xlsx')
    WorksheetExcel = FisieruExcel.active
    print( WorksheetExcel.title)

    coloane = 1
    linii = 1
    while WorksheetExcel.cell(1,coloane).value:
         print("valuare: "+str(WorksheetExcel.cell(1,coloane).value)+"'")
         coloane += 1
    while WorksheetExcel.cell(linii,1).value:
         linii += 1
    #TabelExcel = [[""]*(coloane-1)]*(linii-1)
    TabelExcel = []

    Tabelul = window.findChild(QTableView, "tableWidget")
    Tabelul.setRowCount(linii-2)
    Tabelul.setColumnCount(coloane-1)

    for i in range(1, linii):
        listaRand = []
        for j in range(1, coloane):
            if i-1 > 0:
                celula =WorksheetExcel.cell(i, j).value
                #TabelExcel[i-2][j-1]
                listaRand.append(str(celula))
                Tabelul.setItem(i-2, j-1, QTableWidgetItem(str(celula)))
            else:
                Tabelul.setHorizontalHeaderItem(j-1, QTableWidgetItem(WorksheetExcel.cell(i, j).value))
        if listaRand:
            TabelExcel.append(listaRand)
    print("Am incarcat")

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
    FiltruCombo = window.findChild(QComboBox, "comboBoxSezon")
    FiltruCombo.currentIndexChanged.connect(filtruCrazy)

    if not window:
        print(loader.errorString())
        sys.exit(-1)
    window.show()

    sys.exit(app.exec())