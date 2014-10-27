#!python
#name: processBuilder_v0.2
#author: Gerrit Lükens
#python 3.4.1

#imports
import sys
import os
import configparser
import io
from txtToXlsWriter import convertTXTtoXLS
from PySide.QtCore import *
from PySide.QtGui import *

def getProcessStepFiles(templatePath):
    # crawls subdirectory ProcessSteps and returns a list of all template files
    filesListProcessSteps = []
    for path, dirs, files in os.walk(templatePath):
        sublist = []
        sublist.append(path)
        for f in files:
            sublist.append(f)
        filesListProcessSteps.append(sublist)
    return filesListProcessSteps

class ProcessStepSelectorWidget(QTreeWidget):
    def __init__(self, parent=None):
        super(ProcessStepSelectorWidget, self).__init__(parent)

        self.setColumnCount(1)
        self.setHeaderLabel("Process Steps")

        readIni = configparser.ConfigParser()
        readIni.read("processBuilder.ini")
        templatePath = readIni["DEFAULT"]["templatePath"]

        #get process step files
        filesList = getProcessStepFiles(templatePath)
        countSeparatorsTemplatePath = len(templatePath.split("\\"))

        #create item tree
        branchItems = []
        branchItems.append(self) #set self as root
        tempTreeWidget = 0
        for i in range(1, len(filesList)):
            countSeparators = len(filesList[i][0].split("\\")) - countSeparatorsTemplatePath
            if countSeparators == len(branchItems):
                tempTreeWidget = QTreeWidgetItem(branchItems[-1])
                tempTreeWidget.setText(0, filesList[i][0].split("\\")[-1])
            elif countSeparators > len(branchItems):
                branchItems.append(tempTreeWidget)
                tempTreeWidget = QTreeWidgetItem(branchItems[-1])
                tempTreeWidget.setText(0, filesList[i][0].split("\\")[-1])
            elif countSeparators < len(branchItems):
                while countSeparators < len(branchItems):
                    branchItems.pop()
                tempTreeWidget = QTreeWidgetItem(branchItems[-1])
                tempTreeWidget.setText(0, filesList[i][0].split("\\")[-1])
            for ii in range(1, len(filesList[i])):
                tempLowTreeWidget = QTreeWidgetItem(tempTreeWidget)
                tempLowTreeWidget.setText(0, filesList[i][ii].split(".")[0])
                #set xlswriter command for file in 2nd column of TreeWidgetItem Text
                tempLowTreeWidget.setText(1, ">" + filesList[i][0].replace("\\","/") + "/" + filesList[i][ii].split(".")[0])
                #write parent names in third column
                parentsString = ""
                for iii in range(1,  len(branchItems)):
                    parentsString += branchItems[iii].text(0) + " -> "
                parentsString += tempTreeWidget.text(0)
                tempLowTreeWidget.setText(2, parentsString)


        #sort items
        self.sortItems(0, Qt.AscendingOrder)

        #custom commands
        def createCustomCommands():
            """creates control items in TreeWidget to invoke functions"""
            customCmdRoot = QTreeWidgetItem(self)
            customCmdRoot.setText(0, "Extra")
            customCmdHeading = QTreeWidgetItem(customCmdRoot)
            customCmdHeading.setText(0, "Insert Heading")
            customCmdHeading.setText(1, "COMMAND")
            customCmdHeading.setText(2, "HEADING")
        createCustomCommands()


# noinspection PyUnresolvedReferences
class ProcessBuilderGui(QDialog):
    def __init__(self, parent=None):
        super(ProcessBuilderGui, self).__init__(parent)

        #set window title and load ini configs
        self.setWindowTitle("Process Builder")
        self.readIni = configparser.ConfigParser()
        self.readIni.read("processBuilder.ini")

        #create tree structure with loaded process step templates
        self.selectorWidget = ProcessStepSelectorWidget()

        #create list for process flow
        self.listWidget = QListWidget()
        self.listWidget.setMovement(QListView.Snap)
        self.listWidget.setDragDropMode(QAbstractItemView.InternalMove)
        self.currentListItem = 0
        self.exchangeItem = 0

        #setup process edit window
        self.editOkButton = QPushButton("Ok")
        self.editCancelButton = QPushButton("Cancel")
        self.processEditWidget = QWidget()
        self.processEditWidget.setMinimumSize(400, 200)
        self.tableWidget = QTableWidget()
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.processEditWidgetLayout = QGridLayout()
        self.processEditWidgetSubLayout = QHBoxLayout()
        self.processEditWidgetSubLayout.addWidget(self.editOkButton)
        self.processEditWidgetSubLayout.addWidget(self.editCancelButton)
        self.processEditWidgetLayout.addWidget(self.tableWidget, 0, 0)
        self.processEditWidgetLayout.addLayout(self.processEditWidgetSubLayout, 1, 0)
        self.processEditWidget.setLayout(self.processEditWidgetLayout)

        #setup Main GUI
        #create Main UI buttons
        self.generateXlsButton = QPushButton("Generate")
        self.saveProcessButton = QPushButton("Save")
        self.loadProcessButton = QPushButton("Load")
        self.editProcessButton = QPushButton("Edit")
        self.clearProcessButton = QPushButton("Clear")
        self.layout = QGridLayout()
        self.layoutButtons = QVBoxLayout()
        self.layout.addWidget(self.selectorWidget, 0, 0)
        self.layout.addWidget(self.listWidget, 0, 1)
        self.layout.addWidget(self.generateXlsButton, 1, 0)
        self.layout.addLayout(self.layoutButtons, 0, 2)
        self.layoutButtons.addWidget(self.saveProcessButton)
        self.layoutButtons.addWidget(self.loadProcessButton)
        self.layoutButtons.addWidget(self.editProcessButton)
        self.layoutButtons.addWidget(self.clearProcessButton)
        self.setLayout(self.layout)

        #connect signals
        #connect edit GUI signals
        self.editCancelButton.clicked.connect(self.tableWidget.clear)
        self.editCancelButton.clicked.connect(self.processEditWidget.hide)
        self.editOkButton.clicked.connect(self.writeEditedDatatoProcess)

        #connect Main GUI signals
        self.selectorWidget.itemDoubleClicked.connect(self.translateTreeToList)
        self.listWidget.itemDoubleClicked.connect(self.deleteListItem)
        self.listWidget.itemClicked.connect(self.setActivatedItem)
        self.generateXlsButton.clicked.connect(self.writeToFile)
        self.loadProcessButton.clicked.connect(self.loadProcess)
        self.saveProcessButton.clicked.connect(self.saveProcess)
        self.editProcessButton.clicked.connect(self.editProcess)
        self.clearProcessButton.clicked.connect(self.listWidget.clear)

    ###additional process commands and functions
    def writeToFile(self):
        """writes content of list to iostream for xls translation"""
        excelname = QFileDialog.getSaveFileName(None, "Generate Excel-File", self.readIni["DEFAULT"]["defaultSavePath"], "Excel File (*.xlsx)")
        if excelname[0]:  # user pressed ok
            file = io.StringIO()
            for i in range(0, self.listWidget.count()):
                file.write("%s\n" % self.listWidget.item(i).whatsThis())
            convertTXTtoXLS(file, excelname[0])
            file.close()
    def insertCustomHeading(self):
        """function to insert custom heading into xls"""
        inputHeader = QInputDialog.getText(None, "ProcessBuilder", "Insert Heading:", QLineEdit.Normal)
        if inputHeader[1]:
            tempListItem = QListWidgetItem(self.listWidget)
            tempListItem.setWhatsThis("=" + inputHeader[0] + "\t§header")
            tempListItem.setText("Heading: " + inputHeader[0])

    #translates QTreeWidgetItems to QListWidgetItems and parses additional commands
    def translateTreeToList(self, item, column):
        """translates items from QTreeViewWidget to QListWidget items"""
        if item.text(1) == "COMMAND":
            if item.text(2) == "HEADING": self.insertCustomHeading()

        elif not item.text(1) == "":
            tempListItem = QListWidgetItem(self.listWidget)
            tempListItem.setWhatsThis(item.text(1))
            tempListItem.setText(item.text(2) + " -> " + item.text(0))



    def deleteListItem(self, item):
        """deletes items from QListWidget"""
        self.listWidget.takeItem(self.listWidget.row(item))

    def saveProcess(self):
        """saves process to file"""
        saveFilename = QFileDialog.getSaveFileName(None, "Save Process", sys.path[0], "Process File (*.pro)")
        if saveFilename[0]:
            file = open(saveFilename[0], "w", encoding="UTF-8-sig")
            for i in range(0, self.listWidget.count()):
                file.write("%s\t%s\n" % (self.listWidget.item(i).text(), self.listWidget.item(i).whatsThis()))
            file.close()

    def loadProcess(self):
        """load process from file"""
        inputFilename = QFileDialog.getOpenFileName(None, "Open Process", self.readIni["DEFAULT"]["defaultProcessPath"], "Process File(*.pro)")
        if inputFilename[0]:
            with open(inputFilename[0], encoding="UTF-8-sig") as file:
                for line in file:
                    if "->" in line:
                        tmp = line.split("\t")
                        tmpComm = "\t".join(tmp[1:])
                        tempListItem = QListWidgetItem(self.listWidget)
                        tempListItem.setText(tmp[0])
                        tempListItem.setWhatsThis(tmpComm)
                    else:
                        tmpComm += line
                        tempListItem.setWhatsThis(tmpComm)

    def setActivatedItem(self, item):
        """saves selected item in QListWidget"""
        self.currentListItem = item

    def editProcess(self):
        """function to edit process templates on the fly"""
        if isinstance(self.currentListItem, QListWidgetItem):
            if self.currentListItem.whatsThis()[0] == ">":
                filePath = self.currentListItem.whatsThis().strip(">") + ".txt"
                file = open(filePath, "r", encoding="UTF-8-sig")
            elif self.currentListItem.whatsThis()[0] == "=":
                file = self.currentListItem.whatsThis().splitlines()
            readContent = []
            for line in file:
                readContent.append(line.split("|"))
            if "=" in readContent[0][0]:  # split format line
                readContent[0] = readContent[0][0].split(u"\u00A7")  # u"\u00A7" utf-8 code for paragraph
                readContent[0][1] = u"\u00A7" + readContent[0][1]
            self.tableWidget.setColumnCount(2)
            self.tableWidget.setRowCount(len(readContent))
            tabletWidgetItems = []
            for row in enumerate(readContent):
                for cell in enumerate(row[1]):
                    tempTableItem = QTableWidgetItem(readContent[row[0]][cell[0]].strip())
                    tabletWidgetItems.append(tempTableItem)
                    self.tableWidget.setItem(row[0], cell[0], tempTableItem)
            self.tableWidget.horizontalHeader().resizeSection(0, 150)
            self.processEditWidget.setWindowTitle("Edit " + self.currentListItem.text())
            self.processEditWidget.show()

    def writeEditedDatatoProcess(self):
        """function to write edited cells back to QListWidgetItem"""
        newCommandString = "" + self.tableWidget.item(0, 0).text() + "\t" + self.tableWidget.item(0, 1).text() + "\n"
        for row in range(1, self.tableWidget.rowCount()):
            for col in range(self.tableWidget.columnCount()):
                newCommandString += self.tableWidget.item(row, col).text()
                if col == 0: newCommandString += "\t|\t"
                elif col == 1: newCommandString += "\n"
        if "CUSTOM" in self.currentListItem.text():
            pass
        else:
            self.currentListItem.setText("CUSTOM " + self.currentListItem.text())
        self.currentListItem.setWhatsThis(newCommandString)
        self.processEditWidget.hide()

#mandatory gui lines
app = QApplication(sys.argv)
form = ProcessBuilderGui()
form.show()
app.exec_()