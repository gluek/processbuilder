#!python
#name: processBuilder_v0.2
#author: Gerrit Lükens
#python 3.4.1

#imports
import sys
import os
import configparser
from txtToXlsWriter import TXTtoXLSConverter
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
            customCmdRoot = QTreeWidgetItem(self)
            customCmdRoot.setText(0, "Extra")
            customCmdHeading = QTreeWidgetItem(customCmdRoot)
            customCmdHeading.setText(0, "Insert Heading")
            customCmdHeading.setText(1, "COMMAND")
            customCmdHeading.setText(2, "HEADING")
        createCustomCommands()

class ProcessBuilderGui(QDialog):
    def __init__(self, parent=None):
        super(ProcessBuilderGui, self).__init__(parent)

        self.setWindowTitle("Process Builder")
        readIni = configparser.ConfigParser()
        readIni.read("processBuilder.ini")
        filename = readIni["DEFAULT"]["txtFilename"]

        #create UI buttons
        generateXlsButton = QPushButton("Generate")
        saveProcessButton = QPushButton("Save")
        loadProcessButton = QPushButton("Load")
        clearProcessButton = QPushButton("Clear")

        #create tree structure with loaded process step templates
        selectorWidget = ProcessStepSelectorWidget()

        #create list for process flow
        listWidget = QListWidget()
        listWidget.setMovement(QListView.Snap)
        listWidget.setDragDropMode(QAbstractItemView.InternalMove)

    ###additional process commands and functions
        def writeToFile():
            excelname = QFileDialog.getSaveFileName(None, "Generate Excel-File", "C:\\", "Excel File (*.xlsx)")
            if excelname[0]: #user pressed ok
                file = open(filename, "w", encoding='UTF-8')
                for i in range(0, listWidget.count()):
                    file.write("%s\n" % listWidget.item(i).whatsThis())
                file.close()
                TXTtoXLSConverter.convertTXTtoXLS(filename, excelname[0])
        def insertCustomHeading():
            inputHeader = QInputDialog.getText(None, "ProcessBuilder", "Insert Heading:", QLineEdit.Normal)
            if inputHeader[1]:
                tempListItem = QListWidgetItem(listWidget)
                tempListItem.setWhatsThis("=" + inputHeader[0] + "\t§header")
                tempListItem.setText("Heading: " + inputHeader[0])

        #translates QTreeWidgetItems to QListWidgetItems and parse additional commands
        def translateTreeToList(item, column):
            if item.text(1) == "COMMAND":
                if item.text(2) == "HEADING": insertCustomHeading()

            elif not item.text(1) == "":
                tempListItem = QListWidgetItem(listWidget)
                tempListItem.setWhatsThis(item.text(1))
                tempListItem.setText(item.text(2) + " -> " + item.text(0))



        #func to delete elements inside QListWidget
        def deleteListItem(item):
            listWidget.takeItem(listWidget.row(item))

        #implemente save and load button functionality
        def saveProcess():
            saveFilename = QFileDialog.getSaveFileName(None, "Save Process", sys.path[0], "Process File (*.pro)")
            if saveFilename[0]:
                file = open(saveFilename[0], "w", encoding='UTF-8')
                for i in range(0, listWidget.count()):
                    file.write("%s\t%s\n" % (listWidget.item(i).text(), listWidget.item(i).whatsThis()))
                file.close()

        def loadProcess():
            inputFilename = QFileDialog.getOpenFileName(None, "Open Process", readIni["DEFAULT"]["defaultProcessPath"], "Process File(*.pro)")
            if inputFilename[0]:
                with open(inputFilename[0], encoding='UTF-8') as file:
                    for line in file:
                        tmp = line.replace("\n","").split("\t")
                        tempListItem = QListWidgetItem(listWidget)
                        tempListItem.setText(tmp[0])
                        tempListItem.setWhatsThis(tmp[1])

        #connect signals
        selectorWidget.itemDoubleClicked.connect(translateTreeToList)
        listWidget.itemDoubleClicked.connect(deleteListItem)
        generateXlsButton.clicked.connect(writeToFile)
        loadProcessButton.clicked.connect(loadProcess)
        saveProcessButton.clicked.connect(saveProcess)
        clearProcessButton.clicked.connect(listWidget.clear)

        #set layout
        layout = QGridLayout()
        layoutButtons = QVBoxLayout()
        layout.addWidget(selectorWidget, 0, 0)
        layout.addWidget(listWidget, 0, 1)
        layout.addWidget(generateXlsButton, 1, 0)
        layout.addLayout(layoutButtons, 0, 2)
        layoutButtons.addWidget(saveProcessButton)
        layoutButtons.addWidget(loadProcessButton)
        layoutButtons.addWidget(clearProcessButton)
        self.setLayout(layout)

#mandatory gui lines
app = QApplication(sys.argv)
form = ProcessBuilderGui()
form.show()
app.exec_()