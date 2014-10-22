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

class ProcessBuilderGui(QDialog):
    def __init__(self, parent=None):
        super(ProcessBuilderGui, self).__init__(parent)

        self.setWindowTitle("Process Builder")
        self.readIni = configparser.ConfigParser()
        self.readIni.read("processBuilder.ini")

        #create UI buttons
        self.generateXlsButton = QPushButton("Generate")
        self.saveProcessButton = QPushButton("Save")
        self.loadProcessButton = QPushButton("Load")
        self.editProcessButton = QPushButton("Edit")
        self.clearProcessButton = QPushButton("Clear")

        #create tree structure with loaded process step templates
        self.selectorWidget = ProcessStepSelectorWidget()

        #create list for process flow
        self.listWidget = QListWidget()
        self.listWidget.setMovement(QListView.Snap)
        self.listWidget.setDragDropMode(QAbstractItemView.InternalMove)
        self.currentListItem = 0
        self.exchangeItem = 0

        #initialize process edit window
        self.processEditWidget = QWidget()


        #connect signals
        self.selectorWidget.itemDoubleClicked.connect(self.translateTreeToList)
        self.listWidget.itemDoubleClicked.connect(self.deleteListItem)
        self.listWidget.itemClicked.connect(self.saveActivatedItem)
        self.generateXlsButton.clicked.connect(self.writeToFile)
        self.loadProcessButton.clicked.connect(self.loadProcess)
        self.saveProcessButton.clicked.connect(self.saveProcess)
        self.editProcessButton.clicked.connect(self.editProcess)
        self.clearProcessButton.clicked.connect(self.listWidget.clear)


        #set layout
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

    ###additional process commands and functions
    def writeToFile(self):
        """writes content of list to iostream for xls translation"""
        excelname = QFileDialog.getSaveFileName(None, "Generate Excel-File", r"C:\Users\luekens\PycharmProjects\ProcessBuilder", "Excel File (*.xlsx)")
        if excelname[0]: #user pressed ok
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

    #translates QTreeWidgetItems to QListWidgetItems and parse additional commands
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
            file = open(saveFilename[0], "w", encoding='UTF-8')
            for i in range(0, self.listWidget.count()):
                file.write("%s\t%s\n" % (self.listWidget.item(i).text(), self.listWidget.item(i).whatsThis()))
            file.close()

    def loadProcess(self):
        """load process from file"""
        inputFilename = QFileDialog.getOpenFileName(None, "Open Process", self.readIni["DEFAULT"]["defaultProcessPath"], "Process File(*.pro)")
        if inputFilename[0]:
            with open(inputFilename[0], encoding='UTF-8') as file:
                for line in file:
                    tmp = line.replace("\n","").split("\t")
                    tempListItem = QListWidgetItem(self.listWidget)
                    tempListItem.setText(tmp[0])
                    tempListItem.setWhatsThis(tmp[1])

    def saveActivatedItem(self, item):
        """saves selected item in QListWidget"""
        self.currentListItem = item
        print(self.currentListItem.text())

    def editProcess(self):
        """function to edit process templates on the fly"""
        #TODO implement functionality
        if isinstance(self.currentListItem, QListWidgetItem):
            filePath = self.currentListItem.whatsThis().strip(">") + ".txt"
            file = open(filePath, "r", encoding="UTF-8")
            readContent = []
            for line in file:
                readContent.append(line.split("|"))
            if "=" in readContent[0][0]:  # split format line
                readContent[0] = readContent[0][0].split("§")
                readContent[0][1] = "§" + readContent[0][1]
            #self.processEditWidget.setWindowTitle("Edit " + self.currentListItem.text())
            #self.processEditWidget.show()
            print(readContent)


#mandatory gui lines
app = QApplication(sys.argv)
form = ProcessBuilderGui()
form.show()
app.exec_()