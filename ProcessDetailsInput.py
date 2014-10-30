import sys
from PySide.QtGui import *
from PySide.QtCore import *
import win32com.client as win32
win32.gencache.is_readonly = False
win32.gencache.Rebuild()

class ProcessDetailsInputDialog(QDialog):
    def __init__(self, userListPath, processTypesPath, runSheetPath, parent=None):
        super(ProcessDetailsInputDialog, self).__init__(parent)
        self.runSheetPath = runSheetPath
        self.setWindowTitle("Process Details")
        self.setFixedHeight(100)
        #open excel document
        self.excel = win32.gencache.EnsureDispatch("Excel.Application")
        self.wb = self.excel.Workbooks.Open(self.runSheetPath)
        self.ws = self.wb.Worksheets("Sheet1")

        #create user list dropdown
        userFile = open(userListPath + "\\users.ini", "r", encoding="UTF-8-sig")
        userFileData = userFile.read().split("\n")
        self.users = userFileData[1:]
        self.users.sort()
        self.userList = QComboBox()
        for item in enumerate(self.users):
            self.userList.insertItem(item[0],item[1])

        #create process types dropdown
        typesFile = open(processTypesPath + "\\processtypes.ini", "r", encoding="UTF-8-sig")
        typesFileData = typesFile.read().split("\n")
        self.processTypes = typesFileData[1:]
        self.processTypes.sort()
        self.processTypeList = QComboBox()
        for item in enumerate(self.processTypes):
            self.processTypeList.insertItem(item[0], item[1])

        #header labels
        self.headerProcessId = QLabel()
        self.headerProcessId.setText("Process ID:")
        self.headerProcessType = QLabel()
        self.headerProcessType.setText("Type:")
        self.headerUser = QLabel()
        self.headerUser.setText("User:")
        self.headerDescription = QLabel()
        self.headerDescription.setText("Description:")
        self.headerSamples = QLabel()
        self.headerSamples.setText("Samples:")

        #text edit
        self.lineEditDescription = QLineEdit()
        self.lineEditSamples = QLineEdit()
        self.lineEditSamples.setFixedWidth(100)


        #process id label
        self.currentProcessId = QLabel()
        self.currentProcessId.setText("%i" % self.getNextProcessId())

        #construct gui input dialog
        self.layout = QVBoxLayout()
        self.gridLayout = QGridLayout()
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.setAlignment(Qt.AlignRight)
        self.buttonOk = QPushButton("Ok")
        self.buttonOk.setFixedWidth(100)
        self.buttonCancel = QPushButton("Cancel")
        self.buttonCancel.setFixedWidth(100)
        self.buttonLayout.addWidget(self.buttonOk)
        self.buttonLayout.addWidget(self.buttonCancel)
        #set headers
        self.gridLayout.addWidget(self.headerProcessId, 0, 0)
        self.gridLayout.addWidget(self.headerProcessType, 0, 1)
        self.gridLayout.addWidget(self.headerUser, 0, 2)
        self.gridLayout.addWidget(self.headerDescription, 0, 3)
        self.gridLayout.addWidget(self.headerSamples, 0, 4)
        #set output input
        self.gridLayout.addWidget(self.currentProcessId, 1, 0)
        self.gridLayout.addWidget(self.userList, 1, 1)
        self.gridLayout.addWidget(self.processTypeList, 1, 2)
        self.gridLayout.addWidget(self.lineEditDescription, 1, 3)
        self.gridLayout.addWidget(self.lineEditSamples, 1, 4)
        self.layout.addLayout(self.gridLayout)
        self.layout.addLayout(self.buttonLayout)
        self.setLayout(self.layout)
        #connect signals
        self.buttonOk.clicked.connect(self.writeProcessDetails)
        self.buttonCancel.clicked.connect(self.hide)

    def getNextProcessId(self):
        """returns next free process id in run sheet"""
        #find next empty line
        i = 1
        while self.ws.Cells(i,1).Value is not None:
            i += 1
        #calc next free process id
        processId = int(self.ws.Cells(i-1, 1).Value + 1)
        return processId

    def getNextEmptyLine(self):
        """returns next index of next empty line in run sheet"""
        i = 1
        while self.ws.Cells(i, 1).Value is not None:
            i += 1
            print(self.ws.Cells(i, 1).Value)
        return i

    def writeProcessDetails(self):
        """writes inserted process details to run sheet"""
        line = self.getNextEmptyLine()
        self.ws.Cells(line, 1).Value = self.getNextProcessId()
        self.ws.Cells(line, 2).Value = self.processTypeList.currentText()
        self.ws.Cells(line, 3).Value = self.userList.currentText()
        self.ws.Cells(line, 4).Value = self.lineEditDescription.text()
        self.ws.Cells(line, 5).Value = self.lineEditSamples.text()
        self.wb.SaveAs(self.runSheetPath)
        self.excel.Application.Quit()
        self.hide()

#app = QApplication(sys.argv)
#form = ProcessDetailsInputDialog("Y:\GaN_Device\Laufzettel\ProcessBuilderLists", "Y:\GaN_Device\Laufzettel\ProcessBuilderLists")
#form.show()
#app.exec_()