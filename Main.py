
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
import Algorithms
import pathvalidate
import sys
import os
from PyQt6.QtWidgets import QWizard, QWizardPage, QWidget, QLabel, QFileDialog, QLineEdit, QProgressBar, QPushButton, QApplication, QVBoxLayout
from PyQt6.QtCore import Qt

def Process(directory:str, save_dir:str):
    WorkBook = Algorithms.openWorkbook(directory)
    WorkSheet = WorkBook[(Algorithms.detectAvailableSheet(WorkBook))]
    Result = Algorithms.sortDatas(WorkSheet)
    Result.save(save_dir)
    Result.close()

class MainWizard(QWizard):

    class WelcomePage(QWizardPage):
        def __init__(self, parent: QWidget | None = ...) -> None:
            super().__init__(parent)

            # self.Layout = QVBoxLayout()
            # self.label = QLabel("该向导将协助您智能地完成表格拆分工作.")
            # self.Layout.addWidget(self.label)
            # self.setLayout(self.Layout)

            self.setTitle("欢迎")
            self.setSubTitle("该向导将协助您智能地完成表格拆分工作.")

    class OriginPage(QWizardPage):
        def __init__(self, parent: QWidget | None = ...) -> None:
            super().__init__(parent)

            self.Layout = QVBoxLayout()
            self.label = QLabel("文件路径: ")
            self.label.setTextFormat(Qt.TextFormat.RichText)
            self.Layout.addWidget(self.label)

            self.lineEdit = QLineEdit()
            self.lineEdit.setPlaceholderText("输入文件路径...(或者将文件拖拽到此输入栏中)")
            self.lineEdit.textChanged.connect(self.checkDirValid)
            self.registerField("directory", self.lineEdit)
            self.Layout.addWidget(self.lineEdit)

            self.fileDialogButton = QPushButton("浏览...")
            self.fileDialogButton.setMaximumWidth(200)
            self.fileDialogButton.clicked.connect(self.browseDirectory)
            self.Layout.addWidget(self.fileDialogButton)

            self.setLayout(self.Layout)
            self.setAcceptDrops(True)
            self.setTitle("选择表格文件")
            self.setSubTitle("输入要拆分的单个表格文件的路径, 或者直接将这个文件拖入地址栏中. 这个文件必须是以.xlsx结尾的.")

        def isComplete(self) -> bool:
            return os.path.exists(self.field("directory")) and os.path.isfile(self.field("directory")) and self.field("directory").endswith(".xlsx")
        
        def checkDirValid(self):
            if not self.isComplete():
                self.label.setText("文件路径: <font color='red'>(路径无效)</font>")
            else:
                self.label.setText("文件路径: ")
            self.completeChanged.emit()

        def browseDirectory(self):
            excelfile, _ = QFileDialog.getOpenFileName(self, "选择文件", None, "Microsoft Excel 2010 Workbook File (*.xlsx)")
            if excelfile:
                self.lineEdit.setText(excelfile)

        def dragEnterEvent(self, a0: QDragEnterEvent | None) -> None:
            if a0.mimeData().hasUrls():
                a0.accept()

        def dropEvent(self, a0: QDropEvent | None) -> None:
            if a0.mimeData().hasUrls():
                urls = a0.mimeData().urls()
                if len(urls) > 1:
                    self.label.setText("文件路径: <font color='red'>(仅支持单个文件的拖拽)</font>")
                else:
                    self.lineEdit.setText(urls[0].toString().removeprefix("file:///"))

    class SavePage(QWizardPage):
        def __init__(self, parent: QWidget | None = ...) -> None:
            super().__init__(parent)

            self.Layout = QVBoxLayout()
            self.label = QLabel("保存到文件: ")
            self.label.setTextFormat(Qt.TextFormat.RichText)
            self.Layout.addWidget(self.label)

            self.lineEdit = QLineEdit()
            self.lineEdit.setPlaceholderText("点击\"浏览\"选择保存到的路径...")
            self.lineEdit.setReadOnly(True)
            self.lineEdit.textChanged.connect(self.checkDirValid)
            self.registerField("directory2", self.lineEdit)
            self.Layout.addWidget(self.lineEdit)

            self.fileDialogButton = QPushButton("浏览...")
            self.fileDialogButton.setMaximumWidth(200)
            self.fileDialogButton.clicked.connect(self.browseDirectory)
            self.Layout.addWidget(self.fileDialogButton)

            self.setLayout(self.Layout)
            self.setTitle("保存到表格文件")
            self.setSubTitle("选择要保存到的单个表格文件的路径.")

        def isComplete(self) -> bool:
            return os.path.isdir(os.path.dirname(self.field("directory2"))) and self.field("directory2").endswith(".xlsx") and not(os.path.isdir(self.field("directory2")))
        
        def checkDirValid(self):
            if not self.isComplete():
                self.label.setText("保存到文件: <font color='red'>(路径无效)</font>")
            elif os.path.exists(self.field("directory2")):
                self.label.setText("保存到文件: <font color='red'>(文件已存在, 将被替换)</font>")
            else:
                self.label.setText("保存到文件: ")
            self.completeChanged.emit()

        def browseDirectory(self):
            excelfile, _ = QFileDialog.getSaveFileName(self, "选择文件", None, "Microsoft Excel 2010 Workbook File (*.xlsx)")
            if excelfile:
                self.lineEdit.setText(excelfile)

    class OperationPage(QWizardPage):
        def __init__(self, parent: QWidget | None = ...) -> None:
            super().__init__(parent)
            
            self.Layout = QVBoxLayout()
            self.setLayout(self.Layout)

        def initializePage(self) -> None:
            self.success: bool
            try:
                Process(self.field("directory"), self.field("directory2"))
                self.success = True
            except Exception as err:
                self.errormessage = err
                self.success = False

            if self.success:
                self.setTitle("操作成功完成")
                self.setSubTitle(f"表格已被成功拆分并保存到了: {self.field('directory2')}")
            else:
                self.setTitle("发生错误")
                self.setSubTitle(f"表格在拆分时遇到了问题. 请检查表格格式与路径是否正确. 错误信息: {self.errormessage}")

    def __init__(self):
        super().__init__()

        self.setPage(0, self.WelcomePage(self))
        self.setPage(1, self.OriginPage(self))
        self.setPage(2, self.SavePage(self))
        self.setPage(3, self.OperationPage(self))

        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)
        self.setWindowTitle('表格拆分工具')

        self.setOption(QWizard.WizardOption.NoBackButtonOnStartPage)
        self.setOption(QWizard.WizardOption.NoBackButtonOnLastPage)
        self.setOption(QWizard.WizardOption.NoCancelButton)

        self.setButtonText(QWizard.WizardButton.NextButton, '下一步')
        self.setButtonText(QWizard.WizardButton.BackButton, '上一步')
        self.setButtonText(QWizard.WizardButton.FinishButton, '完成')


app = QApplication(["ScoreSpliter"])
wizard = MainWizard()
wizard.show()
sys.exit(app.exec())