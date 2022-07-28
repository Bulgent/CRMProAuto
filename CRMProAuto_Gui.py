import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from CRMtoEXCEL_Gui import main_CtoE
from AutoPeakCheckkai_Gui import main_APC


class CRMProAuto(QWidget):

    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        #手法選択の生成(横に並べる)
        self.rbtn1 = QRadioButton('AutoPeakCheck')
        self.rbtn2 = QRadioButton('CRMtoEXCEL')

        self.btnLayout = QHBoxLayout()
        self.btnLayout.addWidget(self.rbtn1)
        self.btnLayout.addWidget(self.rbtn2)
        #ボタングループに纏める
        self.btngroup1 = QButtonGroup()
        self.btngroup1.addButton(self.rbtn1)
        self.btngroup1.addButton(self.rbtn2)
        
        self.layout = QVBoxLayout()
        self.layout.addLayout(self.btnLayout)

        #ラジオボタンの実行
        self.rbtn1.toggled.connect(self.switchAPC)
        self.rbtn2.toggled.connect(self.switchCtoE)
        
        self.setGeometry(400, 300, 300, 170)

        self.setLayout(self.layout)
        self.setWindowTitle('CRMProAuto')

        self.show()

    #APCのレイアウト########
    def switchAPC(self):
        radioBtn = self.sender()
        if radioBtn.isChecked():
            #PeakTime
            self.labelPeakTime = QLabel('PeakTime')
            self.labelPeakTime.setStyleSheet("font-weight: bold")
            self.labelCH1 = QLabel('CH1')
            self.textBoxCH1 = QLineEdit(self)
            self.textBoxCH1.setText("6.2")
            self.labelCH2 = QLabel('CH2')
            self.textBoxCH2 = QLineEdit(self)
            self.textBoxCH2.setText("2.3")
            self.checkCH = QCheckBox('Use only CH1', self) #チェックボックス
            self.checkCH.stateChanged.connect(self.removeCH2)
            #横に並べる(CH1)
            self.CH1Layout = QHBoxLayout()
            self.CH1Layout.addWidget(self.labelCH1)
            self.CH1Layout.addWidget(self.textBoxCH1)
            #横に並べる(CH2)
            self.CH2Layout = QHBoxLayout()
            self.CH2Layout.addWidget(self.labelCH2)
            self.CH2Layout.addWidget(self.textBoxCH2)
            #CHのBox作成
            self.channelLayout = QVBoxLayout()
            self.channelLayout.addLayout(self.CH1Layout)
            self.channelLayout.addLayout(self.CH2Layout)
            #縦に並べる(Peak)
            self.PeakLayout = QVBoxLayout()
            self.PeakLayout.addWidget(self.labelPeakTime)
            self.PeakLayout.addLayout(self.channelLayout)
            self.PeakLayout.addWidget(self.checkCH)

            #Option
            self.labelSelect = QLabel('select Time') #チェックボックス
            self.labelSelect.setStyleSheet("font-weight: bold")
            self.rbtn9m = QRadioButton('9 minute') 
            self.rbtn7m = QRadioButton('7 minute(default)') 
            self.btngroup2 = QButtonGroup()
            self.btngroup2.addButton(self.rbtn9m)
            self.btngroup2.addButton(self.rbtn7m)

            self.labelPath = QLabel("If you don't use default folder, please check.")
            self.labelPath.setStyleSheet("font-weight: bold")
            self.checkPath = QCheckBox('Use Different Folder', self)
            self.checkPath.stateChanged.connect(self.addPath)
            self.rbtn7m.setChecked(True)
            #縦に並べる
            self.OptionLayout = QVBoxLayout()
            self.OptionLayout.addWidget(self.labelSelect)
            self.OptionLayout.addWidget(self.rbtn9m)
            self.OptionLayout.addWidget(self.rbtn7m)
            self.OptionLayout.addWidget(self.labelPath)
            self.OptionLayout.addWidget(self.checkPath)

            #pathの箱(消えても同じ場所に出るように)
            self.FolderLayout = QVBoxLayout()

            #実行
            self.buttonEX_APC = QPushButton('Execute', self)
            self.buttonEX_APC.clicked.connect(self.exAPC)

            #APC全体のレイアウト
            self.subAPCLayout = QHBoxLayout()
            self.subAPCLayout.addLayout(self.PeakLayout)
            self.subAPCLayout.addLayout(self.OptionLayout)
            self.APCLayout = QVBoxLayout()
            self.APCLayout.addLayout(self.subAPCLayout)
            self.APCLayout.addLayout(self.FolderLayout)
            self.APCLayout.addWidget(self.buttonEX_APC)

            #挿入
            self.layout.addLayout(self.APCLayout)

        else:
            self.clear_item(self.APCLayout)

    #CtoEのレイアウト########
    def switchCtoE(self):
        radioBtn = self.sender()
        if radioBtn.isChecked():
            #Excel名入力
            self.labelExcelName = QLabel('Excel Name')
            self.labelExcelName.setStyleSheet("font-weight: bold")
            self.textBoxExcelName = QLineEdit(self)
            #横に並べる
            self.Layout_Name_CtoE = QHBoxLayout()
            self.Layout_Name_CtoE.addWidget(self.labelExcelName)
            self.Layout_Name_CtoE.addWidget(self.textBoxExcelName)

            #Option
            self.labelOption_CtoE = QLabel("Option")
            self.labelOption_CtoE.setStyleSheet("font-weight: bold")
            self.checkCH_CtoE = QCheckBox('Use only CH1', self)
            self.checkPath_CtoE = QCheckBox('Use Different Folder', self)
            self.checkPath_CtoE.stateChanged.connect(self.addPath_CtoE)
            #縦に並べる
            self.Layout_Option_CtoE = QVBoxLayout()
            self.Layout_Option_CtoE.addWidget(self.labelOption_CtoE)
            self.Layout_Option_CtoE.addWidget(self.checkCH_CtoE)
            self.Layout_Option_CtoE.addWidget(self.checkPath_CtoE)
            

            #Pathの箱
            self.FolderLayout_CtoE = QVBoxLayout()
            #実行
            self.buttonEX_CtoE = QPushButton('Execute', self)
            self.buttonEX_CtoE.clicked.connect(self.exCtoE)
            #CtoE全体のレイアウト      
            self.CtoELayout = QVBoxLayout() 
            self.CtoELayout.addLayout(self.Layout_Name_CtoE)
            self.CtoELayout.addLayout(self.Layout_Option_CtoE)
            self.CtoELayout.addLayout(self.FolderLayout_CtoE)
            self.CtoELayout.addWidget(self.buttonEX_CtoE)

            #挿入
            self.layout.addLayout(self.CtoELayout)
        else:
            self.clear_item(self.CtoELayout)

    #レイアウトの削除
    def clear_item(self, item):
        if hasattr(item, "layout"):
            if callable(item.layout):
                layout = item.layout()
        else:
            layout = None

        if hasattr(item, "widget"):
            if callable(item.widget):
                widget = item.widget()
        else:
            widget = None

        if widget:
            widget.setParent(None)
        elif layout:
            for i in reversed(range(layout.count())):
                self.clear_item(layout.itemAt(i))

    #channel2の出現
    def removeCH2(self):
        if(not self.checkCH.isChecked()):
            #CH2生成
            self.labelCH2 = QLabel('CH2')
            self.textBoxCH2 = QLineEdit(self)
            self.textBoxCH2.setText("2.3")
            #横に並べる(CH2)
            self.CH2Layout = QHBoxLayout()
            self.CH2Layout.addWidget(self.labelCH2)
            self.CH2Layout.addWidget(self.textBoxCH2)
            #元のレイヤに加える
            self.channelLayout.addLayout(self.CH2Layout)

        else:
            try:
                self.clear_item(self.CH2Layout)
            except:
                pass
    
    #pathの追加
    def addPath(self):
        if(self.checkPath.isChecked()):
            #Folder Name
            self.labelWritePath = QLabel('Write Folder Path')
            self.labelWritePath.setStyleSheet("font-weight: bold")
            self.labelInputPath = QLabel('Input') 
            self.textBoxInputPath = QLineEdit(self)
            self.labelOutputPath = QLabel('Output')
            self.textBoxOutputPath = QLineEdit(self)
            #横に並べる(input)
            self.InputLayout = QHBoxLayout()
            self.InputLayout.addWidget(self.labelInputPath)
            self.InputLayout.addWidget(self.textBoxInputPath)
            #横に並べる(output)
            self.OutputLayout = QHBoxLayout()
            self.OutputLayout.addWidget(self.labelOutputPath)
            self.OutputLayout.addWidget(self.textBoxOutputPath)
            #縦に並べる
            self.subFolderLayout = QVBoxLayout()
            self.subFolderLayout.addWidget(self.labelWritePath)
            self.subFolderLayout.addLayout(self.InputLayout)
            self.subFolderLayout.addLayout(self.OutputLayout)
            #挿入
            self.FolderLayout.addLayout(self.subFolderLayout)
        else:
            try:
                self.clear_item(self.subFolderLayout)
            except:
                pass

    def addPath_CtoE(self):
        if(self.checkPath_CtoE.isChecked()):
            #Folder Name
            self.labelWritePath_CtoE = QLabel('Write Folder Path')
            self.labelWritePath_CtoE.setStyleSheet("font-weight: bold")
            self.labelInputPath_CtoE = QLabel('Input') 
            self.textBoxInputPath_CtoE = QLineEdit(self)
            self.labelOutputPath_CtoE = QLabel('Output')
            self.textBoxOutputPath_CtoE = QLineEdit(self)
            #横に並べる(input)
            self.InputLayout_CtoE = QHBoxLayout()
            self.InputLayout_CtoE.addWidget(self.labelInputPath_CtoE)
            self.InputLayout_CtoE.addWidget(self.textBoxInputPath_CtoE)
            #横に並べる(output)
            self.OutputLayout_CtoE = QHBoxLayout()
            self.OutputLayout_CtoE.addWidget(self.labelOutputPath_CtoE)
            self.OutputLayout_CtoE.addWidget(self.textBoxOutputPath_CtoE)
            #縦に並べる
            self.subFolderLayout_CtoE = QVBoxLayout()
            self.subFolderLayout_CtoE.addWidget(self.labelWritePath_CtoE)
            self.subFolderLayout_CtoE.addLayout(self.InputLayout_CtoE)
            self.subFolderLayout_CtoE.addLayout(self.OutputLayout_CtoE)
            #挿入
            self.FolderLayout_CtoE.addLayout(self.subFolderLayout_CtoE)
        else:
            try:
                self.clear_item(self.subFolderLayout_CtoE)
            except:
                pass

    def exAPC(self):
        #CHで場合分け
        if not self.checkCH.isChecked():
            ch1 = self.textBoxCH1.text()
            ch2 = self.textBoxCH2.text()
            #minuteで場合分け
            if self.rbtn7m.isChecked():
                #Pathで場合分け
                if(self.checkPath.isChecked()):
                    main_APC(ch1, ch2, True, False, True, self.textBoxInputPath.text(), self.textBoxOutputPath.text())
                else:
                    main_APC(ch1, ch2, True, False, False, None, None)
            else:
                #Pathで場合分け
                if(self.checkPath.isChecked()):
                    main_APC(ch1, ch2, True, True, True, self.textBoxInputPath.text(), self.textBoxOutputPath.text())
                else:
                    main_APC(ch1, ch2, True, True, False, None, None)
        else:
            ch1 = self.textBoxCH1.text()
            #minuteで場合分け
            if self.rbtn7m.isChecked():
                #Pathで場合分け
                if(self.checkPath.isChecked()):
                    main_APC(ch1, None, False, False, True, self.textBoxInputPath.text(), self.textBoxOutputPath.text())
                else:
                    main_APC(ch1, None, False, False, False, None, None)
            else:
                #Pathで場合分け
                if(self.checkPath.isChecked()):
                    main_APC(ch1, None, False, True, True, self.textBoxInputPath.text(), self.textBoxOutputPath.text())
                else:
                    main_APC(ch1, None, False, True, False, None, None)
        QMessageBox.information(None, "Information", "Complete!!", QMessageBox.Yes)
    
    def exCtoE(self):
        #CHで場合分け
        if self.checkCH_CtoE.isChecked():
            #Pathの場合分け
            if self.checkPath_CtoE.isChecked():
                main_CtoE(self.textBoxExcelName.text(), False, True, self.textBoxInputPath_CtoE.text(), self.textBoxOutputPath_CtoE.text())
            else:
                main_CtoE(self.textBoxExcelName.text(), False, False, None, None)
        else:
            #Pathの場合分け
            if self.checkPath_CtoE.isChecked():
                main_CtoE(self.textBoxExcelName.text(), True, True, self.textBoxInputPath_CtoE.text(), self.textBoxOutputPath_CtoE.text())
            else:
                main_CtoE(self.textBoxExcelName.text(), True, False, None, None)
        QMessageBox.information(None, "Information", "Complete!!", QMessageBox.Yes)


if __name__ == '__main__':    
    app = QCoreApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    ex = CRMProAuto()
    sys.exit(app.exec_())