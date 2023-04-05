#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project ：installer_new_pyqt
@File    ：test1.py
@IDE     ：PyCharm
@Author  ：Installation
@Date    ：2023/4/3 13:47
"""

import os.path
import zipfile
import winreg
import psutil
import winsound
import ctypes
import Mainwindow
import MQMessageBox
import sys
import json
from typing import Union, Optional, Tuple
from PyQt5 import QtWidgets
from PyQt5 import QtGui
from PyQt5 import QtCore
from PyQt5.QtCore import QFile
from losder.Losder import Losder


def exZipFIle(file, outPutFile):
    with zipfile.ZipFile(file, 'r') as zipRef:
        zipRef.extractall(outPutFile)


def isAdministrator():
    return ctypes.windll.shell32.IsUserAnAdmin()


class MainWindow(QtWidgets.QFrame, Mainwindow.Ui_Mainwindow):
    def __init__(self) -> None:
        super().__init__()

        self.appExe = None
        self._endPos = None
        self._startPos = None
        self._isTracking = None
        self.allConfig = None

        self.version = str()
        self.allInfo = str()
        self.allInfoEnglish = str()
        self.author = str()
        self.zipFile = str()
        self.z7zipfile = str()
        self.needSize = str()
        self.website = str()
        self.QQ = str()
        self.appName = None
        self.mainApp = str()
        self.introduce = str()

        self.setupUi(self)

        # 无边框
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)

        # 去背景
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # 翻译
        self.trans = QtCore.QTranslator(self)

        # 读取配置
        self.readJsonConfig()

        # 连接按钮信号槽
        self.connectSlot()

        # 添加语言框
        self.Language_Choose.addItems(["简体中文", "English"])

        # 返回按钮
        self.Back.setVisible(False)

        # 安装按钮
        self.Install.setVisible(False)

        self.Error = {
            # 复制文件失败
            "copyError": self.tr("Failed to copy file"),
            # 文件损坏
            "zipError": self.tr("File corruption"),
            # 复制文件成功:{newFile}
            "copySuccess": self.tr("The copy files succeeded :{newFile}"),
            # 未知错误
            "unknownError": self.tr("Unknown error"),
            # 配置错误
            "configError": self.tr("There was an error with the configuration file"),
            # 界面索引错误
            "pageIndexError": self.tr("Page index error"),
            # 您必须同意协议
            "agreeAgreementInfos": self.tr("You must agree to the agreement"),
            # 选择的路径为空: {path}
            "pathNone": self.tr("The path selected is empty: {path}"),
            # 文件已存在
            "alreadyExists": self.tr("File already exists"),
            # 请使用管理员权限运行此程序
            "playAdmin": self.tr("Run this program with administrator privileges")
        }
        self.Info = {
            # 错误
            "error": self.tr("Error"),
            # 文件选择
            "fileSelection": self.tr("File selection"),
            # 信息
            "info": self.tr("info"),
            # 检测到此程序未以管理员身份运行，可能存在一些问题
            "admin": self.tr(
                "This program has been detected that you are not running as an administrator, which may have some problems"),
            # 你确定退出吗?
            "exit": self.tr("Are you sure you quit?")

        }
        # 检测管理员
        if isAdministrator() != 1:
            sendMessageBox.info(self.Info["info"], self.Info["admin"])

    def copyQrc(self, oldDir: str, newDir: str, fileName: Union[list[str], str]) -> dict[str, str]:
        """
        将qrc中的复制到newDor
        :param oldDir: 旧路径 (qrc)
        :param newDir: 新的路径
        :param fileName: 文件名字 list就遍历 str直接使用
        :return: 状态 Error:错误 Info:信息
        """
        returnDict = {
            "newFile": None,
            "oldFile": None,
            "Error": None,
            "Info": None
        }

        if isinstance(fileName, list):

            for name in fileName:
                oldFile = oldDir + name
                if not QFile.exists(oldFile):
                    newFile = os.path.join(newDir, name)
                    if QFile.copy(oldFile, newFile):
                        returnDict["newFile"] = newFile

                        returnDict["Info"] = self.Error["copySuccess"].format(newFile)
                    else:
                        returnDict["Error"] = self.Error["copyError"]
                else:

                    returnDict["Error"] = self.Error["zipError"]

                returnDict["oldFile"] = oldFile

        elif isinstance(fileName, str):
            newFile = os.path.join(newDir, fileName)
            oldFile = oldDir + fileName

            if QFile.exists(oldFile):
                if QFile.copy(oldFile, newFile):
                    returnDict["newFile"] = newFile
                    returnDict["Info"] = self.Error["copySuccess"].format(newFile=newFile)
                else:
                    returnDict["Error"] = self.Error["copyError"]
                returnDict["oldFile"] = oldFile
            else:
                returnDict["Error"] = self.Error["zipError"]
        else:
            returnDict["Error"] = self.Error["unknownError"]

        return returnDict

    def browserPath(self) -> None:
        """
        文件浏览
        :return: None
        """
        path = getFileSearch(title=self.Info["fileSelection"], mode="dir", parent=self)
        if path:
            self.path.setText(path[0])
        else:
            self.pathNone()

    def readJsonConfig(self) -> None:
        """
        读取Json
        :return: None
        """
        allConfig: dict = json.loads(open(".\\config\\config.json", "r", encoding="utf-8").read())
        keys: list = list(allConfig.keys())
        temp = os.path.isfile

        # 读取文件修改变量
        for key in keys:
            setattr(self, key, allConfig[key])

        self.checkNone(keys)

        # 7z.exe没有,但是7z后缀
        if self.z7zipfile == "" and os.path.splitext(self.zipFile)[1] == ".7z":
            pass
            # TODO

        if not temp(self.allInfo) and not temp(self.allInfoEnglish):
            self.configError()

        self.initUi()

    def changeLanguage(self) -> None:
        """
        语言切换
        :return: None
        """
        _app = QtWidgets.QApplication.instance()

        if self.Language_Choose.currentText() == "简体中文":
            self.trans.load(".\\QM\\Mainwindow.qm")
            _app.installTranslator(self.trans)
        if self.Language_Choose.currentText() == "English":
            self.appName[0], self.appName[1] = self.appName[1], self.appName[0]
            _app.removeTranslator(self.trans)

        self.retranslateUi(self)
        self.initUi()

    def initUi(self) -> None:
        """
        根据 JSON 格式化UI
        :return: None
        """
        for text in [self.title, self.welcome_title, self.info_install, self.all_info]:
            text.setPlainText(
                text.toPlainText().format(appName=self.appName[0], path=self.path.text()))

        self.TheAgreement.setHtml(self.introduce)

    def changeAllSize(self) -> None:
        """
        输入路径时改变总大小(盘符)
        :return: None
        """
        if os.path.isdir(self.path.text()):
            self.size_group.setPlainText(self.tr("""Space required:{needSize}
Space available:{allSize}""").format(
                allSize=f"""{psutil.disk_usage(os.path.splitdrive(self.path.text())[0]).total / (1024 * 1024 * 1024)
                }GB""",
                needSize=self.needSize))

    def checkNone(self, args) -> None:
        """
        传参判断是否为空
        :param args:
        :return: None
        """
        for var in args:
            if var is None:
                self.configError()

    def connectSlot(self) -> None:
        """
        连接按钮信号槽
        :return: None
        """
        # 最大化按钮:退出代码:0
        self.exit.clicked.connect(lambda: self.close())
        # 最小化按钮
        self.small.clicked.connect(lambda: mainwindow.showNormal())
        # Next按钮
        self.Next.clicked.connect(self.nextPage)
        # Back按钮
        self.Back.clicked.connect(self.backPage)
        # 初始化
        self.mainStack.setCurrentIndex(0)
        # 语言选择
        self.Language_Choose.currentIndexChanged.connect(self.changeLanguage)
        # 文夹输入框输入
        self.path.textChanged.connect(self.changeAllSize)
        # 浏览框
        self.browser.clicked.connect(self.browserPath)
        # 取消
        self.Cancel.clicked.connect(lambda: self.close())
        # 安装
        self.Install.clicked.connect(self.install)

    def backPage(self) -> None:
        """
        Back按钮
        :return: None
        """
        oldPage = self.mainStack.currentIndex()
        newPage = oldPage - 1

        if newPage < 0:
            self.pageIndexError()
            return

        if newPage == 0:
            self.Back.setVisible(False)

        if newPage < self.mainStack.count():
            self.Next.setVisible(True)
            self.Install.setVisible(False)

        self.mainStack.setCurrentIndex(newPage)

    def nextPage(self) -> None:
        """
        Next按钮
        :return: None
        """
        oldPage = self.mainStack.currentIndex()
        newPage = oldPage + 1

        # 如果不在第一个界面显示Back按钮
        if newPage > 0:
            self.Back.setVisible(True)
        if newPage > self.mainStack.count():
            self.pageIndexError()
            return

        if newPage == self.mainStack.count():
            # 最后一个界面改成安装函数
            self.Next.setVisible(False)
            self.Install.setVisible(True)

        if not self.agreeAgreement.isChecked() and newPage == 3:
            self.agreeAgreementInfo()
            return
        self.mainStack.setCurrentIndex(newPage)

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        """
        退出是否确定
        :param a0:
        :return:
        """
        if sendMessageBox.info(self.tr("exit?"), self.Info["exit"]):
            a0.accept()
        else:
            a0.ignore()

    def install(self) -> None:
        """
        安装
        :return: None
        """
        files = self.copyQrc(":/ZIP/", self.path.text(), self.zipFile)

        maybeError = ""

        newFile = files.get("newFile")
        error = files.get("Error")
        info = files.get("Info")

        if newFile:
            fileName, fileExt = os.path.splitext(newFile)
            exZipFIle(newFile, fileName)
            self.createShortcut()
            self.registrationProgram()

        else:
            maybeError += self.Error["alreadyExists"]

        if error:
            sendMessageBox.info(self.Info["error"], error + f":{maybeError}")

        if info:
            sendMessageBox.info(self.Info["info"], info)

        # TODO install

    def registrationProgram(self) -> None:
        """
        修改注册表在控制面板显示
        :return: None
        """
        try:
            key = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
            try:
                winreg.CreateKey(key,
                                 rf"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{self.appName[0]}")
            except PermissionError:
                sendMessageBox.info(self.Info["error"], self.Error["playAdmin"])
                return
            software = winreg.OpenKey(key, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{self.appName[0]}", 0,
                                      winreg.KEY_WRITE)
        except OSError as e:
            sendMessageBox.info(self.Info["error"], e.__str__())
            return

        if software is not None:
            winreg.SetValueEx(software, "DisplayName", 0, winreg.REG_SZ, self.appName[0])
            winreg.SetValueEx(software, "DisplayIcon", 0, winreg.REG_SZ, self.appExe)
            winreg.SetValueEx(software, "UninstallString", 0, winreg.REG_SZ, self.path.text() + "Uninstall.exe")
            winreg.SetValueEx(software, "DisplayVersion", 0, winreg.REG_SZ, self.version)
            winreg.SetValueEx(software, "Publisher", 0, winreg.REG_SZ, self.author)
            winreg.SetValueEx(software, "InstallLocation", 0, winreg.REG_SZ, os.path.splitext(self.appExe)[0])
            winreg.SetValueEx(software, "InstallType", 0, winreg.REG_SZ, "Install")
            software.Close()

    def createShortcut(self) -> None:
        """
        使用bat脚本创建快捷方式
        :return: None
        """
        self.appExe = os.path.join(self.path.text(), os.path.splitext(self.zipFile)[0]) + f"\\{self.mainApp}"
        if os.path.isfile(self.appExe):
            os.system(
                f"""mshta VBScript:Execute("Set a=CreateObject(""WScript.Shell""):Set b=a.CreateShortcut(a.SpecialFolders(""Desktop"") & ""\\{self.appName[0]}.lnk""):b.TargetPath=""{self.appExe}"":b.WorkingDirectory=""%~dp0"":b.Save:close") """
            )
        else:
            self.unKnowError()

    def unKnowError(self) -> None:
        sendMessageBox.info(self.Info["error"], self.Error["unknownError"])

    def agreeAgreementInfo(self) -> None:
        sendMessageBox.info(self.Info["info"], self.Error["agreeAgreementInfos"])

    def pathNone(self) -> None:
        sendMessageBox.info(self.Info["Error"], self.Error["pathNone"])

    def pageIndexError(self) -> None:
        sendMessageBox.info(self.Info["Error"], self.Error["pageIndexError"])

    def configError(self) -> None:
        # 配置文件出错
        sendMessageBox.info(self.Info["Error"], self.Error["configError"])
        sys.exit(-1)

    def mouseMoveEvent(self, a0: QtGui.QMouseEvent) -> None:
        if self._startPos:
            self._endPos = a0.pos() - self._startPos
            self.move(self.pos() + self._endPos)

    def mousePressEvent(self, a0: QtGui.QMouseEvent) -> None:
        if self.childAt(a0.pos().x(), a0.pos().y()).objectName() == "titles":
            if a0.button() == QtCore.Qt.LeftButton:
                self._isTracking = True
                self._startPos = QtCore.QPoint(a0.x(), a0.y())

    def mouseReleaseEvent(self, a0: QtGui.QMouseEvent) -> None:
        if a0.button() == QtCore.Qt.LeftButton:
            self._isTracking = False
            self._startPos = None
            self._endPos = None

    def mouseDoubleClickEvent(self, a0: QtGui.QMouseEvent) -> None:
        if self.childAt(a0.pos().x(), a0.pos().y()).objectName() == "titles":
            if a0.button() == QtCore.Qt.LeftButton:
                self.showNormal()


class playInfo(QtCore.QThread):
    def __init__(self, text) -> None:
        super().__init__()
        self.text = text

    def run(self) -> None:
        winsound.PlaySound(self.text, winsound.SND_ALIAS)


class IQMessageBox(QtWidgets.QDialog, MQMessageBox.Ui_Dialog):
    def __init__(self) -> None:
        super().__init__()
        self._startPos = None
        self._isTracking = None
        self._endPos = None

        self.is_OK = False

        self.setupUi(self)
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)

        self.buttonBox.accepted.connect(self.OK)
        self.buttonBox.rejected.connect(self.NO)

        self.pushButton.clicked.connect(self.close)

        self.sound_ = None

    def OK(self):
        self.is_OK = True

    def NO(self):
        self.is_OK = False

    def mouseMoveEvent(self, a0: QtGui.QMouseEvent) -> None:
        if self._startPos:
            self._endPos = a0.pos() - self._startPos
            self.move(self.pos() + self._endPos)

    def mousePressEvent(self, a0: QtGui.QMouseEvent) -> None:
        if self.childAt(a0.pos().x(), a0.pos().y()).objectName() == "label":
            if a0.button() == QtCore.Qt.LeftButton:
                self._isTracking = True
                self._startPos = QtCore.QPoint(a0.x(), a0.y())

    def mouseReleaseEvent(self, a0: QtGui.QMouseEvent) -> None:
        if a0.button() == QtCore.Qt.LeftButton:
            self._isTracking = False
            self._startPos = None
            self._endPos = None

    def mouseDoubleClickEvent(self, a0: QtGui.QMouseEvent) -> None:
        if self.childAt(a0.pos().x(), a0.pos().y()).objectName() == "label":
            if a0.button() == QtCore.Qt.LeftButton:
                self.showNormal()

    def Play(self, title: str, text: str, type_sound: str) -> bool:
        """
        重新实现QMessageBox
        :param title: 标题
        :param text: 文本
        :param type_sound: 播放音效
        :return: 是否点了确定
        """
        self.close()
        self.setWindowTitle(title)

        self.sound_ = playInfo(type_sound)
        self.sound_.start()
        self.textBrowser.setHtml(f"""<font size="8">{text}</font>""")

        self.show()
        self.exec_()

        return self.is_OK

    def info(self, title: str, text: str) -> bool:
        """
        简单封装IQMessageBox,播放提示音
        :param title: 标题
        :param text: 文本
        :return: 是否点了确定
        """
        return self.Play(title, text, "SystemAsterisk")


def getFileSearch(*args: list, fileCheckType: list = None, title: str = "文件", mode: str = "open",
                  allType: bool = False, parent=None, startDir=os.getcwd()) -> Optional[Tuple[str, str]]:
    """
    :param startDir: 启始路径
    :param parent: 父窗口
    :param args: 更多文件后缀 类型与fileCheckType对应
    :param fileCheckType: 对应后缀类型
    :param title: 对话框标题
    :param mode: 对话框模式
    :param allType: 是否用全后缀
    :return:
    """
    getTypes = ""

    if fileCheckType and len(args) > 0:
        for introduce, arg in zip(fileCheckType, args):
            for i in arg:
                getTypes += f"{introduce}(*.{i});;"
    if allType:
        getTypes += "All File(*.)"
    else:
        getTypes = getTypes[:-2]

    if mode == "save":
        fileName, filetype = QtWidgets.QFileDialog.getSaveFileName(parent, title, startDir,
                                                                   getTypes)
    elif mode == "open":
        fileName, filetype = QtWidgets.QFileDialog.getOpenFileName(parent, title, startDir,
                                                                   getTypes)
    elif mode == "dir":
        fileName = QtWidgets.QFileDialog.getExistingDirectory(parent, title, startDir)
        filetype = str()
    else:
        mainwindow.unKnowError()
        return

    return (fileName, filetype) if fileName else ""


if __name__ == 'main':
    os.chdir(os.path.dirname(__file__))

    app = QtWidgets.QApplication(sys.argv)

    sendMessageBox = IQMessageBox()

    losder = Losder(WhetherItIsEncrypted=False)

    mainwindow = MainWindow()
    mainwindow.show()

    mainwindow.setStyleSheet(losder.read_qss_file(".\\QSS\\Mainwindow.qss"))

    sys.exit(app.exec_())
