import sys
import os
import subprocess
import sys
import win32com.client 
import win32gui

import keyboard
from PyQt5 import QtCore, QtWidgets,QtGui
from PyQt5.Qt import QtWin
from PyQt5.QtCore import Qt, QObject, pyqtSignal,QPoint
from PyQt5.QtGui import QCursor,QIcon
from PyQt5.QtWidgets import QAction, QApplication, QMainWindow, QMenu

# from PySide6 import QtCore, QtWidgets, QtGui
# from PySide6.QtWinExtras import QtWin
# from PySide6.QtCore import Qt, QObject, Signal, QPoint
# from PySide6.QtGui import QCursor, QIcon
# from PySide6.QtWidgets import QAction, QApplication, QMainWindow, QMenu

def read_qss_file(qss_file_name):
    with open(qss_file_name, 'r',  encoding='UTF-8') as file:
        return file.read()

class KeyBoardManager(QObject):
    F1Signal = pyqtSignal()
    F2Signal = pyqtSignal()

    def start(self):
        # keyboard.add_hotkey("alt+z", self.F1Signal.emit, suppress=True)
        # keyboard.add_hotkey("alt+x", self.F2Signal.emit, suppress=True)
        keyboard.add_hotkey("alt+z", self.F1Signal.emit)
        keyboard.add_hotkey("esc", self.F2Signal.emit)

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__(flags=Qt.WindowStaysOnTopHint)
        self.base_dir = "./lnk/"
        self.quick_dir = "./lnk/quick/"
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.menu = QMenu(self)

        # self.menu.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)
        # self.menu.setAttribute(Qt.WA_TranslucentBackground, True)

        self.style_sheet = read_qss_file('./style.qss')
        self.menu.setStyleSheet(self.style_sheet)
        self.menu.triggered.connect(self.exec_lnk_v2)

        

        self.lnk2path = {}
        for j in os.listdir(self.base_dir):
            if j != "quick":
                self.make_menu(os.path.join(self.base_dir,j) ,self.menu)

        for i in  os.listdir(self.quick_dir):
            if i != "readme.md":
                Targetpath = self.shell.CreateShortCut(self.quick_dir+"/"+i).Targetpath
                self.lnk2path[i] = Targetpath
                # print(Targetpath)
                icon = self.get_icon(Targetpath)
                qa = QAction(icon,i[:-4],self.menu)
                self.menu.addAction(qa)

        manager = KeyBoardManager(self)
        manager.F1Signal.connect(self.show_menu)
        manager.F2Signal.connect(self.hide_menu)
        manager.start()

    def make_menu(self,file_path,parent):  
        if os.path.basename(file_path) != "readme.md":  
            sub = parent.addMenu(QIcon("folder.png"),os.path.basename(file_path))
            # sub.setStyleSheet(self.style_sheet)
            files = os.listdir(file_path)
            for fi in files:
                fi_d = file_path+"/"+fi 
                if os.path.isdir(fi_d):
                    self.make_menu(fi_d,sub)
                else:
                    try:
                        if fi.split(".")[-1] == "lnk":
                            Targetpath = self.shell.CreateShortCut(file_path+"/"+fi ).Targetpath
                            self.lnk2path[fi] = Targetpath
                            icon = self.get_icon(Targetpath)
                            qa = QAction(icon,fi[:-4],sub)
                            sub.addAction(qa)
                    except Exception as e:
                        print(e)

    def get_icon(self,icon_path):
        try:
            # print(icon_path)
            icon_path = icon_path.replace("\\","/")
            large, small = win32gui.ExtractIconEx(icon_path, 0)
            
            pixmap = QtWin.fromHICON(large[0])
            win32gui.DestroyIcon(small[0])
            win32gui.DestroyIcon(large[0])
            qi = QIcon()
            qi.addPixmap(pixmap)
            return qi
        except Exception as e:
            print(e)
            return QIcon("folder.png")

    @QtCore.pyqtSlot(QtWidgets.QAction)
    def exec_lnk_v2(self,action):
        exe_path = self.lnk2path[action.text()+".lnk"]
        print(exe_path)
        subprocess.Popen(exe_path,shell=True,stdin=subprocess.PIPE,stdout=subprocess.PIPE)
        return 

    @QtCore.pyqtSlot(QtWidgets.QAction)
    def exec_lnk_v1(self,action):
        shortcut = self.shell.CreateShortCut(self.quick_dir+action.text()+".lnk")
        subprocess.Popen(shortcut.Targetpath,shell=True,stdin=subprocess.PIPE,stdout=subprocess.PIPE)
        return 

    def show_menu(self):
        
        self.menu.popup(QPoint(QCursor.pos().x()-0,QCursor.pos().y()))
        # self.menu.popup(QPoint(0,0))
    
    def hide_menu(self):
        self.menu.hide()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False) 
    win = MainWindow()
    # win.show()
    sys.exit(app.exec_())