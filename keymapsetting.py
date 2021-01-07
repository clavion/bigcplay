# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import QDialog, QPushButton
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QSignalMapper
import win32gui
import math
import ctypes

class KeyMapSetting(QDialog):
    
    def __init__(self, keyMap=[]):
        super(KeyMapSetting, self).__init__()
        self.setWindowTitle('键位设置')
        if(len(keyMap) >= 37):
            self.keyMap = keyMap
        else:
            self.keyMap = [65] * 37
        self.keyCode = []
        for c in keyMap:
            self.keyCode.append(ctypes.windll.User32.VkKeyScanA(c) & 0xffff)
        noteNames = ['C', 'D', 'E', 'F', 'G', 'A', 'B']
        self.noteIndexes = [[0, 2, 4, 5, 7, 9, 11], [1, 3], [6, 8, 10]]
        self.buttons = [None] * 37
        self.noteNames = [''] * 37
        self.isBlackButton = [False, True, False, True, False, False, True, False, True, False, True, False] * 3 + [False]
        self.sm = QSignalMapper(self)
        keyWidth = 50
        keyHeight = 150
        for i in range(3):
            for j in range(7):
                pb = QPushButton()
                pb.setStyleSheet('font-size:11pt;font-family:Consolas;background-color:#eee;color:#000')
                pb.setParent(self)
                pb.move((i * 7 + j) * keyWidth, 0)
                pb.resize(keyWidth, keyHeight)
                index = i*12+self.noteIndexes[0][j]
                self.buttons[index] = pb
                self.noteNames[index] = noteNames[j] + str(i + 3)
                pb.setText(self.buttonText(index))
                self.sm.setMapping(pb, index)
                pb.clicked.connect(self.sm.map)
                pb.keyPressEvent = self.chooseKey
            for j in range(2):
                pb = QPushButton()
                pb.setStyleSheet('font-size:10pt;font-family:Consolas;background-color:#659;color:#fff')
                pb.setParent(self)
                pb.move(int((i * 7) * keyWidth + (0.7 + j) * keyWidth), 0)
                pb.resize(int(0.6 * keyWidth), int(0.6 * keyHeight))
                index = i*12+self.noteIndexes[1][j]
                self.buttons[index] = pb
                self.noteNames[index] = noteNames[j] + '#' + str(i + 3)
                pb.setText(self.buttonText(index))
                self.sm.setMapping(pb, index)
                pb.clicked.connect(self.sm.map)
                pb.keyPressEvent = self.chooseKey
            for j in range(3):
                pb = QPushButton()
                pb.setStyleSheet('font-size:10pt;font-family:Consolas;background-color:#659;color:#fff')
                pb.setParent(self)
                pb.move(int((i * 7) * keyWidth + (3.7 + j) * keyWidth), 0)
                pb.resize(int(0.6 * keyWidth), int(0.6 * keyHeight))
                index = i*12+self.noteIndexes[2][j]
                self.buttons[index] = pb
                self.noteNames[index] = noteNames[j+3] + '#' + str(i + 3)
                pb.setText(self.buttonText(index))
                self.sm.setMapping(pb, index)
                pb.clicked.connect(self.sm.map)
                pb.keyPressEvent = self.chooseKey
        pb = QPushButton()
        pb.setStyleSheet('font-size:11pt;font-family:Consolas;background-color:#eee;color:#000')
        pb.setParent(self)
        pb.move(3 * 7 * keyWidth, 0) 
        pb.resize(keyWidth, keyHeight)
        index = 36
        self.buttons[index] = pb
        self.noteNames[index] = noteNames[0] + str(3 + 3)
        pb.setText(self.buttonText(index))
        self.sm.setMapping(pb, index)
        pb.clicked.connect(self.sm.map)
        pb.keyPressEvent = self.chooseKey
        
        self.sm.mapped.connect(self.buttonClicked)
        self.choosingKeyIndex = -1
            
    def buttonText(self, index):
        if(self.isBlackButton[index]):
            return '\n\n' + self.noteNames[index] + '\n' + chr(self.keyMap[index])
        else:
            return '\n\n\n\n' + self.noteNames[index] + '\n' + chr(self.keyMap[index])
        
    def buttonClicked(self, index):
        self.choosingKeyIndex = index
        
    def chooseKey(self, e):
        if(self.choosingKeyIndex >= 0):
            try:
                self.keyMap[self.choosingKeyIndex] = ord(chr(e.key()).lower())
                self.keyCode[self.choosingKeyIndex] = ctypes.windll.User32.VkKeyScanA(self.keyMap[self.choosingKeyIndex]) & 0xffff
                self.buttons[self.choosingKeyIndex].setText(self.buttonText(self.choosingKeyIndex))
                self.choosingKeyIndex = -1
            except:
                pass
    
if(__name__ == '__main__'):
    from PyQt5.QtWidgets import QApplication
    import sys, os
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    kms = KeyMapSetting()
    kms.show()
    app.exec_()