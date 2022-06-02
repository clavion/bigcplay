# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import QDialog, QHBoxLayout, QLabel, QMessageBox
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import pyqtSignal
import mido
import time
import os
import sys
import _thread

programBasePath = None

def getResourcePath(path):
    global programBasePath
    if(programBasePath is None):
        programBasePath = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(programBasePath, path)  

class MidiMappingSetting(QDialog):
    
    keySelected = pyqtSignal()
    
    def __init__(self, deviceName, isLong):
        super(MidiMappingSetting, self).__init__()
        self.setWindowTitle('映射设置')
        self.deviceName = deviceName
        hbox = QHBoxLayout()
        lbl = QLabel()
        if(isLong):
            lbl.setPixmap(QPixmap(getResourcePath('midiMappingLong.png')))
        else:
            lbl.setPixmap(QPixmap(getResourcePath('midiMapping.png')))
        hbox.addWidget(lbl)
        self.setLayout(hbox)
        self.selectedKey = 0
        self.terminating = False
        self.isLong = isLong
        self.keySelected.connect(self.mappingComplete)
        
    def getMidiKey(self):
        mi = mido.open_input(self.deviceName)
        while(True):
            msg = mi.poll()
            if(msg):
                if(msg.type == 'note_on'):
                    if(msg.note % 12 == 0):
                        self.selectedKey = msg.note
                        self.keySelected.emit()
                        break
            time.sleep(0.05)
            if(self.terminating):
                break
        mi.close()

        
    def begin(self):
        _thread.start_new_thread(self.getMidiKey, ())
        self.exec_()
        
    def closeEvent(self, e):
        self.terminating = True
        
    def mappingComplete(self):
        self.accept()

if(__name__ == '__main__'):
    from PyQt5.QtWidgets import QApplication
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    devices = mido.get_input_names()
    if(len(devices) > 0):
        w = MidiMappingSetting(devices[0], True)
        w.begin()
