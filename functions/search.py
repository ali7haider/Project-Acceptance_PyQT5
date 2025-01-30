import os
import re
import glob
from functions.checkpath import *
from PyQt5.QtWidgets import QMessageBox
from PyQt5 import QtWidgets
from functions.createfolder import *
from functions.mail import mail_signed

def search_clicked(self):
    path = "O:\Field Services Division\Field Support Center\Project Acceptance"
    found = False
    workOrder = self.planfile_entry.text()
    if workOrder != "" and len(workOrder) >= 5:
        for root, subdir, files in os.walk(path):
            for d in subdir:
                if d.find(workOrder) != -1:
                    found = True
                    print(d)
                    print("Im HERE")
                    walk = d
                    self.concat = path + "/" + walk
                    self.mend = self.concat
                    isdir = os.path.isdir(self.concat)
                    if isdir:
                        pump_Found, pressure_Found, gravity_Found, excel = check_path(self.concat)
                        print(str(pump_Found) + " " + str(pressure_Found) + " " + str(gravity_Found))
                        self.development_checkB.setEnabled(True)
                        self.CIP_checkB.setEnabled(True)
                        if pump_Found == 1:
                            self.pump_folder.setChecked(True)
                            
                        if pressure_Found == 1:
                            self.pressure_folder.setChecked(True)
                            
                        if gravity_Found == 1:
                            self.gravity_folder.setChecked(True)
                    else:
                        pass
                
            break   
    else:
        showError()

    if not found and workOrder != "" and len(workOrder) >= 5:
        createNew(self) 


def showError():
    msgBox = QMessageBox()
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def createNew(self):
    workOrder = self.planfile_entry.text()
    self.isFirst = True
    print("in createNew: " + str(self.isFirst))
    msgBox = QMessageBox()
    msgBox.setText("Folder not found. Create new folder with this planfile?")
    msgBox.setWindowTitle("Create new folder")
    msgBox.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel)
    response = msgBox.exec()
    print(response)
    if response == 1024:
        create_planfile_folder(self, workOrder)
        pass
    else:
        pass

def create_planfile_folder(self,workOrder):
    planfile_folder = "O:\Field Services Division\Field Support Center\Project Acceptance"
    # TODO change this with the sql data
    self.path =  planfile_folder + "/" + workOrder + "- PlaceHolder Until I do the sql"
    os.mkdir(self.path)
    makeXL = self.path + "/Excel"
    os.mkdir(makeXL)
    self.development_checkB.setEnabled(True)
    self.CIP_checkB.setEnabled(True)


