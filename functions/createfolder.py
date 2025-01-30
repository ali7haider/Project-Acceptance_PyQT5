import os
import functions.excel as excel
import shutil
import functions.word as word
from PyQt5.QtWidgets import QMessageBox
from PyQt5 import QtWidgets
from functions.search import *


def create_new(self):
        template = "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Templates"
        parent_dir = self.concat
        print("In create New " + parent_dir)
        
        if self.pump_checkB.isChecked():
            if self.CIP_checkB.isChecked():
                variation = "/Cip Pump.csv"
                csv_tranfer = template + variation
            elif self.development_checkB.isChecked():
                variation = "/Development Pump.csv"
                csv_tranfer = template + variation
            self.directoryCode = "/Pump-Station"
            path = parent_dir+ self.directoryCode
            os.mkdir(path) 
            excel.init_table(self, csv_tranfer, variation)
            print("Directory '% s' created" % self.directoryCode) 
            print(path)
            original = template + "/Cip Pump.csv"
            target = parent_dir + "/Excel" + "/Cip Pump.csv"

            shutil.copyfile(original, target)

            original2 = template + "/Development Pump.csv"
            target2 = parent_dir + "/Excel" + "/Development Pump.csv"

            shutil.copyfile(original2, target2)

            self.reset_btn.setEnabled(False)
            self.work_entry_Btn.setEnabled(False)
            self.planfile_Btn.setEnabled(False)
            self.add_row.setEnabled(True)
            self.del_row.setEnabled(True)
            self.send_to_sign.setEnabled(True)
            self.save_btn.setEnabled(True)
            create_successful(self)

        elif self.gravity_checkB.isChecked():
            if self.CIP_checkB.isChecked():
                variation = "/Cip Gravity.csv"
                csv_tranfer = template + variation
            elif self.development_checkB.isChecked():
                variation = "/Development Gravity.csv"
                csv_tranfer = template + variation
            self.directoryCode = "/Wastewater"
            path = parent_dir+ self.directoryCode
            print(path)
            os.mkdir(path) 
            excel.init_table(self, csv_tranfer, variation)
            print("Directory '% s' created" % self.directoryCode) 
            original = template + "/Cip Gravity.csv"
            target = parent_dir + "/Excel" + "/Cip Gravity.csv"

            shutil.copyfile(original, target)

            original2 = template + "/Development Gravity.csv"
            target2 = parent_dir + "/Excel" + "/Development Gravity.csv"

            shutil.copyfile(original2, target2)
            self.reset_btn.setEnabled(False)
            self.work_entry_Btn.setEnabled(False)
            self.planfile_Btn.setEnabled(False)
            self.add_row.setEnabled(True)
            self.send_to_sign.setEnabled(True)
            self.save_btn.setEnabled(True)
            self.del_row.setEnabled(True)
            create_successful(self)

        elif self.pressure_checkB.isChecked():
            if self.CIP_checkB.isChecked():
                variation = "/Cip Pressure.csv"
                csv_tranfer = template + variation
            elif self.development_checkB.isChecked():
                variation = "/Development Pressure.csv"
                csv_tranfer = template + variation
            self.directoryCode = "/Pressurized-Pipe"
            path = parent_dir+ self.directoryCode
            os.mkdir(path) 
            excel.init_table(self, csv_tranfer, variation)
            print("Directory '% s' created" % self.directoryCode) 
            original = template + "/Cip Pressure.csv"
            target = parent_dir + "/Excel" + "/Cip Pressure.csv"

            shutil.copyfile(original, target)

            original2 = template + "/Development Pressure.csv"
            target2 = parent_dir + "/Excel" + "/Development Pressure.csv"

            shutil.copyfile(original2, target2)
            self.reset_btn.setEnabled(False)
            self.work_entry_Btn.setEnabled(False)
            self.planfile_Btn.setEnabled(False)
            self.add_row.setEnabled(True)
            self.send_to_sign.setEnabled(True)
            self.save_btn.setEnabled(True)
            self.del_row.setEnabled(True)
            create_successful(self)

        # first_time(self, path)

# def first_time(self, path):
#     msgBox = QMessageBox()
#     msgBox.setText("Are all deficencies Accepted?")
#     msgBox.setWindowTitle("Create Letter")
#     msgBox.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
#     response = msgBox.exec()
#     print(response)
#     if self.pump_checkB.isChecked():
#         active = "Pump-Station"

#     elif self.gravity_checkB.isChecked():
#         active = "Gravity"
        
#     elif self.pressure_checkB.isChecked():
#         active = "Pressurized-Pipe"

#     if self.CIP_checkB.isChecked():
#         project = "CIP"

#     if self.development_checkB.isChecked():
#         project = "Dev"

#     if response == QtWidgets.QMessageBox.StandardButton.Yes:
#         print("We said yes")
#         print(active)
#         word.acceptance_no_deficiencies(self,active, project, path)
#         return True
        
#     else:
#         return False
    

def create_successful(self):
    msgBox = QMessageBox()
    msgBox.setText("You created " + self.directoryCode)
    msgBox.setWindowTitle("Successful")
    self.create_new_Btn.setEnabled(False)
    self.open_existing_Btn.setEnabled(True)
    msgBox.exec()


