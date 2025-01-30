import os
import re
import glob
from functions.checkpath import *
from PyQt5.QtWidgets import QMessageBox
from PyQt5 import QtWidgets
from functions.createfolder import *
from functions.mail import mail_signed

def search_clicked(self):
    try:
        path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  
        print(path)
        found = False
        workOrder = self.planfile_entry.text()

        if workOrder and len(workOrder) >= 5:
            for root, subdir, files in os.walk(path):
                for d in subdir:
                    if workOrder in d:
                        found = True
                        print(d)
                        print("Im HERE")
                        self.concat = os.path.join(path, 'Data',d)
                        self.mend = self.concat
                        print(f"Normalized path: {self.concat}")

                        if os.path.isdir(self.concat):
                            pump_Found, pressure_Found, gravity_Found, excel = check_path(self.concat)
                            print(f"{pump_Found} {pressure_Found} {gravity_Found}")
                            self.development_checkB.setEnabled(True)
                            self.CIP_checkB.setEnabled(True)
                            
                            if pump_Found:
                                self.pump_folder.setChecked(True)
                            if pressure_Found:
                                self.pressure_folder.setChecked(True)
                            if gravity_Found:
                                self.gravity_folder.setChecked(True)
                        break
            if not found:
                createNew(self)
        else:
            showError("Please enter a valid entry.")

    except Exception as e:
        showError(f"An error occurred: {str(e)}")


def showError(message):
    msgBox = QMessageBox()
    msgBox.setText(message)
    msgBox.setWindowTitle("Error")
    msgBox.exec()


def createNew(self):
    try:
        workOrder = self.planfile_entry.text()
        self.isFirst = True
        print(f"in createNew: {self.isFirst}")
        
        msgBox = QMessageBox()
        msgBox.setText("Folder not found. Create new folder with this planfile?")
        msgBox.setWindowTitle("Create New Folder")
        msgBox.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel)
        response = msgBox.exec()
        
        if response == QtWidgets.QMessageBox.StandardButton.Ok:
            create_planfile_folder(self, workOrder)

    except Exception as e:
        showError(f"Error creating folder: {str(e)}")


def create_planfile_folder(self, workOrder):
    try:
        # Get the parent directory dynamically (one level up)
        parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        project_acceptance_path = os.path.join(parent_dir, "Data")

        # Ensure the "Project Acceptance" directory exists
        if not os.path.exists(project_acceptance_path):
            os.makedirs(project_acceptance_path)  # Creates parent directory if missing

        # Now create the work order folder inside "Project Acceptance"
        self.path = os.path.join(project_acceptance_path, f"{workOrder} - PlaceHolder Until I do the sql")
        os.mkdir(self.path)

        # Create the "Excel" subfolder
        makeXL = os.path.join(self.path, "Excel")
        os.mkdir(makeXL)

        # Enable UI elements
        self.development_checkB.setEnabled(True)
        self.CIP_checkB.setEnabled(True)

    except Exception as e:
        showError(f"Error creating planfile folder: {str(e)}")
