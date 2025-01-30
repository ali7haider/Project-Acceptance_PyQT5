from PyQt5.QtWidgets import QMessageBox

def check_work_area(self):

    pump_string = {'Pump', 'pump', 'pumpstation', 'PumpStation', 'Pumpstation'}
    pressure_string = {"Pressure", "Pressurized","Pressurized Pipe", "Pipe", "pressure", "pressurized"}
    wastewater_string = {"Gravity", "gravity", "wastewater", "WasteWater"}

    work_entry_text = self.work_entry.text()

    if work_entry_text in pump_string:
        print("1")
        self.pump_checkB.setChecked(True)
        self.pressure_checkB.setChecked(False)
        self.gravity_checkB.setChecked(False)
        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)
    elif work_entry_text in pressure_string:
        print("2")
        self.pump_checkB.setChecked(False)
        self.pressure_checkB.setChecked(True)
        self.gravity_checkB.setChecked(False)
        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)
    elif work_entry_text in wastewater_string:
        print("3")
        self.pump_checkB.setChecked(False)
        self.gravity_checkB.setChecked(True)
        self.pressure_checkB.setChecked(False)
        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)
    else:
        showError()
        self.planfile_Btn.setEnabled(False)
        self.planfile_entry.setEnabled(False)
        


def showError():
    msgBox = QMessageBox()
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()
