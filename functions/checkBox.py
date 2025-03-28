from PyQt5.QtWidgets import QMessageBox

def check_work_area(self):
    try:
        pump_string = {'Pump', 'pump', 'pumpstation', 'PumpStation', 'Pumpstation'}
        pressure_string = {"Pressure", "Pressurized", "Pressurized Pipe", "Pipe", "pressure", "pressurized"}
        wastewater_string = {"Gravity", "gravity", "wastewater", "WasteWater"}

        work_entry_text = self.work_entry.text().strip()  # Ensure there are no leading/trailing spaces

        if work_entry_text in pump_string:
            print("1")
            self.pump_checkB.setChecked(True)
            self.pressure_checkB.setChecked(False)
            self.gravity_checkB.setChecked(False)
        elif work_entry_text in pressure_string:
            print("2")
            self.pump_checkB.setChecked(False)
            self.pressure_checkB.setChecked(True)
            self.gravity_checkB.setChecked(False)
        elif work_entry_text in wastewater_string:
            print("3")
            self.pump_checkB.setChecked(False)
            self.gravity_checkB.setChecked(True)
            self.pressure_checkB.setChecked(False)
        else:
            showError("Please enter a valid entry.")
            self.planfile_Btn.setEnabled(False)
            self.planfile_entry.setEnabled(False)
            return  # Exit function to avoid enabling buttons for invalid entry

        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)

    except Exception as e:
        showError(f"An error occurred: {str(e)}")

def showError(message):
    msgBox = QMessageBox()
    msgBox.setText(message)
    msgBox.setWindowTitle("Error")
    msgBox.exec()
