from email import header
from PyQt5 import QtGui
from PyQt5 import QtWidgets
from win32com import client
from functions.input_dialog import InputDialog
from PyQt5.QtCore import QAbstractTableModel, QModelIndex,Qt
from PyQt5.QtWidgets import QComboBox, QItemDelegate,QMessageBox,QInputDialog,QWidget,QDialog
import shutil
import time
import pandas as pd
import configparser
import json
import functions.word as word
import functions.mail as mail
import os
import pythoncom
# Creates the model for the QTableView
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data
        self.__empty_row_count = 0
        self.__total_row_count = 0  # Max visible rows.

    def rowCount(self, index=QModelIndex()):                    
        return  0 if index.isValid() else len(self._data)

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
                value = self._data.iloc[index.row(), index.column()]
                return str(value)

        if index.isValid():
            if role == Qt.ItemDataRole.BackgroundRole:
                value = self._data.iloc[index.row(), index.column()]
                if (
                    (isinstance(value, str))
                    and value == "Accepted"
                ):
                    return QtGui.QColor(207, 225, 167)

                if (
                    (isinstance(value, str))
                    and value == "Rejected"
                ):
                    return QtGui.QColor(187, 28, 42)

                if (
                    (isinstance(value, str))
                    and value == "Removed"
                ):
                    return QtGui.QColor(255, 191, 0)
                if (
                    (isinstance(value, str))
                    and value == "Minor Deficiencies"
                ):
                    return QtGui.QColor(173, 216, 230)

    def setData(self, index, value, role):
        if role == Qt.ItemDataRole.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            return True
        return False

    def emptyRowCount(self):
        return self.__empty_row_count

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]

    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

# Creates the ComboBoxes for the QTableView
class ComboBoxDelegate(QItemDelegate):
    
    def __init__(self, parent=None):
        super(ComboBoxDelegate, self).__init__(parent)
        self.items = []
        

    def setItems(self, items):
        self.items = items

    def createEditor(self, widget, option, index):
        editor = QComboBox(widget)
        editor.addItems(self.items)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.EditRole)

        if value:
            editor.setCurrentText(str(value))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

structure_delegate = ComboBoxDelegate()
category_delegate = ComboBoxDelegate()
category_pump_delegate = ComboBoxDelegate()
approved_cctv_delegate = ComboBoxDelegate()
vendor = ComboBoxDelegate()
reviewer_delegate = ComboBoxDelegate()
delegate = ComboBoxDelegate()
delegate.setItems(["Accepted", "Rejected"])
action_delegate = ComboBoxDelegate()
options_delegate = ComboBoxDelegate()
options_delegate.setItems(["Accepted", "Rejected", "Removed","Minor Deficiencies"])
corrective_delegate = ComboBoxDelegate()
empty_delegate = ComboBoxDelegate()

def filename_checker(self):

        if self.pump_checkB.isChecked() and self.CIP_checkB.isChecked():
            self.filename = "CIP Pump.csv"
            self.code = 1
            self.directory_code = self.concat+"/Pump-Station"
            self.area = "Pump-Station"
            self.cip_dev = "CIP"

        if self.pump_checkB.isChecked() and self.development_checkB.isChecked():
            self.filename = "Development Pump.csv"
            self.code = 2
            self.directory_code = self.concat+"/Pump-Station"
            self.area = "Pump-Station"
            self.cip_dev = "Dev"

        if self.gravity_checkB.isChecked() and self.CIP_checkB.isChecked():
            self.filename = "Cip Gravity.csv"
            self.code = 3
            self.directory_code = self.concat+"/Wastewater"
            self.area = "Gravity"
            self.cip_dev = "CIP"

        if self.gravity_checkB.isChecked() and self.development_checkB.isChecked():
            self.filename = "Development Gravity.csv"
            self.code = 4
            self.directory_code = self.concat+"/Wastewater"
            self.area = "Gravity"
            self.cip_dev = "Dev"

        if self.pressure_checkB.isChecked() and self.CIP_checkB.isChecked():
            self.filename = "CIP Pressure.csv"
            self.code = 5
            self.directory_code = self.concat+"/Pressurized-Pipe"
            self.area = "Pressurized-Pipe"
            self.cip_dev = "CIP"

        if self.pressure_checkB.isChecked() and self.development_checkB.isChecked():
            self.filename = "Development Pressure.csv"
            self.code = 6
            self.directory_code = self.concat+"/Pressurized-Pipe"
            self.area = "Pressurized-Pipe"
            self.cip_dev = "Dev"

        self.excel_filename = self.concat + "/Excel/" + self.filename

# class App(QWidget):
#     global acceptedorrejected
#     global walkthroughorwarranty
#     global workorder

#     def __init__(self):
#         super().__init__()
#         self.title = 'PyQt5 input dialogs - pythonspot.com'
#         self.left = 10
#         self.top = 10
#         self.width = 640
#         self.height = 480
#         self.initUI()
    
#     def initUI(self):
#         # self.setWindowTitle(self.title)
#         # self.setGeometry(self.left, self.top, self.width, self.height)
#         # self.center()

#         self.getChoice()
#         self.getText()
        
#         # self.show()

#     def getChoice(self):
#         accepted_rejected = ("","Acceptance","Rejection")
#         walkthrough_warranty = ("Walkthrough","Warranty")
#         self.item2, okPressed = QInputDialog.getItem(self, "Get item","Walkthrough or Warranty:", walkthrough_warranty, 0, False)
#         self.item, okPressed = QInputDialog.getItem(self, "Get item","Accepted or Rejected:", accepted_rejected, 0, False)
#         if okPressed and self.item and self.item2:
#             print(self.item)
#             print(self.item2)

#     def getText(self):
#         self.text, okPressed = QInputDialog.getText(self, "Get text","Please Enter Inspection Date mm-dd-yyyy:",  QtWidgets.QLineEdit.EchoMode.Normal, "")
#         if okPressed and self.text != '':
#             print(self.text)

#     def center(self):
#         qr = self.frameGeometry()
#         cp = QtGui.QGuiApplication.primaryScreen().availableGeometry().center()
#         qr.moveCenter(cp)
#         self.move(qr.topLeft())

def send(self):
    filename_checker(self)
    print(self.directory_code)
    mail.mail_signed(self,self.directory_code)

# Initialized the table
def init_table(self,template, variation):

    # for proc in psutil.process_iter():
    #     if proc.name() == "excel.exe":
    #         proc.kill()

    filename_checker(self)

    if template!= "":
        self.excel_filename = template
        self.variation1 = variation
    else:
        self.variation1 = "/"+ self.filename
    
    
    print("in excel.py")
    print(self.excel_filename)
    self.df = pd.read_csv(self.excel_filename)
    if self.df.size == 0:
        return

    set_delegates(self, self.tableview)

    self.model = PandasModel(self.df)

    self.tableview.setModel(self.model)
    self.tableview.resizeRowsToContents()
    self.tableview.horizontalHeader().setStretchLastSection(True)
    self.tableview.setWordWrap(True)
    self.tableview.show()


def add_rows(self, table):
    df2 = {'Warranty': ''}
    self.df = self.df.append(df2, ignore_index = True)
    
    self.model = PandasModel(self.df)

    set_delegates(self, table)

    table.setModel(self.model)
    
    table.show()

def remove_row(self, table):
    x = table.selectionModel().currentIndex()
    NewIndex = x.row()
    print("This is the Index: ", NewIndex)
    new_model = self.df.drop(self.df.index[NewIndex])
    self.model = PandasModel(new_model)

    set_delegates(self, table)

    table.setModel(self.model)
    self.df = new_model
    table.show()
    
import json
import configparser
from PyQt5.QtWidgets import QMessageBox
import psutil

def showError(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText("Error")
    msg.setInformativeText(message)
    msg.setWindowTitle("Error")
    msg.exec_()

def set_delegates(self, table):

    config = configparser.RawConfigParser()
    config.read(
        "config\delegates.properties")

    approved_cctv = config.get("GRAVITY", "approved_cctv")
    approved_cctv_list = json.loads(approved_cctv)
    approved_cctv_delegate.setItems(approved_cctv_list)

    vendor_surveyor = config.get("GRAVITY", "oc_surveyor")
    vendor_surveyor_list = json.loads(vendor_surveyor)
    vendor.setItems(vendor_surveyor_list)

    reviewer = config.get("GRAVITY", "reviewer")
    reviewer_list = json.loads(reviewer)
    reviewer_delegate.setItems(reviewer_list)

    gravity_structure = config.get("GRAVITY", "structure")
    gravity_structure_list = json.loads(gravity_structure)
    structure_delegate.setItems(gravity_structure_list)

    action = config.get("GRAVITY", "action")
    action_list = json.loads(action)
    action_delegate.setItems(action_list)

    corrective = config.get("GRAVITY", "corrective")
    corrective_list = json.loads(corrective)
    corrective_delegate.setItems(corrective_list)

    pressure_structure_vals = config.get("PRESSURE", "structure")
    pressure_structure_vals_list = json.loads(pressure_structure_vals)
    structure_delegate.setItems(pressure_structure_vals_list)

    pressure_category_val = config.get("PRESSURE", "category")
    category_list = json.loads(pressure_category_val)
    category_delegate.setItems(category_list)

    pump_values = config.get("PUMP", "category")
    pump_value_list = json.loads(pump_values)
    category_pump_delegate.setItems(pump_value_list)

    if self.code == 1:
        table.setItemDelegateForColumn(0, category_pump_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)

    if self.code == 2:
        table.setItemDelegateForColumn(0, category_pump_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)

    if self.code == 3:
        table.setItemDelegateForColumn(4, structure_delegate)
        table.setItemDelegateForColumn(5, action_delegate)
        table.setItemDelegateForColumn(6, approved_cctv_delegate)
        table.setItemDelegateForColumn(7, options_delegate)
        table.setItemDelegateForColumn(8, vendor)
        table.setItemDelegateForColumn(9, reviewer_delegate)
        table.setItemDelegateForColumn(10, corrective_delegate)
        table.setItemDelegateForColumn(12, options_delegate)
        # table.setItemDelegateForColumn(14, vendor)
        table.setItemDelegateForColumn(14, reviewer_delegate)
        table.setItemDelegateForColumn(15, corrective_delegate)

    if self.code == 4:
        table.setItemDelegateForColumn(2, approved_cctv_delegate)
        table.setItemDelegateForColumn(3, options_delegate)
        table.setItemDelegateForColumn(5, reviewer_delegate)
        table.setItemDelegateForColumn(7, options_delegate)
        table.setItemDelegateForColumn(8, vendor)
        table.setItemDelegateForColumn(9, reviewer_delegate)

    if self.code == 5:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(5, structure_delegate)
        table.setItemDelegateForColumn(6, options_delegate)
        table.setItemDelegateForColumn(8, options_delegate)

    if self.code == 6:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(5, structure_delegate)
        table.setItemDelegateForColumn(6, options_delegate)
        table.setItemDelegateForColumn(8, options_delegate)

# Kill all running Excel instances
def kill_excel():
    for process in psutil.process_iter():
        if process.name().lower() == "excel.exe":
            process.kill()
def kill_word():
    """Forcefully terminate all running instances of Microsoft Word."""
    for process in psutil.process_iter():
        if process.name().lower() == "winword.exe":  # Word process name
            process.kill()

def exportToExcel(self, table):
    columnHeaders = []
    setter = table
    # create column header list
    for j in range(setter.model().columnCount()):
        columnHeaders.append(setter.model().headerData(
            j, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole))

    dfnew = pd.DataFrame(columns=columnHeaders)

    # create dataframe object recordset
    for row in range(setter.model().rowCount()):
        for col in range(setter.model().columnCount()):
            dfnew.at[row, columnHeaders[col]] = setter.model().index(row, col).data()
    
    dfnew.to_csv(self.concat + "/Excel" + self.variation1, index=False)
    print('Excel file exported')


def pandas2word(self):
    # Ensure Excel processes are killed
    kill_excel()
    kill_word()

    # Get the current working directory (path to the script)
    path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


    # Define the temporary directory path
    temp_dir = os.path.join(path, "Temp")  # Create Temp folder path relative to the script directory

    # Ensure the directory exists
    os.makedirs(temp_dir, exist_ok=True)

    # Define the Excel file paths
    excel_filepath = os.path.join(temp_dir, "temp1.xlsx")
    temp_excel_filepath = os.path.join(temp_dir, "temp2.xlsx")

    # Print file paths for debugging (optional)
    print(f"Excel file path: {excel_filepath}")
    print(f"Temporary Excel file path: {temp_excel_filepath}")
    try:
        self.getchoice = QWidget()
        self.dialog = InputDialog() 
        if self.dialog.exec_() == QDialog.Accepted:  # If user clicks 'OK' (btnEdit)
            self.dialog.get_values()  # Fetch values
            
            # Store values in variables
            self.getchoice.item2  = self.dialog.WalkthroughOrButton
            self.getchoice.item = self.dialog.AcceptedOrRejected
            self.getchoice.text= self.dialog.inspection_date
            print(self.getchoice.item,self.getchoice.item2,self.getchoice.text)
            # Assign values in the same format as self.getchoice
            # self.getchoice.item = self.dialog.item
            # self.getchoice.item2 = self.dialog.item2
            # self.getchoice.text = self.dialog.text
        else:
            print("Dialog cancelled")
            return
        # self.getchoice = App()
        # Using pandas ExcelWriter for writing data
        writer = pd.ExcelWriter(excel_filepath, engine='xlsxwriter')
        if self.getchoice.item == "Acceptance" and self.getchoice.item2 == "Walkthrough":
            print("Acceptance and Walkthrough")
            found = self.df[self.df['Walkthrough'].fillna('').str.contains('Rejected', na=False)]
            found2 = len(self.df[self.df['Walkthrough'].fillna('') == ''])
            
            if len(found) == 0 and found2 == 0:
                # Filtered DataFrame that you want to save
                filtered_df = variation(self)
                filtered_df.to_excel(writer, sheet_name='Sheet1', index=False)

                # Get row and column count for formatting
                rowcount = self.tableview.model().rowCount()
                colcount = self.tableview.model().columnCount()

                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # Define formats
                red = workbook.add_format({'bg_color': '#ffcccb'})
                green = workbook.add_format({'bg_color': '#90EE90'})
                orange = workbook.add_format({'bg_color': 'yellow'})
                border = workbook.add_format({'border': 1})
                light_blue = workbook.add_format({'bg_color': '#ADD8E6'})  # Light Blue Background

                header_format = workbook.add_format({'bold': True, 'text_wrap': True})
                
                # Adding text wrapping for cells
                cell_format = workbook.add_format()
                cell_format.set_text_wrap()

                # Apply column widths and conditional formatting based on self.code
                if self.code == 1:
                    worksheet.set_column('A:F', 12, cell_format)
                elif self.code == 2:
                    worksheet.set_column('A:F', 12, cell_format)
                    worksheet.set_column('G:G', 30, cell_format)
                elif self.code == 3:
                    worksheet.set_column('A:K', 10, cell_format)
                    worksheet.set_column('L:L', 25, cell_format)
                elif self.code == 4:
                    worksheet.set_column('A:H', 12, cell_format)
                    worksheet.set_column('I:I', 25, cell_format)
                elif self.code == 5:
                    worksheet.set_column('A:H', 12, cell_format)
                    worksheet.set_column('I:I', 20, cell_format)
                elif self.code == 6:
                    worksheet.set_column('A:H', 12, cell_format)
                    worksheet.set_column('I:I', 20, cell_format)

                # Apply conditional formatting for specific values
                worksheet.conditional_format('A1:Z100', {'type': 'text', 'criteria': 'containing', 'value': 'Rejected', 'format': red})
                worksheet.conditional_format('A1:Z100', {'type': 'text', 'criteria': 'containing', 'value': 'Accepted', 'format': green})
                worksheet.conditional_format('A1:Z100', {'type': 'text', 'criteria': 'containing', 'value': 'Removed', 'format': orange})
                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Minor Deficiencies',
                                                            'format':light_blue
                                                            })
                # Apply border formatting
                worksheet.conditional_format(rowcount, 1, 1, colcount, {'type': 'no_blanks', 'format': border})

                # No need for writer.save() as ExcelWriter will handle it automatically when 'with' block exits.
                print("Excel file saved successfully!")
                # Ensure the file is saved and exists before proceeding
                if os.path.exists(excel_filepath):
                    print(f"File {excel_filepath} exists!")
                    writer.save()
                    excel = client.Dispatch("Excel.Application")
                    print("Opening Excel...")
                    sheets = excel.Workbooks.Open(excel_filepath)
                    
                    work_sheets = sheets.Worksheets[0]
                    work_sheets.PageSetup.Orientation = 2
                    sheets.SaveAs(temp_excel_filepath)
                    sheets.Close(True)
                    excel.Application.Quit()
                    print("Excel closed.")
                    # writer.close()
                    # word_sol(self,name)
                    print("Creating PDF...")
                    create_pdf(self,temp_excel_filepath)
                    print("PDF created.")
                    location = self.directory_code 
                    word.acceptance_no_deficiencies(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item)
                else:
                    print(f"File {excel_filepath} does not exist!")
            else:
                print("Conditions for acceptance not met. Triggering error acceptance.")
                error_acceptance()

        if self.getchoice.item == "Acceptance" and self.getchoice.item2 == "Warranty":
            found = self.df[self.df['Warranty'].str.contains('Rejected')]
            found2 = len(self.df[self.df['Warranty'] == '']) 
            if len(found) == 0 and found2 == 0:
                filtered_df = variation(self)
                filtered_df.to_excel(writer, sheet_name='Sheet1')
                # filtered_df = self.df.loc[:, ~self.df.columns.isin(["OC Surveyor", "Notes.1"])]
                rowcount = self.tableview.model().rowCount()
                colcount = self.tableview.model().columnCount()

                # Get the xlsxwriter workbook and worksheet objects.
                workbook  = writer.book
                cell_format = workbook.add_format()
                cell_format.set_text_wrap()
                red = workbook.add_format({'bg_color': '#ffcccb'})
                green = workbook.add_format({'bg_color': '#90EE90'})
                orange = workbook.add_format({'bg_color': 'yellow'})
                light_blue = workbook.add_format({'bg_color': '#ADD8E6'})  # Light Blue Background

                border = workbook.add_format({'border': 1})

                worksheet = writer.sheets['Sheet1']
                if self.code == 1:
                    worksheet.set_column('A:F', 12)
                    worksheet.set_column('G:G', 60)

                if self.code == 2:
                    worksheet.set_column('A:F', 12)
                    worksheet.set_column('G:G', 60)
            
                if self.code == 3:
                    worksheet.set_column('A:K', 10)
                    worksheet.set_column('L:L', 45)

                if self.code == 4:
                    worksheet.set_column('A:G', 12)
                    worksheet.set_column('H:H', 35)

                if self.code == 5:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 35)

                if self.code == 6:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 40)

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Rejected',
                                                            'format':red
                                                            })

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Accepted',
                                                            'format':green
                                                            })

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Removed',
                                                            'format':orange
                                                            })
                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Minor Deficiencies',
                                                            'format':light_blue
                                                            })

                worksheet.conditional_format(rowcount, 1, 1, colcount, {'type': 'no_blanks','format': border})
                
                writer.save()
                
                # writer.close()
                create_pdf(self,temp_excel_filepath)

                # word_sol(self,excel_filepath)
                location = self.directory_code 
                word.acceptance_no_deficiencies(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item)
            else:
                print("got u again")
                error_acceptance()
                

        if self.getchoice.item == "Rejection" and self.getchoice.item2 == "Walkthrough":
            found2 = len(self.df[self.df['Walkthrough'] == '']) 
            if found2 == 0:
                filtered_df = variation(self)
                filtered_df.to_excel(writer, sheet_name='Sheet1',index=False)
                # filtered_df = self.df.loc[:, ~self.df.columns.isin(["OC Surveyor", "Notes.1"])]
                rowcount = self.tableview.model().rowCount()
                colcount = self.tableview.model().columnCount()

                # Get the xlsxwriter workbook and worksheet objects.
                workbook  = writer.book
                cell_format = workbook.add_format()
                cell_format.set_text_wrap()
                red = workbook.add_format({'bg_color': '#ffcccb'})
                green = workbook.add_format({'bg_color': '#90EE90'})
                orange = workbook.add_format({'bg_color': 'yellow'})
                border = workbook.add_format({'border': 1})
                light_blue = workbook.add_format({'bg_color': '#ADD8E6'})  # Light Blue Background

                worksheet = writer.sheets['Sheet1']
                if self.code == 1:
                    worksheet.set_column('A:F', 12)
                    worksheet.set_column('G:G', 60)

                if self.code == 2:
                    worksheet.set_column('A:F', 12)
                    worksheet.set_column('G:G', 60)
            
                if self.code == 3:
                    worksheet.set_column('A:K', 10)
                    worksheet.set_column('L:L', 45)

                if self.code == 4:
                    worksheet.set_column('A:G', 12)
                    worksheet.set_column('H:H', 35)

                if self.code == 5:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 40)

                if self.code == 6:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 40)
                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Rejected',
                                                            'format':red
                                                            })

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Accepted',
                                                            'format':green
                                                            })

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Removed',
                                                            'format':orange
                                                            })
                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Minor Deficiencies',
                                                            'format':light_blue
                                                            })


                worksheet.conditional_format(rowcount, 1, 1, colcount, {'type': 'no_blanks','format': border})
                
                writer.save()
                
                # writer.close()
                create_pdf(self,temp_excel_filepath)

                # word_sol(self,excel_filepath)
                location = self.directory_code 
                word.rejected_word(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item)
            else:
                    print("got u again")
                    error_rejection()
                    pass

        if self.getchoice.item == "Rejection" and self.getchoice.item2 == "Warranty":
            found2 = len(self.df[self.df['Warranty'] == '']) 
            if found2 == 0:
                filtered_df = variation(self)
                filtered_df.to_excel(writer, sheet_name='Sheet1')
                # filtered_df = self.df.loc[:, ~self.df.columns.isin(["OC Surveyor", "Notes.1"])]
                rowcount = self.tableview.model().rowCount()
                colcount = self.tableview.model().columnCount()

                # Get the xlsxwriter workbook and worksheet objects.
                workbook  = writer.book
                cell_format = workbook.add_format()
                cell_format.set_text_wrap()
                red = workbook.add_format({'bg_color': '#ffcccb'})
                green = workbook.add_format({'bg_color': '#90EE90'})
                orange = workbook.add_format({'bg_color': 'yellow'})
                border = workbook.add_format({'border': 1})
                light_blue = workbook.add_format({'bg_color': '#ADD8E6'})  # Light Blue Background

                worksheet = writer.sheets['Sheet1']
                if self.code == 1:
                    worksheet.set_column('A:F', 12)
                    worksheet.set_column('G:G', 30)

                if self.code == 2:
                    worksheet.set_column('A:F', 12)
                    worksheet.set_column('G:G', 30)
            
                if self.code == 3:
                    worksheet.set_column('A:K', 10)
                    worksheet.set_column('L:L', 25)

                if self.code == 4:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 25)

                if self.code == 5:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 20)

                if self.code == 6:
                    worksheet.set_column('A:H', 12)
                    worksheet.set_column('I:I', 20)

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Rejected',
                                                            'format':red
                                                            })

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Accepted',
                                                            'format':green
                                                            })

                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Removed',
                                                            'format':orange
                                                            })
                worksheet.conditional_format('A1:Z100',
                                                            {'type': 'text',
                                                            'criteria': 'containing',
                                                            'value': 'Minor Deficiencies',
                                                            'format':light_blue
                                                            })

                worksheet.conditional_format(rowcount, 1, 1, colcount, {'type': 'no_blanks','format': border})

                writer.save()
                
                # writer.close()
                create_pdf(self,temp_excel_filepath)

                # word_sol(self,excel_filepath)
                location = self.directory_code 
                word.rejected_word(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item)
            else:
                    print("got u again")
                    error_rejection()
                    pass
    except Exception as e:
        QMessageBox.critical(None, "Error", f"An error occurred:\n{str(e)}")
        print(f"Error: {e}")

def variation(self):
    
    if self.code == 1:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin([ "Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df

    if self.code == 2:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin([ "Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df
    
    if self.code == 3:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin([ "Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df

    if self.code == 4:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["OC Surveyor", "Notes.1", "Warranty", "Reviewer.1"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes", "Vendor Surveyor", "Reviewer"])]
            return filtered_df

    if self.code == 5:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df
    

    if self.code == 6:
    
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df
    

def word_sol(self, name):
    """Converts an Excel table to a Word document and saves it as a PDF."""
    
    pythoncom.CoInitialize()  # Initialize COM

    try:
        path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # Define the temporary directory path
        temp_dir = os.path.join(path, "Temp")
        
        # Define the file paths
        original = os.path.join(temp_dir, "test.docx")
        target = os.path.join(temp_dir, "test_copy.docx")
        
        # Construct the final PDF file name
        final_pdf_path = os.path.join(
            self.directory_code,
            f"{self.area}-{self.getchoice.text}-{self.cip_dev}-{self.getchoice.item2}-SQ{self.planfile_entry.text()}(Asset List).pdf"
        )

        print("FINAL PATH:", final_pdf_path)

        # Ensure template exists
        if not os.path.exists(original):
            raise FileNotFoundError(f"Template file not found: {original}")

        # Copy template
        shutil.copyfile(original, target)

        # Initialize Word
        word = client.Dispatch("Word.Application")
        word.Visible = False  # Keep Word hidden
        doc = word.Documents.Open(target)

        # Open Excel
        excel = client.Dispatch("Excel.Application")
        book = excel.Workbooks.Open(name)
        sheet = book.Worksheets[0]

        # Get row and column count from the table
        rowcount = self.tableview.model().rowCount()
        colcount = self.tableview.model().columnCount()

        # Ensure text wraps properly
        sheet.Columns.WrapText = True
        sheet.PageSetup.Orientation = 2  # Set landscape mode if needed

        # Copy the selected table range
        sheet.Range(sheet.Cells(1, 2), sheet.Cells(rowcount + 2, colcount + 1)).Copy()

        # Paste into Word
        wdRange = doc.Content
        wdRange.Collapse(1)  # Move to the start
        wdRange.PasteExcelTable(False, False, False)

        # Save changes in Word
        doc.Save()  # Save before exporting

        # Convert to PDF (Correct SaveAs syntax)
        doc.SaveAs(final_pdf_path, FileFormat=17)  

        print(f"PDF successfully created: {final_pdf_path}")

    except FileNotFoundError as e:
        print(f"Error: {e}")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Ensure Word and Excel are properly closed
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if word:
                word.Quit()
            if book:
                book.Close(SaveChanges=False)
            if excel:
                excel.Application.Quit()
        except Exception as cleanup_error:
            print(f"Error while closing applications: {cleanup_error}")

        pythoncom.CoUninitialize()  # Uninitialize COM

def create_pdf(self, name2):
    try:
        path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


        # Define the temporary directory path
        temp_dir = os.path.join(path, "Temp")  # Create Temp folder path relative to the script directory
        # Define the Excel file paths
        original = os.path.join(temp_dir, "test.docx")
        target = os.path.join(temp_dir, "test_copy.docx")
        # Define file paths
        
        # Check if the Word template file exists
        if not os.path.exists(original):
            raise FileNotFoundError(f"Template file not found: {original}")

        # Copy template to a temporary document
        shutil.copyfile(original, target)

        # Construct final PDF path
        final_pdf_path = os.path.join(
            self.directory_code,
            f"{self.area}-{self.getchoice.text}-{self.cip_dev}-{self.getchoice.item2}-SQ{self.planfile_entry.text()}(Asset List).pdf"
        )

        # Initialize COM objects
        pythoncom.CoInitialize()

        word = client.Dispatch("Word.Application")
        doc = word.Documents.Open(target)

        # Get row and column count from the table
        rowcount = self.tableview.model().rowCount()
        colcount = self.tableview.model().columnCount()

        # Open Excel and get the worksheet
        excel = client.Dispatch("Excel.Application")
        final_copy = excel.Workbooks.Open(name2)
        final_copy_sheet = final_copy.Worksheets[0]

        # Copy the selected table range
        final_copy_sheet.Range(
            final_copy_sheet.Cells(1, 2),
            final_copy_sheet.Cells(rowcount + 2, colcount)
        ).Copy()

        # Paste the table into the Word document
        wdRange = doc.Content
        wdRange.PasteExcelTable(False, False, False)

        # Save as PDF
        doc.SaveAs(final_pdf_path, FileFormat=17)
        print(f"PDF successfully created: {final_pdf_path}")

    except FileNotFoundError as e:
        print(f"Error: {e}")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Ensure Word and Excel are closed
        try:
            if doc:
                doc.Close()
            if word:
                word.Quit()
            if excel:
                excel.Application.Quit()
        except Exception as cleanup_error:
            print(f"Error while closing applications: {cleanup_error}")

        pythoncom.CoUninitialize()
    

def showError():
    msgBox = QMessageBox()
    msgBox.setText("File Not Found")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def error_acceptance():
    msgBox = QMessageBox()
    msgBox.setText("Rejection or empty space found when trying to send Acceptance Letter")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def error_rejection():
    msgBox = QMessageBox()
    msgBox.setText(" Empty space found when trying to send Rejection Letter")
    msgBox.setWindowTitle("Error")
    msgBox.exec()



def check_b4_save(self):

    msgBox = QMessageBox()
    msgBox.setText("All changes will be saved")
    msgBox.setWindowTitle("Save File")
    msgBox.setStandardButtons(
        QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel
    )
    response = msgBox.exec()
    print(response)
    if response == 1024:
        exportToExcel(self, self.tableview)
    else:
        pass
