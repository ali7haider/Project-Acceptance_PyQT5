from PyQt5.QtWidgets import QDialog
from functions.input_dialog_ui import Ui_Dialog

class InputDialog(QDialog):
    def __init__(self, parent=None):
        super(InputDialog, self).__init__(parent)
        self.ui = Ui_Dialog()  # Initialize UI from the .ui file
        self.ui.setupUi(self)  # Setup UI components from the .ui file

        # Apply the stylesheet directly to the InputDialog to prevent inheritance from the parent
         # Apply styles directly to the combo boxes
        combo_box_style = """
        QComboBox {
            color: black;  /* Normal text color */
            background-color: white;  /* Normal background color */
            padding: 10px;
        }

        QComboBox QAbstractItemView {
        color: black;	
        background-color: white;
        padding: 10px;
        selection-background-color: rgb(39, 44, 54);
    }
        """
        
        # Apply stylesheet directly to the specific combo boxes
        self.ui.cmbxWalkthroughOrButton.setStyleSheet(combo_box_style)
        self.ui.cmbxAcceptedOrRejected.setStyleSheet(combo_box_style)


        self.ui.btnEdit.clicked.connect(self.accept)  # Accept dialog when button is clicked

        # Initialize variables to store values
        self.WalkthroughOrButton = None
        self.AcceptedOrRejected = None
        self.inspection_date = None

    def get_values(self):
        """Retrieve updated values from UI elements and store them."""
        self.WalkthroughOrButton = self.ui.cmbxWalkthroughOrButton.currentText()  # Get selected text from ComboBox
        self.AcceptedOrRejected = self.ui.cmbxAcceptedOrRejected.currentText()  # Get selected text from ComboBox
        self.inspection_date = self.ui.dateInspection.date().toString("yyyy-MM-dd")  # Get date in 'YYYY-MM-DD' format
