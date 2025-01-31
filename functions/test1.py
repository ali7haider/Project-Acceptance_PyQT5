from PyQt5.QtWidgets import (
    QMainWindow, QWidget
)
from functions.test import Ui_MainWindow  # Import the generated UI class

class Screen(QMainWindow):
    def __init__(self):
        super(Screen, self).__init__()
        self.ui = Ui_MainWindow()  # Initialize the UI class
        self.ui.setupUi(self)  # Call the setupUi method to set up the UI

        # Initialize variables to store values
        self.WalkthroughOrButton = None
        self.AcceptedOrRejected = None
        self.inspection_date = None

        # Optional: if you have a button to trigger value fetching (e.g., `btnEdit`)
        self.ui.btnEdit.clicked.connect(self.update_values)

    def update_values(self):
        """Update values from UI elements."""
        self.WalkthroughOrButton = self.ui.cmbxWalkthroughOrButton.currentText()  # Get selected text from ComboBox
        self.AcceptedOrRejected = self.ui.cmbxAcceptedOrRejected.currentText()  # Get selected text from ComboBox
        self.inspection_date = self.ui.dateInspection.date().toString("yyyy-MM-dd")  # Get date in 'YYYY-MM-DD' format

        # Print the values to verify
        print(self.WalkthroughOrButton, self.AcceptedOrRejected, self.inspection_date)

    def get_values(self):
        """Retrieve the stored values."""
        return self.WalkthroughOrButton, self.AcceptedOrRejected, self.inspection_date
