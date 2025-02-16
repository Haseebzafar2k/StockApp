import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QFileDialog,
    QHBoxLayout,
    QSizePolicy,
    QSpacerItem,
)
from PyQt6.QtCore import Qt


class SpreadsheetApp(QWidget):
    def __init__(self):
        super().__init__()

        # Button size control variables
        self.custom_button_size = False  # Default: Buttons occupy all available space
        self.button_width = (
            100  # Default width (used only if custom_button_size is True)
        )
        self.button_height = (
            40  # Default height (used only if custom_button_size is True)
        )
        self.button_spacing = (
            10  # Default spacing (used only if custom_button_size is True)
        )

        self.initUI()

    def initUI(self):
        self.setWindowTitle("PyQt Spreadsheet")
        self.showMaximized()  # Open in full screen

        mainLayout = QVBoxLayout()

        # Top Buttons Layout
        topButtonLayout = QHBoxLayout()
        if self.custom_button_size:
            topButtonLayout.setSpacing(self.button_spacing)

        self.endButton = QPushButton("END PROGRAM")
        self.endButton.clicked.connect(self.close)
        self.set_button_size(self.endButton)
        topButtonLayout.addWidget(self.endButton)

        self.addPositionButton = QPushButton("ADD NEW POSITION")
        self.addPositionButton.clicked.connect(self.add_position)
        self.set_button_size(self.addPositionButton)
        topButtonLayout.addWidget(self.addPositionButton)

        self.restorePositionButton = QPushButton("RESTORE POSITION")
        self.restorePositionButton.clicked.connect(self.restore_position)
        self.set_button_size(self.restorePositionButton)
        topButtonLayout.addWidget(self.restorePositionButton)

        topButtonLayout.addStretch()  # Align buttons to the left

        mainLayout.addLayout(topButtonLayout)

        # Table Widget with Scrollbars
        self.tableWidget = QTableWidget(100, 20)  # More rows and columns
        self.tableWidget.setSizeAdjustPolicy(
            QTableWidget.SizeAdjustPolicy.AdjustToContents
        )
        self.tableWidget.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        self.tableWidget.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        mainLayout.addWidget(self.tableWidget)

        # Buttons
        buttonLayout = QHBoxLayout()
        if self.custom_button_size:
            buttonLayout.setSpacing(self.button_spacing)

        self.loadButton = QPushButton("Load Excel")
        self.loadButton.clicked.connect(self.load_excel)
        self.set_button_size(self.loadButton)
        buttonLayout.addWidget(self.loadButton)

        self.saveButton = QPushButton("Save to Excel")
        self.saveButton.clicked.connect(self.save_excel)
        self.set_button_size(self.saveButton)
        buttonLayout.addWidget(self.saveButton)

        buttonLayout.addStretch()  # Align buttons to the left

        mainLayout.addLayout(buttonLayout)
        self.setLayout(mainLayout)

    def set_button_size(self, button):
        if self.custom_button_size:
            button.setFixedSize(self.button_width, self.button_height)
        else:
            button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def add_position(self):
        # Placeholder for adding a new position logic
        print("Add New Position clicked")

    def restore_position(self):
        # Placeholder for restoring a position logic
        print("Restore Position clicked")

    def load_excel(self):
        options = QFileDialog.Option()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx;*.xls)", options=options
        )

        if file_path:
            df = pd.read_excel(file_path)
            self.tableWidget.setRowCount(df.shape[0])
            self.tableWidget.setColumnCount(df.shape[1])

            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def save_excel(self):
        options = QFileDialog.Option()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)", options=options
        )

        if file_path:
            row_count = self.tableWidget.rowCount()
            col_count = self.tableWidget.columnCount()

            data = []
            for i in range(row_count):
                row_data = []
                for j in range(col_count):
                    item = self.tableWidget.item(i, j)
                    row_data.append(item.text() if item else "")
                data.append(row_data)

            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SpreadsheetApp()
    window.show()
    sys.exit(app.exec())
