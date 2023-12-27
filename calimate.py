"""
Calimate - Automated Tektronics oscilloscope calibration
"""

from PyQt5.QtWidgets import (QAbstractItemView, QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
                             QHBoxLayout, QVBoxLayout, QWidget, QPushButton, QFileDialog, QFrame, QSplitter,
                             QSizePolicy)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtWidgets import QTextEdit

import csv
import locale
import json
import os
import pyvisa
import sys

# Use '' for auto, or force e.g. to 'en_US.UTF-8'
locale.setlocale(locale.LC_ALL, '')

# Prints detailed pyVISA debug info to console
pyvisa.log_to_screen()


class MainWindow(QMainWindow):

    app_ver = "0.1"
    invalid_packet_count = 0
    messages_count = 0
    csv_file_name = None

    def __init__(self):
        super().__init__()

        print(f"Current working directory: {os.getcwd()}")

        self.setWindowTitle(f"Calimate v{MainWindow.app_ver}")
        self.resize(1050, 600)

        self.columns = ("Test", "Result", "Notes")

        # Initialize the table
        self.data_table = QTableWidget(0, len(self.columns))
        self.data_table.setHorizontalHeaderLabels(self.columns)
        # Hide the vertical header (row numbers)
        self.data_table.verticalHeader().hide()
        # Set selection mode to single row selection
        self.data_table.setSelectionMode(QAbstractItemView.SingleSelection)
        # Connect the sectionResized signal to adjust_row_heights
        self.data_table.horizontalHeader().sectionResized.connect(self.adjust_row_heights)
        if sys.platform.startswith("win"):
            self.data_table.setFont(QFont("Courier"))
        else:
            self.data_table.setFont(QFont("Monospace"))
        self.data_table.setSortingEnabled(True)
        self.data_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.data_table.selectionModel().selectionChanged.connect(self.update_inst)

        # Create a QHBoxLayout for buttons
        self.button_layout = QHBoxLayout()
        self.import_button = QPushButton("Import CSV")
        self.import_button.clicked.connect(self.import_table_from_csv)
        self.import_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        # Add button to layout
        self.button_layout.addWidget(self.import_button)

        # Initialize the instrument view, read-only text box with monospace font
        self.inst_textbox = QTextEdit()
        if sys.platform.startswith("win"):
            self.inst_textbox.setFont(QFont("Courier"))
        else:
            self.inst_textbox.setFont(QFont("Monospace"))
        self.inst_textbox.setReadOnly(True)
        self.inst_textbox.setPlainText(f"Calimate\nVersion {MainWindow.app_ver}")

        # Arrange the widgets
        inst_frame = QFrame()
        inst_frame.setContentsMargins(0, 0, 0, 0)
        # inst_layout = QVBoxLayout()
        # inst_layout.addWidget(self.inst_textbox)
        # inst_layout.setContentsMargins(0, 0, 0, 0)
        # inst_frame.setLayout(inst_layout)

        # Create a container for instrument buttons
        self.inst_layout = QVBoxLayout()

        # Create and configure the Search button
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.find_inst)  # Connect the button to find_inst
        self.inst_layout.addWidget(self.search_button)

        # Create a container (scroll area) for instrument buttons
        self.inst_button_container = QVBoxLayout()

        # Add the instrument button container to the main instrument layout
        self.inst_layout.addLayout(self.inst_button_container)

        # Assign the layout to inst_frame
        inst_frame.setLayout(self.inst_layout)

        table_frame = QFrame()
        table_frame.setContentsMargins(0, 0, 0, 0)
        table_layout = QVBoxLayout()
        table_layout.addWidget(self.data_table)
        table_layout.addLayout(self.button_layout)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_frame.setLayout(table_layout)

        # Set the layout margins to zero (or a smaller value)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)  # Set spacing to 0

        # Create the QSplitter with reduced handle width for less space consumption
        main_splitter = QSplitter(Qt.Horizontal)
        # Reduce the splitter handle width to 1 pixel
        main_splitter.setHandleWidth(1)
        main_splitter.addWidget(inst_frame)
        main_splitter.addWidget(table_frame)

        # Adjust these stretch factors to change the initial width of the panels
        # Reduce the stretch factor of the table
        main_splitter.setStretchFactor(0, 2)
        # Increase the stretch factor of the instrument panel
        main_splitter.setStretchFactor(1, 3)

        widget = QWidget()
        widget.setLayout(main_layout)
        widget.layout().addWidget(main_splitter)
        self.setCentralWidget(widget)

        # Set the width of each column here
        column_widths = [80, 80, 440]
        for i, width in enumerate(column_widths):
            self.data_table.setColumnWidth(i, width)

        # Create and set the status bar
        self.status_bar = self.statusBar()
        self.update_status_bar("Welcome to Calimate!")

    def find_inst(self):
        # Clear the previous list of instrument buttons
        # This safely removes all widgets in the inst_button_container layout
        while self.inst_button_container.count():
            child = self.inst_button_container.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        # Now proceed with finding and listing instruments
        rm = pyvisa.ResourceManager()
        inst_list = rm.list_resources()
        for item in inst_list:
            try:
                # Open connection to the instrument
                inst = rm.open_resource(item)
                inst_id = inst.query("*IDN?").strip()
                # TODO: Try other "ID" commands if "IDN?" returns nothing.

                # Return the instrument to local control
                # TODO: This command shoudl be retrieved from the JSON file (done later).
                inst.write(":KEY:FORCe")  # Unlocks the remote control
                inst.close()  # Close the connection after getting ID

                # Construct a filename from the instrument ID
                fn_inst = inst_id.replace(' ', '_').replace(':', '')
                fn_inst_list = fn_inst.split(',')
                json_filename = f"{fn_inst_list[0]}_{fn_inst_list[1]}.json"
                print(f"JSON filename: {json_filename}")

                if os.path.exists(json_filename):
                    with open(json_filename, 'r') as file:
                        data = json.load(file)
                        connect_cmd = data.get("connect")
                        id_cmd = data.get("id")
                        close_cmd = data.get("close")

                        # Use the commands (for demonstration, printing them)
                        print(f"Connect Command: {connect_cmd}")
                        print(f"ID Command: {id_cmd}")
                        print(f"Close Command: {close_cmd}")

                        # Create a button with the instrument's info
                        inst_list = inst_id.split(',')
                        inst_button = QPushButton(
                            f"{inst_list[0]} {inst_list[1]}\nS/N: {inst_list[2]}\nVer: {inst_list[3]}")
                        inst_button.clicked.connect(lambda checked, inst=item: self.select_inst(
                            f"{inst_id}", connect_cmd, id_cmd, close_cmd))
                        self.inst_button_container.addWidget(inst_button)
                else:
                    print(f"Configuration file {json_filename} not found.")

            except Exception as e:
                error_msg = f"Exception: {e}"
                print(error_msg)

    def select_inst(self, inst_id, connect_cmd, id_cmd, close_cmd):
        # Example function showing how you might use the commands
        print(f"Selected: {inst_id}")
        print(f"Connect Command: {connect_cmd}")
        print(f"ID Command: {id_cmd}")
        print(f"Close Command: {close_cmd}")

        self.status_bar.showMessage(f"Selected instrument: {inst_id}")

    def update_status_bar(self, message=None):
        if message:
            self.status_bar.showMessage(message)
        else:
            total_packets = self.data_table.rowCount()
            self.status_bar.showMessage(f"Total packets: {total_packets}")

    def add_data_to_table(self, row_data, adjust_row_height=False):
        self.data_table.setSortingEnabled(False)

        row = self.data_table.rowCount()
        self.data_table.insertRow(row)

        # Variable to store the maximum height required for this row
        max_row_height = 5  # self.data_table.rowHeight(row)

        is_msg_row = False

        for col_index, data_item in enumerate(row_data):
            if data_item == "Msg":
                is_msg_row = True
                MainWindow.messages_count += 1

        for col_index, data_item in enumerate(row_data):
            item = QTableWidgetItem(str(data_item))
            item.setTextAlignment(Qt.AlignLeft)

            # Set the item to be non-editable
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            self.data_table.setItem(row, col_index, item)

            # Check if we're in the "Result" column and set background color
            if col_index == self.columns.index("Result"):
                if is_msg_row:
                    # Semi-transparent 'Material' indigo color for "Msg"
                    item.setBackground(QColor(33, 150, 243, 64))
                elif data_item == "Pass":
                    # Semi-transparent 'Material' green color for Pass
                    item.setBackground(QColor(76, 175, 80, 64))
                elif data_item == "Fail":
                    # Semi-transparent 'Material' red color for Fail
                    item.setBackground(QColor(244, 67, 54, 64))

            # Check if we're in the "Notes" column and if is message row
            if col_index == self.columns.index("Notes"):
                if is_msg_row:
                    # Semi-transparent 'Material' indigo color for "Msg"
                    item.setBackground(QColor(33, 150, 243, 64))

            # Update the row height if needed
            if adjust_row_height:
                self.data_table.resizeRowToContents(row)
                row_height = self.data_table.rowHeight(row)
                if row_height > max_row_height:
                    max_row_height = row_height

        # Set the row to the maximum height required
        if adjust_row_height:
            self.data_table.setRowHeight(row, max_row_height)

        self.data_table.setSortingEnabled(True)

    def adjust_row_heights(self):
        for row in range(self.data_table.rowCount()):
            self.data_table.resizeRowToContents(row)

    def update_inst(self):
        selected_rows = self.data_table.selectionModel().selectedRows()
        if selected_rows:

            # Get the "Notes" column cell's text
            packet_column_index = self.columns.index("Notes")
            packet_cell = self.data_table.item(
                selected_rows[0].row(), packet_column_index).text()

            # Get the "Result" column cell's text
            dir_column_index = self.columns.index("Result")
            dir_cell = self.data_table.item(
                selected_rows[0].row(), dir_column_index).text()
            if "Msg" in dir_cell:
                # Message row, not packet data, so don't attempt parsing.
                self.inst_textbox.clear()
                self.inst_textbox.setPlainText(
                    f"Message:\n- {packet_cell}")
                return

            # Split the packet data into a list and parse it
            # packet_data_list = packet_cell.split()
            # parsed_output = parse_packet(packet_data_list)

            # Update the instrument text box with the parsed data
            # self.inst_textbox.setPlainText(parsed_output)

            self.inst_textbox.setPlainText(packet_cell)

            # Update status bar with selected packet number and total packets
            # Adding 1 for human-readable numbering (1-indexed)
            selected_record_num = selected_rows[0].row() + 1
            total_records = self.data_table.rowCount()
            self.update_status_bar(
                f"File: '{MainWindow.csv_file_name}', "
                f"Selected record: {selected_record_num:n} of {total_records:n}.")

    def import_table_from_csv(self):
        """
        Import data to the table from a CSV file selected by the user.
        """
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(
            self, "QFileDialog.getOpenFileName()", "", "CSV Files (*.csv);;All Files (*)", options=options)
        if fileName:
            MainWindow.csv_file_name = os.path.basename(fileName)
            self.inst_textbox.setPlainText(
                f"Loading file '{MainWindow.csv_file_name}'...")
            self.update_status_bar(
                f"Loading file '{MainWindow.csv_file_name}'...")

            # Force GUI to show our Loading messages.
            QApplication.processEvents()

            try:
                with open(fileName, 'r', newline='') as f:
                    reader = csv.reader(f)

                    # Clear existing data in the table
                    self.data_table.clearContents()
                    self.data_table.setRowCount(0)
                    MainWindow.invalid_packet_count = 0
                    MainWindow.messages_count = 0

                    # Clear the right-pane:
                    self.inst_textbox.clear()

                    self.data_table.setSortingEnabled(False)
                    for row in reader:
                        self.add_data_to_table(row, False)
                    self.adjust_row_heights()
                    self.data_table.setSortingEnabled(True)
                    # Column index 0 (Time)
                    self.data_table.sortItems(0, Qt.AscendingOrder)

                    total_records = self.data_table.rowCount()
                    self.inst_textbox.setPlainText(
                        f"File: '{MainWindow.csv_file_name}' successfully imported\n\n"
                        f"- Total records:   {total_records:8n}\n"
                    )
                    self.update_status_bar(
                        f"File: '{MainWindow.csv_file_name}'. "
                        f"Total records: {total_records:n}."
                    )

            except csv.Error as e:
                error_msg = f"Error reading CSV file at line {reader.line_num}: {e}"
                print(error_msg)
                self.inst_textbox.setPlainText(
                    f"ERROR! File: '{MainWindow.csv_file_name}'\n\n"
                    f"- {error_msg}"
                )
                self.update_status_bar(error_msg)

            except FileNotFoundError as e:
                error_msg = f"File not found: {e}"
                print(error_msg)
                self.inst_textbox.setPlainText(
                    f"ERROR! File: '{MainWindow.csv_file_name}'\n\n"
                    f"- {error_msg}"
                )
                self.update_status_bar(error_msg)

            except ValueError as e:
                error_msg = f"ValueError: {e}"
                print(e)
                print(error_msg)
                self.inst_textbox.setPlainText(
                    f"ERROR! File: '{MainWindow.csv_file_name}'\n\n"
                    f"- {error_msg}"
                )
                self.update_status_bar(error_msg)

            except Exception as e:
                error_msg = f"Exception: {e}"
                print(e)
                print(error_msg)
                self.inst_textbox.setPlainText(
                    f"ERROR! File: '{MainWindow.csv_file_name}'\n\n"
                    f"- {error_msg}"
                )
                self.update_status_bar(error_msg)


app = QApplication([])
window = MainWindow()
window.show()
app.exec_()
