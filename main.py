import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QStackedWidget, QWidget, QVBoxLayout,
    QCheckBox, QPushButton, QLineEdit, QFormLayout, QMessageBox, QLabel,
    QGroupBox, QScrollArea
)
from form_ui import PerformanceForm
from excel_handler import ExcelHandler
from employee_view import EmployeeViewWindow
from curved_performance_view import CurvedPerformanceView
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class AdminColumnManager(QWidget):
    def __init__(self, parent_app):
        super().__init__()
        self.parent_app = parent_app
        self.setMinimumSize(600, 400)
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', sans-serif;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 15px;
                padding: 10px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QCheckBox {
                font-size: 13px;
                padding: 8px;
                margin: 5px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton#backBtn {
                background-color: #ff9800;
            }
            QPushButton#backBtn:hover {
                background-color: #e68a00;
            }
            QLineEdit {
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 8px;
                background-color: white;
                font-size: 13px;
            }
            QLabel {
                font-size: 13px;
                color: #333;
                padding: 5px;
            }
            QLabel#title {
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
            }
        """)
        self.column_checkboxes = {}
        self.dynamic_layout = None  # Store dynamic_layout for adding new checkboxes
        self.create_ui()

    def create_ui(self):
        main_layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel("Manage Columns")
        title_label.setObjectName("title")
        main_layout.addWidget(title_label)

        # Scrollable content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # Non-mandatory columns group
        dynamic_group = QGroupBox("Dynamic Columns (Show/Hide)")
        self.dynamic_layout = QVBoxLayout(dynamic_group)
        
        headers = self.parent_app.excel_handler.get_headers()
        mandatory_columns = self.parent_app.excel_handler.mandatory_columns
        
        for header in headers:
            if header not in mandatory_columns:
                checkbox = QCheckBox(header)
                checkbox.setChecked(header in self.parent_app.excel_handler.visible_columns)
                checkbox.stateChanged.connect(self.update_column_visibility)
                self.column_checkboxes[header] = checkbox
                self.dynamic_layout.addWidget(checkbox)
        
        scroll_layout.addWidget(dynamic_group)

        # Mandatory columns group
        mandatory_group = QGroupBox("Mandatory Columns (Always Visible)")
        mandatory_layout = QVBoxLayout(mandatory_group)
        for header in mandatory_columns:
            label = QLabel(f"• {header}")
            label.setStyleSheet("color: #666; font-style: italic; margin-left: 10px;")
            mandatory_layout.addWidget(label)
        
        scroll_layout.addWidget(mandatory_group)

        # Add new column form
        add_column_group = QGroupBox("Add New Column")
        add_column_layout = QFormLayout(add_column_group)
        self.new_column_name = QLineEdit()
        self.new_column_name.setPlaceholderText("Enter new column name (e.g., Bonus Amount)")
        add_button = QPushButton("Add Column")
        add_button.clicked.connect(self.add_new_column)
        add_column_layout.addRow("Column Name:", self.new_column_name)
        add_column_layout.addRow(add_button)
        
        scroll_layout.addWidget(add_column_group)
        
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)

        # Back button
        back_btn = QPushButton("Back to Form")
        back_btn.setObjectName("backBtn")
        back_btn.clicked.connect(self.parent_app.go_back)
        main_layout.addWidget(back_btn)
        
        main_layout.addStretch()

    def update_column_visibility(self):
        visible_columns = self.parent_app.excel_handler.mandatory_columns + [
            header for header, checkbox in self.column_checkboxes.items() if checkbox.isChecked()
        ]
        self.parent_app.excel_handler.update_visible_columns(visible_columns)
        self.parent_app.reload_form()
        QMessageBox.information(self, "Success", "Column visibility updated!")

    def add_new_column(self):
        column_name = self.new_column_name.text().strip()
        if column_name:
            success, message = self.parent_app.excel_handler.add_column(column_name)
            if success:
                checkbox = QCheckBox(column_name)
                checkbox.setChecked(True)
                checkbox.stateChanged.connect(self.update_column_visibility)
                self.column_checkboxes[column_name] = checkbox
                self.dynamic_layout.addWidget(checkbox)
                self.parent_app.reload_form()
                QMessageBox.information(self, "Success", message)
            else:
                QMessageBox.critical(self, "Error", message)
            self.new_column_name.clear()

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        logging.info("Initializing App...")
        self.setWindowTitle("Performance Review System")
        self.setMinimumSize(1200, 800)

        # Create stacked widget to manage screens
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Initialize ExcelHandler
        self.excel_handler = ExcelHandler()
        logging.info("ExcelHandler created.")

        # Initialize screens
        logging.info("Creating PerformanceForm...")
        self.form = PerformanceForm(self)
        self.form.parent_app = self
        logging.info("Form created.")

        # Add form to stacked widget
        self.stacked_widget.addWidget(self.form)
        logging.info("Form added to stacked widget.")

        # Initialize other screens (will be created on demand)
        self.employee_view = None
        self.curved_view = None
        self.admin_panel = None

        logging.info("Loading employees...")
        self.load_employees()
        logging.info("Employees loaded.")
        logging.info("App initialization complete.")

    def load_employees(self):
        logging.info("Loading employees function called...")
        try:
            employees = self.excel_handler.get_all_employees()
            logging.info(f"Employees fetched: {len(employees)}")
            if employees:
                self.form.populate_employee_dropdown(employees)
                logging.info("Dropdown populated.")
            else:
                logging.info("No employees found in Excel file.")
                self.form.show_error_message("No employees found in the Excel file. Please add employees.")
        except Exception as e:
            logging.error(f"Error loading employees: {e}")
            self.form.show_error_message(f"Error loading employees: {str(e)}")

    def save_form_data(self):
        try:
            form_data = self.form.get_form_data()
            if form_data:
                emp_id = form_data.get("Employee ID")
                existing_data = self.excel_handler.get_employee_data(emp_id)
                final_data = existing_data if existing_data else {}
                final_data.update(form_data)
                logging.info(f"Final data to save for {emp_id} (soft skills): {{k: v for k, v in final_data.items() if 'Rating' in k or 'Score' in k or k == 'Part_B_Total_Score'}}")
                self.excel_handler.save_employee_data(final_data)
                self.form.show_success_message("Data saved successfully!")
                self.form.clear_form()
        except Exception as e:
            self.form.show_error_message(f"Error saving data: {str(e)}")

    def add_new_employee(self, emp_id, name, department, designation, joining_date, contract_expiry, division, exp_pmtf):
        try:
            success, message = self.excel_handler.add_new_employee(
                emp_id, name, department, designation,
                joining_date, contract_expiry, division, exp_pmtf
            )
            return success, message
        except Exception as e:
            return False, f"Error adding employee: {str(e)}"

    def open_employee_view(self, emp_id):
        try:
            logging.info(f"Opening employee view for ID: {emp_id}")
            employee_data = self.excel_handler.get_employee_data(emp_id)
            if employee_data:
                if self.employee_view is None:
                    self.employee_view = EmployeeViewWindow(employee_data, self)
                    self.stacked_widget.addWidget(self.employee_view)
                else:
                    self.employee_view.update_employee_data(employee_data)
                self.stacked_widget.setCurrentWidget(self.employee_view)
            else:
                self.form.show_error_message("Employee not found!")
        except Exception as e:
            self.form.show_error_message(f"Error opening employee view: {str(e)}")

    def open_curved_view(self):
        try:
            logging.info("Opening curved view...")
            if self.curved_view is None:
                logging.info("Creating CurvedPerformanceView...")
                self.curved_view = CurvedPerformanceView(self)
                self.stacked_widget.addWidget(self.curved_view)
                logging.info("Curved view created.")
            self.stacked_widget.setCurrentWidget(self.curved_view)
        except Exception as e:
            self.form.show_error_message(f"Error opening curved view: {str(e)}")

    def open_admin_panel(self):
        try:
            logging.info("Opening admin panel...")
            if self.admin_panel is None:
                logging.info("Creating AdminColumnManager...")
                self.admin_panel = AdminColumnManager(self)
                self.stacked_widget.addWidget(self.admin_panel)
                logging.info("Admin panel created.")
            self.stacked_widget.setCurrentWidget(self.admin_panel)
        except Exception as e:
            self.form.show_error_message(f"Error opening admin panel: {str(e)}")

    def reload_form(self):
        logging.info("Reloading PerformanceForm...")
        try:
            old_form = self.form
            self.form = PerformanceForm(self)
            self.form.parent_app = self
            old_index = self.stacked_widget.indexOf(old_form)
            self.stacked_widget.removeWidget(old_form)
            self.stacked_widget.insertWidget(old_index, self.form)
            self.load_employees()
            logging.info("Form reloaded.")
        except Exception as e:
            logging.error(f"Error reloading form: {str(e)}")
            self.form.show_error_message(f"Error reloading form: {str(e)}")

    def go_back(self):
        self.stacked_widget.setCurrentWidget(self.form)

if __name__ == '__main__':
    logging.info("Main execution starting...")
    app = QApplication(sys.argv)
    logging.info("QApplication created.")
    main_window = App()
    logging.info("App object created, showing window...")
    main_window.show()
    logging.info("Starting event loop...")
    sys.exit(app.exec_())