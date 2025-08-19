from PyQt5.QtWidgets import (
    QWidget, QLabel, QComboBox, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLineEdit, QFormLayout, QTextEdit, QMessageBox, 
    QScrollArea, QGroupBox, QFrame, QSpinBox, QCheckBox, QDialog, QDialogButtonBox, QDateEdit
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont
from datetime import datetime

class PerformanceForm(QWidget):
    def __init__(self, parent_app=None):
        super().__init__()
        self.setMinimumSize(1200, 800)
        self.setStyleSheet("""
            QWidget { background-color: #f5f5f5; font-family: 'Segoe UI', sans-serif; }
            QGroupBox { font-weight: bold; font-size: 12px; border: 2px solid #cccccc; border-radius: 5px; margin-top: 10px; padding-top: 10px; background-color: white; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
            QLineEdit, QTextEdit, QComboBox, QSpinBox { border: 1px solid #ddd; border-radius: 3px; padding: 5px; background-color: white; }
            QPushButton { background-color: #4CAF50; color: white; border: none; padding: 8px 16px; border-radius: 4px; font-weight: bold; }
            QPushButton:hover { background-color: #55DB4D; }
            QPushButton:pressed { background-color: #3d8b40; }
            QPushButton#searchBtn, QPushButton#viewCurvedBtn, QPushButton#adminBtn { background-color: #0886C2; }
            QPushButton#searchBtn:hover, QPushButton#viewCurvedBtn:hover, QPushButton#adminBtn:hover { background-color: #16AAF0; }
            QPushButton#resetBtn { background-color: #0886C2; }
            QPushButton#resetBtn:hover { background-color: #16AAF0; }
            QPushButton#addEmpBtn { background-color: #0886C2; }
            QPushButton#addEmpBtn:hover { background-color: #16AAF0; }
        """)
        
        self.parent_app = parent_app
        self.kpi_widgets = {}
        self.soft_skill_widgets = {}
        self.input_widgets = {}
        self.soft_skills_mapping = {
            "Open & Clear Communication": "Open_Clear_Communication",
            "Attitude, Team Work & Collaboration": "Attitude_Team_Work_Collaboration",
            "Planning & Achievement Focus": "Planning_Achievement_Focus",
            "Creativity & Initiatives": "Creativity_Initiatives",
            "Ownership & Self Accountability": "Ownership_Self_Accountability"
        }
        self.create_ui()

    def create_ui(self):
        main_layout = QVBoxLayout(self)
        
        header_label = QLabel("XYZ\nPerformance Evaluation Form")
        header_label.setAlignment(Qt.AlignCenter)
        header_label.setFont(QFont("Arial", 16, QFont.Bold))
        header_label.setStyleSheet("background-color: #2c3e50; color: white; padding: 15px; border-radius: 5px;")
        main_layout.addWidget(header_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # Get visible headers from ExcelHandler
        headers = self.parent_app.excel_handler.get_visible_headers()

        # Employee Information Group
        emp_info_group = QGroupBox("Employee Information")
        emp_info_layout = QGridLayout(emp_info_group)
        
        row, col = 0, 0
        if "Employee ID" in headers:
            self.employee_combo = QComboBox()
            self.employee_combo.setEditable(True)
            self.employee_combo.currentTextChanged.connect(self.on_employee_selected)
            self.input_widgets["Employee ID"] = self.employee_combo
            emp_info_layout.addWidget(QLabel("Employee ID:"), row, col)
            emp_info_layout.addWidget(self.employee_combo, row, col+1)
            self.search_btn = QPushButton("Search Employee")
            self.search_btn.setObjectName("searchBtn")
            self.search_btn.clicked.connect(self.search_employee)
            emp_info_layout.addWidget(self.search_btn, row, col+2)
            col += 3
            if col >= 4:
                row += 1
                col = 0

        mandatory_fields = [
            "Employee Name", "Division", "Department", "Designation",
            "Date of Joining", "Line Manager", "Date of Evaluation", "Contract Expiry Date", "Exp in xyz"
        ]
        for field in mandatory_fields:
            if field in headers:
                input_widget = QLineEdit()
                if field == "Date of Joining":
                    input_widget.setText("01/07/2025")
                elif field == "Date of Evaluation":
                    input_widget.setText(datetime.now().strftime("%Y-%m-%d"))
                elif field == "Entity Name":
                    input_widget.setText("xyz")
                self.input_widgets[field] = input_widget
                emp_info_layout.addWidget(QLabel(field + ":"), row, col)
                emp_info_layout.addWidget(input_widget, row, col+1)
                col += 2
                if col >= 4:
                    row += 1
                    col = 0

        # Hidden fields
        self.contract_expiry_hidden = QLineEdit()
        self.division_hidden = QLineEdit()
        self.exp_pmtf_hidden = QLineEdit()
        self.contract_expiry_hidden.hide()
        self.division_hidden.hide()
        self.exp_pmtf_hidden.hide()
        
        scroll_layout.addWidget(emp_info_group)

        # Part A: KPIs
        part_a_group = QGroupBox("PART A: Performance Targets / Objectives / KPIs - 70%")
        part_a_layout = QVBoxLayout(part_a_group)
        
        instructions = QLabel("Each target should be SMART (Specific, Measurable, Achievable, Realistic/Relevant, Time bound).\n"
                             "Each Target should carry a weightage of at least 10% and no more than 50%, with a minimum of 4 and maximum of 6 targets.")
        instructions.setWordWrap(True)
        instructions.setStyleSheet("font-style: italic; color: #666; margin-bottom: 10px;")
        part_a_layout.addWidget(instructions)
        
        kpi_table_layout = QGridLayout()
        kpi_headers = ["KPI Title", "Description", "Rating", "Weightage (%)", "Weighted Score"]
        for i, header in enumerate(kpi_headers):
            label = QLabel(header)
            label.setFont(QFont("Arial", 10, QFont.Bold))
            label.setStyleSheet("background-color: #e8e8e8; padding: 5px; border: 1px solid #ccc;")
            kpi_table_layout.addWidget(label, 0, i)
        
        ratings = ["Select Rating", "Serious Performance Concerns (1)", "Below Expectations (2)", 
                  "Meets Expectations (3)", "Exceeds Expectations (4)", "Outstanding (5)"]
        
        for i in range(1, 7):
            kpi_data = {}
            kpi_title = f"KPI_{i}_Title"
            kpi_desc = f"KPI_{i}_Description"
            kpi_rating = f"KPI_{i}_Rating"
            kpi_weight = f"KPI_{i}_Weightage"
            kpi_score = f"KPI_{i}_Weighted_Score"
            
            if kpi_title in headers:
                title_input = QLineEdit()
                title_input.setPlaceholderText(f"KPI {i} Title")
                kpi_data['title'] = title_input
                kpi_table_layout.addWidget(title_input, i, 0)
                self.input_widgets[kpi_title] = title_input
            
            if kpi_desc in headers:
                desc_input = QLineEdit()
                desc_input.setPlaceholderText("Description & Measures")
                kpi_data['description'] = desc_input
                kpi_table_layout.addWidget(desc_input, i, 1)
                self.input_widgets[kpi_desc] = desc_input
            
            if kpi_rating in headers:
                rating_combo = QComboBox()
                rating_combo.addItems(ratings)
                rating_combo.currentTextChanged.connect(self.calculate_scores)
                kpi_data['rating'] = rating_combo
                kpi_table_layout.addWidget(rating_combo, i, 2)
                self.input_widgets[kpi_rating] = rating_combo
            
            if kpi_weight in headers:
                weightage_spin = QSpinBox()
                weightage_spin.setRange(10, 50)
                weightage_spin.setValue(15)
                weightage_spin.setSuffix("%")
                weightage_spin.valueChanged.connect(self.calculate_scores)
                kpi_data['weightage'] = weightage_spin
                kpi_table_layout.addWidget(weightage_spin, i, 3)
                self.input_widgets[kpi_weight] = weightage_spin
            
            if kpi_score in headers:
                score_label = QLabel("0.0")
                score_label.setStyleSheet("background-color: #f0f0f0; padding: 5px; border: 1px solid #ccc;")
                kpi_data['score'] = score_label
                kpi_table_layout.addWidget(score_label, i, 4)
                self.input_widgets[kpi_score] = score_label
            
            self.kpi_widgets[f'kpi_{i}'] = kpi_data
        
        part_a_layout.addLayout(kpi_table_layout)
        
        part_a_total_layout = QHBoxLayout()
        part_a_total_layout.addStretch()
        part_a_total_layout.addWidget(QLabel("Total (Part A):"))
        self.part_a_total_label = QLabel("0.0")
        self.part_a_total_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.part_a_total_label.setStyleSheet("background-color: #e8f5e8; padding: 5px; border: 2px solid #4CAF50;")
        part_a_layout.addLayout(part_a_total_layout)
        if "Part_A_Total_Score" in headers:
            self.input_widgets["Part_A_Total_Score"] = self.part_a_total_label
        
        scroll_layout.addWidget(part_a_group)

        # Part B: Soft Skills
        part_b_group = QGroupBox("PART B: Soft Skills (Key Behavioral Competencies) - 30%")
        part_b_layout = QVBoxLayout(part_b_group)
        
        soft_skills_layout = QGridLayout()
        soft_skills_headers = ["Soft Skills", "Rating", "Weighted Rating (%)"]
        for i, header in enumerate(soft_skills_headers):
            label = QLabel(header)
            label.setFont(QFont("Arial", 10, QFont.Bold))
            label.setStyleSheet("background-color: #e8e8e8; padding: 5px; border: 1px solid #ccc;")
            soft_skills_layout.addWidget(label, 0, i)
        
        soft_skills = [
            "Open & Clear Communication",
            "Attitude, Team Work & Collaboration",
            "Planning & Achievement Focus",
            "Creativity & Initiatives",
            "Ownership & Self Accountability"
        ]
        
        rating_options = ["Select Rating", "Does not Demonstrate (1)", "Developing (2)",
                         "Proficient (3)", "Proficient (4)", "Expert (5)"]
        
        for i, skill in enumerate(soft_skills, 1):
            skill_label = QLabel(skill)
            clean_name = self.soft_skills_mapping[skill]
            rating_key = f"{clean_name}_Rating"
            score_key = f"{clean_name}_Weighted_Score"
            
            if rating_key in headers:
                rating_combo = QComboBox()
                rating_combo.addItems(rating_options)
                rating_combo.currentTextChanged.connect(lambda text, s=skill: self.calculate_scores(s))
                soft_skills_layout.addWidget(skill_label, i, 0)
                soft_skills_layout.addWidget(rating_combo, i, 1)
                self.input_widgets[rating_key] = rating_combo
                self.soft_skill_widgets[clean_name] = {'rating': rating_combo}
                
                if score_key in headers:
                    score_label = QLabel("0.0")
                    score_label.setStyleSheet("background-color: #f0f0f0; padding: 5px; border: 1px solid #ccc;")
                    soft_skills_layout.addWidget(score_label, i, 2)
                    self.input_widgets[score_key] = score_label
                    self.soft_skill_widgets[clean_name]['score'] = score_label
        
        part_b_layout.addLayout(soft_skills_layout)
        
        part_b_total_layout = QHBoxLayout()
        part_b_total_layout.addStretch()
        part_b_total_layout.addWidget(QLabel("Total (Part B):"))
        self.part_b_total_label = QLabel("0.0")
        self.part_b_total_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.part_b_total_label.setStyleSheet("background-color: #e8f5e8; padding: 5px; border: 2px solid #4CAF50;")
        part_b_layout.addLayout(part_b_total_layout)
        if "Part_B_Total_Score" in headers:
            self.input_widgets["Part_B_Total_Score"] = self.part_b_total_label
        
        scroll_layout.addWidget(part_b_group)

        # Overall Performance
        overall_group = QGroupBox("Overall Performance Rating & Recommendations")
        overall_layout = QGridLayout(overall_group)
        
        if "Overall_Rating" in headers:
            overall_layout.addWidget(QLabel("Overall Performance Rating:"), 0, 0)
            self.overall_rating_label = QLabel("0.0")
            self.overall_rating_label.setFont(QFont("Arial", 14, QFont.Bold))
            self.overall_rating_label.setStyleSheet("background-color: #fff3cd; padding: 10px; border: 2px solid #ffc107;")
            overall_layout.addWidget(self.overall_rating_label, 0, 1)
            self.input_widgets["Overall_Rating"] = self.overall_rating_label
        
        if "Overall_Percentage" in headers:
            overall_layout.addWidget(QLabel("Overall Percentage:"), 0, 2)
            self.overall_percentage_label = QLabel("0.0%")
            self.overall_percentage_label.setFont(QFont("Arial", 14, QFont.Bold))
            self.overall_percentage_label.setStyleSheet("background-color: #fff3cd; padding: 10px; border: 2px solid #ffc107;")
            overall_layout.addWidget(self.overall_percentage_label, 0, 3)
            self.input_widgets["Overall_Percentage"] = self.overall_percentage_label
        
        if "Promotion_Recommendation" in headers:
            overall_layout.addWidget(QLabel("Promotion Recommendation:"), 1, 0)
            self.promotion_combo = QComboBox()
            self.promotion_combo.addItems(["Select", "Yes", "No"])
            overall_layout.addWidget(self.promotion_combo, 1, 1)
            self.input_widgets["Promotion_Recommendation"] = self.promotion_combo
        
        if "Retention_Recommendation" in headers:
            overall_layout.addWidget(QLabel("Retention Recommendation:"), 1, 2)
            self.retention_combo = QComboBox()
            self.retention_combo.addItems(["Select", "Yes", "No"])
            overall_layout.addWidget(self.retention_combo, 1, 3)
            self.input_widgets["Retention_Recommendation"] = self.retention_combo
        
        scroll_layout.addWidget(overall_group)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_form)
        button_layout.addWidget(save_btn)
        
        reset_btn = QPushButton("Reset")
        reset_btn.setObjectName("resetBtn")
        reset_btn.clicked.connect(self.clear_form)
        button_layout.addWidget(reset_btn)
        
        add_emp_btn = QPushButton("Add New Employee")
        add_emp_btn.setObjectName("addEmpBtn")
        add_emp_btn.clicked.connect(self.show_add_employee_dialog)
        button_layout.addWidget(add_emp_btn)
        
        view_curved_btn = QPushButton("View Curved")
        view_curved_btn.setObjectName("viewCurvedBtn")
        view_curved_btn.clicked.connect(self.view_curved)
        button_layout.addWidget(view_curved_btn)
        
        admin_btn = QPushButton("Manage Columns")
        admin_btn.setObjectName("adminBtn")
        admin_btn.clicked.connect(self.parent_app.open_admin_panel)
        button_layout.addWidget(admin_btn)
        
        scroll_layout.addLayout(button_layout)
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)

    def calculate_scores(self, skill=None):
        try:
            part_a_total = 0.0
            headers = self.parent_app.excel_handler.get_visible_headers()
            for i in range(1, 7):
                kpi = self.kpi_widgets.get(f'kpi_{i}', {})
                rating_key = f"KPI_{i}_Rating"
                weight_key = f"KPI_{i}_Weightage"
                score_key = f"KPI_{i}_Weighted_Score"
                if rating_key in headers and weight_key in headers and score_key in headers:
                    rating_text = kpi['rating'].currentText()
                    weightage = kpi['weightage'].value()
                    rating_value = self.get_rating_value(rating_text)
                    score = rating_value * (weightage / 100.0) * 70
                    kpi['score'].setText(f"{score:.1f}")
                    part_a_total += score
            self.part_a_total_label.setText(f"{part_a_total:.1f}")
            
            part_b_total = 0.0
            for skill, clean_name in self.soft_skills_mapping.items():
                rating_key = f"{clean_name}_Rating"
                score_key = f"{clean_name}_Weighted_Score"
                if rating_key in headers and score_key in headers:
                    widgets = self.soft_skill_widgets.get(clean_name, {})
                    rating_text = widgets['rating'].currentText()
                    rating_value = self.get_rating_value(rating_text, is_soft_skill=True)
                    score = rating_value * (6 / 100.0) * 30
                    widgets['score'].setText(f"{score:.1f}")
                    part_b_total += score
            self.part_b_total_label.setText(f"{part_b_total:.1f}")
            
            overall_score = part_a_total + part_b_total
            self.overall_percentage_label.setText(f"{overall_score:.1f}%")
            
            if overall_score >= 90:
                rating = "Outstanding (5)"
            elif overall_score >= 80:
                rating = "Exceeds Expectations (4)"
            elif overall_score >= 60:
                rating = "Meets Expectations (3)"
            elif overall_score >= 40:
                rating = "Below Expectations (2)"
            else:
                rating = "Serious Performance Concerns (1)"
            self.overall_rating_label.setText(rating)
        except Exception as e:
            self.show_error_message(f"Error calculating scores: {str(e)}")

    def get_rating_value(self, rating_text, is_soft_skill=False):
        if is_soft_skill:
            rating_map = {
                "Does not Demonstrate (1)": 1,
                "Developing (2)": 2,
                "Proficient (3)": 3,
                "Proficient (4)": 4,
                "Expert (5)": 5
            }
        else:
            rating_map = {
                "Serious Performance Concerns (1)": 1,
                "Below Expectations (2)": 2,
                "Meets Expectations (3)": 3,
                "Exceeds Expectations (4)": 4,
                "Outstanding (5)": 5
            }
        return rating_map.get(rating_text, 0)

    def get_form_data(self):
        try:
            headers = self.parent_app.excel_handler.get_visible_headers()
            emp_id = self.employee_combo.currentText() if "Employee ID" in headers else ""
            if not emp_id:
                self.show_error_message("Please select an Employee ID")
                return None
            
            data = {}
            for header in headers:
                widget = self.input_widgets.get(header)
                if widget:
                    if isinstance(widget, QLineEdit) or isinstance(widget, QLabel):
                        data[header] = widget.text()
                    elif isinstance(widget, QComboBox):
                        data[header] = widget.currentText()
                    elif isinstance(widget, QSpinBox):
                        data[header] = str(widget.value())
            
            return data
        except Exception as e:
            self.show_error_message(f"Error collecting form data: {str(e)}")
            return None

    def populate_employee_dropdown(self, employees):
        self.employee_combo.clear()
        self.employee_combo.addItem("Select Employee ID")
        for emp in employees:
            display_text = f"{emp['Employee ID']} - {emp.get('Employee Name', '')}" if "Employee Name" in emp else emp["Employee ID"]
            self.employee_combo.addItem(display_text, emp["Employee ID"])

    def on_employee_selected(self, display_text):
        if display_text and display_text != "Select Employee ID" and self.parent_app:
            try:
                emp_id = self.employee_combo.itemData(self.employee_combo.currentIndex())
                if not emp_id:
                    emp_id = display_text.split(" - ")[0]
                employee_data = self.parent_app.excel_handler.get_employee_data(emp_id)
                if employee_data:
                    headers = self.parent_app.excel_handler.get_visible_headers()
                    for header, widget in self.input_widgets.items():
                        if header in employee_data:
                            value = employee_data.get(header, "")
                            if isinstance(widget, QLineEdit) or isinstance(widget, QLabel):
                                widget.setText(value)
                            elif isinstance(widget, QComboBox):
                                widget.setCurrentText(value)
                            elif isinstance(widget, QSpinBox):
                                widget.setValue(int(value) if value.isdigit() else 15)
            except Exception as e:
                self.show_error_message(f"Error loading employee data: {str(e)}")

    def search_employee(self):
        emp_id = self.employee_combo.itemData(self.employee_combo.currentIndex())
        if not emp_id:
            emp_id = self.employee_combo.currentText().split(" - ")[0]
        if emp_id and emp_id != "Select Employee ID" and self.parent_app:
            self.parent_app.open_employee_view(emp_id)
        else:
            self.show_error_message("Please select a valid Employee ID")

    def view_curved(self):
        if self.parent_app:
            self.parent_app.open_curved_view()

    def save_form(self):
        if self.parent_app:
            self.parent_app.save_form_data()

    def clear_form(self):
        headers = self.parent_app.excel_handler.get_visible_headers()
        if "Employee ID" in headers:
            self.employee_combo.setCurrentIndex(0)
        for header, widget in self.input_widgets.items():
            if isinstance(widget, QLineEdit):
                widget.clear()
                if header == "Date of Joining":
                    widget.setText("01/07/2025")
                elif header == "Date of Evaluation":
                    widget.setText(datetime.now().strftime("%Y-%m-%d"))
                elif header == "Entity Name":
                    widget.setText("xyz")
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            elif isinstance(widget, QSpinBox):
                widget.setValue(15)
            elif isinstance(widget, QLabel) and header.endswith("Score"):
                widget.setText("0.0")
        self.contract_expiry_hidden.clear()
        self.division_hidden.clear()
        self.exp_pmtf_hidden.clear()

    def show_success_message(self, message):
        QMessageBox.information(self, "Success", message)

    def show_error_message(self, message):
        QMessageBox.critical(self, "Error", message)

    def show_add_employee_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Employee")
        dialog.setMinimumWidth(400)
        layout = QFormLayout(dialog)
        
        emp_id = QLineEdit()
        name = QLineEdit()
        department = QLineEdit()
        designation = QLineEdit()
        joining_date = QDateEdit()
        joining_date.setCalendarPopup(True)
        joining_date.setDate(QDate.currentDate())
        contract_expiry = QDateEdit()
        contract_expiry.setCalendarPopup(True)
        contract_expiry.setDate(QDate.currentDate())
        division = QLineEdit()
        exp_pmtf = QLineEdit()
        line_manager = QLineEdit()
        
        layout.addRow("Employee ID:", emp_id)
        layout.addRow("Name:", name)
        layout.addRow("Department:", department)
        layout.addRow("Designation:", designation)
        layout.addRow("Date of Joining:", joining_date)
        layout.addRow("Contract Expiry Date:", contract_expiry)
        layout.addRow("Division:", division)
        layout.addRow("Experience xyz:", exp_pmtf)
        layout.addRow("Line Manager:", line_manager)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)
        
        if dialog.exec_() == QDialog.Accepted:
            success, message = self.parent_app.add_new_employee(
                emp_id.text(),
                name.text(),
                department.text(),
                designation.text(),
                joining_date.date().toString("dd/MM/yyyy"),
                contract_expiry.date().toString("dd/MM/yyyy"),
                division.text(),
                exp_pmtf.text()
            )
            if success:
                self.show_success_message("Employee added successfully!")
                self.parent_app.load_employees()
            else:
                self.show_error_message(message)