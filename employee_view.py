from PyQt5.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLineEdit, QTextEdit, QGroupBox, QScrollArea, QFrame,
    QTabWidget, QSpinBox, QDoubleSpinBox, QPushButton, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from excel_handler import ExcelHandler

class EmployeeViewWindow(QWidget):
    def __init__(self, employee_data, parent_app=None):
        super().__init__()
        self.employee_data = employee_data
        self.parent_app = parent_app
        self.setMinimumSize(1400, 900)
        self.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                font-family: 'Segoe UI', sans-serif;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px 0 8px;
                color: #495057;
            }
            QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 8px;
                background-color: #e9ecef;
                color: #495057;
            }
            .readonly {
                background-color: #f8f9fa;
                color: #6c757d;
            }
        """)
        
        self.create_ui()
        self.populate_data()
        self.populate_salary_data()

    def create_ui(self):
        main_layout = QVBoxLayout(self)
        
        # Add back button
        back_button = QPushButton("Back to Form")
        back_button.setStyleSheet("""
            QPushButton {
                background-color: #212580;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #3439C7;
            }
        """)
        if self.parent_app:
            back_button.clicked.connect(self.parent_app.go_back)
        main_layout.addWidget(back_button)
        
        header_label = QLabel(f"Employee Performance Review Details\n{self.employee_data.get('Employee Name', 'N/A')} - {self.employee_data.get('Employee ID', 'N/A')}")
        header_label.setAlignment(Qt.AlignCenter)
        header_label.setFont(QFont("Arial", 16, QFont.Bold))
        header_label.setStyleSheet("background-color: #007bff; color: white; padding: 20px; border-radius: 8px; margin-bottom: 10px;")
        main_layout.addWidget(header_label)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        tab_widget = QTabWidget()
        performance_tab = QWidget()
        self.create_performance_tab(performance_tab)
        tab_widget.addTab(performance_tab, "Performance Review")
        salary_tab = QWidget()
        self.create_salary_tab(salary_tab)
        tab_widget.addTab(salary_tab, "Salary & Adjustments")
        scroll_layout.addWidget(tab_widget)
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)

    def create_performance_tab(self, parent):
        layout = QVBoxLayout(parent)
        basic_info_group = QGroupBox("Employee Information")
        basic_info_layout = QGridLayout(basic_info_group)
        self.emp_id_field = QLineEdit()
        self.emp_id_field.setReadOnly(True)
        self.emp_name_field = QLineEdit()
        self.emp_name_field.setReadOnly(True)
        self.emp_dept_field = QLineEdit()
        self.emp_dept_field.setReadOnly(True)
        self.emp_designation_field = QLineEdit()
        self.emp_designation_field.setReadOnly(True)
        self.joining_date_field = QLineEdit()
        self.joining_date_field.setReadOnly(True)
        self.contract_expiry_field = QLineEdit()
        self.contract_expiry_field.setReadOnly(True)
        self.division_field = QLineEdit()
        self.division_field.setReadOnly(True)
        basic_info_layout.addWidget(QLabel("Employee ID:"), 0, 0)
        basic_info_layout.addWidget(self.emp_id_field, 0, 1)
        basic_info_layout.addWidget(QLabel("Name:"), 0, 2)
        basic_info_layout.addWidget(self.emp_name_field, 0, 3)
        basic_info_layout.addWidget(QLabel("Department:"), 1, 0)
        basic_info_layout.addWidget(self.emp_dept_field, 1, 1)
        basic_info_layout.addWidget(QLabel("Designation:"), 1, 2)
        basic_info_layout.addWidget(self.emp_designation_field, 1, 3)
        basic_info_layout.addWidget(QLabel("Date of Joining:"), 2, 0)
        basic_info_layout.addWidget(self.joining_date_field, 2, 1)
        basic_info_layout.addWidget(QLabel("Contract Expiry Date:"), 2, 2)
        basic_info_layout.addWidget(self.contract_expiry_field, 2, 3)
        basic_info_layout.addWidget(QLabel("Division:"), 3, 0)
        basic_info_layout.addWidget(self.division_field, 3, 1)
        layout.addWidget(basic_info_group)
        kpi_group = QGroupBox("Performance KPIs (Read-Only)")
        kpi_layout = QGridLayout(kpi_group)
        headers = ["KPI", "Title", "Rating", "Weightage", "Score"]
        for i, header in enumerate(headers):
            label = QLabel(header)
            label.setFont(QFont("Arial", 10, QFont.Bold))
            label.setStyleSheet("background-color: #e9ecef; padding: 8px; border: 1px solid #dee2e6;")
            kpi_layout.addWidget(label, 0, i)

        self.kpi_fields = {}
        for i in range(1, 7):
            kpi_label = QLabel(f"KPI {i}")
            kpi_title = QLineEdit()
            kpi_title.setReadOnly(True)
            kpi_rating = QLineEdit()
            kpi_rating.setReadOnly(True)
            kpi_weightage = QLineEdit()
            kpi_weightage.setReadOnly(True)
            kpi_score = QLineEdit()
            kpi_score.setReadOnly(True)
            
            kpi_layout.addWidget(kpi_label, i, 0)
            kpi_layout.addWidget(kpi_title, i, 1)
            kpi_layout.addWidget(kpi_rating, i, 2)
            kpi_layout.addWidget(kpi_weightage, i, 3)
            kpi_layout.addWidget(kpi_score, i, 4)
            
            self.kpi_fields[f'kpi_{i}'] = {
                'title': kpi_title,
                'rating': kpi_rating,
                'weightage': kpi_weightage,
                'score': kpi_score
            }
        layout.addWidget(kpi_group)
        soft_skills_group = QGroupBox("Soft Skills (Read-Only)")
        soft_skills_layout = QGridLayout(soft_skills_group)
        soft_skills_headers = ["Skill", "Rating", "Weighted Score"]
        for i, header in enumerate(soft_skills_headers):
            label = QLabel(header)
            label.setFont(QFont("Arial", 10, QFont.Bold))
            label.setStyleSheet("background-color: #e9ecef; padding: 8px; border: 1px solid #dee2e6;")
            soft_skills_layout.addWidget(label, 0, i)
        self.soft_skill_fields = {}
        soft_skills = [
            "Open & Clear Communication",
            "Attitude, Team Work & Collaboration",
            "Planning & Achievement Focus",
            "Creativity & Initiatives",
            "Ownership & Self Accountability"
        ]
        for i, skill in enumerate(soft_skills, 1):
            skill_label = QLabel(skill)
            rating_field = QLineEdit()
            rating_field.setReadOnly(True)
            score_field = QLineEdit()
            score_field.setReadOnly(True)
            soft_skills_layout.addWidget(skill_label, i, 0)
            soft_skills_layout.addWidget(rating_field, i, 1)
            soft_skills_layout.addWidget(score_field, i, 2)
            
            clean_skill_name = skill.replace(" ", "_").replace("&", "").replace(",", "")
            self.soft_skill_fields[clean_skill_name] = {
                'rating': rating_field,
                'score': score_field
            }
        layout.addWidget(soft_skills_group)
        summary_group = QGroupBox("Performance Summary")
        summary_layout = QGridLayout(summary_group)
        self.overall_rating_field = QLineEdit()
        self.overall_rating_field.setReadOnly(True)
        self.overall_percentage_field = QLineEdit()
        self.overall_percentage_field.setReadOnly(True)
        self.promotion_field = QLineEdit()
        self.promotion_field.setReadOnly(True)
        self.retention_field = QLineEdit()
        self.retention_field.setReadOnly(True)
        summary_layout.addWidget(QLabel("Overall Rating:"), 0, 0)
        summary_layout.addWidget(self.overall_rating_field, 0, 1)
        summary_layout.addWidget(QLabel("Overall Percentage:"), 0, 2)
        summary_layout.addWidget(self.overall_percentage_field, 0, 3)
        summary_layout.addWidget(QLabel("Promotion Rec:"), 1, 0)
        summary_layout.addWidget(self.promotion_field, 1, 1)
        summary_layout.addWidget(QLabel("Retention Rec:"), 1, 2)
        summary_layout.addWidget(self.retention_field, 1, 3)
        layout.addWidget(summary_group)

    def save_employee_data(self):
        """Save employee data to Excel - safer version"""
        try:
            excel_handler = ExcelHandler()
            emp_id = self.emp_id_field.text()
            existing_data = excel_handler.get_employee_data(emp_id)
            if not existing_data:
                raise ValueError("Employee not found!")
            existing_data.update({
                "Employee ID": emp_id,
                "Employee Name": self.emp_name_field.text(),
                "Department": self.emp_dept_field.text(),
                "Designation": self.emp_designation_field.text(),
                "Date of Joining": self.joining_date_field.text(),
                "Contract Expiry Date": self.contract_expiry_field.text(),
                "Division": self.division_field.text(),
                "Last_Year_Rating": self.last_year_rating.text(),
                "Last_Year_Increment": self.last_year_increment.text(),
                "Basic_Salary": self.basic_salary.value(),
                "Gross_Amount": self.gross_amount.value(),
                "Car_Allowance": self.car_allowance.value(),
                "Fuel_Litre": self.fuel_litre.value(),
                "Fuel_Price": self.fuel_price.value(),
                "House_Rent": self.house_rent.value(),
                "Medical": self.medical.value(),
                "Utilities": self.utilities.value(),
                "Total_Salary": (
                    self.utilities.value() + self.medical.value() + self.basic_salary.value() +
                    self.gross_amount.value() + self.car_allowance.value() +
                    (self.fuel_litre.value() * self.fuel_price.value()) + self.house_rent.value()
                ),
                "Diff_Salary": self.diff_salary.value(),
                "Diff_Conveyance": self.diff_conveyance.value(),
                "Fuel_Litre_Adj": self.fuel_litre_adj.value(),
                "Fuel_Price_Adj": self.fuel_price_adj.value(),
                "Diff_Car_Allowance": self.diff_car_allowance.value(),
                "Amount_Diff_Fuel": self.fuel_litre_adj.value() * self.fuel_price_adj.value(),
                "Total_Salary_Adj": (
                    self.diff_salary.value() + self.diff_conveyance.value() + self.diff_car_allowance.value()
                ),
                "Allowance_Adj": (
                    self.diff_car_allowance.value() + self.diff_conveyance.value()
                ),
                "Salary_Adj_Impact": self.salary_adj_impact.text(),
                "Salary_Increment_2425": self.salary_increment_2425.value(),
                "Training_Recommendations": self.training_recommendations.toPlainText(),
                "KPI_1_Title": existing_data.get("KPI_1_Title", ""),
                "KPI_1_Rating": existing_data.get("KPI_1_Rating", ""),
                "KPI_1_Weightage": existing_data.get("KPI_1_Weightage", ""),
                "KPI_1_Weighted_Score": existing_data.get("KPI_1_Weighted_Score", ""),
                "KPI_2_Title": existing_data.get("KPI_2_Title", ""),
                "KPI_2_Rating": existing_data.get("KPI_2_Rating", ""),
                "KPI_2_Weightage": existing_data.get("KPI_2_Weightage", ""),
                "KPI_2_Weighted_Score": existing_data.get("KPI_2_Weighted_Score", ""),
                "KPI_3_Title": existing_data.get("KPI_3_Title", ""),
                "KPI_3_Rating": existing_data.get("KPI_3_Rating", ""),
                "KPI_3_Weightage": existing_data.get("KPI_3_Weightage", ""),
                "KPI_3_Weighted_Score": existing_data.get("KPI_3_Weighted_Score", ""),
                "KPI_4_Title": existing_data.get("KPI_4_Title", ""),
                "KPI_4_Rating": existing_data.get("KPI_4_Rating", ""),
                "KPI_4_Weightage": existing_data.get("KPI_4_Weightage", ""),
                "KPI_4_Weighted_Score": existing_data.get("KPI_4_Weighted_Score", ""),
                "KPI_5_Title": existing_data.get("KPI_5_Title", ""),
                "KPI_5_Rating": existing_data.get("KPI_5_Rating", ""),
                "KPI_5_Weightage": existing_data.get("KPI_5_Weightage", ""),
                "KPI_5_Weighted_Score": existing_data.get("KPI_5_Weighted_Score", ""),
                "KPI_6_Title": existing_data.get("KPI_6_Title", ""),
                "KPI_6_Rating": existing_data.get("KPI_6_Rating", ""),
                "KPI_6_Weightage": existing_data.get("KPI_6_Weightage", ""),
                "KPI_6_Weighted_Score": existing_data.get("KPI_6_Weighted_Score", ""),
                "Part_A_Total_Score": existing_data.get("Part_A_Total_Score", ""),
                "Open_Clear_Communication_Rating": existing_data.get("Open_Clear_Communication_Rating", ""),
                "Open_Clear_Communication_Weighted_Score": existing_data.get("Open_Clear_Communication_Weighted_Score", ""),
                "Attitude_Team_Work_Collaboration_Rating": existing_data.get("Attitude_Team_Work_Collaboration_Rating", ""),
                "Attitude_Team_Work_Collaboration_Weighted_Score": existing_data.get("Attitude_Team_Work_Collaboration_Weighted_Score", ""),
                "Planning_Achievement_Focus_Rating": existing_data.get("Planning_Achievement_Focus_Rating", ""),
                "Planning_Achievement_Focus_Weighted_Score": existing_data.get("Planning_Achievement_Focus_Weighted_Score", ""),
                "Creativity_Initiatives_Rating": existing_data.get("Creativity_Initiatives_Rating", ""),
                "Creativity_Initiatives_Weighted_Score": existing_data.get("Creativity_Initiatives_Weighted_Score", ""),
                "Ownership_Self_Accountability_Rating": existing_data.get("Ownership_Self_Accountability_Rating", ""),
                "Ownership_Self_Accountability_Weighted_Score": existing_data.get("Ownership_Self_Accountability_Weighted_Score", ""),
                "Part_B_Total_Score": existing_data.get("Part_B_Total_Score", ""),
                "Overall_Rating": self.overall_rating_field.text(),
                "Overall_Percentage": self.overall_percentage_field.text(),
                "Promotion_Recommendation": self.promotion_field.text(),
                "Retention_Recommendation": self.retention_field.text(),
            })
            excel_handler.save_employee_data(existing_data)
            QMessageBox.information(self, "Success", "Employee data saved successfully!")
            if self.parent_app:
                self.parent_app.go_back()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save data: {str(e)}")

    def create_salary_tab(self, parent):
        layout = QVBoxLayout(parent)
        last_year_group = QGroupBox("Last Year Ratings")
        last_year_layout = QGridLayout(last_year_group)
        last_year_layout.addWidget(QLabel("Rating:"), 0, 0)
        self.last_year_rating = QLineEdit()
        last_year_layout.addWidget(self.last_year_rating, 0, 1)
        last_year_layout.addWidget(QLabel("Salary Increment:"), 0, 2)
        self.last_year_increment = QLineEdit()
        last_year_layout.addWidget(self.last_year_increment, 0, 3)
        layout.addWidget(last_year_group)
        current_salary_group = QGroupBox("Salary as of Previous Month")
        current_salary_layout = QGridLayout(current_salary_group)
        self.basic_salary = QDoubleSpinBox()
        self.basic_salary.setRange(0, 999999)
        self.basic_salary.valueChanged.connect(self.calculate_totals)
        self.gross_amount = QDoubleSpinBox()
        self.gross_amount.setRange(0, 999999)
        self.gross_amount.valueChanged.connect(self.calculate_totals)
        self.car_allowance = QDoubleSpinBox()
        self.car_allowance.setRange(0, 999999)
        self.car_allowance.valueChanged.connect(self.calculate_totals)
        self.fuel_litre = QDoubleSpinBox()
        self.fuel_litre.setRange(0, 999)
        self.fuel_litre.valueChanged.connect(self.calculate_totals)
        self.fuel_price = QDoubleSpinBox()
        self.fuel_price.setRange(0, 999)
        self.fuel_price.setValue(267)  # Default price
        self.fuel_price.valueChanged.connect(self.calculate_totals)
        self.house_rent = QDoubleSpinBox()
        self.house_rent.setRange(0, 999999)
        self.house_rent.valueChanged.connect(self.calculate_totals)
        self.medical = QDoubleSpinBox()
        self.medical.setRange(0, 999999)
        self.medical.valueChanged.connect(self.calculate_totals)
        self.utilities = QDoubleSpinBox()
        self.utilities.setRange(0, 999999)
        self.utilities.valueChanged.connect(self.calculate_totals)
        self.total_salary = QLineEdit()
        self.total_salary.setReadOnly(True)
        self.total_salary.setStyleSheet("background-color: #fff3cd; font-weight: bold;")
        current_salary_layout.addWidget(QLabel("Basic Salary:"), 0, 0)
        current_salary_layout.addWidget(self.basic_salary, 0, 1)
        current_salary_layout.addWidget(QLabel("Gross Amount:"), 0, 2)
        current_salary_layout.addWidget(self.gross_amount, 0, 3)
        current_salary_layout.addWidget(QLabel("Car Allowance:"), 1, 0)
        current_salary_layout.addWidget(self.car_allowance, 1, 1)
        current_salary_layout.addWidget(QLabel("Fuel (Litre):"), 1, 2)
        current_salary_layout.addWidget(self.fuel_litre, 1, 3)
        current_salary_layout.addWidget(QLabel("Fuel Price/Litre:"), 2, 0)
        current_salary_layout.addWidget(self.fuel_price, 2, 1)
        current_salary_layout.addWidget(QLabel("House Rent:"), 2, 2)
        current_salary_layout.addWidget(self.house_rent, 2, 3)
        current_salary_layout.addWidget(QLabel("Medical:"), 3, 0)
        current_salary_layout.addWidget(self.medical, 3, 1)
        current_salary_layout.addWidget(QLabel("Utilities:"), 3, 2)
        current_salary_layout.addWidget(self.utilities, 3, 3)
        current_salary_layout.addWidget(QLabel("Total Salary:"), 4, 0)
        current_salary_layout.addWidget(self.total_salary, 4, 1, 1, 3)
        
        layout.addWidget(current_salary_group)
        salary_adjustment_group = QGroupBox("Salary Adjustment")
        salary_adjustment_layout = QGridLayout(salary_adjustment_group)
        
        self.diff_salary = QDoubleSpinBox()
        self.diff_salary.setRange(-999999, 999999)
        self.diff_salary.valueChanged.connect(self.calculate_impact)
        
        self.diff_conveyance = QDoubleSpinBox()
        self.diff_conveyance.setRange(-999999, 999999)
        self.diff_conveyance.valueChanged.connect(self.calculate_impact)
        
        self.fuel_litre_adj = QDoubleSpinBox()
        self.fuel_litre_adj.setRange(0, 999)
        self.fuel_litre_adj.valueChanged.connect(self.calculate_impact)
        
        self.fuel_price_adj = QDoubleSpinBox()
        self.fuel_price_adj.setRange(0, 999)
        self.fuel_price_adj.setValue(267)
        self.fuel_price_adj.valueChanged.connect(self.calculate_impact)
        
        self.diff_car_allowance = QDoubleSpinBox()
        self.diff_car_allowance.setRange(-999999, 999999)
        self.diff_car_allowance.valueChanged.connect(self.calculate_impact)
        
        self.amount_diff_fuel = QLineEdit()
        self.amount_diff_fuel.setReadOnly(True)
        self.amount_diff_fuel.setStyleSheet("background-color: #e9ecef; font-weight: bold;")
        
        salary_adjustment_layout.addWidget(QLabel("Difference in Salary:"), 0, 0)
        salary_adjustment_layout.addWidget(self.diff_salary, 0, 1)
        salary_adjustment_layout.addWidget(QLabel("Difference Conveyance:"), 0, 2)
        salary_adjustment_layout.addWidget(self.diff_conveyance, 0, 3)
        
        salary_adjustment_layout.addWidget(QLabel("Fuel (Ltr):"), 1, 0)
        salary_adjustment_layout.addWidget(self.fuel_litre_adj, 1, 1)
        salary_adjustment_layout.addWidget(QLabel("Fuel Price Rs @ 267:"), 1, 2)
        salary_adjustment_layout.addWidget(self.fuel_price_adj, 1, 3)
        
        salary_adjustment_layout.addWidget(QLabel("Difference Car Allowance:"), 2, 0)
        salary_adjustment_layout.addWidget(self.diff_car_allowance, 2, 1)
        salary_adjustment_layout.addWidget(QLabel("Amount Diff in Fuel:"), 2, 2)
        salary_adjustment_layout.addWidget(self.amount_diff_fuel, 2, 3)
        
        layout.addWidget(salary_adjustment_group)
        
        impact_group = QGroupBox("IMPACT")
        impact_layout = QGridLayout(impact_group)
        
        self.total_salary_adj = QLineEdit()
        self.total_salary_adj.setReadOnly(True)
        self.total_salary_adj.setStyleSheet("background-color: #d1ecf1; font-weight: bold;")
        
        self.allowance_adj = QLineEdit()
        self.allowance_adj.setReadOnly(True)
        self.allowance_adj.setStyleSheet("background-color: #d1ecf1; font-weight: bold;")
        
        self.salary_adj_impact = QLineEdit()
        self.salary_adj_impact.setReadOnly(True)
        self.salary_adj_impact.setStyleSheet("background-color: #d1ecf1; font-weight: bold;")
        
        impact_layout.addWidget(QLabel("Total Salary Adjustment + Allowance:"), 0, 0)
        impact_layout.addWidget(self.total_salary_adj, 0, 1)
        impact_layout.addWidget(QLabel("Allowance Adjustment:"), 0, 2)
        impact_layout.addWidget(self.allowance_adj, 0, 3)
        
        impact_layout.addWidget(QLabel("Salary Adjustment:"), 1, 0)
        impact_layout.addWidget(self.salary_adj_impact, 1, 1)
        
        layout.addWidget(impact_group)
        increment_group = QGroupBox("Salary Increment 24/25 %")
        increment_layout = QGridLayout(increment_group)
        
        self.salary_increment_2425 = QDoubleSpinBox()
        self.salary_increment_2425.setRange(0, 100)
        self.salary_increment_2425.setSuffix("%")
        
        increment_layout.addWidget(QLabel("Increment Percentage:"), 0, 0)
        increment_layout.addWidget(self.salary_increment_2425, 0, 1)
        
        layout.addWidget(increment_group)
        training_group = QGroupBox("Training Recommendations")
        training_layout = QVBoxLayout(training_group)
        
        self.training_recommendations = QTextEdit()
        self.training_recommendations.setMaximumHeight(100)
        self.training_recommendations.setPlaceholderText("Enter training recommendations here...")
        
        training_layout.addWidget(self.training_recommendations)
        
        layout.addWidget(training_group)
        
        save_button = QPushButton("Save Changes")
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                padding: 12px 30px;
                border: none;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        save_button.clicked.connect(self.save_employee_data)
        layout.addWidget(save_button)

    def calculate_impact(self):
        """Calculate impact values"""
        try:
            fuel_diff = self.fuel_litre_adj.value() * self.fuel_price_adj.value()
            self.amount_diff_fuel.setText(f"Rs. {fuel_diff:,.2f}")
            total_adj = (self.diff_salary.value() + self.diff_conveyance.value() + 
                        self.diff_car_allowance.value())
            self.total_salary_adj.setText(f"Rs. {total_adj:,.2f}")
            allowance_adj = self.diff_conveyance.value() + self.diff_car_allowance.value()
            self.allowance_adj.setText(f"Rs. {allowance_adj:,.2f}")
            self.salary_adj_impact.setText(f"Rs. {self.diff_salary.value():,.2f}")
            
        except:
            self.amount_diff_fuel.setText("Rs. 0.00")
            self.total_salary_adj.setText("Rs. 0.00")
            self.allowance_adj.setText("Rs. 0.00")
            self.salary_adj_impact.setText("Rs. 0.00")

    def calculate_totals(self):
        """Calculate total salary and adjustments"""
        try:
            fuel_cost = self.fuel_litre.value() * self.fuel_price.value()
            total = (self.basic_salary.value() + self.gross_amount.value() + 
                    self.car_allowance.value() + fuel_cost + self.house_rent.value() + 
                    self.medical.value() + self.utilities.value())
            self.total_salary.setText(f"Rs. {total:,.2f}")
        except:
            self.total_salary.setText("Rs. 0.00")

    def populate_data(self):
        """Populate fields with employee data"""
        self.emp_id_field.setText(str(self.employee_data.get('Employee ID', '')))
        self.emp_name_field.setText(str(self.employee_data.get('Employee Name', '')))
        self.emp_dept_field.setText(str(self.employee_data.get('Department', '')))
        self.emp_designation_field.setText(str(self.employee_data.get('Designation', '')))
        self.joining_date_field.setText(str(self.employee_data.get('Date of Joining', '')))
        self.contract_expiry_field.setText(str(self.employee_data.get('Contract Expiry Date', '')))
        self.division_field.setText(str(self.employee_data.get('Division', '')))
        for i in range(1, 7):
            if f'KPI_{i}_Title' in self.employee_data:
                self.kpi_fields[f'kpi_{i}']['title'].setText(str(self.employee_data.get(f'KPI_{i}_Title', '')))
                self.kpi_fields[f'kpi_{i}']['rating'].setText(str(self.employee_data.get(f'KPI_{i}_Rating', '')))
                self.kpi_fields[f'kpi_{i}']['weightage'].setText(str(self.employee_data.get(f'KPI_{i}_Weightage', '')))
                self.kpi_fields[f'kpi_{i}']['score'].setText(str(self.employee_data.get(f'KPI_{i}_Weighted_Score', '')))
        for skill, fields in self.soft_skill_fields.items():
            fields['rating'].setText(str(self.employee_data.get(f'{skill}_Rating', '')))
            fields['score'].setText(str(self.employee_data.get(f'{skill}_Weighted_Score', '')))
        self.overall_rating_field.setText(str(self.employee_data.get('Overall_Rating', '')))
        self.overall_percentage_field.setText(str(self.employee_data.get('Overall_Percentage', '')))
        self.promotion_field.setText(str(self.employee_data.get('Promotion_Recommendation', '')))
        self.retention_field.setText(str(self.employee_data.get('Retention_Recommendation', '')))

    def populate_salary_data(self):
        """Populate salary related data if available"""
        if 'Last_Year_Rating' in self.employee_data:
            self.last_year_rating.setText(str(self.employee_data.get('Last_Year_Rating', '')))
        if 'Last_Year_Increment' in self.employee_data:
            self.last_year_increment.setText(str(self.employee_data.get('Last_Year_Increment', '')))
        
        self.basic_salary.setValue(float(self.employee_data.get('Basic_Salary', 0) or 0))
        self.gross_amount.setValue(float(self.employee_data.get('Gross_Amount', 0) or 0))
        self.car_allowance.setValue(float(self.employee_data.get('Car_Allowance', 0) or 0))
        self.fuel_litre.setValue(float(self.employee_data.get('Fuel_Litre', 0) or 0))
        self.house_rent.setValue(float(self.employee_data.get('House_Rent', 0) or 0))
        self.medical.setValue(float(self.employee_data.get('Medical', 0) or 0))
        self.utilities.setValue(float(self.employee_data.get('Utilities', 0) or 0))
        self.salary_adj_impact.setText(str(self.employee_data.get('Salary_Adj_Impact', 'Rs. 0.00')))
        self.training_recommendations.setText(str(self.employee_data.get('Training_Recommendations', '')))
        self.salary_increment_2425.setValue(float(self.employee_data.get('Salary_Increment_2425', 0) or 0))
        
        self.calculate_totals()

    def update_employee_data(self, employee_data):
        """Update the widget with new employee data"""
        self.employee_data = employee_data
        self.populate_data()
        self.populate_salary_data()
        # Update header label
        header_label = self.findChild(QLabel)
        if header_label:
            header_label.setText(f"Employee Performance Review Details\n{self.employee_data.get('Employee Name', 'N/A')} - {self.employee_data.get('Employee ID', 'N/A')}")