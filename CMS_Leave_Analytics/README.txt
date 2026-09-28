# CMS Leave Analytics — Version 1

## 1. Install Python packages
Open Command Prompt in this folder and run:

pip install flask pandas lxml openpyxl

## 2. Run
python app.py

Then open:
http://127.0.0.1:5000

## 3. Use
1. Click "Import HRMS Report".
2. Select the HRMS .xls report.
3. The dashboard imports the report into a local SQLite database.
4. Use filters for month, leave type, branch, designation, status and employee.
5. "Monthly Leave Above Threshold" shows employees whose total APPROVED leave days in a month are greater than the selected threshold.
6. "Individual Leave > Threshold" shows individual leave applications longer than the threshold.
7. Use Export Report to export the current filtered report as CSV.

## Important
This version treats "more than 2 days in a month" as SUM of approved leave days for each employee in each month. A separate tab handles individual leave applications above 2 days.

The imported report replaces the previous report in the local database, so the dashboard always represents the latest uploaded HRMS report.


Updated: threshold is now interpreted as AT LEAST the selected number of approved days (>=). The dashboard labels were aligned with this behavior.
