import openpyxl

#Create a new workbook
workbook = openpyxl.Workbook()

#Select the active sheet
worksheet = workbook.active

#Write data to specific cells
#Add headers to the sheet
#worksheet["A1"] = "Employee ID"
#worksheet["B1"] = "First Name"
#worksheet["C1"] = "Last Name"
#worksheet["D1"] = "Department"
#worksheet["E1"] = "DOJ"
#worksheet["F1"] = "Salary"

#Add employee data to each row
employee_data = [
                 ["Employee ID", "First Name", "Last Name", "Department", "DOJ", "Salary"],
                 ["1", "Mohammad", "Zabiullah", "Engineering", "2011-10-10", "137500"],
                 ["2", "Baduruddin", "Syed", "Leadership", "2011-11-01", "177500"],
                 ["3", "Shibu", "Mathew", "SRE", "2012-05-05", "137500"],
                 ["4", "Matt", "C", "DataCenter", "2020-10-10", "157500"],
                 ["5", "Nate", "Holler", "Engineering", "2023-04-10", "97500"],
                 ["6", "Kenton", "Porter", "SRE", "2023-04-10", "87500"]]


for employee in employee_data:
    worksheet.append(employee)

#Save the workbook
workbook.save("employee.xlsx")
