import openpyxl

# create a new excel file with coulms: Arrorney, Link, Name, Membership level, First name, Last name, City, State/Province, Firm website, Firm name, Speciality, Membership
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Attorneys"

sheet["A1"] = "Attorney"
sheet["B1"] = "Link"
sheet["C1"] = "Name"
sheet["D1"] = "Membership level"
sheet["E1"] = "First name"
sheet["F1"] = "Last name"
sheet["G1"] = "City"
sheet["H1"] = "State/Province"
sheet["I1"] = "Firm website"
sheet["J1"] = "Firm name"
sheet["K1"] = "Speciality"
sheet["L1"] = "Membership"

wb.save("attorneys.xlsx")

