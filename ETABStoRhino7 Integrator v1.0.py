import comtypes.client
import xlwings as xw

# Initialize ETABS API
ETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")

# Get design forces using cDesignResults
DesignForces = ETABSObject.SapModel.DesignResults.DesignForces

# Create a new Excel workbook
wb = xw.Book()

# Define a sheet in the workbook
sheet = wb.sheets.add('Design Forces')

# Check if there are design forces
if DesignForces.Count > 0:
    force_names = DesignForces.ForceNames

    # Write headers based on available force names
    sheet.range('A1').value = 'Member Name'
    for col, name in enumerate(force_names, start=2):
        sheet.range(1, col).value = name

    # Initialize row counter
    row = 2

    # Iterate through design forces
    for i in range(1, DesignForces.Count + 1):
        member_name = DesignForces.Name(i)

        # Get forces for the member
        member_forces = DesignForces.Values(i)
        for col, force_value in enumerate(member_forces, start=2):
            sheet.range(row, col).value = force_value

        # Increment row counter
        row += 1

    # Save the Excel workbook
    wb.save('DesignForces.xlsx')

    print("Design forces exported to Excel.")
else:
    print("No design forces found in the model")

# Clean up ETABS API (No need to close the model as it's already open)
ETABSObject.ApplicationExit()
