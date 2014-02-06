# Code for creating an excel file and populating it with information about the drives on each 
# computer name specified in a given file.

# This file demonstrates creating an excel document and filling it with information dynamically from several computers on the network. 
# Very useful script that can be adapted for inventory of specific components in a work network.
# AUTHOR: Keith S.


# Create an excel document
$Excel = New-Object -Com Excel.Application
$Excel.visible = $True
$Excel = $Excel.Workbooks.Add()

# Add headings to the page
$Sheet = $Excel.WorkSheets.Item(1)
$Sheet.Cells.Item(1,1) = "Computer"
$Sheet.Cells.Item(1,2) = "Drive Letter"
$Sheet.Cells.Item(1,3) = "Description"
$Sheet.Cells.Item(1,4) = "FileSystem"
$Sheet.Cells.Item(1,5) = "Size in GB"
$Sheet.Cells.Item(1,6) = "Free Space in GB"

# Add some color to the headings to makes them stick out
$WorkBook = $Sheet.UsedRange
$WorkBook.Interior.ColorIndex = 56
$WorkBook.Font.ColorIndex = 2
$WorkBook.Font.Bold = $True

# Name the current sheet
$Sheet.Name = "Drive Info"

# Specify row to begin adding information (so we don't overwrite the headings
$intRow = 2

# Get a list of workstations that will be queried for drive space
$servers = Get-Content -Path "Computers.txt"

# Iterate through each workstation and query the requested information from each drive on that machine
# Output information to excel spreadsheet as it is gathered
foreach ($server in $servers) {
	$colItems = Get-wmiObject -class "Win32_LogicalDisk" -namespace "root\CIMV2" -computername $server
	foreach ($objItem in $colItems) {
		$Sheet.Cells.Item($intRow,1) = $objItem.SystemName
		$Sheet.Cells.Item($intRow,2) = $objItem.DeviceID
		$Sheet.Cells.Item($intRow,3) = $objItem.Description
		$Sheet.Cells.Item($intRow,4) = $objItem.FileSystem
		$Sheet.Cells.Item($intRow,5) = $objItem.Size / 1GB
		$Sheet.Cells.Item($intRow,6) = $objItem.FreeSpace / 1GB

		$intRow = $intRow + 1
	}
}

# Autofit the entire document so information is not hidden under columns
$WorkBook.EntireColumn.AutoFit()

# Save this file to the location below (note, this will save to a default save location [such as Documents], not the current directory)
$Excel.SaveAs("DriveInfo.xlsx")

# Clear any errors that popup
Clear