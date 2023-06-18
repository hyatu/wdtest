# Add the required assembly
Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

# Create a new Word application object
$wordApp = New-Object -ComObject Word.Application

# Create a new document
$document = $wordApp.Documents.Add()

# Set the title
$titleText = "System documentation for $($env:COMPUTERNAME)"
$titleRange = $document.Range()
$titleRange.Text = $titleText
$titleRange.Font.Size = 18
$titleRange.Font.Bold = $true
$titleRange.ParagraphFormat.Alignment = 1 # Center alignment
$titleRange.InsertParagraphAfter()

# Insert a table of contents
$tocRange = $document.Range()
$wordApp.Selection = $tocRange
$wordApp.Selection.ParagraphFormat.Alignment = 0 # Left alignment
$wordApp.Selection.TypeText("Table of Contents")
$wordApp.Selection.TypeParagraph()
$wordApp.Selection.TypeText("{TOC \o '1-3' \h \z \u}")
$wordApp.Selection.TypeParagraph()

# Insert page numbers
$wordApp.ActiveWindow.View.Type = 3 # Print Layout view
$wordApp.Selection.GoTo(1, 2, 1) # Go to the header section
$wordApp.ActiveWindow.ActivePane.Selection.HeaderFooter.PageNumbers.Add()

# Insert server information section headers
$sectionHeaders = @(
    "Server Information",
    "Operating System",
    "Network Information",
    "Running Services",
    "Installed Software",
    "Installed Windows Roles"
)

foreach ($header in $sectionHeaders) {
    $headerRange = $document.Range()
    $headerRange.Text = $header
    $headerRange.Font.Size = 14
    $headerRange.Font.Bold = $true
    $headerRange.InsertParagraphAfter()
    $headerRange.Collapse(0)
}

# Get server information (use your own script to retrieve the information)
$serverInfo = Get-ServerInformation

# Populate the document with server information
foreach ($property in $serverInfo.PSObject.Properties) {
    $propertyName = $property.Name
    $propertyValue = $property.Value

    $infoRange = $document.Range()
    $infoRange.Text = "$propertyName : $propertyValue"
    $infoRange.InsertParagraphAfter()
    $infoRange.Collapse(0)
}

# Insert table for running services
$servicesTable = $document.Tables.Add($infoRange, $services.Count + 1, 3)
$servicesTable.Borders.Enable = $true
$servicesTable.Cell(1, 1).Range.Text = "Service Name"
$servicesTable.Cell(1, 2).Range.Text = "Display Name"
$servicesTable.Cell(1, 3).Range.Text = "Status"

$row = 2
foreach ($service in $services) {
    $servicesTable.Cell($row, 1).Range.Text = $service.Name
    $servicesTable.Cell($row, 2).Range.Text = $service.DisplayName
    $servicesTable.Cell($row, 3).Range.Text = $service.Status
    $row++
}

$infoRange.Collapse(0)

# Insert table for installed software
$softwareTable = $document.Tables.Add($infoRange, $software.Count + 1, 3)
$softwareTable.Borders.Enable = $true
$softwareTable.Cell(1, 1).Range.Text = "Software Name"
$softwareTable.Cell(1, 2).Range.Text = "Vendor"
$softwareTable.Cell(1, 3).Range.Text = "Version"

$row = 2
foreach ($sw in $software) {
    $softwareTable.Cell($row, 1).Range.Text = $sw.Name
    $softwareTable.Cell($row, 2).Range.Text = $sw.Vendor
    $softwareTable.Cell($row, 3).Range.Text = $sw.Version
    $row++
}

$infoRange.Collapse(0)

# Save the document
$docPath = "C:\Temp\server_document.docx"
$document.SaveAs([ref]$docPath)

# Close the document and Word application
$document.Close()
$wordApp.Quit()

Write-Host "Server document created and saved at: $docPath"
