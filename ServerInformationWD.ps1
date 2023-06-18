# Server Information
$serverName =  $env:COMPUTERNAME  # Replace with the actual server name "SERVER_NAME"

# Create a Word Application object
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $true

# Create a new Word document
$document = $wordApp.Documents.Add()

# Set up document formatting
$document.Styles["Normal"].Font.Name = "Arial"
$document.Styles["Normal"].Font.Size = 11

# Add a title
$title = $document.Styles["Title"]
$title.Font.Size = 18
$title.Font.Bold = $true

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "Server Information"
$paragraph.Range.set_Style("Title")
$paragraph.Range.InsertParagraphAfter()

# Add a table of contents
$tocRange = $document.Range()
$tocRange.Collapse(0)
$tocRange.Text = "Table of Contents"
$tocRange.set_Style("Heading 1")
$tocRange.InsertParagraphAfter()

$document.TablesOfContents.Add($tocRange, 1, $true)

# Set up section headers
$sectionHeaders = $document.Styles["Heading 2"]
$sectionHeaders.Font.Bold = $true

# Server Name and Description
$server = Get-WmiObject Win32_ComputerSystem -ComputerName $serverName
$serverName = $server.Name
$serverDescription = $server.Description

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "1. Server Name and Description"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "Server Name: $serverName"
$paragraph.Range.InsertParagraphAfter()

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "Description: $serverDescription"
$paragraph.Range.InsertParagraphAfter()

# IP Address
$ipAddresses = (Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $serverName | Where-Object { $_.IPAddress -ne $null }).IPAddress

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "2. IP Addresses"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "IP Addresses: $ipAddresses"
$paragraph.Range.InsertParagraphAfter()

# CPU information
$cpu = Get-WmiObject Win32_Processor -ComputerName $serverName
$cpuName = $cpu.Name
$cpuCores = $cpu.NumberOfCores

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "3. CPU Information"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "CPU: $cpuName"
$paragraph.Range.InsertParagraphAfter()

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "CPU Cores: $cpuCores"
$paragraph.Range.InsertParagraphAfter()

# RAM information
$ram = Get-WmiObject Win32_ComputerSystem -ComputerName $serverName
$ramTotal = [math]::Round($ram.TotalPhysicalMemory / 1GB, 2)

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "4. RAM Information"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "RAM (GB): $ramTotal"
$paragraph.Range.InsertParagraphAfter()

# Disk information
$disks = Get-WmiObject Win32_LogicalDisk -ComputerName $serverName

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "5. Disk Information"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

foreach ($disk in $disks) {
    $drive = $disk.DeviceID
    $size = [math]::Round($disk.Size / 1GB, 2)
    $freeSpace = [math]::Round($disk.FreeSpace / 1GB, 2)

    $paragraph = $document.Content.Paragraphs.Add()
    $paragraph.Range.Text = "Drive: $drive"
    $paragraph.Range.InsertParagraphAfter()

    $paragraph = $document.Content.Paragraphs.Add()
    $paragraph.Range.Text = "Size (GB): $size"
    $paragraph.Range.InsertParagraphAfter()

    $paragraph = $document.Content.Paragraphs.Add()
    $paragraph.Range.Text = "Free Space (GB): $freeSpace"
    $paragraph.Range.InsertParagraphAfter()
}

# Installed Software
$software = Get-WmiObject Win32_Product -ComputerName $serverName | Select-Object Name

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "6. Installed Software"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

foreach ($sw in $software) {
    $installedSoftware = $sw.Name

    $paragraph = $document.Content.Paragraphs.Add()
    $paragraph.Range.Text = $installedSoftware
    $paragraph.Range.InsertParagraphAfter()
}

# Installed Roles
$roles = Get-WindowsFeature -ComputerName $serverName | Where-Object { $_.Installed -eq "True" } | Select-Object Name

$paragraph = $document.Content.Paragraphs.Add()
$paragraph.Range.Text = "7. Installed Roles"
$paragraph.Range.set_Style("Heading 2")
$paragraph.Range.InsertParagraphAfter()

foreach ($role in $roles) {
    $installedRole = $role.Name

    $paragraph = $document.Content.Paragraphs.Add()
    $paragraph.Range.Text = $installedRole
    $paragraph.Range.InsertParagraphAfter()
}

# Add page numbers
$document.Sections.Headers.Footer.Range.Select()
$wordApp.Selection.Fields.Add($wordApp.Selection.Range, 33)

# Save and close the Word document
#$document.SaveAs("ServerInformation.docx")
#$document.Close()

# Save the document
$docPath = "C:\Temp\server_document.docx"
$document.SaveAs([ref]$docPath)
$document.Close()

# Quit the Word application
$wordApp.Quit()
