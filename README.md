# powershell-examples
Powershell Examples

http://ramblingcookiemonster.github.io/images/Cheat-Sheets/powershell-cheat-sheet.pdf

## Help

```
$PSVersionTable

get-module -listavailable

update-help

Get-Help about*
Get-Help about_commonparameters

Get-Help *comparison*

Get-Help -showwindow get-service   # v3

get-command *cert*
get-command *-Object
get-command -Module PKI,WebAdministration
get-command -verb new
get-command -noun proc*

get-alias

Get-Process (gps, ps)

get-psdrive | get-member
```

## Types

```
"abc".GetType().name
$a -is [int]
$false -is [bool]

$a = 123  # [Int] [Int32] [Int64] [Double]
$a = "ABC", 'ABC'
$a = "ABC"[0]
$a = [byte]97
$a = "体"

# Arrays
$arr = 1,2,3
$arr = @(1)
@(@(@(1,2,3,4))).length  # Automatically flatten

# Sets
$s = @{a = 1; b = 2.0; c = "33"}

$d = Get-Date # [DateTime]
$d = [DateTime]"2001-12-31"
$d = [DateTime]"2001-12-31 23:59:45"

# Cast
$a = [String]123
$a = [Int]"123"
```

## Basics

```
1..20 | select -Skip 5 -First 3
6
7
8

foreach ($item in 1,2,3) {$item}
1,2,3 | ForEach-Object{ $_ }
1,2,3 | %{ $_ }
$s = 1
foreach ($var in $s) { $var } # scalar handled as a collection

# Empty set will iterate once
# -> You cannot call a method on a null-valued expression.
foreach ($file in @{}) {$item.ToString()}
```
## Input

```
Pipeline: Get-Process accepts input object with PropertyName=Name
"a*","b*","c*" | select @{n='Name';e={$_}} | Get-Process

Via parenthesis
Get-WmiObject -Class Win32_BIOS -ComputerName ($mylist = @('mach1.domain.com','mach2.domain.com'))
```

## Formatting

```
Get-Process explorer | format-list *
get-service | get-member      # alias gm
get-service | Select-Object Name,Status
get-service | export-csv -Path e:\services.csv
get-service | out-gridview

get-process | where {$_.handles -gt 1000}
get-process | where handles -gt 1000 | sort handles  # v3

get-process | select -ExpandProperty name
(get-process).name   # v3

dir e:\temp | select name,length | sort length
```

## Jobs

```
help about_jobs

$j = Start-Job {sleep 3; echo "Finished!"}
Get-Job
Receive-Job $j > result.txt
$res = Receive-Job $j
```

## XML

```
$myxml = [xml](cat E:\data\xml\metadata.xml)
$myxml | gm

$myxml = [xml]'<a><b>123</b><b>456</b></a>'
$myxml.a.b[1]           # 456

$e = $myxml.CreateElement('c')
```

## JSON

```
$v = '{"abc": 123, "def": 456}' | ConvertFrom-Json
```

## HTTP

```
(Invoke-RestMethod 'http://en.wikipedia.org/w/api.php?action=query&meta=siteinfo&siprop=namespaces&format=json').query.namespaces
(Invoke-RestMethod "http://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/significant_month.geojson").features
```

## Certificates

```
get-command -Module PKI
get-command *cert* | sort -Property Source

Get-ChildItem Cert:\CurrentUser\My\   | where {$_.Thumbprint -eq 'B3C28A5D56141F6107AC0F49CDE7ED50B940E39E'} | Test-Certificate -Policy SSL
Get-ChildItem Cert:\LocalMachine\My\ | Remove-Item -Confirm

cd cert:\
dir -recurse
```

## GUI

```
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Operation Completed",0,"Done",0x1)
```

```
Add-Type -AssemblyName System.Windows.Forms
$Form = New-Object system.Windows.Forms.Form
$Form.ShowDialog()
```

https://blogs.technet.microsoft.com/stephap/2012/04/23/building-forms-with-powershell-part-1-the-form/

```
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
$form= New-Object Windows.Forms.Form
$button = New-Object Windows.Forms.Button
$button.text = "Click Here!"
$button.add_click({ $form.close() })
$form.controls.add($button)
$form.ShowDialog()
```

```
# Load the Winforms assembly
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")

# Create the form
$form = New-Object Windows.Forms.Form

#Set the dialog title
$form.text = "PowerShell WinForms Example"

# Create the label control and set text, size and location
$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point 50,30
$label.Size = New-Object Drawing.Point 200,15
$label.text = "Enter your name and click the button"

# Create TextBox and set text, size and location
$textfield = New-Object Windows.Forms.TextBox
$textfield.Location = New-Object Drawing.Point 50,60
$textfield.Size = New-Object Drawing.Point 200,30

# Create Button and set text and location
$button = New-Object Windows.Forms.Button
$button.text = "Greeting"
$button.Location = New-Object Drawing.Point 100,90

# Set up event handler to extarct text from TextBox and display it on the Label.
$button.add_click({

$label.Text = "Hello " + $textfield.text

})

# Add the controls to the Form
$form.controls.add($button)
$form.controls.add($label)
$form.controls.add($textfield)

# Display the dialog
$form.ShowDialog()
```

```
<#
.SYNOPSIS
    Demonstration of System.Windows.Forms.ListView sorting via PowerShell.
.DESCRIPTION
    Displays a basic Windows Form with a ListView control and some items. Each column is sortable.
.LINK
    https://etechgoodness.wordpress.com/2014/02/25/sort-a-windows-forms-listview-in-powershell-without-a-custom-comparer/
.NOTES
    By Eric Siron
    Version 1.0.1 June 6th 2014:  Slightly modified to work with a wider range of PS hosts and versions.
    Version 1.1 August 12th 2015: Improved test procedure per feedback from reader Stanley
#>
 
## Set up the environment
Add-Type -AssemblyName System.Windows.Forms
$LastColumnClicked = 0 # tracks the last column number that was clicked
$LastColumnAscending = $false # tracks the direction of the last sort of this column
 
## Create a form and a ListView
$Form = New-Object System.Windows.Forms.Form
$ListView = New-Object System.Windows.Forms.ListView
 
## Configure the form
$Form.Text = "ListView Sort Demo"
 
## Configure the ListView
$ListView.View = [System.Windows.Forms.View]::Details
$ListView.Width = $Form.ClientRectangle.Width
$ListView.Height = $Form.ClientRectangle.Height
$ListView.Anchor = "Top, Left, Right, Bottom"
 
# Add the ListView to the Form
$Form.Controls.Add($ListView)
 
# Add columns to the ListView
$ListView.Columns.Add("Item Name", -2) | Out-Null
$ListView.Columns.Add("Color") | Out-Null
$ListView.Columns.Add("Size") | Out-Null
$ListView.Columns.Add("Weight") | Out-Null
 
# Add list items
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Barlav")
$ListViewItem.Subitems.Add("Green") | Out-Null
$ListViewItem.Subitems.Add("Tiny") | Out-Null
$ListViewItem.Subitems.Add("11") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Floomquet")
$ListViewItem.Subitems.Add("Red") | Out-Null
$ListViewItem.Subitems.Add("Large") | Out-Null
$ListViewItem.Subitems.Add("2") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Gardgel")
$ListViewItem.Subitems.Add("Yellow") | Out-Null
$ListViewItem.Subitems.Add("Jumbo") | Out-Null
$ListViewItem.Subitems.Add("7") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Wilbit")
$ListViewItem.Subitems.Add("Turquoise") | Out-Null
$ListViewItem.Subitems.Add("Gigantic") | Out-Null
$ListViewItem.Subitems.Add("1") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
$ListViewItem = New-Object System.Windows.Forms.ListViewItem("Zasitch")
$ListViewItem.Subitems.Add("Beige") | Out-Null
$ListViewItem.Subitems.Add("Small") | Out-Null
$ListViewItem.Subitems.Add("5") | Out-Null
$ListView.Items.Add($ListViewItem) | Out-Null
 
## Set up the event handler
$ListView.add_ColumnClick({SortListView $_.Column})
 
## Event handler
function SortListView
{
 param([parameter(Position=0)][UInt32]$Column)
 
$Numeric = $true # determine how to sort
 
# if the user clicked the same column that was clicked last time, reverse its sort order. otherwise, reset for normal ascending sort
if($Script:LastColumnClicked -eq $Column)
{
    $Script:LastColumnAscending = -not $Script:LastColumnAscending
}
else
{
    $Script:LastColumnAscending = $true
}
$Script:LastColumnClicked = $Column
$ListItems = @(@(@())) # three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
 
foreach($ListItem in $ListView.Items)
{
    # if all items are numeric, can use a numeric sort
    if($Numeric -ne $false) # nothing can set this back to true, so don't process unnecessarily
    {
        try
        {
            $Test = [Double]$ListItem.SubItems[[int]$Column].Text
        }
        catch
        {
            $Numeric = $false # a non-numeric item was found, so sort will occur as a string
        }
    }
    $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
}
 
# create the expression that will be evaluated for sorting
$EvalExpression = {
    if($Numeric)
    { return [Double]$_[0] }
    else
    { return [String]$_[0] }
}
 
# all information is gathered; perform the sort
$ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscending}
 
## the list is sorted; display it in the listview
$ListView.BeginUpdate()
$ListView.Items.Clear()
foreach($ListItem in $ListItems)
{
    $ListView.Items.Add($ListItem[1])
}
$ListView.EndUpdate()
}
 
## Show the form
$Response = $Form.ShowDialog()
```

## Win32

```
Get-WmiObject -List | sort
```

## Active Directory

```
(New-Object DirectoryServices.DirectorySearcher "ObjectClass=user").FindAll() | Select path

([ADSISearcher]"Name=John*").FindAll() | Select Path

([ADSISearcher]"(|(Name=John*)(Name=Jane*))").FindAll() | Select Path

([ADSISearcher]"(Name=John*)").FindAll() | Select Properties | %{ $_.properties}

([ADSISearcher]"(Name=John*)").FindAll() | Select -ExpandProperty Properties

([ADSISearcher]"(Name=Xa*)").FindOne().Properties.directreports
```

## Internet Explorer

```
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate("http://www.microsoft.com/technet/scriptcenter/default.mspx")
$ie.Document.Body.InnerText
$ie.Quit()
Remove-Variable ie
```
