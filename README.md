# powershell-examples
Powershell Examples

http://ramblingcookiemonster.github.io/images/Cheat-Sheets/powershell-cheat-sheet.pdf

## Help

```
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
