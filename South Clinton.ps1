using namespace System.Management.Automation.Host

<#
.Synopsis
 View congregation data and personal info
.DESCRIPTION
 Use external files for congregation and personal info
.EXAMPLE
'.\South Clinton.ps1'
.EXAMPLE
'.\South Clinton.ps1 -find TEXT    - search the CSV'
.EXAMPLE
'.\South Clinton.ps1 -group NAME   - show the FS group'
.EXAMPLE
'.\South Clinton.ps1 -zip TEXT     - search the encrypted file'
#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory=$false, Position=0, ValueFromPipeline=$false)]
    [string]
    $find="",
    [Parameter(Mandatory=$false, Position=1, ValueFromPipeline=$false)]
    [string]
    $group="",
    [Parameter(Mandatory=$false, Position=2, ValueFromPipeline=$false)]
    [string]
    $zipped="",
    [Parameter(Mandatory=$false, Position=3, ValueFromPipeline=$false)]
    [switch]
    $excel,
    [Parameter(Mandatory=$false, Position=4, ValueFromPipeline=$false)]
    [switch]
    $passthru
)



Begin {

Write-Verbose "Starting '$($myinvocation.mycommand)'"
Write-Verbose "Running under PowerShell $($psversiontable.PSVersion)"
# Write-Verbose "Operating System: $((Get-Ciminstance -classname win32_operatingsystem).caption)" # Verbose info from Get-CimInstance
Write-Verbose "PSBoundparameters:"
Write-Verbose ($PSBoundParameters | Out-String)

LogPSStart -Message "[$((Get-Date).TimeofDay) BEGIN   ] Begining $($myinvocation.mycommand)"

Pause

} # begin

Process {

function Do-Popup {
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory=$false, Position=0, ValueFromPipeline=$false)]
        [string]
        $Text = "Text to Display",
        
        [Parameter(Mandatory=$false, Position=1, ValueFromPipeline=$false)]
        [Int32]
        $Duration = 5,
        
        [Parameter(Mandatory=$false, Position=2, ValueFromPipeline=$false)]
        [string]
        $Title = "TitleBar"
    )
    $ws = New-Object -ComObject Wscript.Shell
    $ws.popup($Text, $Duration, $Title) | Out-Null
}
function DoExcel
{
    $cols = @("Name", "GRP", "Phone", "Email", "Street", "City", "State", "Zip", "Elder", "Pioneer", "MS", "Notes") 

    # "First Name", "Last Name", 
    # "Phone1", "Type1", "Phone2", "Type2", 
    # "Labels", 
    # "E-mail 1 - Type", 
    # "E-mail 2 - Type", 
    # "E-mail 2 - Value", 
    # "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2 - Value", # "Phone 3 - Type", "Phone 3 - Value", 
    # "Address 1 - Type",  
    # "Address 1 - Formatted", 
    # "Address 1 - Country",
    # "Address 2 - Type",  
    # new properties/columns besides 'Notes'

    $fsgroup_s = @("Bolden","Brown","Holmes","Kirkland","Lewis","Prince","Ruiz","Vann")
            
    $ExcelParams = @{ Path = $env:TEMP + '\Excel.xlsx'; Show = $true; Verbose = $true; now = $true; FreezetopRow = $true}
    cd C:\Users\keith\iCloudDrive\csv #### PROBLEMATIC  -  NEEDS TO WORK ANYWHERE, NOT JUST AT HOME ####
    Remove-Item -Path $ExcelParams.Path -Force -EA Ignore


    $cont1 = Import-Csv .\junk.csv # | Where-Object {  ( $fsgroup_s -contains $_."Address 2 - Street"  ) } # remove rows that MATCH "z0"  
    
    <#
    Google Contact Changes 2024
    "First Name" old "Given Name"
    "Last Name" old "Family Name"
    "Labels" old "Group Membership"
    "E-mail 1 - Value" old "Email 1 - Value"
    #>

    # $_.Name -notmatch "z0" -or 
    # ! ([string]::IsNullOrEmpty($_."Address 2 - Street"))
    # ! ( $_."Address 2 - Street" -eq "" )
    # ( $_."Address 2 - Street" -notmatch "^\s*$" )

    $cont = $cont1 | Select-Object *, $groupoverseers_, $name_ # add a property 'OVERSEER' 9/24/2023

    $cont | Add-Member -MemberType AliasProperty -name "GRP" -value "OVERSEER" # no longer using "Address 2 - Street" 9/24/2023
    $cont | Add-Member -MemberType AliasProperty -name "Street" -value "Address 1 - Street"
    $cont | Add-Member -MemberType AliasProperty -name "City" -value "Address 1 - City"
    $cont | Add-Member -MemberType AliasProperty -name "State" -value "Address 1 - Region"
    $cont | Add-Member -MemberType AliasProperty -name "Zip" -value "Address 1 - Postal Code"
    $cont | Add-Member -MemberType AliasProperty -name "Email" -value "E-mail 1 - Value"
    $cont | Add-Member -MemberType AliasProperty -name "Phone" -value "Phone 1 - Value"
    $cont | Add-Member -MemberType AliasProperty -name "Phone2" -value "Phone 2 - Value"
    $cont | Add-Member -MemberType AliasProperty -name "Type1" -value "Phone 1 - Type"
    $cont | Add-Member -MemberType AliasProperty -name "Type2" -value "Phone 2 - Type"
    
    $cont | Add-Member -MemberType ScriptProperty -name "Elder" -value { ($this."Labels" -match "Elders") ? "Yes" : ""  } 
    $cont | Add-Member -MemberType ScriptProperty -name "Pioneer" -value { ($this."Labels" -match "Pioneer") ? "Yes" : "" } 
    $cont | Add-Member -MemberType ScriptProperty -name "MS" -value { ($this."Labels" -match "MS") ? "Yes" : "" } 

    # $cont | Add-Member -MemberType ScriptProperty -name "Main"  -value { "$($this."Phone1") $($this."Type1"[0])"}
    # $cont | Add-Member -MemberType ScriptProperty -name "Other" -value { "$($this."Phone2") $($this."Type2"[0])"}

    $cont | Select-Object $cols | Group-Object "GRP" | ForEach-Object { $_.Group | Export-Excel -path ($env:TEMP + '\Excel.xlsx') -WorksheetName $_.Name -TableName $_.Name -AutoSize -AutoFilter }
    $cont | Select-Object $cols | Export-Excel @ExcelParams -IncludePivotChart -WorksheetName Master -MoveToStart -TableName contacts -AutoSize

    # $cont | Select $cols  |Export-Excel @ExcelParams -IncludePivotChart -TableName contacts -AutoSize
}

# Desktop Shortcut
# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -ExecutionPolicy Bypass -File "C:\Users\keith\iCloudDrive\csv\South Clinton.ps1"

# Clear-Host

Set-Strictmode -Version 3

$dbg = "debug"
Set-PSBreakpoint -Variable dbg -Mode Read | Out-Null

Add-Type -AssemblyName System.Windows.Forms
# $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
# $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $PSScriptRoot }
# $null = $FileBrowser.ShowDialog()
# $filePicker = $FileBrowser.FileName

# $host.UI.RawUI.WindowSize.Height = 200

# Get-PSReadLineOption
${[} = "$([char]0x1b)[30;47m"
${]} = "$([char]0x1b)[0m"

$enw = '^\s*$'



$fileloc1 = $PSScriptRoot + "\contacts.csv" 
$fileloc2 = $PSScriptRoot + "\schedule.csv"
$fileloc3 = $PSScriptRoot + "\junk.csv"

<#
Google Contact Changes 2024
"First Name" old "Given Name"
"Last Name" old "Family Name"
"Labels" old "Group Membership"
"E-mail 1 - Value" old "Email 1 - Value"
#>

$contacts = Import-csv $fileloc3 | Sort-Object "Family Name", "Given Name"  
                                               # Name(Full Name) Given Name(First Name) Family Name(Last Name)  
                                               
$rowindex = 0
$contacts = $contacts | Select-Object *, RowNum_, Prop1_, Prop2_, Prop3_ # NOT Calcuated Properties >>> RowNum_, Prop1_, Prop2_, Prop3_
foreach ( $row in $contacts) {$row.RowNum_ = $rowindex; $rowindex++ } 

$sched = Import-Csv $fileloc2

# $list = [System.Collections.Generic.List[pscustomobject]]::new()
$list = [System.Collections.Generic.List[pscustomobject]](Import-CSV $fileloc3)

$cols = $contacts[0].psobject.properties.name # all of the junk.csv column names
$colhash = [ordered]@{}; $ndx = 0
foreach( $val in $cols ) { $colhash.Add( $val, $ndx++ )}

$hset = [System.Collections.Generic.HashSet[string]]@() # unique array [hashset]

$contacts | Foreach-Object {  # process all rows, every column in each row at a time 
    
    foreach ($prop in $_.PSObject.Properties)
    {
        # doSomething $prop.Name, $prop.Value
        # if ( $prop.value -notmatch '^\s*$' ) { "$($prop.name):$($prop.value)" }  # display were cells have data/value
        # if ( $prop.value -match '^\s*$' ) { "$($prop.name):$($prop.value)" }     # display were cells DONT have data/value
    } 

}

$contacts | ConvertTo-Json | Out-File ($PSScriptRoot + "\junk.json") # added 4/3/2024
$pscustomjson = $contacts | ConvertTo-Json | ConvertFrom-Json # added 4/3/2024

#
## process all rows, every column in each row at a time
#
# $contacts.foreach( { foreach( $prop in $_.psobject.properties ) { if ($prop.value -notmatch '^\s*$') {"$($prop.name):$($prop.value)"} } } )
# $contacts.foreach( { foreach( $prop in $_.psobject.properties ) { if ($prop.value -match '^\s*$') {"$($prop.name):$($prop.value)"} } } )


$calevent_= @{ 
    N = "Date          ";
    E={ "{3:ddd} {0:d2}/{1:d2}/{2:d4}  " -f [int]$_.MONTH, [int]$_.DAY, [int]$_.YEAR, 
    $(([datetime]"$($_.YEAR) $($_.MONTH) $($_.DAY)"))}}
    # $(([datetime]"$(YEAR) $(MONTH) $(DAY)"))}}

$calevent2_= @{ 
    N = "Date_          ";
    E={ "{3:ddd} {0:d2}/{1:d2}/{2:d4}  " -f [int]$_.M, [int]$_.D, [int]$_.Y, 
    $(([datetime]"$($_.Y) $($_.M) $($_.D)"))}}
    # $(([datetime]"$(YEAR) $(MONTH) $(DAY)"))}}
    


$roles_array = ("Elder", "Pioneer", "MS", "SMPW", "Cart", "LDC", "Old-Contacts", "No-FSG", "DI", "Inactive")
$worksheets = ("RUIZ-FSG", "BOLDEN-FSG", "BROWN-FSG", "HOLMES-FSG", "KIRKLAND-FSG", "LEWIS-FSG", "PRINCE-FSG", "VANN-FSG", "MISC", "DI", "updated_", "OLD-CONTACTS") # 6/15/23 removed WILKINSON - added HOLMES & DI & updated_

$roles_ = 
@{
    N="Roles"
    E={
        $found = ""
        foreach ( $rol in $roles_array)
        {
            if ($($_."Labels") -match $rol) {$found = $found + "$rol"[0]} # First letter
        }
        $found
    }
}

$groupoverseers_ = # 6/15/23 CREATED
@{
    N="OVERSEER"
    E={
        $found = ""
        foreach ( $over in $worksheets)
        {
            if ($($_."Labels") -match $over) {$found = $over}
        }
        $found
    }
}



$name_ = @{ N = 'Name'; # subsequent calls reverses the name again
    E = {
        
    # $first, $last = $_."name" -split " ";

    # $first = $first.replace(',',''); $last = $last.replace(',','');        
    
    $last = $_."Last Name"; $first = $_."First Name";
    
    "$last, $first";

    }
}


$noswapnameP_ = @{ N = 'REGULAR PIONEERS  '; 
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$first, $last"}}

$noswapnameM_ = @{ N = 'MINISTERIAL SERV'; 
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$first, $last"}}  
    
$noswapnameE_ = @{ N = 'ELDERS'; 
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$first, $last"}}

$phone_ = @{ N = 'Phone';  E = {$_."phone 1 - value"}}
$addr_  = @{ N = 'Address';  E = {($_."address 1 - formatted").replace("`n"," ").replace("US","")}}
$email_ = @{ N = 'Email';  E = {$_."e-mail 1 - value"}}

$notes_ = @{ N = 'Notes'; E = {$_."Notes"}}
# $notes_ = @{ N = 'Notes'; E = {($_."Notes").replace("`n"," ")}}

$fsgroup_ = @{ N = 'FSGroup';  E = {($_."address 2 - street").ToUpper()}}

#return true/false
$pioneer_ = @{ N = 'Pioneer';  E = {$_."Labels" -match "Pioneer"}}
$elder_ =  @{ N = 'Elder';  E = {$_."Labels" -match "Elders"}}
$ms_ = @{ N = 'MS';  E = {$_."Labels" -match "MS"}}
$old_ = @{ N = 'Old';  E = {$_."Labels" -match "Old-Contacts"}}
$nofsg_ = @{ N = 'No-FSG';  E = {$_."Labels" -match "No-FSG"}}

$updated_ = @{ N = 'Updated';  E = {if($_."Custom Field 1 - Label" -match "Updated"){$_."Custom Field 1 - Value"} }}

# $contacts | Add-Member -MemberType AliasProperty -name "updated_" -value "Custom Field 1 - Value" # 6/15/23 


$gridview = $contacts | 
Select-Object $name_, $phone_, $addr_, $email_, $fsgroup_, $roles_, $groupoverseers_, $pioneer_, $elder_, $ms_, $old_, $nofsg_, RowNum_, Prop1_, Prop2_, Prop3_ # $notes_
# use the 'N' of the calculated properties going forward when selecting objects from '$gridview'
# Name  Phone  Addesss  Email  Notes  FSGroup  Pioneer  Elders  MS  Old  No-FSG  Roles


$search = $contacts | 
Select-Object $name_, $phone_, $addr_, $email_, $roles_, $groupoverseers_, $notes_, $pioneer_, $elder_, $ms_, RowNum_, Prop1_, Prop2_, Prop3_ # 6/15/23 removed $fsgroup_ 
# use the 'N' of the calculated properties going forward when selecting objects from '$search'
# Name  Phone  Addesss  Email  Notes  FSGroup  Pioneer  Elders  MS  Old  No-FSG  


$limited = $contacts | 
Select-Object $name_, $phone_, $addr_, $email_, $roles_, $groupoverseers_, $updated_, RowNum_, Prop1_, Prop2_, Prop3_ # 6/15/23 removed $fsgroup_, $pioneer_, $elder_, $notes_,
# use the 'N' of the calculated properties going forward when selecting objects from '$search'
# Name  Phone  Addesss  Email  Notes  FSGroup  Pioneer  Elders  MS  Old  No-FSG 



# $IsWindows ? ($opsys = "Windows") : ($opsys = "Linux")
# $opsys

# mode 300  # may leave artifacts

# $wshell = New-Object -ComObject wscript.shell;
# $wshell.SendKeys('%~') # ALT-ENTER

# Add-Type -AssemblyName System.Windows.Forms
# [System.Windows.Forms.SendKeys]::SendWait('%~')

# $wshell = New-Object -ComObject wscript.shell;
# $wshell.AppActivate('South Clinton')
# Sleep 1
# $wshell.SendKeys('%~')

$params = @{
    MemberType = "ScriptMethod"
    Name       = "OutHashtable"
    Value      = {
        $hash = [ordered]@{}
        $this.psobject.properties.name.foreach({
                $hash[$_] = $this.$_
            })
        return $hash
    }
}


$params2 = @{
    MemberType = "ScriptMethod"
    Name       = "ColsWithData"
    Value      = {
        
        # $this.foreach() does not work in 5.1 but OK in 7.xxx
        $this | foreach-object( { 
                
                foreach( $p in $_.psobject.properties ) 
                { 
                    if ($p.Value -notmatch '^\s*$') # -notmatch: columns with data
                    {
                        # "$($p.name) : $($p.value)"
                        # $this -eq $_  so  $this.$_ -eq $this.$this
                        $hset.Add($p.Name) | Out-Null
                    } 
                }
                                
            } )
          
    }
    
}

    
$contacts | Add-Member -MemberType ScriptMethod -Name info -Value { param([string]$find = "Default" )   if ( $this.name -match $find ) 
{ 
    "(#:{0:D3}) {1} {2}" -f $this.RowNum_, ( ( $this."Phone 1 - Value" ) ? $($this."Phone 1 - Value" -replace " ","")  : "NO-PHONE" ) , $this.name
    if ($this."E-mail 1 - Value") {$this."E-mail 1 - Value"} else {"NO-EMAIL"}

    # if ($this."Labels" -match "\w*-FSG") {$matches[0];""} else {"NO-FSG`n"} 
    
    $val = [regex]::Match($this."Labels","\w*-FSG")
    if ( $val.Success ) {$val.value;""} else {"NO-FSG`n"}

    }
}

$limited | Add-Member -MemberType ScriptMethod -Name info -Value { param([string]$find = "Default" )   if ( $this.name -match $find ) 
{ 
    "(#:{0:D3}) {1} {2}" -f $this.RowNum_, ( [string]::IsNullOrEmpty( $this."Phone" ) ? "NO-PHONE" : $this."Phone" -replace " ","" ) , $this.name
    if ($this."Email") {$this."Email"} else {"NO-EMAIL"}

    # if ($this."Labels" -match "\w*-FSG") {$matches[0];""} else {"NO-FSG`n"} 
    
    # $val = [regex]::Match($this."OVERSEER","\w*-FSG")
    if ( $this."OVERSEER" ) {$this."OVERSEER";""} else {"NO-FSG`n"}
    "Notes:`n{0}" -f $this.Notes
    "<------------>`n"

    }
}

$search   | Add-Member -MemberType ScriptMethod -Name PP -Value { $this.psobject.properties }
$contacts | Add-Member -MemberType ScriptMethod -Name PP -Value { $this.psobject.properties }
$search   | Add-Member -MemberType ScriptMethod -Name PPN -Value { $this.psobject.properties.name }  # name is an array  ....name.foreach ( $_ ... )
$contacts | Add-Member -MemberType ScriptMethod -Name PPN -Value { $this.psobject.properties.name }  # name is an array  ....name.foreach ( $_ ... )
$search   | Add-Member -MemberType ScriptMethod -Name PPV -Value { $this.psobject.properties.value } # value is an array ...value.foreach ( $_ ... )
$contacts | Add-Member -MemberType ScriptMethod -Name PPV -Value { $this.psobject.properties.value } # value is an array ...value.foreach ( $_ ... )

Add-Member -InputObject $contacts -MemberType ScriptMethod -Name PPPN -Value { $this.psobject.properties.name }  # name is an array  ....name.foreach ( $_ ... )

$search   | Add-Member @params  # .OutHashTable()
$contacts | Add-Member @params  # .OutHashTable()
$search   | Add-Member @params2 # .ColsWithData()
$contacts | Add-Member @params2 # .ColsWithData()
$contacts | Add-Member -MemberType ScriptMethod -Name OutJson -Value {$this | Convertto-json }
$search   | Add-Member -MemberType ScriptMethod -Name OutJson -Value {$this | Convertto-json }

$CSVData =  { $this <# $this | gm -MemberType NoteProperty #> }
$search   | Add-Member -MemberType ScriptMethod -Name Func -Value {$this; 1..5 | % {& $CSVData} }
$contacts | Add-Member -MemberType ScriptMethod -Name Func -Value {$this}

foreach ($row in $sched) # add four more columns to schedule.csv
{
    $row | Add-Member -NotePropertyName 'Y' -NotePropertyValue ''
    $row | Add-Member -NotePropertyName 'M' -NotePropertyValue ''
    $row | Add-Member -NotePropertyName 'D' -NotePropertyValue ''
    $row | Add-Member -NotePropertyName 'MESS' -NotePropertyValue ''
}

[int]$inc = 25

foreach ( $row in $sched[0..24]) # fill in the new columns with recs 25-49
{
    $row.Y = $sched[$inc].YEAR
    $row.M = $sched[$inc].MONTH
    $row.D = $sched[$inc].DAY
    $row.MESS = $sched[$inc].MESSAGE
    $inc++
}

# $contacts | Add-Member -MemberType AliasProperty -Name "GMEM" -Value "Labels"
# $contacts | gm

# $contacts.OutHashTable()

#
#
#
#
clear
# Do-Popup "Welcome" "5" "SouthClinton"
# winget list "JW Library"
# (winget list "JW Library" | Select-String 'JW.*').Matches.Value
$JW = "`nCurrent JW Library Version:`n"
#  $JW += ((winget list "JW Library" | Select-String 'Watch.*').Matches.Value).Replace(" ","`n")  # 'Watchtower' no longer returned
$JW += ((winget list "JW Library" | Select-String 'JW.*').Matches.Value).Replace(" ","`n")
$contacts.ColsWithData()
$v = $PSVersionTable
$os = (Get-CimInstance Win32_OperatingSystem).Caption 
$plat = [System.Environment]::OSVersion.Platform
# "Platform:$plat  OS:$os  PS:$($v.PSEdition) v$($v.PSVersion.ToString())  Columns:$($cols.Count):$($hset.Count)" # $hset.Clear()
$s = "Platform:$plat`nOS:$os`nPS:$($v.PSEdition)`nv$($v.PSVersion.ToString())`nColumns:$($cols.Count):$($hset.Count)`n$JW"
$pop = { Do-Popup $s "5" "SouthClinton" }
& $pop
# $PSCmdlet.MyInvocation.BoundParameters
# $PSBoundParameters
#
#
#
#


do # MAIN LOOP
{ 
    
    # command line inputs:  $find  $group  $zipped $excel $passthru
    if ( ! [string]::IsNullOrEmpty($find)  ) 
        {"We have FIND input"; $ans = "F"+$find; $find = "" }
    
    elseif ( ! [string]::IsNullOrEmpty($group)  ) 
        {"We have GROUP input"; $ans = "&"+$group; $group = "" }

    elseif ( ! [string]::IsNullOrEmpty($zipped)  ) 
        {"We have ZIPPED input"; $ans = "S"+$zipped; $zipped = "" }
    
    elseif ( [bool]$excel -eq $true  ) 
        {"We have EXCEL input"; $ans = "E"+$excel; [bool]$excel = $false }
    
    elseif ( [bool]$passthru -eq $true  ) 
        {"PASSTHRU not yet implemented"; $ans = ""; pause; [bool]$passthru = $false; systeminfo; continue}
    # command line inputs:  $find  $group  $zipped  $excel  $passthru
    else 
    {
    
    
        $ans = Read-Host "${[}'Fxxx':${]} Find(csv) | ${[}'CAL':${]} Dates | ${[}'View':${]} GridView | ${[}'Gxxx':${]} Grep | ${[}'&[FSO]':${]} Groups | ${[}'%':${]} Role Totals | ${[}'R[CDEILMNOPS]':${]} Roles | ${[}'.':${]} App | ${[}'QUIT':${]}" # $ans or $ans[0] is a REGEX used with -match
    
    
    }

    if ($ans.Length -eq 0 ) { $ans = '*' } # $ans = "[a-z0-9]"} # an empty string matches ANYTHING
    
    Clear-Host
    [System.Console]::Clear() # redundant

      
    if ( $ans[0] -match 'g') # greps
    { 
        $ans = $ans.Remove(0,1)
        if ( [string]::IsNullOrEmpty($ans) ) 
        {
            "Nothing to search for:"
            continue
        }
        
        $unzipped = ( .\7za.exe e .\whatever.7z -p"${c:\zealous}" -so  )
        ( $we = $unzipped ) | Out-file temp.fille
        # ( $we = $unzipped ) -> variable squeezing
        # $we is used for debbuging
        $unzipped = ""
        # function read-unzipped { & $unzipped }
        $arr = @(); $arr = (Get-Content -raw .\temp.fille) -split '\r?\n\r?\n' -match $ans
        (Get-Content -raw .\temp.fille) -split "`r?`n`r?`n" -match $ans  # ` backtick 'escape char' must be in doublequotes  
                                                                         # Escape sequences are only interpreted when contained in double-quoted (") strings.
        rm .\temp.fille
        continue

        # The above is same as for '/' or 's' below - only seaches 'whatever.txt'.
        # TODO - Incorporate the Open File Dialog for any file

        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $PSScriptRoot }
        $null = $FileBrowser.ShowDialog()
        # Get-Content $FileBrowser.FileName
        # Select-String -Path $FileBrowser.FileName -Pattern $ans
        (Get-Content -Raw $FileBrowser.FileName) -split '\r?\n\r?\n' -match $ans
        
        continue

    }    
    
    if ( $ans[0] -match '\$') # display variable
    {
        "in variable"
        
        if ( $ans.Length -eq 1 ) { ls variable:/ }
        
        Invoke-Expression $ans
        continue
    }
    
    if ( $ans[0] -match '\?') # Help
    { 
        $ans = $ans.Remove(0,1)

        @"
         ${[} E             :${]} DoExcel()
         ${[} Fxxx          :${]} Find in 'junk.csv'
         ${[} Gxxx          :${]} Grep
         ${[} S/xxx or /xxx :${]} Find in 'whatever.7z'
         ${[} Cal           :${]} Dates
         ${[} &[FSO]        :${]} List FS Group members
         ${[} %             :${]} Role Totals
         ${[} R[CDEILMNOPS] :${]} Roles 
         ${[} U             :${]} View updated_ contacts
         ${[} Quit          :${]} Exit/Bye
         ${[} View          :${]} GridViewqqq
         ${[} .             :${]} App 
         ${[} ?             :${]} Help
         ${[} -Excel        :${]} DoExcel() <switch>
         ${[} -Passthru     :${]} System Information <switch>
         ${[} -Find         :${]} Find in 'junk.csv'
         ${[} -Group        :${]} List FS Group members
         ${[} -Zipped       :${]} Find in 'whatever.7z'
"@

        continue
    }

    if ( $ans[0] -match 'd') # DEBUG
    { 
        $dbg
        $ans = $ans.Remove(0,1)
        continue
    } 
    
    if ( $ans[0] -match 'c') # CAL
    {
        
        $sched | Select-Object $calevent_, MESSAGE, $calevent2_, MESS -first 25 | out-host | ft; wsl cal -A 2; read-Host -Prompt "`r"; Clear-Host; continue 
    }

    if ( $ans[0] -match 'v') { $gridview | Out-GridView -PassThru | Select-Object Name, Phone, Email, Address | Out-String ;  continue } # GRID

    if ( $ans[0] -match 'q') { exit } # QUIT
    
    # $choice = $search | Select-Object  Name, Phone, Address, Email, FSGroup, $roles_, $notes_

    if ( $ans[0] -eq "E") 
    { 
        DoExcel; continue 
    }

    if ( $ans[0] -eq "u") # 6/15/23  Show edit history of contacts if added to google contacts
    { 
        $limited | Where-Object { !([string]::IsNullOrEmpty($_."updated_")) } | ft
        continue 
    }
    
    if ( $ans[0] -eq "F") # search CSV for any part of a cell number / email address
    {
        # $ans = $ans.trim($ans[0])
        $ans = $ans.remove(0,1)
        # $title = @{ N='Cell'; E={$_."phone 1 - value" -replace '[()\s\s-\.]', ''} }  # strip all chars except 10 digit cell#
        
        $number = @{ N='Cell'; E={$_."phone 1 - value" -replace '[^0-9]' } }           # strip all chars except 10 digit cell#
        $email_address = @{ N='Email'; E={$_."e-mail 1 - value" } }     
        
        $cell = $contacts | Select-Object $name_, $number, $email_address
        $result = $cell | Where-Object { $_.Name -match $ans -or $_.Cell -match $ans -or $_.Email -match $ans }

        if ( [string]::IsNullOrEmpty($result) ) { continue }

        $result | Format-Table

        # "Number of finds: {0}" -f @($result).Count        
            # @(   ...   )       is needed if only ONE object is returned - Count looks for # of ARRAY items
            
        "Number of finds: {0}" -f ([array]$result).Count  
            # ([array]   ...   ) is needed if only ONE object is returned - Count looks for # of ARRAY items
        continue
    }

    if ( $ans[0] -eq "R" ) # Roles
    {
        $ans = $ans.remove(0,1)
        if ( [string]::IsNullOrEmpty($ans) ) { continue}
        # $roles_array = ("Elder", "Pion 
        switch ( [char[]]$ans ) # char array - If you give a switch an array, it will process each element in that collection.
        {
            'C' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "C" ) { $found.Name } } "${[}Cart-$(@($c).count)${]}"; $c }
                     
            'D' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "D" ) { $found.Name } } "${[}Disfellowshiped-$(@($c).count)${]}"; $c }
                       
            'E' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "E" ) { $found.Name } } "${[}Elders-$(@($c).count)${]}"; $c }  
                
            'L' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "L" ) { $found.Name } } "${[}LDC-$(@($c).count)${]}"; $c }
                
            'M' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "M" ) { $found.Name } } "${[}MS-$(@($c).count)${]}"; $c }
                
            'N' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "N" ) { $found.Name } } "${[}No-FSG-$(@($c).count)${]}"; $c }
                
            'O' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "O" ) { $found.Name } } "${[}Old-$(@($c).count)${]}"; $c }
                
            'P' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "P" ) { $found.Name } } "${[}Pioneers-$(@($c).count)${]}"; $c }
                
            'S' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "S" ) { $found.Name } } "${[}SMPW-$(@($c).count)${]}"; $c }

            'I' {   $c = foreach ( $found in $search) { if ( $found.Roles -match "I" ) { $found.Name } } "${[}Inactive-$(@($c).count)${]}"; $c }
                
            
            default { "${[}Not Found: $_${]}" }
        }
        continue
    }


    if ( $ans[0] -eq "S" -or $ans[0] -eq "/") # search encrypted file
    {
        
        $ans = $ans.remove(0,1)
        
        $unzipped = ( .\7za.exe e .\whatever.7z -p"${c:\zealous}" -so  )
        if ($ans.Length -eq 0 ) 
        { 
            "Nothing to search for:"
            continue
            
            # $unzipped | Select-String -Pattern ".*" -AllMatches
            ($unzipped) | Out-File .\searchme
            (Get-Content -Raw .\searchme) -split '\r?\n\r?\n' -match ""
            rm .\searchme
            $unzipped = ""
        }
        else {
            # $unzipped | Select-String -Pattern $ans -AllMatches -Context 3
            ($unzipped) | Out-File .\searchme
            (Get-Content -Raw .\searchme) -split '\r?\n\r?\n' -match $ans
            rm .\searchme
            $unzipped = ""
        }
        
        continue
        
        <#
        if ( ($unzipped | Select-String "kibrown031"-AllMatches) -ne $null ) # match found
                { [array]$ndxarray = ($unzipped | Select-String "kibrown031" -AllMatches).Matches.Index }

        #>
        
        <#if ( $script:ndx -lt $script:ndxarray.Count )
        {
            $TextBox8.SelectionStart = ndxarray[$script:ndx]
            $TextBox8.SelectionLength = stringToFind.Length
            $TextBox8.ScrollToCaret()
            ndx++
            inloop = $true
            return
        }
        #>
    }

    if ( $ans[0] -eq ".") # open TestForm app
    {  
        
        . ($PSScriptRoot + '.\TestForm.ps1')
    }
    elseif ( $ans[0] -eq "&" ) # FSGroup
    {
        $ans = $ans.remove(0,1).toupper()
        if ($ans.Length -eq 0 ) { continue } # an empty string matches ANYTHING

        # 6/15/23  $result = $search | Where-Object { $_.FSGroup -match "\b$ans\b" }       # EXACT match  OLD and not BOLDEN ( OLD is inside of BOLDEN)
        $result = $limited | Where-Object { $_.OVERSEER -match "\b$ans\b" }       # EXACT match  OLD and not BOLDEN ( OLD is inside of BOLDEN)
        if ( [string]::IsNullOrEmpty($result) )
            {
                # 6/15/23  $result = $search | Where-Object { $_.FSGroup -match "$ans" }  # not EXACT match
                $result = $limited | Where-Object { $_.OVERSEER -match "$ans" }  # not EXACT match
            }    

        Write-Output("Total # in the $ans FSGroup: $($result.count)")
        Clear-Host
        $display = Read-Host "Table or List?: ( 'T' for Table 'L' for List )"

        $result = $result | Sort-Object Name

        if ( $display -match 'l' ) 
            { $result | Format-List * }
        else
            { $result | Format-Table * }

        ([array]$result).count
       
    }
    elseif ( $ans[0] -eq "%" ) # Roles
    {
        # $ans = $ans.trim("%").toupper()
        # $ans = $ans.remove(0,1)

        foreach ( $rol in $roles_array )
        {  
            $result = $search | Where-Object { $_.Roles -match $rol[0]} # First letter
                $num = 0; if ( @($result).count -eq 0 ) { $num = 0} else { $num = ([array]$result).Count }
                Write-Output("Total # of ${rol}: $num")
            # DisplayPioneer
        }
        foreach ( $fsg in $worksheets )
        {
            $result = $search | Where-Object { $_.OVERSEER -match $fsg}
            $num = 0; if ( @($result).count -eq 0 ) { $num = 0} else { $num = ([array]$result).Count }
            Write-Output("Total # of $fsg-FSG: $num")
        }
        do
        {  
            $search | Where-Object { $_."Elder" -match $true } | 
                Format-Table -AutoSize -Property $noswapnameE_, Email, Phone, Address 
            break
        }
        while ($true)
        do
        {  
            $search | Where-Object { $_."MS" -match $true } | 
                Format-Table -AutoSize -Property $noswapnameM_, Email, Phone, Address 
            break
        }
        while ($true)
        do
        {  
            $search | Where-Object { $_."Pioneer" -match $true } | 
                Format-Table -AutoSize -Property $noswapnameP_, Email, Phone, Address 
            break
        }
        while ($true)
    }
    else {
        if ( $ans -eq '*' ) { continue}
        "Unknown value for `$ans[0]: {0}" -f $ans[0]
        
        $display = Read-Host "Table or List?: ( 'T' for Table 'L' for List )"
        Clear-Host    
        if ( $display -match 'l' )
        {        
            $search |
            Where-Object { $_."Name" -match $ans -or $_.Address -match $ans -or $_.OVERSEER -match $ans} | Format-List
        }
        elseif ( $display -match 't' )
        {        
            $search | 
            Where-Object { $_."Name" -match $ans -or $_.Address -match $ans -or $_.OVERSEER -match $ans} | Format-Table
        }
        else # 't' or 'l' not chosen
        {
            continue
            $result = $search | 
            Where-Object { $_."Name" -match $ans -or $_.Address -match $ans -or $_.OVERSEER -match $ans}
            $result | Format-Table
            ([array]$result).Count
        }
    }

} while ($true)

} #process

End { # shouldn't be called

    Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending $($myinvocation.mycommand)"
    Write-Verbose "END: That's All Folks!!!"
    Pause
}

Clean {

    Clear-Host
    $answer = $Host.UI.PromptForChoice('Backup key files?', '', @('&Yes', '&No', '&Cancel'), 1)
    if ($answer -eq 0) {
        Write-Verbose "YES - Saving Files to $env:HOMEPATH\Documents\Backup"
        $files=
        (
            "whatever.7z","South Clinton.ps1", "junk.csv", "schedule.csv", 
            "wslnano.ps1", "welcome.ps1", "dotsource.ps1", 
            "KB-GNUCash_Register.xml",
            "NW Scheduler South Clinton Field Service Reports.xlsx" 
        )

        foreach ($file in $files ) 
        { Copy-Item $($PSScriptRoot + "\" + $file) $env:HOMEPATH\Documents\Backup }
        
        $files = Get-ChildItem $env:HOMEPATH\Documents\Backup\ | Sort-Object -Descending -Property LastWriteTime
        foreach ( $file in $files ) { Write-Host -ForegroundColor Green "$($file.LastWriteTime) <> $($file.Name)" }
}else{
        Write-Verbose 'NO - NOT Saving Files ...'
        $files = Get-ChildItem $env:HOMEPATH\Documents\Backup\ | Sort-Object -Descending -Property LastWriteTime
        foreach ( $file in $files ) { Write-Host -ForegroundColor Green "$($file.LastWriteTime) <> $($file.Name)" }
}

    Write-Verbose "[$((Get-Date).TimeofDay) CLEAN    ] Ending $($myinvocation.mycommand)"
    LogPSStart -Message "[$((Get-Date).TimeofDay) CLEAN    ] Ending $($myinvocation.mycommand)"
}

# $contacts columns
<#
Name
Given Name
Additional Name
Family Name
Yomi Name
Given Name Yomi
Additional Name Yomi
Family Name Yomi
Name Prefix
Name Suffix
Initials
Nickname
Short Name
Maiden Name
Birthday
Gender
Location
Billing Information
Directory Server
Mileage
Occupation
Hobby
Sensitivity
Priority
Subject
Notes
Language
Photo
Group Membership
E-mail 1 - Type
E-mail 1 - Value
E-mail 2 - Type
E-mail 2 - Value
E-mail 3 - Type
E-mail 3 - Value
E-mail 4 - Type
E-mail 4 - Value
Phone 1 - Type
Phone 1 - Value
Phone 2 - Type
Phone 2 - Value
Phone 3 - Type
Phone 3 - Value
Address 1 - Type
Address 1 - Formatted
Address 1 - Street
Address 1 - City
Address 1 - PO Box
Address 1 - Region
Address 1 - Postal Code
Address 1 - Country
Address 1 - Extended Address
Address 2 - Type
Address 2 - Formatted
Address 2 - Street
Address 2 - City
Address 2 - PO Box
Address 2 - Region
Address 2 - Postal Code
Address 2 - Country
Address 2 - Extended Address
Organization 1 - Type
Organization 1 - Name
Organization 1 - Yomi Name
Organization 1 - Title
Organization 1 - Department
Organization 1 - Symbol
Organization 1 - Location
Organization 1 - Job Description
Custom Field 1 - Type
Custom Field 1 - Value
#>


# $gridview | Out-GridView -PassThru | Select-Object Name, Phone, Email, Address | Out-Printer -Name "Microsoft Print to PDF"


<#
$gridview = $contacts | Select-Object `
@{ N = 'Name'; E = {$first, $last = $_."name" -split " "; "$last, $first ";}},
@{ N = 'Phone';  E = {$_."phone 1 - value" + "  "}},
@{ N = 'Addr';  E = {($_."address 1 - formatted").replace("`n"," ").replace("US","")}},
@{ N = 'Email';  E = {$_."e-mail 1 - value"  + "  "}},
@{ N = 'FS Group';  E = {($_."address 2 - street").ToUpper()}},
@{ N = 'Pioneer';  E = {$_."Group Membership" -match "Pioneer"}},
@{ N = 'Elder';  E = {$_."Group Membership" -match "Elders"}},
@{ N = 'MS';  E = {$_."Group Membership" -match "MS"}},
@{ N = 'Old';  E = {$_."Group Membership" -match "Old-Contacts"}},
@{ N = 'No-FSG';  E = {$_."Group Membership" -match "No-FSG"}}

$gridview | Out-GridView -PassThru | Select-Object Name, Phone, Email, Addr | Out-Printer -Name "Microsoft Print to PDF"
#>


 # $gridview | Out-GridView -PassThru | Select name, Phone, Email, Address | clip
 # $gridview | Out-GridView -PassThru | Select name, Phone, Email, Address | Out-Printer -Name "Microsoft Print to PDF"
 # $gridview | Out-GridView -PassThru | Select name, Phone, Address, Email | Format-Table

    # RecNo, 
    # @{ N = 'NUMM';  E = {$script:cpt++;$script:cpt}},
    # Notes 
    # "Group Membership",
    # Notes | Out-GridView -PassThru | clip.exe
    # Notes | Out-GridView -PassThru | Out-Printer -Name "Microsoft Print to PDF"

# $contacts | select NUM, name, "phone 1 - value", "address 1 - formatted", "e-mail 1 - value", @{ Name = 'FS Group';  Expression = {$_."address 2 - street"}}, "Group Membership" | Out-GridView


<#
$printer = Get-Printer -Name "Microsoft Print to PDF" -ErrorAction SilentlyContinue
if (!$?)
{
    Write-Warning "Your PDF Printer is not yet available!"
}
else
{
    Write-Warning "PDF printer is ready for use."
}
#>

<#
1..100 | ForEach-Object {
        Write-Progress -Activity "Copying files" -Status "$_ %" -Id 1 -PercentComplete $_ -CurrentOperation "Copying file file_name_$_.txt"
        Start-Sleep -Milliseconds 500    
        # sleep simulates working code, replace this line with your executive code (i.e. file copying)
    }
#>

<#
psEdit # $profile
#>

<#
Get-CimInstance Win32_PhysicalMemory     Get-WmiObject -Class Win32_PhysicalMemory                                                                                      
Get-CimInstance Win32_Processor          Get-WmiObject -Class Win32_Processor                                                                                           
Get-CimInstance Win32_LogicalDisk        Get-WmiObject -Class Win32_LogicalDisk                                                                                         
Get-CimInstance Win32_DiskDrive          Get-WmiObject -Class Win32_DiskDrive                                                                                           
Get-CimInstance Win32_OperatingSystem    Get-WmiObject -Class Win32_OperatingSystem 
#>

