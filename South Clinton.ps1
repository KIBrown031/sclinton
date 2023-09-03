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

function DoExcel
{
    $cols = @("Name", "Group", "E-mail 1 - Value", "Main", "Other", "Street", "City", "State", "Zip", "Elder", "Pioneer", "MS", "Notes") 
    # "Given Name", "Family Name", 
    # "Phone1", "Type1", "Phone2", "Type2", 
    # "Group Membership", 
    # "E-mail 1 - Type", 
    # "E-mail 2 - Type", 
    # "E-mail 2 - Value", 
    # "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2 - Value", # "Phone 3 - Type", "Phone 3 - Value", 
    # "Address 1 - Type",  
    # "Address 1 - Formatted", 
    # "Address 1 - Country",
    # "Address 2 - Type",  
    # new properties/columns besides 'Notes'

    $fsgroups = @("Alston","Bolden","Brown","Holmes","Kirkland","Lewis","Prince","Vann")
            
    $ExcelParams = @{ Path = $env:TEMP + '\Excel.xlsx'; Show = $true; Verbose = $true; now = $true; FreezetopRow = $true}
    # cd C:\Users\keith\iCloudDrive\csv #### PROBLEMATIC  -  NEEDS TO WORK ANYWHERE, NOT JUST AT HOME ####
    Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
    $cont = Import-Csv .\junk.csv | ? {  ( $fsgroups -contains $_."Address 2 - Street"  ) } # remove rows that MATCH "z0"  

    # $_.Name -notmatch "z0" -or 
    # ! ([string]::IsNullOrEmpty($_."Address 2 - Street"))
    # ! ( $_."Address 2 - Street" -eq "" )
    # ( $_."Address 2 - Street" -notmatch "^\s*$" )

    $cont | Add-Member -MemberType AliasProperty -name "Group" -value "Address 2 - Street"
    $cont | Add-Member -MemberType AliasProperty -name "Street" -value "Address 1 - Street"
    $cont | Add-Member -MemberType AliasProperty -name "City" -value "Address 1 - City"
    $cont | Add-Member -MemberType AliasProperty -name "State" -value "Address 1 - Region"
    $cont | Add-Member -MemberType AliasProperty -name "Zip" -value "Address 1 - Postal Code"
    $cont | Add-Member -MemberType AliasProperty -name "Phone1" -value "Phone 1 - Value"
    $cont | Add-Member -MemberType AliasProperty -name "Phone2" -value "Phone 2 - Value"
    $cont | Add-Member -MemberType AliasProperty -name "Type1" -value "Phone 1 - Type"
    $cont | Add-Member -MemberType AliasProperty -name "Type2" -value "Phone 2 - Type"

    $cont | Add-Member -MemberType ScriptProperty -name "Elder" -value { ($this."Group Membership" -match "Elders") ? "Yes" : ""  } 
    $cont | Add-Member -MemberType ScriptProperty -name "Pioneer" -value { ($this."Group Membership" -match "Pioneer") ? "Yes" : "" } 
    $cont | Add-Member -MemberType ScriptProperty -name "MS" -value { ($this."Group Membership" -match "MS") ? "Yes" : "" } 

    $cont | Add-Member -MemberType ScriptProperty -name "Main"  -value { "$($this."Phone1") $($this."Type1"[0])"}
    $cont | Add-Member -MemberType ScriptProperty -name "Other" -value { "$($this."Phone2") $($this."Type2"[0])"}

    $cont | Select-Object $cols | Group-Object "Group" | ForEach-Object { $_.Group | Export-Excel -path ($env:TEMP + '\Excel.xlsx') -WorksheetName $_.Name -TableName $_.Name -AutoSize -AutoFilter }
    $cont | Select-Object $cols | Export-Excel @ExcelParams -IncludePivotChart -WorksheetName Master -MoveToStart -TableName contacts -AutoSize

    # $cont | Select $cols  |Export-Excel @ExcelParams -IncludePivotChart -TableName contacts -AutoSize
}

$fileloc1 = $PSScriptRoot + "\contacts.csv" 
$fileloc2 = $PSScriptRoot + "\schedule.csv"
$fileloc3 = $PSScriptRoot + "\junk.csv"

$contacts = Import-csv $fileloc3 | Sort-Object "Family Name", "Given Name"  
                                               # Name(Full Name) Given Name(First Name) Family Name(Last Name)     
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

#
## process all rows, every column in each row at a time
#
# $contacts.foreach( { foreach( $prop in $_.psobject.properties ) { if ($prop.value -notmatch '^\s*$') {"$($prop.name):$($prop.value)"} } } )
# $contacts.foreach( { foreach( $prop in $_.psobject.properties ) { if ($prop.value -match '^\s*$') {"$($prop.name):$($prop.value)"} } } )


$calevent= @{ 
    N = "Date          ";
    E={ "{3:ddd} {0:d2}/{1:d2}/{2:d4}  " -f [int]$_.MONTH, [int]$_.DAY, [int]$_.YEAR, 
    $(([datetime]"$($_.YEAR) $($_.MONTH) $($_.DAY)"))}}
    # $(([datetime]"$(YEAR) $(MONTH) $(DAY)"))}}

$calevent2= @{ 
    N = "Date_          ";
    E={ "{3:ddd} {0:d2}/{1:d2}/{2:d4}  " -f [int]$_.M, [int]$_.D, [int]$_.Y, 
    $(([datetime]"$($_.Y) $($_.M) $($_.D)"))}}
    # $(([datetime]"$(YEAR) $(MONTH) $(DAY)"))}}
    


$rolesarray = ("Elder", "Pioneer", "MS", "SMPW", "Cart", "LDC", "Old-Contacts", "No-FSG", "DI", "Inactive")
$fsgrouparray = ("ALSTON", "BOLDEN", "BROWN", "HOLMES", "KIRKLAND", "LEWIS", "PRINCE", "VANN", "MISC", "DI", "UPDATED") # 6/15/23 removed WILKINSON - added HOLMES & DI & UPDATED

$roles = 
@{
    N="Roles"
    E={
        $found = ""
        foreach ( $rol in $rolesarray)
        {
            if ($($_."Group Membership") -match $rol) {$found = $found + "$rol"[0]} # First letter
        }
        $found
    }
}

$groupoverseers = # 6/15/23 CREATED
@{
    N="OVER"
    E={
        $found = ""
        foreach ( $over in $fsgrouparray)
        {
            if ($($_."Group Membership") -match $over) {$found = $over}
        }
        $found
    }
}


$name = @{ N = 'Name'; # subsequent calls reverses the name again
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$last, $first"}}

$noswapnameP = @{ N = 'REGULAR PIONEERS  '; 
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$first, $last"}}

$noswapnameM = @{ N = 'MINISTERIAL SERV'; 
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$first, $last"}}  
    
$noswapnameE = @{ N = 'ELDERS'; 
    E = {$first, $last = $_."name" -split " ";
    $first = $first.replace(',',''); $last = $last.replace(',',''); "$first, $last"}}

$phone = @{ N = 'Phone';  E = {$_."phone 1 - value" + "  "}}
$addr = @{ N = 'Address';  E = {($_."address 1 - formatted").replace("`n"," ").replace("US","")}}
$email = @{ N = 'Email';  E = {$_."e-mail 1 - value"  + "  "}}

$notes = @{ N = 'Notes'; E = {$_."Notes"}}
# $notes = @{ N = 'Notes'; E = {($_."Notes").replace("`n"," ")}}

$fsgroup = @{ N = 'FSGroup';  E = {($_."address 2 - street").ToUpper()}}

#return true/false
$pioneer = @{ N = 'Pioneer';  E = {$_."Group Membership" -match "Pioneer"}}
$elder =  @{ N = 'Elder';  E = {$_."Group Membership" -match "Elders"}}
$ms = @{ N = 'MS';  E = {$_."Group Membership" -match "MS"}}
$old = @{ N = 'Old';  E = {$_."Group Membership" -match "Old-Contacts"}}
$nofsg = @{ N = 'No-FSG';  E = {$_."Group Membership" -match "No-FSG"}}

$updated = @{ N = 'UPDATED';  E = {if($_."Custom Field 1 - Type" -match "UPDATED"){$_."Custom Field 1 - Value"} }}

# $contacts | Add-Member -MemberType AliasProperty -name "UPDATED" -value "Custom Field 1 - Value" # 6/15/23 


$gridview = $contacts | 
Select-Object $name, $phone, $addr, $email, $fsgroup, $roles, $groupoverseers, $pioneer, $elder, $ms, $old, $nofsg # $notes
# use the 'N' of the calculated properties going forward when selecting objects from '$gridview'
# Name  Phone  Addesss  Email  Notes  FSGroup  Pioneer  Elders  MS  Old  No-FSG  Roles

$search = $contacts | 
Select-Object $name, $phone, $addr, $email, $roles, $groupoverseers, $notes, $pioneer, $elder, $ms # 6/15/23 removed $fsgroup 
# use the 'N' of the calculated properties going forward when selecting objects from '$search'
# Name  Phone  Addesss  Email  Notes  FSGroup  Pioneer  Elders  MS  Old  No-FSG  

$limited = $contacts | 
Select-Object $name, $phone, $addr, $email, $roles, $groupoverseers, $updated, $notes # 6/15/23 removed $fsgroup, $pioneer, $elder
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

# $contacts | Add-Member -MemberType AliasProperty -Name "GMEM" -Value "Group Membership"
# $contacts | gm

# $contacts.OutHashTable()


winget list "JW Library"
$contacts.ColsWithData()
$v = $PSVersionTable
$os = (Get-CimInstance Win32_OperatingSystem).Caption 
$plat = [System.Environment]::OSVersion.Platform
"Platform:$plat  OS:$os  PS:$($v.PSEdition) v$($v.PSVersion.ToString())  Columns:$($cols.Count):$($hset.Count)"; # $hset.Clear()



do
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
         ${[} U             :${]} View UPDATED contacts
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
        
        $sched | Select-Object $calevent, MESSAGE, $calevent2, MESS -first 25 | out-host | ft; wsl cal -A 2; read-Host -Prompt "`r"; Clear-Host; continue 
    }

    if ( $ans[0] -match 'v') { $gridview | Out-GridView -PassThru | Select-Object Name, Phone, Email, Address | Out-String ;  continue } # GRID

    if ( $ans[0] -match 'q') { exit } # QUIT
    
    # $choice = $search | Select-Object  Name, Phone, Address, Email, FSGroup, $roles, $notes

    if ( $ans[0] -eq "E") 
    { 
        DoExcel; continue 
    }

    if ( $ans[0] -eq "u") # 6/15/23  Show edit history of contacts if added to google contacts
    { 
        $limited | Where-Object { !([string]::IsNullOrEmpty($_."UPDATED")) } | ft
        continue 
    }
    
    if ( $ans[0] -eq "F") # search CSV for any part of a cell number / email address
    {
        # $ans = $ans.trim($ans[0])
        $ans = $ans.remove(0,1)
        # $title = @{ N='Cell'; E={$_."phone 1 - value" -replace '[()\s\s-\.]', ''} }  # strip all chars except 10 digit cell#
        
        $number = @{ N='Cell'; E={$_."phone 1 - value" -replace '[^0-9]' } }           # strip all chars except 10 digit cell#
        $emailaddress = @{ N='Email'; E={$_."e-mail 1 - value" } }     
        
        $cell = $contacts | Select-Object $name, $number, $emailaddress
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
        # $rolesarray = ("Elder", "Pion 
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
        $result = $limited | Where-Object { $_.OVER -match "\b$ans\b" }       # EXACT match  OLD and not BOLDEN ( OLD is inside of BOLDEN)
        if ( [string]::IsNullOrEmpty($result) )
            {
                # 6/15/23  $result = $search | Where-Object { $_.FSGroup -match "$ans" }  # not EXACT match
                $result = $limited | Where-Object { $_.OVER -match "$ans" }  # not EXACT match
            }    

        Write-Output("Total # in the $ans FSGroup: $($result.count)")
        Clear-Host
        $display = Read-Host "Table or List?: ( 'T' for Table 'L' for List )"
        if ( $display -match 'l' ) 
            { $result | Format-List }
        else
            { $result | Format-Table }

        ([array]$result).count
       
    }
    elseif ( $ans[0] -eq "%" ) # Roles
    {
        # $ans = $ans.trim("%").toupper()
        # $ans = $ans.remove(0,1)

        foreach ( $rol in $rolesarray )
        {  
            $result = $search | Where-Object { $_.Roles -match $rol[0]} # First letter
                $num = 0; if ( @($result).count -eq 0 ) { $num = 0} else { $num = ([array]$result).Count }
                Write-Output("Total # of ${rol}: $num")
            # DisplayPioneer
        }
        foreach ( $fsg in $fsgrouparray )
        {
            $result = $search | Where-Object { $_.OVER -match $fsg}
            $num = 0; if ( @($result).count -eq 0 ) { $num = 0} else { $num = ([array]$result).Count }
            Write-Output("Total # of $fsg-FSG: $num")
        }
        do
        {  
            $search | Where-Object { $_."Elder" -match $true } | 
                Format-Table -AutoSize -Property $noswapnameE, Email, Phone, Address 
            break
        }
        while ($true)
        do
        {  
            $search | Where-Object { $_."MS" -match $true } | 
                Format-Table -AutoSize -Property $noswapnameM, Email, Phone, Address 
            break
        }
        while ($true)
        do
        {  
            $search | Where-Object { $_."Pioneer" -match $true } | 
                Format-Table -AutoSize -Property $noswapnameP, Email, Phone, Address 
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
            Where-Object { $_."Name" -match $ans -or $_.Address -match $ans -or $_.OVER -match $ans} | Format-List
        }
        elseif ( $display -match 't' )
        {        
            $search | 
            Where-Object { $_."Name" -match $ans -or $_.Address -match $ans -or $_.OVER -match $ans} | Format-Table
        }
        else # 't' or 'l' not chosen
        {
            continue
            $result = $search | 
            Where-Object { $_."Name" -match $ans -or $_.Address -match $ans -or $_.OVER -match $ans}
            $result | Format-Table
            ([array]$result).Count
        }
    }

} while ($true)

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

