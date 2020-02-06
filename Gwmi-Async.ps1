#requires -Version 2



<#
.SYNOPSIS
Svendsen Tech's generic Get-WmiObject wrapper. Primarily designed for data
collection from a list of computers, with output to XML. There's a parser
designed for the XML as well.

Copyright (c) 2012, Svendsen Tech.
All rights reserved.
Author: Joakim Svendsen

.DESCRIPTION
The script also offers a WMI timeout parameter in the form of a time span object,
but be aware that this only affects successful queries. You might still
experience lengthy timeouts.

See the comprehensive online documentation and examples at:
http://www.powershelladmin.com/Get-wmiobject_wrapper

.PARAMETER ComputerName
Computer or computers to process. Use a file with "(gc computerfile.txt)".
.PARAMETER OutputFile
Output file name. NB! A full path is needed. You will be asked to overwrite it
if it exists unless -Clobber is specified. If the script detects you didn't provide
a full path, you will be asked to use the current working directory prepended
to the filename you provided.
.PARAMETER MultiClassProperty
Multiple WMI classes with the property or properties to extract. Designed for XML output.
A string in the format: "wmi_class1:prop1,prop2|wmi_class2:prop1|wmi_class3:prop1,prop2,prop3",
and so on.
.PARAMETER Timeout
The WMI timeout, represented as a time span object. Default "0:0:10" (10 seconds).
.PARAMETER NoPingTest
Specify this if you do NOT want to skip computers that do not respond to ICMP ping/echo.
.PARAMETER Clobber
Overwrite output file or files if they exist without prompting.
.PARAMETER Scope
Usually not necessary. By default "\\${Computer}\root\cimv2" will be used, while
this lets you replace the part "root\cimv2" with what you specify instead.
.PARAMETER CustomWql
The default WQL query is "SELECT prop1, prop2 FROM Win32_ClassHere", while this parameter
lets you append something like: WHERE DriveType="3". This "custom WQL" will be used in
all the queries, so if the property/condition doesn't apply to other classes, you will
see errors.
.PARAMETER Credential
Specify alternate credentials using a PSCredentials object (Get-Help Get-Credential).
.PARAMETER Domain
Specify the target computers' domain for use with alternative credentials specified with
-Credential. Both short and long forms should work. If you specify a domain in the credential
object passed to -Credential, you can skip specifying it again with this parameter (omit it),
since I built in a check for it.
#>

param(
    [Parameter(Mandatory=$true,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true,
               Position=0)]
        [Alias('PSComputerName', 'Cn', 'Hostname')]
        [string[]] $ComputerName,
    [Parameter(Mandatory=$true)][string] $OutputFile,
    [Parameter(Mandatory=$true)][string] $MultiClassProperty,
    [int32] $InvokeAsync = 32,
    [timespan] $Timeout = '0:0:30',
    [string] $Scope = 'root\cimv2',
    [string]   $CustomWql = '',
    [switch] $NoPingTest,
    [switch] $Clobber,
    [System.Management.Automation.Credential()] $Credential = [System.Management.Automation.PSCredential]::Empty,
    [string] $Domain = ''
    #[switch] $DoNotCloseRunspacePool
)

Set-StrictMode -Version Latest

## DEBUG
$ErrorActionPreference = 'Stop'

$StartTime = Get-Date
Write-Output "`nScript start time: $StartTime"

function Get-Result {
    
    [CmdletBinding()]
    param([switch] $Wait)
    do
    {
        $More = $false
        
        ## DEBUG
        # $Global:rst = $RunspaceTimers
        #'check $rst'
        
        foreach ($Runspace in $Runspaces) {
            # For some reason using "$RunspaceTimers.($Runspace.ID).Date failed. I spent ages on this...
            $StartTime = $RunspaceTimers[($Runspace.ID)]
            if ($Runspace.Handle.isCompleted) {
                Write-Verbose -Message ('Thread done for {0}' -f $Runspace.ComputerName)
                $Runspace.PowerShell.EndInvoke($Runspace.Handle)
                $Runspace.PowerShell.Dispose()
                $Runspace.PowerShell = $null
                $Runspace.Handle = $null
            }
            elseif ($Runspace.Handle -ne $null) {
                $More = $true
            }
            if ($Timeout -and $StartTime) {
                if ((New-TimeSpan -Start $StartTime) -ge $Timeout -and $Runspace.PowerShell) {
                    Write-Warning -Message ('Timeout {0}' -f $Runspace.ComputerName)
                    $Runspace.PowerShell.Dispose()
                    $Runspace.PowerShell = $null
                    $Runspace.Handle = $null
                    #Write-Verbose "Got here"
                    foreach ($Class in $ClassPropertyHash.Keys) {
                        Add-Member -Name $Class -Value $(New-Object PSObject) -MemberType NoteProperty -InputObject $Data.($Runspace.ComputerName) -Force
                        Add-Member -Name 'Error' -Value $(New-Object PSObject -Property @{
                            0 = "Timeout"
                        }) -MemberType NoteProperty -InputObject $Data.($Runspace.ComputerName).$Class # ... fixing in Gwmi-Wrapper-Report instead. # nope
                    }
                }
            }
        }
        if ($More -and $Wait)
        {
            Start-Sleep -Milliseconds 100
        }
        foreach ($Thread in $Runspaces.Clone()) {
            if ( -not $Thread.Handle) {
                #Write-Verbose -Message ('Removing {0} from runspaces' -f $Thread.ComputerName)
                $Runspaces.Remove($Thread)
            }
        }
        $ProgressSplatting = @{
            Activity = 'Running WMI queries'
            Status = '{0} of {1} total threads done' -f ($RunspaceCounter - $Runspaces.Count), $RunspaceCounter
            PercentComplete = ($RunspaceCounter - $Runspaces.Count) / $RunspaceCounter * 100
        }
        Write-Progress @ProgressSplatting
    }
    while ($More -and $Wait)
}

# Do some parameter checks
if ($Credential.Username -ne $null -and -not $Domain) {
    if ($Credential.Username.Contains('\')) {
        $Domain = $Credential.Username.Split('\')[0]
        Write-Warning -Message "-Domain not specified, but found one in supplied credentials. Using '${Domain}'. ."
    }
    else {
        Write-Error -Message "-Domain not specified and none found in supplied credentials. Exiting with code 2."
        exit 2
    }
}

if (-not $Clobber -and (Test-Path -Path $OutputFile -PathType Leaf)) {
    $Answer = Read-Host "XML output file, '$OutputFile', already exists and -Clobber not specified. Overwrite? (y/n) [yes]"
    if ($Answer -imatch 'n') {
        Write-Output 'Aborted. Exiting with code 3.'
        exit 3
    }
}

$ClassPropertyHash = @{}
    
# The strings should be in the form:
# "class1:prop1,prop2|class2:prop3,prop4,prop5|class3:prop6" - and so on...
$ClassesAndProperties = $MultiClassProperty -split '\s*\|\s*'
    
foreach ($ClassAndProperties in $ClassesAndProperties) {
    if ($ClassAndProperties -match '\A([^:]+?)\s*:\s*([^:]+)\z') {
        $ClassPropertyHash.($Matches[1]) = ,@($Matches[2] -split '\s*,\s*')
    }
    else {
        Write-Warning -Message "The following invalid class and property element was ignored: $ClassAndProperties"
    }
}

#$ClassPropertyHash = Get-ClassPropertyHash $MultiClassProperty
#"ClassPropertyHash:"
#$ClassPropertyHash.GetEnumerator()

# Check for duplicates and provide some feedback.
$ComputerNameCount = $ComputerName.Count
$UniqueComputerNameCount = @($ComputerName | Get-Unique -AsString).Count

$DuplicateCount = $ComputerNameCount - $UniqueComputerNameCount

if ($DuplicateCount -gt 0) {
    "Found $DuplicateCount duplicates. Removing duplicates."
    $ComputerName = $ComputerName | Get-Unique -AsString
}

# This is the script block that's run for each runspace thread.
$ScriptBlock = {
    
    [CmdletBinding()]
    param($ClassPropertyHash, $Computer, $CustomWql, $Timeout, $NoPingTest, $Scope, $MyCredential, $Domain, $RunspaceCounter)

    # Display progress on one line.
    #Write-Host -NoNewline -ForegroundColor Green "`rProcessing ${Computer}...                       "
    
    $RunspaceTimers.$RunspaceCounter = Get-Date # New-Object PSObject -Property @{ Date = Get-Date; Classes = @() }
    $Classes = $ClassPropertyHash.Keys
    $Data.$Computer = New-Object PSObject
    $PingDone = $false
    
    foreach ($Class in $Classes) {
        
        # Push these to the shared variable RunspaceTimers (now misnomered). This is in order to handle timeouts later on.
        #$RunspaceTimers.$RunspaceCounter.Classes += $Class

        # There's something odd going on. This shouldn't be necessary,
        # but apparently is... Better not have spaces in properties.
        $PropertiesString = (([string[]] $ClassPropertyHash.$Class) -split '\s+') -join ','
        $Properties = [string[]] ($PropertiesString -split ',')
        
        #'PropertiesString: "' + $PropertiesString + '"'
        
        # Create a new object for each class, as a note property, to hold the properties
        Add-Member -InputObject $Data.$Computer -MemberType NoteProperty -Name $Class -Value $(New-Object PSObject)
        
        ## DEBUG
        #[string[]]$ClassPropertyHash.$Class -join ','
        
        $Query = "SELECT $PropertiesString FROM $Class"
        
        if ($CustomWql) {
            $Query += " $CustomWql"
        }

        # Ping stuff. I had to put it in here or rewrite quite a bit, in order for
        # Gwmi-Wrapper-Report to work properly, and to get it easily parseable... Hmm.
        # At least I managed to avoid pinging more than once with two bools.
        if (-not $PingDone) {
            if (-not $NoPingTest -and -not (Test-Connection -ComputerName $Computer -Quiet -Count 1)) {
                Add-Member -Name 'NoPing' -Value $(New-Object PSObject -Property @{
                    0 = 'No ping reply'
                }) -MemberType NoteProperty -InputObject $Data.$Computer.$Class
                $PingReply = $false
                continue
            }
            else {
                $PingReply = $true
            }
            $PingDone = $true
        }
        else {
            if ($PingReply -eq $false) {
                Add-Member -Name 'NoPing' -Value @(New-Object PSObject -Property @{
                    0 = 'No ping reply'
                }) -MemberType NoteProperty -InputObject $Data.$Computer.$Class
                continue
            }
        }
        $Searcher = $null
        $ErrorActionPreference = 'Stop'
        try {
            
            $Searcher = [WmiSearcher] $Query
            # Not as useful as I had hoped...
            $Searcher.Options.Timeout = $Timeout
            $ConnectionOptions = New-Object Management.ConnectionOptions
            if ($MyCredential.Username -ne $null) {
                $ConnectionOptions.Authority = "ntlmdomain:${Domain}"
                $ConnectionOptions.Username = $MyCredential.Username.Split('\')[1]
                $ConnectionOptions.SecurePassword = $MyCredential.Password
            }
            $ManagementScope = New-Object Management.ManagementScope -ArgumentList "\\${Computer}\${Scope}", $ConnectionOptions
            $Searcher.Scope = $ManagementScope
            
            # Having to loop and extract the properties individually with -ExpandProperty
            # is ugly, but I simply can't find any other way while maintaining dynamic
            # parameters, properties and classes... I've been banging my head against
            # the wall for a while with this.
                        
            # I think all errors are terminating with $Searcher.Get()
            $SearcherResults = $Searcher.Get()
            $PropertyValues  = @()
            
            foreach ($Prop in $Properties) {
                
                # Preserve array structure for those with multiple values and enforce it
                # for one-element values.
                # This is what's supposed to work, but it's a bug: http://connect.microsoft.com/PowerShell/feedback/details/657211/select-object-expand-property-throws-and-exception-and-dies-when-property-is-null
                #$PropertyValues += ,@($SearcherResults | Select-Object -ExpandProperty $Prop)
                $PropertyValues += ,@($SearcherResults | ForEach-Object {
                    if (-not $_.$Prop) { '' }
                    else { $_.$Prop }
                    })
                
            }
            
            if ($Properties.Count -ne $PropertyValues.Count) {
                #Write-Host -ForegroundColor Red "Error: ${Computer}: ${Class}: Count of properties does not match count of retrieved values! This shouldn't happen. Skipping..."
                Add-Member -Name 'Error' -Value $(New-Object PSObject -Property @{
                    0 = "Error: ${Computer}: ${Class}: Count of properties does not match count of retrieved values!"
                }) -MemberType NoteProperty -InputObject $Data.$Computer.$Class
                continue
            }
            
            $PropertyCount = $Properties.Count
            
            for ( $i = 0; $i -lt $PropertyCount; $i++ ) {
                Add-Member -Name $Properties[$i] -Value $(New-Object PSObject) -MemberType NoteProperty -InputObject $Data.$Computer.$Class
                $ValueCount = @($PropertyValues[$i]).Count
                # Don't break the XML parsing if there are no values by inserting at least one empty value.
                if ($ValueCount -eq 0) {
                    Add-Member -Name 0 -Value 'EMPTY' -MemberType NoteProperty -InputObject $Data.$Computer.$Class.($Properties[$i])
                    continue
                }
                for ($k=0; $k -lt $ValueCount; ++$k) {
                    #'Adding ' + $i + ' (i) ' + $Value +  ' (value) ' + $Class + ' (class)'
                    $Value = @($PropertyValues[$i])[$k]
                    Add-Member -Name $k -Value $Value -MemberType NoteProperty -InputObject $Data.$Computer.$Class.($Properties[$i])
                }
            }
            
        } # end of try statement
        catch {
            #Write-Host -ForegroundColor Red "WMI error: $($Error[0].ToString())"
            Add-Member -Name 'Error' -Value $(New-Object PSObject -Property @{
                0 = $_.ToString() -replace '[\r\n]+'
            }) -MemberType NoteProperty -InputObject $Data.$Computer.$Class
        }
        $ErrorActionPreference = 'Continue'

    } # end of class foreach
    
} # end of script block

# Set up runspace stuff.
$SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$Data = [hashtable]::Synchronized(@{})
$SessionState.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry ‘Data’, $Data, ''))
$RunspaceTimers = [HashTable]::Synchronized(@{})
$SessionState.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry ‘RunspaceTimers’, $RunspaceTimers, ''))
$RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $InvokeAsync, $SessionState, $Host)
$RunspacePool.ApartmentState = "STA"
#$RunspacePool.ThreadOptions = "ReuseThread"
$RunspacePool.Open()

#$RunspacePool.SessionStateProxy.SetVariable("Data", $Data)

Write-Verbose -Message ('Starting runspaces... ' + (Get-Date))

$Runspaces = New-Object -TypeName System.Collections.ArrayList
[int] $RunspaceCounter = 0

foreach ($Computer in $ComputerName) {
    $RunspaceCounter++
    $PSScript = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
    [void] $PSScript.AddArgument($ClassPropertyHash)
    [void] $PSScript.AddArgument($Computer)
    [void] $PSScript.AddArgument($CustomWql)
    [void] $PSScript.AddArgument($Timeout)
    [void] $PSScript.AddArgument($NoPingTest)
    [void] $PSScript.AddArgument($Scope)
    [void] $PSScript.AddArgument($Credential)
    [void] $PSScript.AddArgument($Domain)
    [void] $PSScript.AddArgument($RunspaceCounter)
    [void] $PSScript.AddParameter('Verbose',$VerbosePreference)
    $PSScript.RunspacePool = $RunspacePool
    
    #[void] $Runspaces.Add((New-Object psobject -Property @{
    #    Object = $PSScript
    #    Result = $PSScript.BeginInvoke()
    #}))
    [void]$Runspaces.Add(@{
        Handle = $PSScript.BeginInvoke()
        PowerShell = $PSScript
        ComputerName = $Computer
        ID = $RunspaceCounter
    })
    Get-Result
} # end of computers foreach

Write-Verbose -Message ('Finished starting jobs. ' + (Get-Date))
Write-Verbose -Message ' ### Waiting for jobs to finish ### '
Get-Result -Wait
$RunspacePool.Close()
$RunspacePool.Dispose()
Remove-Variable RunspacePool

Write-Verbose -Message 'Exporting data to a global hash called $WmiData'
$Global:WmiData = $Data
Write-Verbose -Message ("Writing XML... " + (Get-Date))
$XmlString = "<computers>`n"

foreach ($Key in $Data.Keys | Sort) {
    $Computer = $Key
    $XmlString += "  <computer>`n    <name>$Computer</name>`n"
    foreach ($ClassName in $Data.$Computer | Get-Member -MemberType NoteProperty | Select -Exp Name) {
        $XmlString += "    <class>`n      <name>$ClassName</name>`n      <classentries>`n"
        #'Class count: ' + @($Data.$Computer.$ClassName | Get-Member -MemberType NoteProperty).Count
        foreach ($Property in $Data.$Computer.$ClassName | Get-Member -MemberType NoteProperty | Select -Exp Name) {
            #"Property: $Property"
            #$Data.$Computer.$ClassName.$Property | Out-Host
            $Count = @($Data.$Computer.$ClassName.$Property | Get-Member -MemberType NoteProperty).Count
            #'Property count: ' + $Count
            $XmlString += "        <classentry>`n          <property>$Property</property>`n          <values>`n"
            for ($i=0; $i -lt $Count; ++$i) {
                $Value = $Data.$Computer.$ClassName.$Property.$i
                $XmlString += "            <value>$Value</value>`n"
            }
            $XmlString += "          </values>`n        </classentry>`n"
        }
        $XmlString += "      </classentries>`n    </class>`n"
    }
    $XmlString += "  </computer>`n"
}

$XmlString += "</computers>`n"

#$Xml = [xml] $XmlString
#$Xml.Save($XmlFilename)

# Had some encoding issues, so I'm now using this to save it as UTF-8
# Had to play with this a bit to keep the line breaks preserved.
$XmlString = '<?xml version="1.0" encoding="utf-8"?>' + "`n" + $XmlString
$XmlString -split "`n" | Out-File -Width 10000000 -Encoding UTF8 -FilePath $OutputFile

if ($?) {
    "Successfully saved '$OutputFile'"
}
else {
    "Failed to save '$OutputFile': $($Error[0].ToString())"
}

@"
Script start time: $StartTime
Script end time:   $(Get-Date)

Exposed data hash as `$Global:WmiData.
Access it with "`$WmiData.GetEnumerator()" from the shell.

"@
