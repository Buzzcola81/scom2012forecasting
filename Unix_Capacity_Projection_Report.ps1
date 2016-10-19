
######################################################################################################
#
# Script: Unix_Capacity_Projection_Report.ps1
#
# Author: Martin Sustaric
#
# Date: 14/07/2014
#
#
# About: This script extracts the list of Unix servers that are monitored buy the
#        SCOM Gateway Server (Using the Unix respource pool name). 
#
#
# Usage: Script to be setup as a scheduled task. The account used to execure the script
#        needs permission to the SCOM datawarehouse and access to the fileshare to save the report.
#        The folders/file locations need to be specified.
#
#
# Versions:
# 14/07/2014 - 1.0 - Inital release creation by Martin Sustaric
#
#
#####################################################################################################

$startDTM = (Get-Date)
write-host "Script started $startDTM"

$ScomMgtServer = "Server"
$tablecsvpath = "C:\TrendingReport\UnixRawDataTable.csv"                             #Raw data to be saved
$tableCapacity1 = "C:\TrendingReport\UnixRawDataCapacity.csv"                        #Processed Raw data to be saved
$ArchiveFolder = "C:\TrendingReport\Unix"                                            #Archive Folder for Reports
[string]$reportpath1="C:\TrendingReport\Unix_CapacityProvisioning_report.htm"        #Report Path for Capacity Provisioning
[string]$reportpath2="C:\TrendingReport\Unix_CapacityData_report.htm"                #Report Path for Raw Data
$ReportSaveFolder = "\\fileserver\SCOM Scheduled Folder"                             #Fileshare for report to be copied to
$Customer = "Customer/Company"                                                       #Customer/Company - not mandatory to update/enter
$GatewayServerUsed = "Server"                                                        #Gateway Server that is used to get the list of Windows servers for report 
$ResourcePoolUsed = "Unix/Linux Resource Pool name"                                  #Resource Pool name that is used to get the list of Unix servers for report




#rename file names with date and time
[string]$extdate = Get-Date -format 'yyMMddhhmmss'
[string]$extdate2 = (get-date).AddMonths(-1).ToString("MMMyyyy")
$tablecsvpathwithdate = $tablecsvpath.trimend('.csv') + "-" + "$extdate" + ".csv"
$tableCapacity1withdate = $tableCapacity1.trimend('.csv') + "-" + "$extdate" + ".csv"
$reportpath1 = $reportpath1.trimend("htm") 
$reportpath1 = $reportpath1.trimend(".")
$reportpath1 = $reportpath1 + "-" + "$extdate2" + ".htm"
$reportpath2 = $reportpath2.trimend("htm") 
$reportpath2 = $reportpath2.trimend(".")
$reportpath2 = $reportpath2 + "-" + "$extdate2" + ".htm"

$TD = 0

#Function to execute SQL query and return data
#Updated
function query-sql($sqlText, $database = "OperationsManagerDW", $server = "dcvicscomrpdb01.datacom.com.au")
{
    $connection = new-object System.Data.SqlClient.SqlConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database");
    $connection.Open();

    #Create a command object
    $cmd = $connection.CreateCommand()
    $cmd.CommandText = $sqlText

    #Execute the Command
    $Reader = $cmd.ExecuteReader()

    $Datatable = New-Object System.Data.DataTable
    $Datatable.Load($Reader)

    # Close the database connection
    $connection.Close()

    $datatable

}



#Function to get the raw data from SCOM Datawarehouse and save it into a csv
function fetch-data ()
{
    #Get the list of Windows Servers and Unix Servers from SCOM
    Import-Module OperationsManager
    New-SCOMManagementGroupConnection -ComputerName $ScomMgtServer
    
    $UnixMembers = Get-SCOMResourcePool  -DisplayName "$ResourcePoolUsed" | Get-SCXAgent | Select Name

    $sqlresults =$null
    #$UnixMembers

    $list = $UnixMembers
    $a = "("
    $b = $null
    $c = ")"

    foreach ($server in $list)
    {
        $svr = $server.name
        $server = "'$svr', "
        $b = $b + $server
    }

    $UnixServers = "$a" + "$b" + "$c"

    $UnixServers = $UnixServers -replace '\, \)',')'



    #Check if there are Unix servers to fetch Performance Data if so get data
    $i = 0
    $end = $UnixMembers.count

    $UnixQuery = @"
use OperationsManagerDW
SELECT
vPerf.DateTime,
vPerf.AverageValue,
vPerformanceRuleInstance.InstanceName,
vManagedEntity.Path,
vPerformanceRule.ObjectName,
vPerformanceRule.CounterName

FROM Perf.vPerfDaily AS vPerf with (NOLOCK)
INNER JOIN vPerformanceRuleInstance ON vPerformanceRuleInstance.PerformanceRuleInstanceRowId = vPerf.PerformanceRuleInstanceRowId 
INNER JOIN vManagedEntity ON vPerf.ManagedEntityRowId = vManagedEntity.ManagedEntityRowId 
INNER JOIN vPerformanceRule ON vPerformanceRuleInstance.RuleRowId = vPerformanceRule.RuleRowId

WHERE

vManagedEntity.Path in $UnixServers
AND (vPerformanceRule.ObjectName IN ('Processor', 'Memory', 'Logical Disk'))
AND (vPerformanceRule.CounterName IN ('% Processor Time', '% Available Memory', '% Free Space'))
AND vPerf.DateTime >= (DATEADD(m,-6,DATEADD(mm, DATEDIFF(m,0,GETDATE()), 0)))
AND vPerf.DateTime < (DATEADD(d,-1,DATEADD(mm, DATEDIFF(m,0,GETDATE()),0)) + '23:59:00.000')

--ORDER BY vPerf.DateTime DESC
"@

    $sqlresults = query-sql -sqlText $UnixQuery

    $steptime = get-date
    Write-host "Query completed - $steptime"



    #Export raw data
    $sqlresults | export-csv $tablecsvpathwithdate -notypeinformation
    
    $steptime = get-date
    Write-host "Raw Data saved - $steptime"
}



#Function to determin gradient of trend line given an array of values
function Get-Trendline 
{
<#  
  .SYNOPSIS   
    Calculate the linear trend line from a series of numbers
  .DESCRIPTION
    Assume you have an array of numbers
      $inarr = @(15,16,12,11,14,8,10)
    and the trendline is represented as
      y = a + bx
    where y is the element in the array
          x is the index of y in the array.
          a is the intercept of trendline at y axis
          b is the slope of the trendline

    Calling the function with
    PS> Get-Trendline -data $inarr
    will return an array of a and b.
  .PARAMETER data
   A one dimensional array containing the series of numbers
  .EXAMPLE  
   Get-Trendline -data @(15,16,12,11,14,8,10)
#> 

    param ($data)

    $n = $data.count

    $sumX=0
    $sumX2=0
    $sumXY=0
    $sumY=0

    for ($i=1; $i -le $n; $i++) 
    {
        $sumX+=$i
        $sumX2+=([Math]::Pow($i,2))
        $sumXY+=($i)*($data[$i-1])
        $sumY+=$data[$i-1]
    }

    $b = ($sumXY - $sumX*$sumY/$n)/($sumX2 - $sumX*$sumX/$n)
    $a = $sumY / $n - $b * ($sumX / $n)

    
    return @($a,$b)

}



#Function that processes the raw data into the raw data for reporting
function process-data ()
{
    #Import Data
    $importdata = Import-Csv -Path $tablecsvpathwithdate
    $steptime = get-date
    write-host "Import data - completed $steptime"
    #Update the blank Memory Instance fields with Memory
    $importdata | ForEach-Object ($_){if ($_.ObjectName -eq "Memory"){$_.InstanceName = "Memory"}}
    $steptime = get-date
    write-host "Import data - memory data cleaned $steptime"
    #List of servers
    $ServerList = $importdata | select path | Sort-Object Path -Unique
    $steptime = get-date
    write-host "Import data - serverlist gathered $steptime"

    #Set variables/arrays to empty (needed or will cause loop to fail)
    $ServerInstanceTable = $Null
    $Server=$Null
    $CapacityTable = @()
    [int32]$numberofinstances = 0


    ForEach ($Server in $ServerList) 
    {
        #Add to ServerInstanceList table
        #$ServerInstanceTable += $importdata | where {$_.Path -eq $Server.Path} | select Path, InstanceName | Sort-Object InstanceName -Unique
        #Store current Unique Instances
        $UniqueInstances = $importdata | where {$_.Path -eq $Server.Path} | select InstanceName | Sort-Object InstanceName -Unique

        foreach ($Instance in $UniqueInstances)
        {

        $numberofinstances = $numberofinstances + 1
        $steptime = get-date
        write-host "Trending data - $Server and $Instance - Instance number is $numberofinstances - $steptime"

            #Determin Average
            $CalcData = $importdata | where {(($_.Path -eq $Server.Path) -and ($_.InstanceName -eq $Instance.InstanceName))} | select DateTime, AverageValue, CounterName
            $Average = ($CalcData.AverageValue | Measure-Object -Average).average
            $CounterName = $CalcData | select CounterName | Sort-Object CounterName -Unique
            $CounterName = $CounterName.CounterName

            #Determin Projections
            $data = @()
            $data = Get-Trendline -data $CalcData.AverageValue
            [string]$rate = $data[1]

            [string]$A90Days = (90 * $rate) + $Average
            [string]$A180Days = (180 * $rate) + $Average
            [string]$A365Days = (365 * $rate) + $Average

            #Remove negative projected values and format results
            if ($A90Days -match '-'){$A90Days = '0'}
            if ($A180Days -match '-'){$A180Days = '0'}
            if ($A365Days -match '-'){$A365Days = '0'}


            #Calculate Days to Upgrade
            #Capacity provisioning days to upgrade (cleaning up data)
            $DaysToUpgrade = $Null

            if($rate -eq '0')
            {
                #No growth + or -
                $DaysToUpgrade = 'No Growth'
            }
            else
            {
        
                $DaysToUpgrade = ((90 - $Average)/$rate)
                $DaysToUpgrade = "{0:N2}" -f $DaysToUpgrade
                if (($Average -gt 90) -and ($DaysToUpgrade -match '-'))
                {
                    $DaysToUpgrade = '0'
                }
                if (($Average -lt 90) -and ($DaysToUpgrade -match '-'))
                {
                    $DaysToUpgrade = 'Negative Growth'
                }
            
            $Average = "{0:N2}" -f $Average
            [single]$A90Days  = "{0:N2}" -f $A90Days
            [single]$A180Days = "{0:N2}" -f $A180Days
            [single]$A365Days = "{0:N2}" -f $A365Days


            #Save results into table (Powershell object)
            $TargetProperties = @{Server=$Server.Path; InstanceName=$Instance.InstanceName; CounterName=$CounterName; AverageForPeriod=$Average; Projection90Days=$A90Days; Projection180Days=$A180Days; Projection365Days=$A365Days; DaysToUpgrade=$DaysToUpgrade}
            $TargetObject = New-Object PSObject –Property $TargetProperties
            $CapacityTable +=  $TargetObject
        
        }
        }

    }



    $steptime = get-date
    write-host "Analyize data - completed $steptime"

    #Write Utilization results to csv
    $CapacityTable | select Server, InstanceName, CounterName, AverageForPeriod, Projection90Days, Projection180Days, Projection365Days, DaysToUpgrade | Export-Csv -path $tableCapacity1withdate
    
    #Return processed table
    Return $CapacityTable
}



#Function to move file and rename if file exists
function MoveFileToFolder ([string]$source,[string]$destination)
{
    #Test if file exists and if so move it
    if (Test-Path $source)
    {
    
        $files = Get-Item -Path $source
 
        #verify if the list of source files is empty
        if ($files -ne $null) 
        {
     
            foreach ($file in $files) 
            {
                $filename = $file.Name
                $destinationfilename = "$destination" + "\" + "$filename"
 
                #verify if destination file exists and rename
                if (Test-Path $destinationfilename) 
                {
                    [string]$ext = Get-Date -format 'yyMMddhhmmss'
                    $NewFileName = "$ext" + "-" + "$filename" 
                    $NewDestination = "$ArchiveFolder" + "\" + $NewFileName
                    Move-Item -path $source -destination $NewDestination -ea silentlycontinue
                }
                else
                {
                    Move-Item $source $destination -ea silentlycontinue
                }
            }
        }
    }
    else
    {
        Write-host "File Does not exist"
    }
 }


 
#Function Archive Data
function archive-data ()
{
    #Archive Raw Data
    MoveFileToFolder -source $tablecsvpathwithdate -destination $ArchiveFolder
    MoveFileToFolder -source $tableCapacity1withdate -destination $ArchiveFolder
    
    #Copy report to fileshare
    Copy-Item -Path $reportpath1 -Destination $ReportSaveFolder
    Copy-Item -Path $reportpath2 -Destination $ReportSaveFolder

    #Archive Report
    MoveFileToFolder -source $reportpath1 -destination $ArchiveFolder
    MoveFileToFolder -source $reportpath2 -destination $ArchiveFolder
}



function report-all($CapacityData)
{
    [string]$datereport =  (get-date).AddMonths(-1).ToString("MMMM yyyy")

    ForEach ($_ in $CapacityData)
    {
        if ($_.daystoupgrade -notmatch "N")
        {
            $_.daystoupgrade = $_.daystoupgrade -as [decimal]
            [decimal]$_.daystoupgrade  = "{0:N2}" -f $_.daystoupgrade
        }
        
        $_.projection90days = $_.projection90days -as [decimal]
        [decimal]$_.projection90days  = "{0:N2}" -f $_.projection90days
        $_.projection180days = $_.projection180days -as [decimal]
        [decimal]$_.projection180days  = "{0:N2}" -f $_.projection180days
        $_.projection365days = $_.projection365days -as [decimal]
        [decimal]$_.projection365days  = "{0:N2}" -f $_.projection365days
        $_.averageforperiod = $_.averageforperiod -as [decimal]
        [decimal]$_.averageforperiod  = "{0:N2}" -f $_.averageforperiod
    }

    #Define reports

    #Report1 is the Capacity Provisioning 0-45 Day Lead-time, 90% Upgrade Point
    $Report1 = $CapacityData | select  Server, Instancename, Countername, Averageforperiod, daystoupgrade | where {$_.daystoupgrade -notmatch "N"}
    $Report1 = $Report1 | where {$_.daystoupgrade -lt '46'}
    $Report1 = $Report1 | Select @{Name="Server Name";Expression={$_."Server"}}, @{Name="Instance Name";Expression={$_."Instancename"}}, @{Name="Average for period '%'";Expression={$_."averageforperiod"}}, @{Name="Days to Upgrade";Expression={$_."daystoupgrade"}}
    [xml]$Report1html = $Report1 | Sort-Object daystoupgrade | ConvertTo-Html -fragment

    #Report2 is the Capacity Provisioning 46-200 Day Lead-time, 90% Upgrade Point
    $Report2 = $CapacityData | select  Server, Instancename, Countername, Averageforperiod, daystoupgrade | where {$_.daystoupgrade -notmatch "N"}
    $Report2 = $Report2 | where {(($_.daystoupgrade -ge '46') -and ($_.daystoupgrade -le '200'))}
    $Report2 = $Report2 | Select @{Name="Server Name";Expression={$_."Server"}}, @{Name="Instance Name";Expression={$_."Instancename"}}, @{Name="Average for period '%'";Expression={$_."averageforperiod"}}, @{Name="Days to Upgrade";Expression={$_."daystoupgrade"}}
    [xml]$Report2html = $Report2 | Sort-Object daystoupgrade | ConvertTo-Html -fragment

    #Report3 CPU Utilization Capacity Projection
    $Report3 = $CapacityData | select  Server, Countername, Averageforperiod, projection90days, projection180days, projection365days
    $Report3 = $Report3 | where {$_.countername -match 'Processor'} | Sort-Object Server 
    $Report3 = $Report3 | Select @{Name="Server Name";Expression={$_."Server"}}, @{Name="Counter";Expression={$_."countername"}}, @{Name="Average for period '%'";Expression={$_."averageforperiod"}}, @{Name="Projection 90 Days";Expression={$_."projection90days"}}, @{Name="Projection 180 Days";Expression={$_."projection180days"}}, @{Name="Projection 365 Days";Expression={$_."projection365days"}}
    [xml]$Report3html = $Report3 | ConvertTo-Html -fragment

    #Report4 Memory Utilization Capacity Projection
    $Report4 = $CapacityData | select  Server, Countername, Averageforperiod, projection90days, projection180days, projection365days
    $Report4 = $Report4 | where {$_.CounterName -match 'Memory'} | Sort-Object Server 
    $Report4 = $Report4 | Select @{Name="Server Name";Expression={$_."Server"}},  @{Name="Counter";Expression={$_."countername"}}, @{Name="Average for period '%'";Expression={$_."averageforperiod"}}, @{Name="Projection 90 Days";Expression={$_."projection90days"}}, @{Name="Projection 180 Days";Expression={$_."projection180days"}}, @{Name="Projection 365 Days";Expression={$_."projection365days"}}
    [xml]$Report4html = $Report4 | ConvertTo-Html -fragment

    #Report5 Disk Utilization Capacity Projection
    $Report5 = $CapacityData | select  Server, Instancename, Countername, Averageforperiod, projection90days, projection180days, projection365days
    $Report5 = $Report5 | where {$_.CounterName -match 'Space'} | Sort-Object Server 
    $Report5 = $Report5 | Select @{Name="Server Name";Expression={$_."Server"}},  @{Name="Counter";Expression={$_."countername"}}, @{Name="Instance Name";Expression={$_."Instancename"}}, @{Name="Average for period '%'";Expression={$_."averageforperiod"}}, @{Name="Projection 90 Days";Expression={$_."projection90days"}}, @{Name="Projection 180 Days";Expression={$_."projection180days"}}, @{Name="Projection 365 Days";Expression={$_."projection365days"}}
    [xml]$Report5html = $Report5 | ConvertTo-Html -fragment

    $endDTM = (Get-Date)

    #build HTML Body for report Capacity Provisioning
    $fragments1 += "<h1>Monthly Health Report</h1>"
    $fragments1 += "<h5>Server list used from '$Gatewayserverused' and '$ResourcePoolUsed'</h5>"
    $fragments1 += "<h2>Unix Report - $Customer for Month: $datereport </h2>"
    $fragments1 += "<h3>Capacity Provisioning 0-45 Day Lead-time, 90% Upgrade Point</h3>"
    $fragments1 += $Report1html.innerxml
    $fragments1 += "<h5><i>Note: Days to upgrade is referenced from the 1st of the current month.</i></h5>"
    $fragments1 += "<h5> </h5>"
    $fragments1 += "<h3>Capacity Provisioning 46-200 Day Lead-time, 90% Upgrade Point</h3>"
    $fragments1 += $Report2html.innerxml
    $fragments1 += "<h5><i>Note: Days to upgrade is referenced from the 1st of the current month.</i></h5>"
    $fragments1 += "<h5> </h5>"
    
    
    #build HTML Body for report Raw Data
    $fragments2 += "<h1>Monthly Health Report</h1>"
    $fragments2 += "<h5>Server list used from '$Gatewayserverused' and '$ResourcePoolUsed'</h5>"
    $fragments2 += "<h2>Unix Report - $Customer for Month: $datereport </h2>"
    $fragments2 += "<h3>CPU Utilization Capacity Projections</h3>"
    $fragments2 += $Report3html.innerxml
    $fragments2 += "<h5> </h5>"
    $fragments2 += "<h3>Memory Utilization Capacity Projections</h3>"
    $fragments2 += $Report4html.innerxml
    $fragments2 += "<h5> </h5>"
    $fragments2 += "<h3>Disk Utilization Capacity Projections</h3>"
    $fragments2 += $Report5html.innerxml
    $fragments2 += "<h5> </h5>"
    

    # Build the Header
    $head = "<style type=`"text/css`">
h1
{
font-family: `"Calibri`";
font-size: 30px;
font-weight:bold;
}
h2
{
font-family: `"Calibri`";
font-size: 20px;
font-weight:bold;
}
h3
{
font-family: `"Calibri`";
font-size: 18px;
font-weight:bold;
}
h4
{
font-family: `"Calibri`";
font-size: 16px;
font-weight:bold;
}
h5
{
font-family: `"Calibri`";
font-size: 14px;
font-weight:bold;
}
p
{
font-family: `"Calibri`";
font-size: 10px;
}
table {
width:100%;
}
table.internal {
}
th {
font: bold 11px `"Calibri`";
sans-serif;
color: #6D929B;
border-right: 1px solid #C1DAD7;
border-bottom: 1px solid #C1DAD7;
border-top: 1px solid #C1DAD7;
letter-spacing: 2px;
text-transform: uppercase;
text-align: left;
padding: 6px 6px 6px 12px;
background: #b1f098;
width: 18%;
}
th.nobg {
border-top: 0;
border-left: 0;
border-right: 1px solid #C1DAD7;
background: none;
}
th.spec {     
border-left: 1px solid #C1DAD7;
border-top: 0;
background: #fff url(images/bullet1.gif) no-repeat;
font: bold 10px `"Calibri`", Verdana, Arial, Helvetica,       sans-serif;
}

th.specalt {
border-left: 1px solid #C1DAD7;
border-top: 0;
background: #f5fafa url(images/bullet2.gif) no-repeat;
font: bold 10px `"Calibri`", Verdana, Arial, Helvetica,       sans-serif;
color: #B4AA9D;
}
td {
border-right: 1px solid #C1DAD7;
border-bottom: 1px solid #C1DAD7;
background: #fff;
padding: 6px 6px 6px 12px;
color: #797979;
}
td.alt {
background: #F5FAFA;
color: #797979;
}

td.Disabled {
background: #DDDDDD;
}

td.Online {
background: #FFFFFF;
}

caption {
padding: 0 0 5px 0;
width: 700px; 
font: italic 11px `"Calibri`", Verdana, Arial, Helvetica, sans-serif;
text-align: right;
}
.danger {background-color: red}.warn {background-color: yellow}</style>"

    #create the HTML document for Capacity Provisioning Report
    ConvertTo-HTML -Head $head -Body $fragments1 -PostContent "</br><i>Baseline: 6 months + report period. Report generated: $(Get-Date)</i>" | Out-File -FilePath $reportpath1 -Encoding ascii

    #create the HTML document for Raw Data Report
    ConvertTo-HTML -Head $head -Body $fragments2 -PostContent "</br><i>Baseline: 6 months + report period. Report generated: $(Get-Date)</i>" | Out-File -FilePath $reportpath2 -Encoding ascii

}




#Get the Performance Data
fetch-data

#Process the Raw data and get averages and projections and save results into file and variable
$CapacityTable = process-data

#Generate Reports in HTML
report-all $CapacityTable

#Archive data files and report and placy acopy of report on the fileshare
archive-data
