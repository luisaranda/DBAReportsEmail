function Send-DBAEmail {
    [CmdletBinding()]
    param (
        [Parameter(Position=0,HelpMessage="Enter the name of a SQL server")]
        [ValidateNotNullorEmpty()]
        [Alias("ComputerName")]
        [string]$SqlServer=$env:ComputerName,

        [string]$Database='dbareports',

        [string]$Path='~\Documents'
    )
    
    begin {
        Import-Module -Name EnhancedHTML2
        Import-Module -Name SqlServer

    }
    
    process {

$style = @"
<style>
body {
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}
h1 {
    text-align:center;
}
h2 {
    border-top:1px solid #666666;
}

th {
    background-color: #4CAF50;
    color: white;
}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
.paginate_enabled_next, .paginate_enabled_previous {
    cursor:pointer; 
    border:1px solid #222222; 
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.paginate_disabled_previous, .paginate_disabled_next {
    color:#666666; 
    cursor:pointer;
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.dataTables_info { margin-bottom:4px; }
.sectionheader { cursor:pointer; }
.sectionheader:hover { color:red; }
.grid { width:100% }
.danger { background-color: red }
.warn { background-color: yellow }
.red {
    color:red;
    font-weight:bold;
} 
</style>
"@        
        $everything_ok = $true

        #Test connectivity to instance. Will stop script if this fails
        try {
            Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Query 'SELECT @@version' -ErrorAction Stop | Out-Null
        }
        catch {
            $everything_ok = $false
            Write-Error -Message "Failed connecting to instance $SqlServer and database $Database"
        }
        
        if($everything_ok){

            $filepath = Join-Path -Path $Path -ChildPath "DatabaseReport.html"

            $params = @{
                        'As'               = 'Table';
                        'PreContent'       = '<h2>&diams; SQL Agent Summary</h2>';
                        'EvenRowCssClass'  = 'even';
                        'OddRowCssClass'   = 'odd';
                        'MakeHiddenSection'= $false;
                        'TableCssClass'    = 'grid';
                        'Properties'       = 'Environment', 
                                             'ServerName',
                                             'NumberOfJobs',
                                             'SuccessfullJobs',
                                             'FailedJobs',
                                             'DisabledJobs',
                                             'UnknownJobs'
                       }

            $html_SAJS = Get-SQLAgentJobSummary -SqlServer $SqlServer -Database $Database | ConvertTo-EnhancedHTMLFragment @params

            $params = @{
                        'As'               = 'Table';
                        'PreContent'       = '<h2>&diams; Databases not backed up in the last 24 Hours</h2>';
                        'EvenRowCssClass'  = 'even';
                        'OddRowCssClass'   = 'odd';
                        'MakeHiddenSection'= $false;
                        'TableCssClass'    = 'grid';
                        'Properties'       = 'ServerName',
                                             'Name',
                                             'DaysSinceBackup',
                                             'LastBackupDate',
                                             'LastDifferentialBackupDate',
                                             'LastLogBackupDate'
                       }

            $html_SDMB = Get-SQLDatabaseMissingBackup -SqlServer $SqlServer -Database $Database | ConvertTo-EnhancedHTMLFragment @params

            $params = @{
                        'As'               = 'Table';
                        'PreContent'       = '<h2>&diams; Low Disk Space</h2>';
                        'EvenRowCssClass'  = 'even';
                        'OddRowCssClass'   = 'odd';
                        'MakeHiddenSection'= $false;
                        'TableCssClass'    = 'grid';
                        'Properties'       = 'ServerName',
                                             'DiskName',
                                             'Label',
                                             'Capacity',
                                             'FreeSpace',												
                                             @{n='Percentage';e={$_.Percentage}; css={if ($_.Percentage -lt 11) {'danger'} else {'warn'}}}

                       }

            $html_SLDS = Get-SQLLowDiskSpace -SqlServer $SqlServer -Database $Database | ConvertTo-EnhancedHTMLFragment @params

            $params = @{
                        'As'               = 'Table';
                        'PreContent'       = '<h2>&diams; Top 5 Fastest Growing Disks in Last 24 Hours</h2>';
                        'EvenRowCssClass'  = 'even';
                        'OddRowCssClass'   = 'odd';
                        'MakeHiddenSection'= $false;
                        'TableCssClass'    = 'grid';
                        'Properties'       = 'ServerName',
                                             'DiskName',
                                             'Label',
                                             'Capacity',
                                             'Growth',
                                             'FreeSpace',
                                             'Percentage'
                       }

            $html_SFGD = Get-SQLFastestGrowingDisks -SqlServer $SqlServer -Database $Database | ConvertTo-EnhancedHTMLFragment @params
            

            $params = @{
                        'CssStyleSheet'     = $style;
                        'Title'             = "Database Report for BCC domain";
                        'PreContent'        = "<h1>Database Report for BCC domain</h1>";
                        'HTMLFragments'     = @($html_SAJS,$html_SDMB,$html_SLDS,$html_SFGD);
                       }
            
            ConvertTo-EnhancedHTML @params | Out-File -FilePath $filepath

        }        
    }
    
    end {

    }
    
}

function Get-SQLAgentJobSummary {
    [CmdletBinding()]
    param (
        [string]$SqlServer=$env:ComputerName,
        [string]$Database='dbareports'

    )
    
    begin {
$Query = @"
SELECT AJS.DATE
	,IL.ServerName
	,IL.Environment
	,NumberOfJobs
	,SuccessfulJobs
	,FailedJobs
	,DisabledJobs
	,UnknownJobs
FROM Info.AgentJobServer AJS
INNER JOIN InstanceList IL ON AJS.InstanceID = IL.InstanceID
WHERE AJS.DATE > DATEADD(Day, - 1, GetDate())
"@      
    }
    
    process {
        $results = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Query $Query

        foreach($result in $results){
            $props = @{'Date'           = $result.DATE;
                       'ServerName'     = $result.ServerName;
                       'Environment'    = $result.Environment;
                       'NumberOfJobs'   = $result.NumberOfJobs;
                       'SuccessfulJobs' = $result.SuccessfulJobs;
                       'FailedJobs'     = $result.FailedJobs;
                       'DisabledJobs'   = $result.DisabledJobs;
                       'UnknownJobs'    = $result.UnknownJobs
            }

            New-Object -TypeName PSObject -Property $props
        }
    }
    
    end {
    }
}

function Get-SQLLowDiskSpace {
    [CmdletBinding()]
    param (
        [string]$SqlServer=$env:ComputerName,
        [string]$Database='dbareports'                
    )
    
    begin {
$Query = @"
SELECT [Date]
	,(
		SELECT ServerName
		FROM dbo.Instancelist
		WHERE InstanceID = [DiskSpace].ServerID
		) AS ServerName
	,[DiskName]
	,[Label]
	,[Capacity]
	,[FreeSpace]
	,[Percentage]
FROM [Info].[DiskSpace]
WHERE DATE > DATEADD(Day, - 1, GETDATE())
	AND Percentage <= 15
"@        
    }
    
    process {
        $results = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Query $Query

        foreach($result in $results){
            $props = @{'Date'           = $result.DATE;
                       'ServerName'     = $result.ServerName;
                       'DiskName'       = $result.DiskName;
                       'Label'          = $result.Label;
                       'Capacity'       = $result.Capacity;
                       'FreeSpace'      = $result.FreeSpace;
                       'Percentage'     = $result.Percentage
            }

            New-Object -TypeName PSObject -Property $props
        }        
    }
    
    end {
    }
}

function Get-SQLDatabaseMissingBackup {
    [CmdletBinding()]
    param (
        [string]$SqlServer=$env:ComputerName,
        [string]$Database='dbareports'                
    )
    
    begin {
$Query = @"
DECLARE @CurrDate DATETIME

SET @CurrDate = GETDATE()

SELECT il.[ServerName]
	,D.NAME
	,CASE 
		WHEN D.LastBackupDate = '0001-01-01 00:00:00.0000000'
			THEN NULL
		ELSE D.LastBackupDate
		END AS LastBackupDate
	,CASE 
		WHEN D.LastDifferentialBackupDate = '0001-01-01 00:00:00.0000000'
			THEN NULL
		ELSE D.LastDifferentialBackupDate
		END AS LastDifferentialBackupDate
	,CASE 
		WHEN D.LastLogBackupDate = '0001-01-01 00:00:00.0000000'
			THEN NULL
		ELSE D.LastLogBackupDate
		END AS LastLogBackupDate
	,(
		SELECT MAX(BackupDate)
		FROM (
			VALUES (D.LastBackupDate)
				,(D.LastDifferentialBackupDate)
			) AS VALUE(BackupDate)
		) AS MaxDate
	,DATEDIFF(DAY, (
			SELECT MAX(BackupDate)
			FROM (
				VALUES (D.LastBackupDate)
					,(D.LastDifferentialBackupDate)
				) AS VALUE(BackupDate)
			), @CurrDate) AS DaysSinceBackup
	,DATEDIFF(HOUR, (
			SELECT MAX(BackupDate)
			FROM (
				VALUES (D.LastBackupDate)
					,(D.LastDifferentialBackupDate)
				) AS VALUE(BackupDate)
			), @CurrDate) AS HoursinceBackup
	,CASE 
		WHEN (DATEDIFF(HOUR, D.LastBackupDate, @CurrDate) > 24)
			AND (DATEDIFF(HOUR, D.LastDifferentialBackupDate, @CurrDate) > 24)
			THEN 1
		ELSE 0
		END AS Olderthan24
FROM [dbo].[InstanceList] il
JOIN [Info].[Databases] D ON IL.InstanceID = D.InstanceID
JOIN [Info].[SQLInfo] SQL ON IL.InstanceID = SQL.InstanceID
WHERE il.Inactive = 0
	AND il.NotContactable = 0
	AND D.Inactive <> 1
	AND D.STATUS = 'Normal'
	AND SQL.SQLVersion <> 'SQL 2000'
	AND D.IsUpdateable = 1
	AND D.NAME != 'tempdb'
	AND (
		(DATEDIFF(HOUR, D.LastBackupDate, @CurrDate) > 24)
		AND (DATEDIFF(HOUR, D.LastDifferentialBackupDate, @CurrDate) > 24)
		)
ORDER BY DaysSinceBackup DESC
	,HoursinceBackup DESC
"@        
    }
    
    process {
        $results = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Query $Query

        foreach($result in $results){
            $props = @{'ServerName'                 = $result.ServerName;
                       'Name'                       = $result.Name;
                       'LastBackupDate'             = $result.LastBackupDate;
                       'LastDifferentialBackupDate' = $result.LastDifferentialBackupDate;
                       'LastLogBackupDate'          = $result.LastLogBackupDate;
                       'DaysSinceBackup'            = $result.DaysSinceBackup
            }

            New-Object -TypeName PSObject -Property $props
        }        
    }
    
    end {
    }
}

function Get-SQLFastestGrowingDisks {
    [CmdletBinding()]
    param (
        [string]$SqlServer=$env:ComputerName,
        [string]$Database='dbareports'                
    )
    
    begin {
$Query = @"
EXECUTE [dbo].[Get_FastestGrowingDisks]
"@        
    }
    
    process {
        $results = Invoke-Sqlcmd -ServerInstance $SqlServer -Database $Database -Query $Query

        foreach($result in $results){
            $props = @{'ServerName' = $result.Server;
                       'DiskName'   = $result.DiskName;
                       'Date'       = $result.Date;
                       'Label'      = $result.Label;
                       'Capacity'   = $result.Capacity;
                       'FreeSpace'  = $result.FreeSpace;
                       'Growth'     = $result.Growth;
                       'Percentage' = $result.Percentage
            }

            New-Object -TypeName PSObject -Property $props
        }        
    }
    
    end {
    }
}

Send-DBAEmail