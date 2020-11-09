$config = ConvertFrom-Json $configuration;
$connectionString =  "DRIVER={Progress OpenEdge $($config.driver_version) driver};HOST=$($config.host_name);PORT=$($config.port);DB=$($config.database);UID=$($config.user);PWD=$($config.password);DIL=$($config.isolation_mode);AS=$($config.array_size);"

if($config.enableETWT) { $connectionString += "ETWT=1;" }
if($config.enableUWCT) { $connectionString += "UWCT=1;" }
if($config.enableKA) { $connectionString += "KA=1;" }
 
function get_data_objects {
[cmdletbinding()]
Param (
[string]$connectionString,
[string]$query
   )
    Process
    {
        $conn = (new-object System.Data.Odbc.OdbcConnection);
        $conn.connectionstring = $connectionString;
        $conn.open();
 
        $cmd = (New-object System.Data.Odbc.OdbcCommand($query,$conn));
        $dataSet = (New-Object System.Data.DataSet);
        $dataAdapter = (New-Object System.Data.Odbc.OdbcDataAdapter($cmd));
        $dataAdapter.Fill($dataSet) | Out-Null
        $conn.Close()
 
        $result = $dataset.Tables[0];
 
        @($result);
    }
}
 
$employees = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "HAAPRO-PROFILE"."NAME-ID"
						        , "HAAPRO-PROFILE"."HAAPRO-OTHER-ID"
								, "NAME"."ALTERNATE-ID"
						        , "NAME"."FIRST-NAME"
						        ,  "NAME"."MIDDLE-NAME"
						        , "NAME"."LAST-NAME"
						        , "NAME"."NALPHAKEY"
						        , "NAME"."PRIMARY-PHONE"
						        , "NAME"."SECOND-PHONE"
						        , CAST("NAME"."BIRTHDATE" as date) "BIRTHDATE"
						        , "NAME"."INTERNET-ADDRESS"
                                , "NAME"."INTERNET-ADDRESS-2"
                                , "NAME"."INTERNET-ADDRESS-3"
                                , "NAME"."INTERNET-ADDRESS-4"
						        , "NAME-DUSER"."DUSER-ID"
						        , "HAAPRO-PROFILE"."HAAPRO-ACTIVE"
						        , "HAAPRO-PROFILE"."HAAPRO-START-DTE"
						        , "HAAPRO-PROFILE"."HAAPRO-TERM-DTE"
						        , "HAAPRO-PROFILE"."HAAETY-EMP-TYPE-CODE"
						        , "HAAPRO-PROFILE"."HAABLD-BLD-CODE"
						        , "HAAPRO-PROFILE"."HPADCL-CHK-LOC-CODE"
						FROM "PUB"."HAAPRO-PROFILE"
						INNER JOIN "PUB"."NAME" ON "NAME"."NAME-ID" = "HAAPRO-PROFILE"."NAME-ID"
						LEFT JOIN "PUB"."NAME-DUSER" ON "NAME-DUSER"."NAME-ID" = "HAAPRO-PROFILE"."NAME-ID"';
 
 $assignments  = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT  DISTINCT
						          "HPMASN-ASSIGNMENTS"."HPMASN-ID"
						        , "HPMBRK-ASN-BRKDWN"."HPMBRK-ID"
						        , "HPMASN-ASSIGNMENTS"."NAME-ID"
						        , "HPMASN-ASSIGNMENTS"."HPADCP-PAY-CODE"
						        , "HPMASN-ASSIGNMENTS"."HAADSC-ID-ASN"
						        , "HPMASN-ASSIGNMENTS"."HAADSC-ID-POS"
						        , "HPMASN-ASSIGNMENTS"."HAADSC-ID-GRP"
						        , "HPMASN-ASSIGNMENTS"."HAACAM-DESC"
						        , "HPMASN-ASSIGNMENTS"."HAAPLC-DESC"
						        , "HPMASN-ASSIGNMENTS"."HAAMXM-DESC"
						        , "HPMASN-ASSIGNMENTS"."HAADSC-ID-JOB-TYPE"
						        , "HPMASN-ASSIGNMENTS"."HPMASN-FTE"
						        , "HPMASN-ASSIGNMENTS"."HPMASN-FTE-CALC"
						        , "HPMASN-ASSIGNMENTS"."HAABLD-BLD-CODE"
						        , "HPMASN-ASSIGNMENTS"."HPMASN-FIS-YEAR"
						        , TO_CHAR("HPMASN-ASSIGNMENTS"."HPMASN-CON-START-DATE",''yyyy-mm-dd'') "HPM-ASN-CON-START-DATE"
						        , TO_CHAR("HPMASN-ASSIGNMENTS"."HPMASN-CON-STOP-DATE",''yyyy-mm-dd'') "HPM-ASN-CON-STOP-DATE"
						        , TO_CHAR("HPMASN-ASSIGNMENTS"."HPMASN-START-DATE",''yyyy-mm-dd'') "HPMASN-START-DATE"
						        , TO_CHAR("HPMASN-ASSIGNMENTS"."HPMASN-END-DATE",''yyyy-mm-dd'') "HPMASN-END-DATE"
						        , "HPMBRK-ASN-BRKDWN"."HPMBRK-PCT"
						        , "HPMBRK-ASN-BRKDWN"."HAABLD-BLD-CODE"
						        , "HPMBRK-ASN-BRKDWN"."HAADSC-ID-ASN"
						FROM "PUB"."HAAPRO-PROFILE"
						INNER JOIN "PUB"."HPMASN-ASSIGNMENTS" ON "HPMASN-ASSIGNMENTS"."NAME-ID" = "HAAPRO-PROFILE"."NAME-ID"
						LEFT JOIN "PUB"."HPMBRK-ASN-BRKDWN" ON "HPMBRK-ASN-BRKDWN"."HPMASN-ID" = "HPMASN-ASSIGNMENTS"."HPMASN-ID"
						INNER JOIN "PUB"."HPMPLN-PLAN" ON "HPMPLN-PLAN"."HPMPLN-ID" = "HPMASN-ASSIGNMENTS"."HPMPLN-ID" AND "HPMPLN-SN-PLAN-X" = 0
						WHERE "HAAPRO-PROFILE"."HAAPRO-ACTIVE" = 1 
						AND "HPMPLN-PLAN"."HPMPLN-YEAR"  >= CAST(TO_CHAR(CURDATE(),''yyyy'') as int)-1'

$assignmentdescriptions = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "HAADSC-CODE"
						        , "HAADSC-ID"
						        , "HAADSC-DESC"
						FROM "PUB"."HAADSC-DESCS" 
						WHERE "HAADSC-IND" = ''ASSIG'''

$positions = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "HAADSC-CODE"
						        , "HAADSC-ID"
						        , "HAADSC-DESC"
						FROM "PUB"."HAADSC-DESCS" 
						WHERE "HAADSC-IND" = ''POSIT'''

$jobTypes = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "HAADSC-CODE"
						        , "HAADSC-ID"
						        , "HAADSC-DESC"
						FROM "PUB"."HAADSC-DESCS" 
						WHERE "HAADSC-IND" = ''JOBTP'''
 
 $buildingCodes = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT   "HAABLD-BLD-CODE"
						       , "HAABLD-SDESC"
						       , "HAABLD-DESC"
						       , "HAABLD-STATE-CODE"
						       , "HAABLD-SCHOOL-TYPE"
						FROM "PUB"."HAABLD-BLD-CODES"'

foreach($employee in $employees)
{
    $person = @{};
    $person["ExternalId"] = $employee.'NAME-ID';
    $person["DisplayName"] = "$($employee.'FIRST-NAME') $($employee.'LAST-NAME') ($($employee.'NAME-ID'))"
    $person["Role"] = "Employee"
 
    foreach($prop in $employee.PSObject.properties)
    {
        if(@("RowError","RowState","Table","HasErrors","ItemArray") -contains $prop.Name) { continue; }
        $person[$prop.Name.replace('-','_')] = "$($prop.Value)";
    }
 
    $person["Contracts"] = [System.Collections.ArrayList]@();
 
    #Employee Profile
    $contract = @{};
    $contract["ExternalId"] = "$($employee.'NAME-ID')"
    $contract["Role"] = "Employee"
    $contract["ContractType"] = "Profile"
    $contract["StartDate"] = $employee.'HAAPRO-START-DTE'
    $contract["EndDate"] = $employee.'HAAPRO-TERM-DTE'
    $contract["TitleId"] = "Employee"
    $contract["Title"] = "Employee"
    [void]$person.Contracts.Add($contract);

    #Assignments
    foreach($assign in $assignments)
    {
        if($assign.'NAME-ID' -ne $employee.'NAME-ID') { continue; }

        $contract = @{};
        $contract["ExternalId"] = "$($assign.'HPMBRK-ID')"
        $contract["Role"] = "Employee"
        $contract["ContractType"] = "Assignment"
        $contract["StartDate"] = $employee.'HAAPRO-START-DTE'
        $contract["EndDate"] = $employee.'HAAPRO-TERM-DTE'
        $contract["TitleId"] = "Employee"
        $contract["Title"] = "Employee"

        foreach($prop in $assign.PSObject.properties)
        {
            if(@("RowError","RowState","Table","HasErrors","ItemArray") -contains $prop.Name) { continue; }
            $contract[$prop.Name.replace('-','_')] = "$($prop.Value)";
        }

        foreach($assignDesc in $assignmentdescriptions)
        {
           if($assignDesc.'HAADSC-ID' -ne $assign.'HAADSC-ID-ASN') { continue; }
           $contract['AssignmentDescription'] = $assignDesc.'HAADSC-DESC';
           $contract["TitleId"] = $assign.'HAADSC-ID-ASN';
           $contract["Title"] = $assignDesc.'HAADSC-DESC';
           break;
        }

        foreach($jobType in $jobTypes)
        {
           if($jobType.'HAADSC-ID' -ne $assign.'HAADSC-ID-JOB-TYPE') { continue; }
           $contract['JobType'] = $jobType.'HAADSC-DESC'
           break;
        }
  
        [void]$person.Contracts.Add($contract);
    }

    Write-Output ($person | ConvertTo-Json -Depth 50);
}
