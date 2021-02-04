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
 
$students = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "STUDENT"."NAME-ID"
                                , "STUDENT"."OTHER-ID"
                                , "NAME"."ALTERNATE-ID"
                                , "NAME"."FIRST-NAME"
                                , "NAME"."MIDDLE-NAME"
                                , "NAME"."LAST-NAME"
                                , "NAME"."NALPHAKEY"
                                , "NAME"."PRIMARY-PHONE"
                                , "NAME"."SECOND-PHONE"
                                , CAST("NAME"."BIRTHDATE" as date) "BIRTHDATE"
                                , "NAME"."INTERNET-ADDRESS"
                                , "NAME"."INTERNET-ADDRESS-2"
                                , "NAME"."INTERNET-ADDRESS-3"
                                , "NAME"."INTERNET-ADDRESS-4"
                                , "STUDENT"."MN-EDE-NBR"
                                , "NAME-DUSER"."DUSER-ID"
                                , "STUDENT-ENTITY"."STUDENT-STATUS"
                                , "ENT-GRD-GY-XREF"."GRADE"
                                , "STUDENT"."GRAD-YR"
                                , "STUDENT"."NXT-GRAD-YR"
                                , "FS-CUST"."FS-CUST-KEY-PAD-NBR"
                        FROM "PUB"."STUDENT"
                        INNER JOIN "PUB"."NAME" ON "NAME"."NAME-ID" = "STUDENT"."NAME-ID"
                        INNER JOIN (SELECT * FROM "PUB"."STUDENT-ENTITY" WHERE "X-DEFAULT-ENTITY" = 1) "STUDENT-ENTITY" ON "STUDENT"."NAME-ID" = "STUDENT-ENTITY"."STUDENT-ID"
                        INNER JOIN "PUB"."ENTITY" ON "STUDENT-ENTITY"."ENTITY-ID" = "ENTITY"."ENTITY-ID"
                        LEFT JOIN "PUB"."ENT-GRD-GY-XREF" ON "ENTITY"."SCHOOL-YEAR" = "ENT-GRD-GY-XREF"."SCHOOL-YEAR" AND "STUDENT"."GRAD-YR" = "ENT-GRD-GY-XREF"."GRAD-YR"
                        LEFT JOIN "PUB"."FS-CUST" ON "STUDENT"."NAME-ID" = "FS-CUST"."NAME-ID"
                        LEFT JOIN "PUB"."NAME-DUSER" ON "NAME-DUSER"."NAME-ID" = "STUDENT"."NAME-ID"
                        WHERE "STUDENT-ENTITY"."STUDENT-STATUS" = ''A''';
 
 
$studentEntities = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT "STUDENT-ID"
                            , "ENTITY-ID"
                            , "CALENDAR-ID"
                            , "SCHOOL-ID"
                            , "HOMEROOM-NUMBER"
                            , "STUDENT-PERCENT-ENROLLED"
                    FROM "PUB"."STUDENT-ENTITY"
                    WHERE "STUDENT-ENTITY"."STUDENT-STATUS" = ''A''';
 
 
$calenderMaster = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "CALENDAR-MASTER"."SCHOOL-YEAR"
                                , "CALENDAR-MASTER"."ENTITY-ID"
                                , "CALENDAR-MASTER"."TRACK"
                                , "CALENDAR-MASTER"."CALENDAR-ID"
                                , "CALENDAR-MASTER"."CAL-STR-DTE"
                                , "CALENDAR-MASTER"."CAL-STP-DTE"
                                , "CALENDAR-DESC"."CALENDAR-SDESC"
                                , "CALENDAR-DESC"."CALENDAR-LDESC"
                        FROM "PUB"."CALENDAR-MASTER"
                        INNER JOIN "PUB"."CALENDAR-DESC" ON "CALENDAR-DESC"."X-DEFAULT-CALENDAR" = 1 AND "CALENDAR-MASTER"."CALENDAR-ID" = "CALENDAR-DESC"."CALENDAR-ID" AND "CALENDAR-DESC"."ENTITY-ID" = "CALENDAR-MASTER"."ENTITY-ID"
                        WHERE "CAL-STP-DTE" >= CURDATE()';

 
foreach($student in $students)
{
    $person = @{};
    $person["ExternalId"] = $student.'NAME-ID';
    $person["DisplayName"] = "$($student.'FIRST-NAME') $($student.'LAST-NAME')"
    $person["Role"] = "Student"
 
    foreach($prop in $student.PSObject.properties)
    {
        if(@("RowError","RowState","Table","HasErrors","ItemArray") -contains $prop.Name) { continue; }
        $person[$prop.Name.replace('-','_')] = "$($prop.Value)";
    }
 
    $person["Contracts"] = [System.Collections.ArrayList]@();
 
 
    foreach($entity in $studentEntities)
    {
        if($entity.'STUDENT-ID' -eq $student.'NAME-ID')
        {
            $contract = @{};
            $contract["ExternalId"] = "$($student.'NAME-ID').$($entity.'SCHOOL-ID')"
            $contract["Role"] = "Student"
            foreach($prop in $entity.PSObject.properties)
            {
                if(@("RowError","RowState","Table","HasErrors","ItemArray") -contains $prop.Name) { continue; }
                $contract[$prop.Name.replace('-','_')] = "$($prop.Value)";
            }
 
            foreach($calendar in $calenderMaster)
            {
                 
                if($entity.'CALENDAR-ID' -eq $calendar.'CALENDAR-ID')
                {
                    $contract["START_DATE"] = $calendar.'CAL-STR-DTE' | Get-Date -Format "MM/dd/yyyy";
                    $contract["END_DATE"] = $calendar.'CAL-STP-DTE' | Get-Date -Format "MM/dd/yyyy";
                    $contract["SCHOOL_YEAR"] = $calendar.'SCHOOL-YEAR';
                }
            }
 
            [void]$person.Contracts.Add($contract);
        }
    }
 
 
    Write-Output ($person | ConvertTo-Json -Depth 50);
}
