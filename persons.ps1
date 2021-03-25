Write-Information "Processing Persons"

#region Configuration
$config = ConvertFrom-Json $configuration;
$connectionString =  "DRIVER={Progress OpenEdge $($config.driver_version) driver};HOST=$($config.host_name);PORT=$($config.port);DB=$($config.database);UID=$($config.user);PWD=$($config.password);DIL=$($config.isolation_mode);AS=$($config.array_size);"

if($config.enableETWT) { $connectionString += "ETWT=1;" }
if($config.enableUWCT) { $connectionString += "UWCT=1;" }
if($config.enableKA) { $connectionString += "KA=1;" }
#endregion Configuration

#region Functions
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
#endregion Functions

#region Open VPN
if($config.enableVPN) {
    Write-Information "Opening VPN"
    #Ensure VPN Connection is closed
    &"$($config.vpnClosePath)" > $null 2>&1
    
    #Reopen VPN Connection
    &"$($config.vpnOpenPath)" > $null 2>&1
}
#endregion Open VPN

#region Execute
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
 Write-Information "$($students.count) Student Records";
 
$studentEntities = get_data_objects `
            -connectionString $connectionString `
            -query '
                        SELECT  DISTINCT
						  "STUDENT-ENTITY"."STUDENT-ID"
						, "STUDENT-ENTITY"."ENTITY-ID"
						, "STUDENT-ENTITY"."CALENDAR-ID"
						, "STUDENT-ENTITY"."SCHOOL-ID"
						, "STUDENT-ENTITY"."HOMEROOM-NUMBER"
						, "STUDENT-ENTITY"."STUDENT-PERCENT-ENROLLED"
						, "CALENDAR-MASTER"."SCHOOL-YEAR"
                        , "CALENDAR-MASTER"."TRACK"
                        , "CALENDAR-MASTER"."CAL-STR-DTE" "START_DATE"
                        , "CALENDAR-MASTER"."CAL-STP-DTE" "END_DATE"
                        , "CALENDAR-DESC"."CALENDAR-SDESC"
                        , "CALENDAR-DESC"."CALENDAR-LDESC"
                    FROM "PUB"."STUDENT-ENTITY"
                    INNER JOIN "PUB"."ENTITY" ON "ENTITY"."ENTITY-ID" = "STUDENT-ENTITY"."ENTITY-ID"
                    INNER JOIN "PUB"."CALENDAR-MASTER" ON "CALENDAR-MASTER"."ENTITY-ID" = "ENTITY"."ENTITY-ID" AND "CALENDAR-MASTER"."SCHOOL-YEAR" = "ENTITY"."SCHOOL-YEAR" AND "CALENDAR-MASTER"."CALENDAR-ID" = "STUDENT-ENTITY"."CALENDAR-ID"
                    INNER JOIN "PUB"."CALENDAR-DESC" ON "CALENDAR-DESC"."X-DEFAULT-CALENDAR" = 1 AND "CALENDAR-MASTER"."CALENDAR-ID" = "CALENDAR-DESC"."CALENDAR-ID" AND "CALENDAR-DESC"."ENTITY-ID" = "CALENDAR-MASTER"."ENTITY-ID"
                    WHERE "STUDENT-ENTITY"."STUDENT-STATUS" = ''A''';
  Write-Information "$($studentEntities.count) Student Entities Records";


foreach($student in $students)
{
    $person = @{};
    $person["ExternalId"] = $student.'NAME-ID';
    $person["DisplayName"] = "$($student.'FIRST-NAME') $($student.'LAST-NAME') ($($student.'NAME-ID'))"
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
			$contract["ExternalId"] = "$($student.'NAME-ID').$($entity.'ENTITY-ID').$($entity.'SCHOOL-YEAR').$($entity.'CALENDAR-ID')"
			$contract["Role"] = "Student"
			$contract["Type"] = "Student"
			
			foreach($prop in $entity.PSObject.properties)
			{
				if(@("RowError","RowState","Table","HasErrors","ItemArray") -contains $prop.Name) { continue; }
				$contract[$prop.Name.replace('-','_')] = "$($prop.Value)";
			}

			[void]$person.Contracts.Add($contract);
        }
    }
 
 
    Write-Output ($person | ConvertTo-Json -Depth 50);
}
#endregion Execute

#region Close VPN
if($config.enableVPN) {
    Write-Information "Closing VPN"
    &"$($config.vpnClosePath)" > $null 2>&1
}
#endregion Close VPN

Write-Information "Finished Processing Persons"
