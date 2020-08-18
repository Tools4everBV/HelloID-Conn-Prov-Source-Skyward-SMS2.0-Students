## Settings ##
$configuration = @{
    driver_version = "11.7";
    host_name = "<HOST NAME OR IP";
    port = "<PORT>";
    database = "<DB NAME>";
    user = "SKYDBUSER";
    password = "SKYDBPASS";
    isolation_mode = "READ UNCOMMITED";
    array_size = "50";
    enableETWT = $true;
    enableUWCT = $false;
    enableKA = $true;
}
 
 
$connectionString =  "DRIVER={Progress OpenEdge $($configuration.driver_version) driver};HOST=$($configuration.host_name);PORT=$($configuration.port);DB=$($configuration.database);UID=$($configuration.user);PWD=$($configuration.password);DIL=$($configuration.isolation_mode);AS=$($configuration.array_size);"
 
if($configuration.enableETWT) { $connectionString += "ETWT=1;" }
if($configuration.enableUWCT) { $connectionString += "UWCT=1;" }
if($configuration.enableKA) { $connectionString += "KA=1;" }
     
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
 
 
$courses = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "COR-NUM-ID"
                                , "COR-ALPHAKEY"
                                , "COURSE-TYPE-ID"
                                , "DEPARTMENT-ID"
                                , "ENTITY-ID"
                                , "SCHOOL-YEAR"
                                , "COR-GRD-RNG-LOW"
                                , "COR-GRD-RNG-HIGH"
                                , "COR-LENGTH-SET-ID"
                                , "COR-STATUS"
                                , "COR-SDESC"
                                , "COR-LDESC"
                                , "SUBJECT-ID"
                                , CASE
                                          WHEN "SUBJECT-ID" IS NULL OR LENGTH("SUBJECT-ID") < 1 THEN NULL
                                                                  ELSE CAST("ENTITY-ID" as varchar(8)) + ''.'' + CAST("SCHOOL-YEAR" as varchar(8)) + ''.'' + CAST("SUBJECT-ID" as varchar(8))
                                                          END "SUBJECT-CALC-ID"
                        FROM "PUB"."COURSE"
                        WHERE "SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'')';
 
 
$sections = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "COR-NUM-ID"
                                , "TRACK"
                                , "CLAS-SECTION"
                                , "SCH-STR-TRM"
                                , "SCH-STP-TRM"
                                , "DSP-PERIOD"
                                , "ATN-PERIOD"
                                , "NAME-ID"
                                , "BUILDING-ID"
                                , "ENTITY-ID"
                                , "ROOM-NUMBER"
                                , "TCHR-PRIME-FLAG"
                                , "SCHOOL-YEAR"
                                , "COR-ALPHAKEY"
                                , CAST("COR-NUM-ID" as varchar(8)) + ''.'' + CAST("TRACK" as varchar(2)) + ''.'' + CAST("CLAS-SECTION" as varchar(8)) "SECTION-ID"
                        FROM "PUB"."CLASS-MEET"
                        WHERE "SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'')';
 
$enrollments = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT            "STUDENT-CLASS"."STUDENT-ID"
                                        , "STUDENT-CLASS"."COR-NUM-ID"
                                        , "CLASS-MEET"."DSP-PERIOD"
                                        , "STUDENT-CLASS"."CLAS-SECTION"
                                        , "CLASS-MEET"."ROOM-NUMBER"
                                        , "CLASS-MEET"."TRACK"
                                        , CASE WHEN "COURSE"."COR-LENGTH-SET-ID" = ''YR'' THEN CAST("CALENDAR"."CAL-STR-DTE" as varchar(16)) ELSE CAST("STUDENT-CLASS"."SCHD-STR-TRM" as varchar(16)) END "SCHD-STR-TRM"
                                        , CASE WHEN "COURSE"."COR-LENGTH-SET-ID" = ''YR'' THEN CAST("CALENDAR"."CAL-STP-DTE" as varchar(16)) ELSE CAST("STUDENT-CLASS"."SCHD-STP-TRM" as varchar(16)) END "SCHD-STR-TRM"
                                        , "STUDENT-CLASS"."SCHOOL-YEAR"
                                        , CASE
                                                WHEN "COURSE"."COR-LENGTH-SET-ID" = ''YR'' THEN CAST("STUDENT-CLASS"."ENTITY-ID" as varchar(4)) + ''.'' + CAST("STUDENT-CLASS"."SCHOOL-YEAR" as varchar(4))
                                                ELSE CAST("STUDENT-CLASS"."ENTITY-ID" as varchar(4)) + ''.'' + CAST("STUDENT-CLASS"."SCHOOL-YEAR" as varchar(4)) + ''.'' + CAST("TERM-NBR" as varchar(3)) + ''.'' + CAST("STUDENT-CLASS"."SCHD-STR-TRM" as varchar(2)) + ''.'' + CAST("STUDENT-CLASS"."SCHD-STP-TRM" as varchar(2))
                                          END ''TERM-ID''
                                        , "STUDENT-CLASS"."ENTITY-ID"
                                        , "COURSE"."COR-LENGTH-SET-ID"
                                        , "STUDENT-CLASS"."SCHD-STATUS"
                                        , "COURSE"."COR-STATUS"
                                        , CAST("CLASS-MEET"."COR-NUM-ID" as varchar(8)) + ''.'' + CAST("CLASS-MEET"."TRACK" as varchar(2)) + ''.'' + CAST("CLASS-MEET"."CLAS-SECTION" as varchar(8)) "SECTION-ID"
                                        , ''Student'' "Role"
                        FROM "PUB"."STUDENT-CLASS"
                        INNER JOIN "PUB"."STUDENT" ON "STUDENT-CLASS"."STUDENT-ID" = "STUDENT"."STUDENT-ID"
                        INNER JOIN "PUB"."CLASS-MEET" ON    "CLASS-MEET"."COR-NUM-ID" = "STUDENT-CLASS"."COR-NUM-ID"
                                                        AND "CLASS-MEET"."TRACK" = "STUDENT-CLASS"."TRACK"
                                                        AND "CLASS-MEET"."CLAS-SECTION" = "STUDENT-CLASS"."CLAS-SECTION"
                        INNER JOIN "PUB"."COURSE" ON "STUDENT-CLASS"."COR-NUM-ID" = "COURSE"."COR-NUM-ID"
                        LEFT JOIN (
                                      SELECT    "CALENDAR-MASTER"."SCHOOL-YEAR"
                                                                        , "CALENDAR-MASTER"."ENTITY-ID"
                                                                        , "CALENDAR-MASTER"."TRACK"
                                                                        , "CALENDAR-MASTER"."CALENDAR-ID"
                                                                        , "CALENDAR-MASTER"."CAL-STR-DTE"
                                                                        , "CALENDAR-MASTER"."CAL-STP-DTE"
                                                                        , "CALENDAR-DESC"."CALENDAR-SDESC"
                                                                        , "CALENDAR-DESC"."CALENDAR-LDESC"
                                                                FROM "PUB"."CALENDAR-MASTER"
                                                                INNER JOIN "PUB"."CALENDAR-DESC" ON "CALENDAR-DESC"."X-DEFAULT-CALENDAR" = 1 AND "CALENDAR-MASTER"."CALENDAR-ID" = "CALENDAR-DESC"."CALENDAR-ID" AND "CALENDAR-DESC"."ENTITY-ID" = "CALENDAR-MASTER"."ENTITY-ID"
                                                                WHERE "CALENDAR-MASTER"."CAL-STR-DTE" <= CURDATE()
                                                                AND "CALENDAR-MASTER"."CAL-STP-DTE" >= CURDATE()
                                  ) "CALENDAR" ON "CALENDAR"."SCHOOL-YEAR" = "STUDENT-CLASS"."SCHOOL-YEAR" AND "CALENDAR"."ENTITY-ID" = "STUDENT-CLASS"."ENTITY-ID" AND "CALENDAR"."TRACK" = "CLASS-MEET"."TRACK"
                        WHERE "STUDENT-CLASS"."SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'')
                        AND "STUDENT-CLASS"."SCHD-STATUS" = ''A'' 
                        AND "COURSE"."COR-STATUS" = ''A''';
 
$entities = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "ENTITY-ID"
                                , "ENTITY-NAME"
                                , "SCHOOL-YEAR"
                                , "ENTITY-STATUS"
                        FROM "PUB"."ENTITY"'
 
 
$schools = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "SCHOOL-ID"
                                , "SCHOOL-PRINCIPAL"
                                , "DISTRICT-CODE"
                                , "SCHOOL-NAME"
                                , "SCHOOL-BEG-GRADE"
                                , "SCHOOL-END-GRADE"
                                , "SCHOOL-PHONE"
                                , "SCHOOL-FAX"
                                , "SCHOOL-NUMBER"
                                , "SCHOOL-ORG-NUMBER"
                                , "BUILDING-ID"
                        FROM "PUB"."SCHOOL"'
 
$subjects = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT     "ENTITY-ID"
                                ,  "SCHOOL-YEAR"
                                ,  "SUBJECT-ID"
                                ,  "SUBJECT-SDESC"
                                ,  "SUBJECT-LDESC"
                                , CAST("ENTITY-ID" as varchar(8)) + ''.'' + CAST("SCHOOL-YEAR" as varchar(8)) + ''.'' + CAST("SUBJECT-ID" as varchar(8)) "SUBJECT-CALC-ID"
                        FROM "PUB"."SUBJECT"
                        WHERE "SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'')'
 
 
 
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
 
$terms = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT "ENTITY-ID" ,"SCHOOL-YEAR" ,CAST("ENTITY-ID" as varchar(64)) + ''.'' + CAST("SCHOOL-YEAR" as varchar(64)) + ''.'' + CAST("TERM-NBR" as varchar(3)) + ''.'' + CAST("SCHD-TRM-STR" as varchar(64)) + ''.'' + CAST("SCHD-TRM-STP" as varchar(64)) "Term-ID" ,CASE WHEN "TERM-NBR" = 0 THEN "SEMESTER-DESC" ELSE "TERM-DESC" END "TERM-NAME" ,"SEM-TRM-STR-DATE" ,"SEM-TRM-STP-DATE" ,CAST("ENTITY-ID" as varchar(64)) + ''.'' + CAST("SCHOOL-YEAR" as varchar(64)) "PARENT-TERM" FROM "PUB"."TERM-DEFINITION" WHERE "SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'') UNION SELECT DISTINCT "CLAS-CONTROL-SET"."ENTITY-ID" , "CLAS-CONTROL-SET"."SCHOOL-YEAR" , CAST("CLAS-CONTROL-SET"."ENTITY-ID" as varchar(64)) + ''.'' + CAST("CLAS-CONTROL-SET"."SCHOOL-YEAR" as varchar(64)) "Term-ID" , CAST("CLAS-CONTROL-SET"."SCHOOL-YEAR" as varchar(4)) + '' School Year'' "TERM-NAME" , CAST("CALENDAR"."CAL-STR-DTE" as date) "start" , CAST("CALENDAR"."CAL-STP-DTE" as date) "end" , CAST(null as varchar(4)) "PARENT-TERM" FROM "PUB"."CLAS-CONTROL-SET" LEFT JOIN( SELECT "CALENDAR-MASTER"."SCHOOL-YEAR" ,"CALENDAR-MASTER"."ENTITY-ID" ,"CALENDAR-MASTER"."TRACK" ,"CALENDAR-MASTER"."CALENDAR-ID" ,"CALENDAR-MASTER"."CAL-STR-DTE" ,"CALENDAR-MASTER"."CAL-STP-DTE" ,"CALENDAR-DESC"."CALENDAR-SDESC" ,"CALENDAR-DESC"."CALENDAR-LDESC" FROM "PUB"."CALENDAR-MASTER" INNER JOIN "PUB"."CALENDAR-DESC" ON "CALENDAR-DESC"."X-DEFAULT-CALENDAR" = 1 AND "CALENDAR-MASTER"."CALENDAR-ID" = "CALENDAR-DESC"."CALENDAR-ID" AND "CALENDAR-DESC"."ENTITY-ID" = "CALENDAR-MASTER"."ENTITY-ID") "CALENDAR" ON "CALENDAR"."SCHOOL-YEAR" = "CLAS-CONTROL-SET"."SCHOOL-YEAR" AND "CALENDAR"."ENTITY-ID" = "CLAS-CONTROL-SET"."ENTITY-ID" WHERE "CLAS-CONTROL-SET"."SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'') UNION SELECT DISTINCT "ENTITY-ID" ,"SCHOOL-YEAR" ,CAST("ENTITY-ID" AS VARCHAR(64)) + ''.'' + CAST("SCHOOL-YEAR" AS VARCHAR(64)) + ''.'' + CAST("TRACK" AS VARCHAR(3)) + ''.'' + CAST("CCS-SCH-STR-TRM" AS VARCHAR(64)) + ''.'' + CAST("CCS-SCH-STP-TRM" AS VARCHAR(64)) "TERM-ID" ,CAST("CCS-DESC" AS VARCHAR(64)) "TERM-NAME" ,CAST("CCS-ATND-STR-DTE" AS DATE) "start" ,CAST("CCS-ATND-STP-DTE" AS DATE) "end" ,CAST("ENTITY-ID" AS VARCHAR(64)) + ''.'' + CAST("SCHOOL-YEAR" AS VARCHAR(64)) "PARENT-TERM" FROM "PUB"."CLAS-CONTROL-SET" WHERE "SCHOOL-YEAR" >= TO_CHAR(CURDATE(),''yyyy'')'
 
 
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
 
 
    foreach($class in $enrollments)
    {
        if($class.'STUDENT-ID' -eq $student.'NAME-ID')
        {
            $contract = @{};
            $contract["ExternalId"] = "$($student.'NAME-ID').$($class.'SECTION-ID')"
            $contract["Role"] = "Student"
 
            foreach($prop in $class.PSObject.properties)
            {
                if(@("RowError","RowState","Table","HasErrors","ItemArray") -contains $prop.Name) { continue; }
                $contract[$prop.Name.replace('-','_')] = "$($prop.Value)";
            }
 
            foreach($term in $terms)
            {
                if($term.'Term-ID' -eq $class.'TERM-ID')
                {
                    $contract['START_DATE'] = $term.'SEM-TRM-STR-DATE' | Get-Date -Format "MM/dd/yyyy"
                    $contract['END_DATE'] = $term.'SEM-TRM-STP-DATE' | Get-Date -Format "MM/dd/yyyy";
                }
            }
 
            [void]$person.Contracts.Add($contract);
        }
    }
 
    Write-Output ($person | ConvertTo-Json -Depth 50);
}
