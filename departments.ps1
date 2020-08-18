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
 
$entities = get_data_objects `
            -connectionString $connectionString `
            -query 'SELECT    "ENTITY-ID"
                                , "ENTITY-NAME"
                                , "SCHOOL-YEAR"
                                , "ENTITY-STATUS"
                        FROM "PUB"."ENTITY"'
 
 
foreach($entity in $entities)
{
     $row = @{
              ExternalId = $entity.'ENTITY-ID';
              DisplayName = $entity.'ENTITY-NAME';
    }
 
    $row | ConvertTo-Json -Depth 10
}
