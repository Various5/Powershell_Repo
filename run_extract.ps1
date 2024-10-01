# filename: run_extract.ps1

# Database connection details
$dbhost = "localhost"
$dbport = "7618"
$username = "isl"
$password = "password"
$database = "isl"
$psqlpath = "C:\Program Files\ISL Conference Proxy\postgresql_bin\bin\psql.exe"

# Paths
$sqlFile = "C:\isllogs\query_sessions_isl.sql"
$outputCsv = "C:\isllogs\output.csv"
$existingCsv = "\\smbpath\output.csv"

# Set the PGPASSWORD environment variable
$env:PGPASSWORD = $password

# Run the SQL script and redirect output to CSV file
$psqlExitCode = & "$psqlpath" -h $dbhost -p $dbport -U $username -d $database -f "$sqlFile" > "$outputCsv" 2> "C:\isllogs\psql_errors.log"

# Clear the PGPASSWORD environment variable
$env:PGPASSWORD = $null

# Check for psql errors
if ($LASTEXITCODE -ne 0) {
    Write-Error "psql command failed with exit code $LASTEXITCODE. Check the log file at C:\isllogs\psql_errors.log"
    exit $LASTEXITCODE
}

# Check if the output CSV exists
if (Test-Path $outputCsv) {
    try {
        # Import the new data from output.csv
        $newData = Import-Csv -Path $outputCsv -Encoding UTF8

        # Initialize an array to store processed new data
        $processedData = @()

        foreach ($row in $newData) {
            # Get the 'messages' field
            $messages = $row.messages

            # Proceed only if 'messages' is not null or empty
            if ($messages -and $messages -ne 'No messages') {
                # Split the 'messages' string into key-value pairs
                $params = $messages -split '&'

                # Initialize a hashtable to store parameters
                $dict = @{}

                foreach ($param in $params) {
                    # Split each parameter into key and value
                    $kv = $param -split '=', 2
                    if ($kv.Length -eq 2) {
                        $key = $kv[0]
                        $value = $kv[1]
                        # Store the key-value pair in the hashtable
                        $dict[$key] = $value
                    }
                }

                # Extract the desired values from messages
                $memo = $dict['MSGDATA_Memo']
                $nameFirma = $dict['MSGDATA_NameFirma']
                $rechnungErstellt = $dict['MSGDATA_RechnungErstellt']

                # Set 'Verrechnet' to 'False' if it's empty
                if ([string]::IsNullOrEmpty($rechnungErstellt)) {
                    $rechnungErstellt = "False"
                }

                # URL-decode the values and replace '+' with spaces
                Add-Type -AssemblyName System.Web
                if ($memo) {
                    $memo = $memo -replace '\+', ' '
                    $memo = [System.Web.HttpUtility]::UrlDecode($memo)
                }
                if ($nameFirma) {
                    $nameFirma = $nameFirma -replace '\+', ' '
                    $nameFirma = [System.Web.HttpUtility]::UrlDecode($nameFirma)
                }

                # Create a unique key for each record
                $uniqueKey = $row.d_username + '_' + $row.created_time_readable + '_' + $nameFirma

                # Create a new object with the desired properties, including UniqueKey
                $newRow = [PSCustomObject]@{
                    'UniqueKey'    = $uniqueKey   # Use as unique identifier
                    'Startzeit'    = $row.created_time_readable
                    'Dauer'        = $row.duration_readable
                    'Benutzer'     = $row.d_username
                    'Firma_Name'   = $nameFirma
                    'Verrechnet'   = $rechnungErstellt
                    'Memo'         = $memo
                }

                # Add the new row to the processed data array
                $processedData += $newRow
            }
        }

        # Import existing data if the file exists
        $existingData = @()
        if (Test-Path $existingCsv) {
            $existingData = Import-Csv -Path $existingCsv -Encoding UTF8
        }

        # Combine existing data and new processed data
        $combinedData = $existingData + $processedData

        # Remove duplicates based on 'UniqueKey'
        $uniqueData = $combinedData | Sort-Object -Property UniqueKey -Unique

        # Export the unique data back to the CSV file
        $uniqueData | Select-Object -Property Startzeit,Dauer,Benutzer,Firma_Name,Verrechnet,Memo | Export-Csv -Path $existingCsv -NoTypeInformation -Encoding UTF8

    } catch {
        Write-Error "An error occurred during processing: $_"
    }
} else {
    Write-Error "Output CSV file not found at $outputCsv"
}

