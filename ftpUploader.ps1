# Konfiguration
$localFolder = "D:\source"   # Quellordner
$ftpServer   = "ftp://ftp.com/In/"  # Ziel-FTP-Pfad (abschließender Slash!)

# Pfad zu den gespeicherten FTP-Anmeldedaten
$ftpCredPath = "C:\Users\wefe_adm\gt_creds\ftpcred.xml"

# Laden der FTP-Zugangsdaten aus der verschlüsselten Datei
if (Test-Path $ftpCredPath) {
    $ftpCredential = Import-Clixml $ftpCredPath
    $ftpUsername   = $ftpCredential.UserName
    $ftpPassword   = $ftpCredential.GetNetworkCredential().Password
} else {
    Write-Error "FTP Credential-Datei nicht gefunden: $ftpCredPath"
    exit
}

# Logdatei im Skriptordner
$logFile = Join-Path $PSScriptRoot "FtpUpload.log"

# Funktion zum Schreiben von Logeinträgen (mit Zeitstempel)
function Write-Log($message) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
}

Write-Log "Script gestartet. Quellordner: '$localFolder' | Ziel-FTP: '$ftpServer'"

# Funktion: Datei via FTP hochladen
function Upload-FileToFtp($filePath) {
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $ftpUri   = $ftpServer + $fileName
    Write-Log "Beginne Upload der Datei '$fileName' an '$ftpUri'."
    try {
        $ftpRequest = [System.Net.FtpWebRequest]::Create($ftpUri)
        $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($ftpUsername, $ftpPassword)
        $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
        $ftpRequest.UseBinary = $true
        $ftpRequest.KeepAlive = $false

        # Dateiinhalt lesen
        $fileContent = [System.IO.File]::ReadAllBytes($filePath)
        $ftpRequest.ContentLength = $fileContent.Length

        # Upload durchführen
        $requestStream = $ftpRequest.GetRequestStream()
        $requestStream.Write($fileContent, 0, $fileContent.Length)
        $requestStream.Close()

        $response = $ftpRequest.GetResponse()
        $response.Close()

        Write-Log "Datei '$fileName' erfolgreich hochgeladen."
    }
    catch {
        $errorMsg = "Fehler beim Upload der Datei '$fileName': $_"
        Write-Log $errorMsg
        throw $_
    }
}

# Archiv-Ordner festlegen und erstellen, falls nicht vorhanden
$archiveFolder = Join-Path $localFolder "Archiv"
if (-not (Test-Path $archiveFolder)) {
    try {
        New-Item -ItemType Directory -Path $archiveFolder | Out-Null
        Write-Log "Archiv-Ordner wurde erstellt: $archiveFolder"
    }
    catch {
        Write-Log "Fehler beim Erstellen des Archiv-Ordners: $_"
        throw $_
    }
} else {
    Write-Log "Archiv-Ordner existiert bereits: $archiveFolder"
}

# Hauptlogik: Verarbeite nur .pdf und .txt Dateien, die direkt im Quellordner liegen
try {
    # Erzeuge einen Suchpfad mit Wildcard, damit -Include funktioniert:
    $searchPath = Join-Path $localFolder "*"
    Write-Log "Suche nach Dateien in: $searchPath"
    $files = Get-ChildItem -Path $searchPath -Include *.pdf,*.txt -File
    $fileCount = $files.Count
    Write-Log "$fileCount Datei(en) gefunden, die verarbeitet werden sollen."
    
    foreach ($file in $files) {
        Write-Log "Verarbeite Datei: '$($file.Name)'."
        try {
            Upload-FileToFtp $file.FullName
            # Nach erfolgreichem Upload: Datei in den Archiv-Ordner verschieben
            Move-Item -Path $file.FullName -Destination $archiveFolder -Force
            Write-Log "Datei '$($file.Name)' wurde ins Archiv verschoben."
        }
        catch {
            Write-Log "Fehler bei der Verarbeitung von Datei '$($file.Name)': $_"
        }
    }
}
catch {
    Write-Log "Fehler in der Hauptlogik: $($_.Exception.Message)"
}

Write-Log "Script beendet."
