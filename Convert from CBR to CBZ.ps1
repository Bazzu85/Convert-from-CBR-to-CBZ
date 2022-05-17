Function CheckDirectories (){
    if (!(Test-Path -LiteralPath "$PSScriptRoot\configuration")){
        New-Item -Path ("$PSScriptRoot\configuration") -ItemType Directory -Force
    }
    if (!(Test-Path -LiteralPath "$PSScriptRoot\logs")){
        New-Item -Path ("$PSScriptRoot\logs") -ItemType Directory -Force
    }
}
Function LogWrite ($logString){
    $date = Get-Date -Format "yyyyMMdd HHmm"
    $logFile = "$PSScriptRoot\logs\$date.log"
    $logstring = "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss") $logstring"
    Write-Host $logstring
    Add-content $logFile -value $logstring
}
Function Get-ConfigurationFromJson {
    $configurationFile = "$PSScriptRoot\configuration\configuration.json"

    #Default object
    $configuration = [ordered]@{
        testing = $true;
        testingPath = "$PSScriptRoot\Test";
        zipPath = "$env:ProgramFiles\7-Zip\7z.exe";
        deleteOriginalFile = $false;
        removeComicInfoXml = $false;
        extensionList=@("cbr","cb7");
    }
    # If not found, create the .config file
    if (!(Test-Path -Path $configurationFile)){
        ConvertTo-Json -InputObject $configuration | Out-File $configurationFile
        exit 
    }
    #region Read it and build the $globalOptions object
    $configuration = Get-Content -LiteralPath $configurationFile | ConvertFrom-Json
    LogWrite "***** Configuration from File *****"
    LogWrite "* Testing flag: $($configuration.testing)"
    LogWrite "* Testing path: $($configuration.testingPath)"
    LogWrite "* 7zip path: $($configuration.zipPath)"
    LogWrite "* Delete original File flag: $($configuration.deleteOriginalFile)"
    LogWrite "* Remove ComicInfo.xml: $($configuration.removeComicInfoXml)"
    $composedString = ""
    foreach ($extension in $configuration.extensionList){
        if ($composedString -eq ""){
            $composedString = "$composedString$extension"
        } else {
            $composedString = "$composedString,$extension"
        }
    }
    LogWrite "* Extension to work on: $($composedString)"
    LogWrite "***********************************"
    Write-Host ""
    Write-Host "If ok press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    return $configuration
}

# check if the directories exists or create it
CheckDirectories

# get the configuration from json
$configuration = Get-ConfigurationFromJson

# work on 7zip path and check if exists
if (-not (Test-Path -Path $configuration.zipPath -PathType Leaf)) {
    LogWrite "7 zip file '$($configuration.zipPath)' not found"
    LogWrite -NoNewLine 'Press any key to continue...';
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    exit
}
Set-Alias 7zip $configuration.zipPath

# get the working path
$done = $false
if ($configuration.testing -eq $false){
    while(!$done){
        $path = Read-Host "Please digit a folder to convert CBR to CBZ recursively. Pay attention to special chars like – that are transformed by powershell"

        if (Test-Path -LiteralPath $($path)){
            $done = $true
        } else {
            LogWrite "Folder $path not valid. Retry"
        }
    }
} else {
    $path = $configuration.testingPath
}

# work on cbr file on the path provided
#$files = Get-ChildItem -LiteralPath $path -Include "*.cbr,*.cb7" -Recurse -Force
$files = Get-ChildItem -LiteralPath $path -Recurse -Force
$count = 0
foreach($file in $files){
    LogWrite "Working on $file"
    Write-Host $file.Extension
    # check if extension is allowed
    $extensionOk = $false
    foreach ($extension in $configuration.extensionList){
        $workExtension = ".$extension"
        if ($file.Extension.ToLower() -eq $workExtension.ToLower()){
            $extensionOk = $true
            break
        }
    }
    if (!$extensionOk){
        continue
    }

    $GUID = New-Guid
    $tempDir = "$($file.Directory)\$GUID"
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $cbzFile = "$($file.Directory)\$($fileNameWithoutExtension).cbz"
    if (!(Test-Path -LiteralPath $tempDir)){
        New-Item -Path $tempDir -ItemType Directory -Force
    }

    7zip x -o"$tempDir" $file -r
    if ($LASTEXITCODE -ne 0){
        LogWrite "Error: extract exit code: $LASTEXITCODE"
    }
    LogWrite "Extracted content to $tempDir"

    if ($LASTEXITCODE -eq 0){
        if ((Test-Path -LiteralPath "$cbzFile" -PathType Leaf)){
            Remove-Item -LiteralPath $cbzFile -Confirm:$false -Force -Recurse | Out-Null
        }

        if ($configuration.removeComicInfoXml){
            7zip a $cbzFile "$tempDir\*" -xr!"ComicInfo.xml"
        } else {
            7zip a $cbzFile "$tempDir\*"
        }
        
        if ($LASTEXITCODE -ne 0){
            LogWrite "Error: CBZ creation exit code: $LASTEXITCODE"
        } else {
            LogWrite "Created CBZ file $cbzFile"

            $count++
            if ($configuration.deleteOriginalFile){
                Remove-Item -LiteralPath $file -Confirm:$false -Force -Recurse | Out-Null
            }
        }
    }
    Remove-Item -LiteralPath $tempDir -Confirm:$false -Force -Recurse | Out-Null
}
LogWrite "All done. Number of file converted: $count"
LogWrite "Press any key to continue...";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

