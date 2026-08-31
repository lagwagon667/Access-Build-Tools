foreach ($class in Get-Item -Path "$PSScriptRoot\Classes\*.ps1") {
    . $class
}

function Build-Accdb {
    <#
.SYNOPSIS
Builds an Access database (ACCDB) from source files exported via the
ms-access-vcs add-in.

.DESCRIPTION
Builds an Access database (ACCDB) from a source folder exported through the
ms-access-vcs add-in. The cmdlet reconstructs the database structure, applies
optional versioning, and initializes database connections to prevent ODBC login
dialogs during the build.

When versioning parameters are provided, the cmdlet uses GitVersion to generate
the current version and injects it into the database. Any supplied database
connections are established before the build so that no ODBC authentication
prompts interrupt or block the build process.

The resulting ACCDB can be written to a custom output path, and existing files
may be overwritten when the -Force switch is used.

.PARAMETER SourceFolder
Path to the folder containing the exported Access source files. This parameter
is mandatory.

.PARAMETER VersionFile
Optional path to the file into which the generated GitVersion version number
will be injected during the build process.

.PARAMETER VersionPattern
Optional regular expression pattern that identifies the portion of the
VersionFile to be replaced with the VersionReplacementPattern during version
injection.

.PARAMETER VersionReplacementPattern
Optional string used to generate the final replacement text for version
injection. The pattern may contain the placeholders $version and $date, which
are substituted with the actual GitVersion-generated version number and the
commit date. The resulting string is then used together with VersionPattern to
replace all matches in the VersionFile

.PARAMETER Connections
An array of database-connection descriptor objects used to initialize and cache
connections to prevent ODBC authentication prompts that block the build process.
Each object provides a DSN, username, and password. 

.PARAMETER VcsInstallPath
Specifies an optional custom installation path for the VCS add‑in. When not
provided, the module attempts to locate the add‑in in the standard default
location.

.PARAMETER Output
Optional path where the built ACCDB should be written. If omitted, the database 
name as specified in dbs-properties.json is used.

.PARAMETER Force
Allows overwriting an existing output file.

.EXAMPLE
Build-Accdb -SourceFolder "my.accdb.src" 
Builds the database from the source folder and writes the result to the
output path specified in my.accdb.src/dbs-properties.json.

.EXAMPLE
Build-Accdb -SourceFolder .\my.accdb.src\ -Force  -Connections @((New-Connection -DSN "mydb_dsn" -User "username" -Password "secret"))  -VersionFile .\my.accdb.src\forms\frmMain.form -VersionPattern 'Caption\s*=\s*"Version A\.B\.C"' -VersionReplacementPattern 'Caption ="Version $version ($date)"'
Builds the database, caches a database connection and injects a version number extracted from version.txt.

.NOTES
This cmdlet is part of the AccessBuildTools and is typically used as the
first step in the build pipeline before running unit tests, publishing, or
compiling to ACCDE.
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolder,
        [string]$VersionFile,
        [string]$VersionPattern,
        [string]$VersionReplacementPattern,
        [object[]]$Connections,
        [string]$VcsInstallPath,
        [string]$Output,
        [switch]$Force)

    try {
        $start = Get-Date
        $conns = [object[]]$Connections
        $SourceFolder = [System.IO.Path]::GetFullPath($SourceFolder, (Get-Location))
        if (-not  (Test-Path $SourceFolder)) {
            throw "Source folder $SourceFolder not found. Aborting..."    
        }
        if ($VersionFile) {
            $VersionFile = [System.IO.Path]::GetFullPath($VersionFile, (Get-Location))
            if (-not  (Test-Path $VersionFile)) {
                throw "Version file $VersionFile not found. Aborting..."    
            }
        }
    
        Write-Progress -Activity "Build-Accdb" -Status "Checking prerequisites..." -PercentComplete 10
        Write-Debug "Building from source files located in $SourceFolder"
    
        if (-not $Output) {
            $json = Get-Content (Join-Path $SourceFolder "dbs-properties.json") -Raw | ConvertFrom-Json
            $rawName = $json.Items.Name.Value
    
            $Output = Join-Path (Split-Path $sourceFolder -Parent) $rawName.Split(':')[-1]
        }
        $Output = [System.IO.Path]::GetFullPath($Output, (Get-Location))
        Write-Debug "Set output to $Output"

        if (Test-Path $Output) {
            if ($Force) {
                $filename = [System.IO.Path]::GetFileNameWithoutExtension($Output)
                $extension = [System.IO.Path]::GetExtension($output)
                $backup = "${filename}_backup$extension"
                $i = 1

                $outputFolder = [System.IO.Path]::GetDirectoryName($Output)

                while (Test-Path (Join-Path $outputFolder $backup)) {
                    $filename = [System.IO.Path]::GetFileNameWithoutExtension($Output)
                    $extension = [System.IO.Path]::GetExtension($output)
                    $backup = "${filename}_backup_$i$extension"
                    $i++
                }
                Rename-Item $Output $backup
                Write-Debug "Created backup file $(Join-Path $outputFolder $backup)"
            }
            else {
                throw "File '$Output' already exists. Use -Force to overwrite."
            }
        }

        $vcsApi = Get-VcsApi -VcsInstallPath $VcsInstallPath
            
        if ($VersionFile -and $VersionPattern -and $VersionReplacementPattern) {
            Write-Progress -Activity "Build-Accdb" -Status "Determining version..." -PercentComplete 20
            $versionJson = Get-Version
            $version = $versionJson.SemVer
            $date = ([datetime]::Parse($versionJson.CommitDate)).ToString("dd.MM.yyyy")
            Write-Debug "Setting version to $version"

            $VersionReplacementPattern = $VersionReplacementPattern -replace '\$version', $version
            $VersionReplacementPattern = $VersionReplacementPattern -replace '\$date', $date
    
            $script:count = 0
            $content = Get-Content $VersionFile -Raw

            $evaluator = {
                param($match)
                $script:count++
                return $match.Result($VersionReplacementPattern)
            }

            $contentWithVersion = [regex]::Replace($content, $VersionPattern, $evaluator)
            Set-Content $VersionFile $contentWithVersion

            if ($script:count -eq 0) {
                Write-Warning "No version strings were replaced in $VersionFile"
            }
        }
    
        $accessApp = Open-AccessWithoutStartupCommands -AccessDBPath $Output
            
        Write-Progress -Activity "Build-Accdb" -Status "Caching DB Connections..." -PercentComplete 30
        CacheDBConnections -AccessApp $accessApp -Connections $conns
        
        Write-Debug "Building database $output"
        Write-Progress -Activity "Build-Accdb" -Status "Building..." -PercentComplete 100

        $null = $accessApp.Run($vcsApi, "BuildHeadless", [ref]"$sourceFolder")

        $elapsed = (Get-Date) - $start
        $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed
        Write-Host "Database $output built successfully. Duration: $elapsedStr"
    }
    finally {
        if ($accessApp) {
            Close-AccessInstance -AccessApp $accessApp
        }
    }
}

function IsDebugMode {
    return $DebugPreference -eq 'Continue' -or $DebugPreference -eq 'Inquire'
}

function Get-VcsApi {
    param([string]$VcsInstallPath)

    Write-Debug "Determining the path to the VCS addin..."

    if ($VcsInstallPath) {
        Write-Debug "VcsInstallPath provided: $VcsInstallPath."
        $VcsInstallPath = [System.IO.Path]::GetFullPath($VcsInstallPath, (Get-Location))
        $vcsInstallDirectory = [System.IO.Path]::GetDirectoryName($VcsInstallPath)
        $vcsAccdaName = [System.IO.Path]::GetFileNameWithoutExtension($VcsInstallPath)
        $accdaPath = Join-Path $vcsInstallDirectory $vcsAccdaName    
        Write-Debug "Checking if VCS addin exists in $accdaPath..."
        if (-not (Test-Path $accdaPath)) {
            throw "VCS addin not found at provided install path $accdaPath"
        }

        Write-Debug "Found Addin: $(Join-Path $vcsInstallDirectory $vcsAccdaName).API"
        return "$(Join-Path $vcsInstallDirectory $vcsAccdaName).API"
    }

    $vcsAddinFolder = (Join-Path $env:AppData "MSAccessVCS")
    $candidate = (Join-Path $vcsAddinFolder "Version Control.accda")
    Write-Debug "Checking if VCS addin is installed at default location $candidate."
    
    if ((Test-Path $candidate)) {
        Write-Debug "Found Addin: $(Join-Path $vcsAddinFolder "Version Control.API")"
        return "$vcsAddinFolder\Version Control.API"
    }

    $tmpAddinFolder = (Join-Path $env:AppData "MSAccessVCS.Build")
    $tmpCandidate = (Join-Path $tmpAddinFolder "Version Control.accda")
    Write-Debug "VCS addin not installed. Downloading addin to $tmpCandidate for installation."
    
    $sourceUrl = "https://github.com/lagwagon667/msaccess-vcs-addin/releases/download/v5.1.0-alpha-1/Version.Control.zip"

    if (-not (Test-Path $tmpAddinFolder)) {
        New-Item -ItemType Directory -Path $tmpAddinFolder | Out-Null
    }
        
    $zip = (Join-Path $tmpAddinFolder "Version Control.zip")
    Invoke-WebRequest -Uri $sourceUrl -OutFile $zip
    Expand-Archive -Path $zip -DestinationPath $tmpAddinFolder -Force
    Remove-Item $zip

    Add-AccessTemporaryTrustedLocation -Path $tmpAddinFolder -Key "AccessBuildToolsInstallVcsAddin"

    Write-Debug "Installing VCS addin to default location..."
    $windowStyle = if ($IsDebugMode) { 'Normal' } else { 'Hidden' }
    $process = Start-Process -FilePath "msaccess.exe" `
        -ArgumentList """$tmpCandidate"" /cmd INSTALL SILENT" `
        -WindowStyle $windowStyle `
        -PassThru
    $process.WaitForExit()
        
    Remove-AccessTemporaryTrustedLocation -Key "AccessBuildToolsInstallVcsAddin"
    
    Write-Debug "Cleaning up..."
    Remove-Item $tmpAddinFolder -Recurse -Force
    # Write-Debug "Found VCS addin in $vcsAddinFolder..."
    return "$vcsAddinFolder\Version Control.API"
}
function CacheDBConnections {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Connection[]]$Connections)

    if ($Connections) {
        Write-Debug "Caching DB connections..."
        New-Module -AccessApp $AccessApp -ModuleName "DbConnectModule" -Content (Get-DbConnectModule)
        foreach ($connection in $Connections) {
            Write-Debug "Establishing db connection for $($connection.User) to $($connection.DSN)..."
            $accessApp.Run("DbConnect", $connection.DSN, $connection.User, $connection.Password)
        }
        Remove-Module -AccessApp $AccessApp -ModuleName "DbConnectModule"
        Write-Debug "Finished caching DB connections..."
    }
    else {
        Write-Debug "No DB connections for caching provided..."
    }   
}

function Invoke-UnitTests {
    <#
.SYNOPSIS
Runs all tests in the project.

.DESCRIPTION
Executes unit tests defined in the Access project and reports results. 
Tests must be defined according to the format expected by ms-access-vcs`s own 
test format. The ms-access-vcs test runner will be used for running the tests.

.PARAMETER AccessDBPath
Path to the Access Database containing the tests.

.PARAMETER Connections
An array of database‑connection descriptor objects used to initialize and cache
connections for the unit‑test run. Each object provides a DSN, username, and
password. Before executing the unit tests, the module establishes a connection
for every supplied descriptor so the tests can reuse these connections without
re‑authentication.

.PARAMETER VcsInstallPath
Specifies an optional custom installation path for the VCS add‑in. When not
provided, the module attempts to locate the add‑in in the standard default
location.

.EXAMPLE
Invoke-UnitTests -AccessDBPath .\myunittestdb.accdb -Connections @((New-Connection -DSN "mydb_dsn" -User "username" -Password "secret"))
Runs all tests in the myunittestdb.accdb folder.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessDBPath,
        [object[]]$Connections,
        [string]$VcsInstallPath)
    try {
        $start = Get-Date
        $conns = [Connection[]]$Connections
        $AccessDBPath = [System.IO.Path]::GetFullPath($AccessDBPath, (Get-Location))
        if (-not (Test-Path $AccessDBPath)) {
            throw "Acccess DB $AccessDBPath not found."
        }
        Write-Progress -Activity "Invoke-UnitTests" -Status "Checking prerequisites..." -PercentComplete 10 
        $vcsApi = Get-VcsApi -VcsInstallPath $VcsInstallPath
        $accessApp = Open-AccessWithoutStartupCommands -AccessDBPath $AccessDBPath
    
        Write-Progress -Activity "Invoke-UnitTests" -Status "Caching DB Connections..." -PercentComplete 20 
        CacheDBConnections -AccessApp $accessApp -Connections $conns

        Write-Progress -Activity "Invoke-UnitTests" -Status "Running Unit Tests..." -PercentComplete 100 
        $result = $accessApp.Run($vcsApi, "RunTestsHeadless") | ConvertFrom-Json
        $elapsed = (Get-Date) - $start
        $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed
        if ($result.allPassed) {
            $test = $result.summary.subs -eq 1 ? "test" : "tests"
            Write-Host "$($result.summary.subs) unit $test from $AccessDBPath passed. Duration: $elapsedStr"
        }
        else {
            $failCount = $result.summary.failed + $result.summary.errored
            $test = $failCount -eq 1 ? "test" : "tests"
            Write-Error "$failCount unit $test from $AccessDBPath finished with errors. Duration: $elapsedStr"
        }
    }
    finally {
        if ($accessApp) {
            Close-AccessInstance -AccessApp $accessApp
        }
    }
}

function Get-Version {
    Write-Debug "Determining version..."
    $gitVersionFolder = Join-Path -Path $env:LOCALAPPDATA -ChildPath "gitVersion\6.8.2"
    $gitVersionTool = Join-Path -Path $gitVersionFolder -ChildPath "gitVersion.exe"
    if (!(Test-Path $gitVersionTool)) {
        Write-Debug "GitVersion not found. Downloading from GitHub..."
        $sourceUrl = "https://github.com/GitTools/GitVersion/releases/download/6.8.2/gitversion-win-x64-6.8.2.zip"
        $zip = "gitVersion.zip"
        Invoke-WebRequest -Uri $sourceUrl -OutFile $zip
        Expand-Archive -Path $zip -DestinationPath $gitVersionFolder -Force
        Remove-Item $zip
    }

    $gitVersion = & $gitversionTool @('/l', 'gitVersion.log') | ConvertFrom-Json
    return $gitVersion
}

function Publish-Accdb {
    <#
.SYNOPSIS
Embeds external references into an Access ACCDB and injects startup code to
unpack them at runtime.

.DESCRIPTION
Processes a Microsoft Access database (ACCDB) by embedding external references
directly into the file and injecting the required startup code to unpack and
load these references when the database is opened. This allows the ACCDB to run
on machines where the referenced libraries are not available.

In addition to embedding references, the cmdlet can optionally apply the
current theme of the Access database to any ACCDB references before embedding
it.

You must provide the path to the ACCDB and a list of reference specifications
describing which references should be embedded. The cmdlet writes the modified
ACCDB to the specified output path. If a file already exists at that location, 
it can be overwritten using the -Force switch. If Output is omitted, the provided 
Access DB is overwritten.

.PARAMETER AccessDBPath
Path to the ACCDB file that should be processed. This parameter is mandatory.

.PARAMETER ReferencesToEmbed
Array of reference specification objects describing which external references
should be embedded into the ACCDB. Each specification contains:
- Name - the name of the reference to embed. Use Get-References to display 
  a list of all suitable references
- ApplyTheme - whether the current Access theme should be applied to this
  reference (only relevant for ACCDB references)

This parameter is mandatory.

.PARAMETER Output
Optional path where the processed ACCDB should be written. If omitted, the database
specified in AccessDBPath is used.

.PARAMETER Force
Allows overwriting an existing ACCDB file at the output location. Only relevant 
in conjunction with the Output parameter

.EXAMPLE
Publish-Accdb -AccessDBPath "my.accdb" -ReferencesToEmbed  `
    @((New-ReferenceSpec -Name "refA" -ApplyTheme $true), (New-ReferenceSpec -Name "refB" -ApplyTheme $false)) `
    -Output "dist\my.accdb" -Force
Embeds the specified references into the ACCDB and writes the published version
to the given output path, overwritting it if it already exists.

.NOTES
Databases processed with Publish-Accdb cannot be compiled into an ACCDE using
Convert-ToAccde. Publishing injects unpack/startup code that modifies the ACCDB
internally, and such modifications are not compatible with ACCDE compilation.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessDBPath,
        [Parameter(Mandatory)]
        [object[]]$ReferencesToEmbed,
        [string]$Output,
        [switch]$Force
    )
    try {
        $start = Get-Date
                
        Write-Debug "Start publishing Access DB"
        if (-not $ReferencesToEmbed -or $ReferencesToEmbed.Count -eq 0) {
            throw "No references to embed provided. Exiting..."
        }
        $refsToEmbed = [ReferenceSpec[]]$ReferencesToEmbed

        Write-Progress -Activity "Publish-Accdb" -Status "Checking prerequisites..." -PercentComplete 10 
        
        $AccessDBPath = [System.IO.Path]::GetFullPath($AccessDBPath, (Get-Location))
        if (-not $Output) {
            Write-Debug "Setting output to $AccessDBPath..."
            $Output = $AccessDBPath
        }
        else {
            $Output = [System.IO.Path]::GetFullPath($Output, (Get-Location))
            Write-Debug "Using provided output at $Output..."
        }
    
        Add-AccessTemporaryTrustedLocation -Path ([System.IO.Path]::GetDirectoryName($Output))
        if ($AccessDbPath -ne $Output) {
            Copy-Item -Path $AccessDBPath -Destination $Output -Force
        }
        
        $accessApp = Open-AccessWithoutStartupCommands -AccessDbPath $Output
        $dbName = [System.IO.Path]::GetFileNameWithoutExtension($Output)
        
        Publish -AccessApp $accessApp -DbName $dbName -ReferencesToEmbed $refsToEmbed -CurrentProgress 10
        
        $elapsed = (Get-Date) - $start
        $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed
        Write-Host "Finished publishing to $Output. Duration: $elapsedStr"
    }
    finally {
        if ($accessApp) {
            Close-AccessInstance -AccessApp $accessApp
        }
        Remove-AccessTemporaryTrustedLocation
    }
}

function Convert-ToAccde {
    <#
.SYNOPSIS
Compiles an Access database (ACCDB) into an ACCDE file.

.DESCRIPTION
Converts a Microsoft Access database (ACCDB) into a compiled ACCDE file. The
cmdlet invokes Access, producing an ACCDE suitable for deployment where source
code should be protected and only compiled objects are required.

The command requires the path to the source ACCDB and an output path for the
generated ACCDE. If an ACCDE already exists at the output location, it can be
overwritten using the -Force switch.

.PARAMETER AccessDBPath
Path to the ACCDB file that should be compiled. This parameter is mandatory.

.PARAMETER Output
Path where the resulting ACCDE file will be written. This parameter is
mandatory.

.PARAMETER Force
Allows overwriting an existing ACCDE file at the output location.

.EXAMPLE
Convert-ToAccde -AccessDBPath "my.accdb" -Output "my.accde"
Compiles the specified ACCDB into an ACCDE and writes it to the given output
path.

.NOTES
This cmdlet cannot be used on ACCDB files that were previously processed with
Publish-Accdb. Publishing embeds references and injects code that modifies the database
internally, and such modified databases cannot be run when compiled into an ACCDE.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessDBPath,
        [Parameter(Mandatory)]
        [string]$Output,
        [switch]$Force
    )
    try {
        $start = Get-Date
        $AccessDBPath = [System.IO.Path]::GetFullPath($AccessDBPath, (Get-Location))
        if (-not (Test-Path $AccessDBPath)) {
            throw "Acccess DB $AccessDBPath not found."
        }
            
        Write-Progress -Activity "Convert-ToAccde" -Status "Checking prerequisites..." -PercentComplete 10 

        Write-Debug "Start compiling to accde"
        
        $Output = [System.IO.Path]::GetFullPath($Output, (Get-Location))
        if (Test-Path $Output) {
            if ($Force) {
                $filename = [System.IO.Path]::GetFileNameWithoutExtension($Output)
                $extension = [System.IO.Path]::GetExtension($output)
                $backup = "${filename}_backup$extension"
                $i = 1

                $outputFolder = [System.IO.Path]::GetDirectoryName($Output)

                while (Test-Path (Join-Path $outputFolder $backup)) {
                    $filename = [System.IO.Path]::GetFileNameWithoutExtension($Output)
                    $extension = [System.IO.Path]::GetExtension($output)
                    $backup = "${filename}_backup_$i$extension"
                    $i++
                }
                Rename-Item $Output $backup
                Write-Debug "Created backup file $(Join-Path $outputFolder $backup)"
            }
            else {
                throw "Accde $Output already exists. Use -Force to overwrite."
            }
        }

    
        Add-AccessTemporaryTrustedLocation -Path ([System.IO.Path]::GetDirectoryName($AccessDBPath))
        # these two commands are required, otherwise SysCmd 603 fails
        $accessApp = Open-AccessWithoutStartupCommands -AccessDBPath $AccessDBPath
        $accessApp.CloseCurrentDatabase()

        Write-Progress -Activity "Convert-ToAccde" -Status "Compiling to .accde..." -PercentComplete 100 
        $result = $accessApp.SysCmd(603, $AccessDBPath, $Output)
        if ($result -eq 0) {
            throw 'SysCmd returned 0. This means an error occurred. Check if the database can be compiled.'
        }
        $elapsed = (Get-Date) - $start
        $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed
        Write-Host "Finished compiling .accde to $Output. Duration: $elapsedStr"
    }
    finally {
        if ($accessApp) {
            Close-AccessInstance -AccessApp $accessApp
        }
        Remove-AccessTemporaryTrustedLocation
    }
}

function Get-References {
    <#
.SYNOPSIS
Retrieves all references defined in an Access ACCDB file that are not builtin.

.DESCRIPTION
Reads the reference information from a Microsoft Access database (ACCDB) and
displays the list of references currently registered in the file. This includes
the reference name, version information, and full path.

The cmdlet is typically used to inspect the references of an ACCDB before embedding 
references. It provides a clear overview of which external libraries the database 
depends on.

.PARAMETER AccessDBPath
Path to the ACCDB file whose references should be retrieved. This parameter is
mandatory.

.EXAMPLE
Get-References -AccessDBPath "my.accdb"
Returns all references defined in the specified ACCDB.
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AccessDBPath
    )
    try {
        $start = Get-Date
        $AccessDBPath = [System.IO.Path]::GetFullPath($AccessDBPath, (Get-Location))
        if (-not (Test-Path $AccessDBPath)) {
            throw "Acccess DB $AccessDBPath not found."
        }
    
        Write-Debug "Listing all references from $AccessDBPath..."

        Add-AccessTemporaryTrustedLocation -Path ([System.IO.Path]::GetDirectoryName($AccessDBPath))
        $accessApp = Open-AccessWithoutStartupCommands -AccessDBPath $AccessDBPath
        $notBuiltinRefs = $accessApp.References | Where-Object { -not $_.Builtin }
        Write-Debug "Access DB has $($notBuiltinRefs.Count) reference(s)..."
        
        $notBuiltinRefs | Format-Table @("Name", "Major", "Minor", "FullPath") -AutoSize
    
        $elapsed = (Get-Date) - $start
        $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed
        Write-Host "Finished listing references. Duration: $elapsedStr"
        Write-Debug "Finished listing references..."
    }
    finally {
        if ($accessApp) {
            Close-AccessInstance -AccessApp $accessApp
        }
        Remove-AccessTemporaryTrustedLocation
    }
}

function Open-AccessWithoutStartupCommands {
    param(
        [Parameter(Mandatory)]
        [string]$AccessDbPath,
        [string]$Parameters
    )

    $env:ACCESS_NO_AUTOEXEC = "1"
    $accessApp = New-Object -ComObject "Access.Application"
    if (IsDebugMode) {
        $accessApp.Visible = $true
    }
    $accessApp.SetOption("Error Trapping", [ref]2) # 2 = on unhandled errors
    if ($Parameters) {
        Write-Debug "Parameters set to $Parameters..."
        $accessApp.Command = $Parameters
    }

    if (Test-Path $AccessDbPath) {
        Write-Debug "Opening existing database $AccessDbPath"
        $accessApp.OpenCurrentDatabase($AccessDbPath)
    }
    else {
        Write-Debug "Creating empty database $AccessDbPath"
        $accessApp.NewCurrentDatabase($AccessDbPath)
    }

    $env:ACCESS_NO_AUTOEXEC = "0"
    
    return $accessApp
}

function New-TempFile {
    param(
        [Parameter(Mandatory)]
        [string]$Filename,
        [Parameter(Mandatory)]
        [string]$DBName
    )
    Write-Debug "Getting temp file name based on $Filename..."
    $Filename = Split-Path $Filename -Leaf

    $base = Join-Path $env:LOCALAPPDATA "AccessTemp" $DBName

    if (-not (Test-Path $base)) {
        New-Item -ItemType Directory -Path $base | Out-Null
    }

    $idx = 1
    $path = Join-Path $base $Filename
    
    $ext = [System.IO.Path]::GetExtension($Filename)
    $prefix = [System.IO.Path]::GetFileNameWithoutExtension($Filename)

    while (Test-Path $path) {
        $path = Join-Path $base "${prefix}_$idx$ext"
        $idx++
    } 
    
    Write-Debug "Temp file name is $path."
    return $path
}


function Add-AccessTemporaryTrustedLocation {
    param(
        [string]$Path,
        [string]$Key
    )
    if (-not $Path) {
        $Path = Join-Path $env:LOCALAPPDATA "AccessTemp"
    }
    if (-not $Key) {
        $Key = "TemporaryBuildTrustedLocation"
    }
    
    $officeVersions = "16.0", "15.0", "14.0", "12.0"
    $baseKey = $null

    foreach ($v in $officeVersions) {
        $test = "HKCU:\Software\Microsoft\Office\$v\Access\Security\Trusted Locations"
        if (Test-Path $test) {
            $baseKey = $test
            break
        }
    }

    if (-not $baseKey) {
        throw "Could not find Access Trusted Locations registry path."
    }

    $newKey = Join-Path $baseKey $Key

    New-Item -Path $newKey -Force | Out-Null
    New-ItemProperty -Path $newKey -Name "Path" -Value "$Path\" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $newKey -Name "Description" -Value "Temporary Trusted Location" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $newKey -Name "AllowSubFolders" -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $newKey -Name "Date" -Value (Get-Date).ToString() -PropertyType String -Force | Out-Null

    Write-Debug "Set trusted location $Path to $Key..."
}

function Remove-AccessTemporaryTrustedLocation {
    param(
        [string]$Key
    )
    if (-not $Key) {
        $Key = "TemporaryBuildTrustedLocation"
    }

    $officeVersions = "16.0", "15.0", "14.0", "12.0"
    $baseKey = $null

    foreach ($v in $officeVersions) {
        $test = "HKCU:\Software\Microsoft\Office\$v\Access\Security\Trusted Locations"
        if (Test-Path $test) {
            $baseKey = $test
            break
        }
    }

    if (-not $baseKey) {
        throw "Could not find Access Trusted Locations registry path."
    }

    $newKey = Join-Path $baseKey $Key

    if (Test-Path $newKey) {
        Remove-Item -Path $newKey -Recurse -Force -ErrorAction SilentlyContinue
        Write-Debug "Removed trusted location $Key..."
    }
    else {
        Write-Debug "Trusted location $Key not found. Doing nothing..."
    }
}

function Close-AccessInstance {
    param(
        [Parameter(Mandatory)]
        $AccessApp
    )

    Write-Debug "Closing Access instance..."
    try {
        try {
            $AccessApp.CloseCurrentDatabase()
        }
        catch {
        }

        try {
            $AccessApp.Quit()
        }
        catch {
        }
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($AccessApp) | Out-Null

        Remove-Variable AccessApp -ErrorAction SilentlyContinue

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Publish {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$DBName,
        [Parameter(Mandatory)]
        [ReferenceSpec[]]$ReferencesToEmbed,
        [Parameter(Mandatory)]
        [int]$CurrentProgress
    )

    Write-Debug "Publishing $DbName..."
    Test-ReferencesTable -AccessApp $AccessApp
    $progressStepSize = (100 - $CurrentProgress) / $ReferencesToEmbed.Count

    foreach ($refSpec in $ReferencesToEmbed) {
        Write-Debug "Embedding reference $($refSpec.Name)..."
        $CurrentProgress += $progressStepSize
        Write-Progress -Activity "Publish-Accdb" -Status "Embedding reference $($refSpec.Name)..." -PercentComplete $CurrentProgress
        $reference = Get-Reference -AccessApp $AccessApp -ReferenceName $refSpec.Name
        $referencePath = $reference.FullPath
        $refName = Split-Path $referencePath -Leaf 
        Remove-Reference -AccessApp $AccessApp -ReferenceName $refSpec.Name

        $tempFile = New-TempFile -FileName $refName -DBName $DBName

        try {
            Write-Debug "Copying ref to temporary file $tempFile..."
            Copy-Item -Path $referencePath -Destination $tempFile -Force
            if ($refSpec.ApplyTheme) {
                Write-Debug "Applying theme to $refName..."
                $path = Split-Path $tempFile -Parent
                Add-AccessTemporaryTrustedLocation -Path $path -Key $DBName
                $themingApp = Open-AccessWithoutStartupCommands -AccessDbPath $tempFile
                $themeData = Get-Theme -AccessApp $accessApp -DBName $DBName
                Set-Theme -AccessApp $themingApp -ThemeData $themeData -DBName $DBName
                Close-AccessInstance -AccessApp $themingApp
                Remove-AccessTemporaryTrustedLocation -Key $DBName 
                Write-Debug "Theme successfully applied to $refName..."
            }
            Add-AssemblyToResources -AccessApp $AccessApp -ReferencePath $tempFile
        }
        finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }

    Add-AutoExecMacro -AccessApp $AccessApp -DBName $DBName -ReferencesToEmbed $ReferencesToEmbed

    Write-Debug "Finished publishing $DBName"
}

function Get-Reference {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$ReferenceName
    )

    Write-Debug "Searching reference $ReferenceName..."
    $references = $AccessApp.References |
    Where-Object { $_.Name -eq $ReferenceName }
    if ($references.Count -eq 0) {
        throw "Reference matching name $ReferenceName not found"
    }

    $reference = $references |
    Select-Object -First 1

    Write-Debug "Found reference $ReferenceName..."
    return $reference
}

function Remove-Reference {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$ReferenceName
    )
    Write-Debug "Removing reference $ReferenceName..."
    $reference = $AccessApp.References |
    Where-Object { $_.Name -eq $ReferenceName } |
    Select-Object -First 1

    $AccessApp.References.Remove($reference)
    Write-Debug "Finished removing reference $ReferenceName."
}

function Test-ReferencesTable {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp
    )

    Write-Debug "Checking if RuntimeReferences table exists..."
    $db = $AccessApp.CurrentDb()

    $exists = $false
    foreach ($td in $db.TableDefs) {
        if ($td.Name -eq "RuntimeReferences") {
            $exists = $true
            Write-Debug "RuntimeReferences table exists..."
            break
        }
    }

    if ($exists) { return }

    Write-Debug "RuntimeReferences table does not exist. Creating a new one..."

    $sql = @"
CREATE TABLE RuntimeReferences (
    Name TEXT PRIMARY KEY,
    Data LONGBINARY
)
"@
    $db.Execute($sql)
    $db.TableDefs.Refresh()
    
    # Hide table
    $AccessApp.SetHiddenAttribute(0, "RuntimeReferences", $true)  # acTable = 0
    Write-Debug "Finished checking if RuntimeReferences table exists..."
}

function Add-AssemblyToResources {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$ReferencePath
    )

    Write-Debug "Adding $ReferencePath to Access DB resources"

    $db = $accessApp.CurrentDb()

    $rs = $db.OpenRecordset("RuntimeReferences", 2)  # dbOpenDynaset = 2

    $rs.AddNew()
    $filename = [System.IO.Path]::GetFileName($ReferencePath)
    $rs.Fields["Name"].Value = $filename

    # Read file bytes
    $bytes = [System.IO.File]::ReadAllBytes($ReferencePath)

    # Append into LONGBINARY
    $rs.Fields["Data"].AppendChunk($bytes)

    $rs.Update()
    $rs.Close()

    Write-Debug "Finished adding $ReferencePath to Access DB resources"
}

function Set-Theme {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [object[]]$ThemeData,
        [Parameter(Mandatory)]
        [string]$DBName
    )

    Write-Debug "Applying theme to database..."
    $themeName = $ThemeData[0]
    $extension = $ThemeData[1]
    $name = $ThemeData[2]
    $type = $ThemeData[3]
    $fileName = $ThemeData[4]
    $fullFilePath = $ThemeData[5]

    $db = $AccessApp.CurrentDb()
    $rs = $db.OpenRecordset("MSysResources", 2)  # dbOpenDynaset

    $rs.AddNew()
    $rs.Fields["Name"].Value = $name
    $rs.Fields["Type"].Value = $type
    $rs.Fields["Extension"].Value = $extension

    $child = $rs.Fields["Data"].Value
    $child.AddNew()
    Write-Debug "Loading Theme data from $fullFilePath"
    $child.Fields["FileData"].LoadFromFile($fullFilePath)
    $child.Fields["FileName"].Value = $fileName
    $child.Update()

    $rs.Update()
    $rs.Close()

    # Cleanup
    Remove-Item $fullFilePath -Force

    try {
        $db.Properties["Theme Resource Name"].Value = $themeName
    }
    catch {
        # Create property if it doesn't exist
        $dbText = 10
        $p = $db.CreateProperty("Theme Resource Name", $dbText, $themeName)  
        $db.Properties.Append($p)
    }
    Write-Debug "Finished applying theme to database..."
}

function Get-Theme {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$DBName
    )

    Write-Debug "Retrieving theme from database..."

    $db = $AccessApp.CurrentDb()

    $rs = $db.OpenRecordset("SELECT * FROM MSysResources WHERE Extension='thmx' ORDER BY id DESC", 2)

    if ($rs.EOF) {
        throw "No THMX resource found."
    }

    $extension = $rs.Fields["Extension"].Value
    $name = $rs.Fields["Name"].Value
    $type = $rs.Fields["Type"].Value
    $child = $rs.Fields["Data"].Value

    $child.MoveFirst()
    $fileName = $child.Fields["FileName"].Value
    $themeFile = New-TempFile -Filename $fileName -DBName $DBName
    Write-Debug "Writing theme data to $themeFile..."
    $child.Fields["FileData"].SaveToFile($themeFile)
    $themeName = $db.Properties["Theme Resource Name"].Value

    $rs.Close()
    
    Write-Debug "Finished retrieving theme from database."

    return $themeName, $extension, $name, $type, $fileName, $themeFile
}


function Add-AutoExecMacro {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$DBName,
        [Parameter(Mandatory)]
        [ReferenceSpec[]]$ReferencesToEmbed
    )

    $macroFunction = Get-ReferenceLoaderModule
    $macroFunction = $macroFunction.Replace('$dbName', $DBName)

    $macro = Get-AutoExecOrDefault -AccessApp $AccessApp

    if ($macro) {
        Write-Debug "AutoExec macro exists. Renaming it..."

        $macroName = "AutoExecReferences"
        $macroName = Rename-Macro -AccessApp $AccessApp -OldName "AutoExec" -NewName $macroName

        $macroFunction = $macroFunction.Replace(
            "'@placeholder RunMacro",
            "DoCmd.RunMacro `"$macroName`""
        )

        Write-Debug "Renamed AutoExec macro will be called from ReferenceLoader..."
    }

    $startupForm = Get-StartupFormOrDefault -AccessApp $AccessApp

    if ($startupForm) {
        Clear-StartupForm -AccessApp $AccessApp

        $macroFunction = $macroFunction.Replace(
            "'@placeholder StartupForm",
            "DoCmd.OpenForm `"$startupForm`""
        )

        Write-Debug "Added call to load original startup form from ReferenceLoader..."
    }

    New-Module -AccessApp $AccessApp -ModuleName "BootstrapModule" -Content $macroFunction
    $AccessApp.SetHiddenAttribute(5, "BootstrapModule", $true)   # 5 = acModule

    $autoExecContent = Get-AutoExecMacro
    New-Macro -AccessApp $AccessApp -MacroName "AutoExec" -MacroContent $autoExecContent -DBName $DBName
    $AccessApp.SetHiddenAttribute(4, "AutoExec", $true)      # 4 = acMacro
}

function Get-AutoExecOrDefault {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp
    )

    Write-Debug "Searching AutoExec macro..."
    foreach ($obj in $AccessApp.CurrentProject.AllMacros) {
        if ($obj.Name -eq "AutoExec") {
            Write-Debug "Found AutoExec macro."
            return $obj.Name
        }
    }

    Write-Debug "No AutoExec macro found."
    return $null
}

function Test-MacroExists {
    param(
        [Parameter(Mandatory)]
        $accessApp,

        [Parameter(Mandatory)]
        [string]$MacroName
    )

    foreach ($m in $accessApp.CurrentProject.AllMacros) {
        if ($m.Name -eq $MacroName) {
            return $true
        }
    }

    return $false
}

function Rename-Macro {
    param(
        [Parameter(Mandatory)]
        [object]$AccessApp,
        [Parameter(Mandatory)]
        [string]$OldName,
        [Parameter(Mandatory)]
        [string]$NewName
    )

    Write-Debug "Renaming macro $OldName to $NewName..."
    $index = 0
    $macroName = $NewName

    while (Test-MacroExists -AccessApp $AccessApp -MacroName $macroName) {
        $macroName = "$NewName$index"
        $index++
    }

    # 4 = acMacro
    $AccessApp.DoCmd.Rename($macroName, 4, $oldName)

    Write-Debug "Finished renaming macro $OldName to $NewName..."
    return $macroName
}

function Get-StartupFormOrDefault {
    param(
        [Parameter(Mandatory)]
        $AccessApp
    )

    Write-Debug "Searching for a startup form..."

    $db = $AccessApp.CurrentDb()

    foreach ($p in $db.Properties) {
        if ($p.Name -eq "StartupForm") {
            Write-Debug "Found startup form $($p.Value)..."
            return $p.Value
        }
    }

    Write-Debug "No startup form found."
    return $null
}

function Clear-StartupForm {
    param(
        [Parameter(Mandatory)]
        $AccessApp
    )

    Write-Debug "Removing startup form from database..."
    $AccessApp.CurrentDb().Properties.Delete("StartupForm")
    Write-Debug "Finished removing startup form from database..."
}

function New-Module {
    param(
        [Parameter(Mandatory)]
        $AccessApp,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [string]$Content
    )

    Write-Debug "Adding module $ModuleName to database..."
    $project = Get-VbProjectOfCurrentDb -AccessApp $accessApp

    $component = $project.VBComponents.Add(1) # 1 = vbext_ct_StdModule
    $component.Name = $ModuleName
    $component.CodeModule.AddFromString($Content)
    
    try {
        $accessApp.DoCmd.Save(5, $component.Name) # 5 = acModule
    }
    catch {
        Write-Error "COM Exception: $($_.Exception.Message)"
        Write-Error "Access Errors:"

        $errors = $accessApp.Errors
        for ($i = 0; $i -lt $errors.Count; $i++) {
            $err = $errors.Item($i)
            Write-Error "[$i] $($err.Number): $($err.Description)"
        }

        throw
    }
    Write-Debug "Finished adding module $ModuleName to database..."
}

function Remove-Module {
    param(
        [Parameter(Mandatory)]
        $AccessApp,
        [Parameter(Mandatory)]
        [string]$ModuleName
    )

    Write-Debug "Removing module $ModuleName..."
    $project = Get-VbProjectOfCurrentDb -AccessApp $AccessApp

    $component = $project.VBComponents.Item($ModuleName)
    $project.VBComponents.Remove($component) 
    Write-Debug "Finished removing module $ModuleName..."
}

function New-Macro {
    param(
        [Parameter(Mandatory)]
        $AccessApp,
        [Parameter(Mandatory)]
        [string]$MacroName,
        [Parameter(Mandatory)]
        [string]$MacroContent,
        [Parameter(Mandatory)]
        [string]$DBName
    )

    Write-Debug "Creating macro $MacroName..."
    $temp = New-TempFile -Filename "$MacroName.mcr" -DBName $DBName
    [System.IO.File]::WriteAllText($temp, $MacroContent, [System.Text.Encoding]::UTF8)

    $AccessApp.LoadFromText(4, $MacroName, $temp) # 4 = acMacro

    Remove-Item $temp -Force
    Write-Debug "Finished creating macro $MacroName."
}

function Get-VbProjectOfCurrentDB {
    param(
        [Parameter(Mandatory)]
        $AccessApp
    )

    Write-Debug "Retrieving VB project of current Access DB..."
    $accdbPath = $AccessApp.CurrentProject.FullName

    foreach ($proj in $AccessApp.VBE.VBProjects) {
        try {
            if ($proj.FileName -eq $accdbPath) {
                Write-Debug "Finished retrieving VB project of current Access DB..."
                return $proj
            }
        }
        catch {
            # Ignore exceptions
        }
    }

    throw "VBProject for ACCDB not found."
}

function Get-AutoExecMacro {
    @'
Version =196611
PublishOption =1
ColumnsShown =0
Begin
    Action ="RunCode"
    Argument ="=BootstrapReferences()"
End
Begin
    Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
        "nterfaceMacro MinimumClientDesignVersion=\"14.0.0000.0000\" xmlns=\"http://schem"
        "as.microsoft.com/office/accessservices/2009/11/application\"><Statements><Action"
        " Name=\"RunCode\"><Argument Nam"
End
Begin
    Comment ="_AXL:e=\"FunctionName\">=BootstrapReferences()</Argument></Action></Statements><"
        "/UserInterfaceMacro>"
End
'@
}

function Get-DbConnectModule {
    @'
Public Sub DbConnect(ByVal pDsn As String, ByVal pUser As String, ByVal pPass As String)
    Dim connStr As String: connStr = "ODBC;DSN=" & pDsn & ";UID=" & pUser & ";PWD=" & pPass
    DBEngine.Workspaces(0).OpenDatabase "", dbDriverCompleteRequired, False, connStr
End Sub
'@
}

function Get-ReferenceLoaderModule {
    @'
Private Sub AddReference(pFilePath As String)
    Dim ref As Reference

    On Error Resume Next
    Set ref = Application.References.AddFromFile(pFilePath)
    
    If Err.Number <> 0 Then
        MsgBox "Failed to add reference:  " & pFilePath & vbCrLf & Err.Description, vbExclamation
        Err.Clear
    End If

    On Error GoTo 0
End Sub

Private Sub EnsureRemoveReference(pPath As String)
    Dim refItem As Reference
    For Each refItem In Application.References
        If Right(refItem.FullPath, Len(pPath)) = pPath Then
            Application.References.Remove refItem
        End If
    Next
End Sub

Private Function GetFolderFromPath(ByVal pFullPath As String) As String
    Dim fso As Object :   Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderFromPath = fso.GetParentFolderName(pFullPath)
End Function

Private Function JoinPaths(ByVal pPathA As String, ByVal pPathB As String) As String
    Dim fso As Object :   Set fso = CreateObject("Scripting.FileSystemObject")
    JoinPaths = fso.BuildPath(pPathA, pPathB)
End Function

Private Sub CreateFolderRecursive(pPath As String)
    Dim fso As Object
    Dim parts() As String
    Dim current As String
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    parts = Split(pPath, "\")
    current = parts(0) & "\"

    For i = 1 To UBound(parts)
        current = current & parts(i) & "\"
        If Not fso.FolderExists(current) Then
            fso.CreateFolder current
        End If
    Next i
End Sub

Public Function BootstrapReferences() As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset2
    Dim fld As DAO.Field2
    Dim tempPath As String
    Dim filePath As String
    Dim fileName As String

    Const ChunkSize As Long = 32768   ' 32 KB chunks

    Dim fileNum As Integer
    Dim offset As Long
    Dim totalSize As Long
    Dim chunk() As Byte
    
    Set db = CurrentDb()
    
    Set rs = db.OpenRecordset( _
        "SELECT Name, Data FROM RuntimeReferences", _
        dbOpenDynaset)

    tempPath = JoinPaths(Environ$("TEMP"), "AccessReferences\$dbName\")

    While Not rs.EOF
        fileName = rs.Fields("Name")
        filePath = JoinPaths(tempPath, fileName)
        Dim fullFolderPath As String : fullFolderPath = GetFolderFromPath(filePath)
    
        If Dir(fullFolderPath, vbDirectory) = "" Then CreateFolderRecursive fullFolderPath
        Set fld = rs.Fields("Data")
        totalSize = fld.FieldSize
    
        fileNum = FreeFile
        Open filePath For Binary Access Write As #fileNum
    
        offset = 0
        Do While offset < totalSize
            chunk = fld.GetChunk(offset, ChunkSize)
            Put #fileNum, , chunk
            offset = offset + UBound(chunk) + 1
        Loop
    
        Close #fileNum
            
        EnsureRemoveReference fileName
        AddReference filePath
        rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    '@placeholder RunMacro

    '@placeholder StartupForm
End Function
'@
}

Export-ModuleMember -Function Build-Accdb
Export-ModuleMember -Function Publish-Accdb
Export-ModuleMember -Function Invoke-UnitTests
Export-ModuleMember -Function Convert-ToAccde
Export-ModuleMember -Function Get-References

# $typeAcceleratorsClass = [psobject].Assembly.GetType(
#     'System.Management.Automation.TypeAccelerators'
# )

# $typesToExport = @("ReferenceSpec", "Connection")
# foreach ($typeToExport in  $typesToExport) {
#     $type = $typeToExport -as [System.Type]
#     if (-not $type) {
#         Write-Error -Message (
#             'Unable to export {0}. Type not found.' -f $typeToExport
#         )
#     }
#     else {
#         $null = $TypeAcceleratorsClass::Add($typeToExport, $type)
#     }
# }

Write-Host "Access Build Tools ready to use"