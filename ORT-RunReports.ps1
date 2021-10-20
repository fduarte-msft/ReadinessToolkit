$script:rootFolderPath = "D:\SourceFiles"
$script:outputFolderPath = "D:\OutputReports"

function Get-ProcessOutput {
    [CmdletBinding()]
    Param(
        # Path to executtable
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [string]
        $Path,

        # Parameters 
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [string]
        $Parameters

    )

    # Setup the Process startup info 
    $local:processInfo = New-Object System.Diagnostics.ProcessStartInfo 
    $local:processInfo.FileName = "$local:Path"
    $local:processInfo.Arguments = "$local:Parameters"
    $local:processInfo.UseShellExecute = $false
    $local:processInfo.CreateNoWindow = $false
    $local:processInfo.RedirectStandardOutput = $true
    $local:processInfo.RedirectStandardError = $true

    # Create a process object using the startup info
    $local:process= New-Object System.Diagnostics.Process 
    $local:process.StartInfo = $processInfo

    # Start the process 
    $local:process.Start() | Out-Null

    # get output from stdout and stderr
    [PSCustomObject] $local:output = New-Object PSObject -Property @{
        stdout =$local:process.StandardOutput.ReadToEnd()
        stderr =$local:process.StandardError.ReadToEnd()
    }

    return $local:Output
}


# Get path to ReadinessReportCreateor.exe
$script:ortPaths = "${env:ProgramFiles(x86)}\Microsoft Readiness Toolkit for Office\ReadinessReportCreator.exe","$env:ProgramFiles\Microsoft Readiness Toolkit for Office\ReadinessReportCreator.exe"
foreach ($script:ortPath in $script:ortPaths) {
    if (Test-Path -Path $ortPath -PathType Leaf) {
        $script:ortFilePath = $script:ortPath
        break
    }
}

if ($script:ortFilePath) {

    # Grab each folder (recursive)
    $script:subFolders = Get-ChildItem -Path $script:rootFolderPath -Recurse -Directory -Force -ErrorAction SilentlyContinue | Select-Object FullName

    # scan each folder using ORT
    foreach ($script:subFolder in $script:subFolders) {
        try {
            Write-Output "Processing: $($script:subFolder.FullName)"
            $script:process = Get-ProcessOutput -Path $script:ortFilePath -Parameters "-p `"$($script:subFolder.FullName)`" -r -output `"$($script:outputFolderPath)`"" -ErrorAction Stop
            Write-Output "Output: $($script:process.stdout)"
            
        } catch {
            Write-Output "Error detected: $($Error[0].Exception)"
            Write-Output "Output: $($script:process.stderr)"
        }
    }

} else {
    Write-Output "Readiness Toolkit not installed. Install and re-run script"
}