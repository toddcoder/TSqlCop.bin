param ([switch]$Install, [switch]$Uninstall)

$visualStudio = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\Extensions\Enterprise\TSqlCop"
$visualStudioLabel = "Visual Studio 2017"
$ssms130 = "C:\Program Files (x86)\Microsoft SQL Server\130\Tools\Binn\ManagementStudio\Extensions\TSqlCop"
$ssms130Label = "SQL Server Management Studio 2016"
$ssms140 = "C:\Program Files (x86)\Microsoft SQL Server\140\Tools\Binn\ManagementStudio\Extensions\TSqlCop"
$ssms140Label = "SQL Server Management Studio 7.x"

$labels = @{$visualStudio = $visualStudioLabel; $ssms130 = $ssms130Label; $ssms140 = $ssms140Label; "13.0" = $ssms130Label; "14.0" = $ssms140Label}

$appDataLocal = "$env:USERPROFILE\AppData\Local\TSqlCop"

if ($Install) {
    $visualStudio, $ssms130, $ssms140 | ForEach-Object -Process {
        Write-Output "Copying files for $($labels[$_])"
        New-Item -Path $_ -ItemType Directory -Force | Out-Null
        Copy-Item -Path "\\l5cg7451fl9\TSqlCop extension\*" -Destination $_ -Recurse -Force
    }
    "13.0", "14.0" | ForEach-Object -Process {
        Write-Output "Modifying registry for $($labels[$_])"
        $key = "HKCU:\Software\Microsoft\SQL Server Management Studio\$_\Packages\{B94020EE-B348-4105-A362-A2ED89388CC8}"
        New-Item -Path $key -Force | Out-Null
        New-ItemProperty -Path $key -Name "SkipLoading" -Value "1" -PropertyType DWORD -Force | Out-Null
    }
    New-Item -Path $appDataLocal -ItemType Directory -Force | Out-Null
    Copy-Item -Path "\\l5cg7451fl9\TSqlCop extension\tsqlcop.json" -Destination $appDataLocal -Force
    Copy-Item -Path "\\l5cg7451fl9\TSqlCop extension\tsqlcop.sql.format.options.xml" -Destination $appDataLocal -Force
    Write-Output "Starting VSIX installer..."
    pushd 'C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE'
    &".\VSIXInstaller.exe" ".\Extensions\Enterprise\TSqlCop\TSqlCop.vsix"
    popd
}
elseif ($Uninstall) {
    Write-Output "Starting VSIX uninstaller..."
    pushd 'C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE'
    &".\VSIXInstaller.exe" /u:10fda389-bb88-4343-8dea-34c98e2e404a | Out-Null
    popd
    $visualStudio, $ssms130, $ssms140 | ForEach-Object -Process {
        Write-Output "Deleting files for $($labels[$_])"
        if (Test-Path -Path $_) {
            Remove-Item -Path $_ -Recurse -Force
        }
    }
    "13.0", "14.0" | ForEach-Object -Process {
        Remove-Item -Path "HKCU:\Software\Microsoft\SQL Server Management Studio\$_\Packages\{B94020EE-B348-4105-A362-A2ED89388CC8}" -Force | Out-Null
    }
    Remove-Item -Path $appDataLocal -Recurse -Force
}