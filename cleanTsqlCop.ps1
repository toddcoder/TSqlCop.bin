# load Visual Studio's private registry hive
$name = (dir ~\AppData\Local\Microsoft\VisualStudio | Where-Object {$_.Name -like '15.0*'} | Select-Object -Property Name).Name
pushd "~\AppData\Local\Microsoft\VisualStudio\$name"
copy .\privateregistry.bin .\privateregistry.backup.bin
reg load HKLM\isolatedHive .\privateregistry.bin
pushd "HKLM:\isolatedHive\Software\Microsoft\VisualStudio\${name}_Config\"

# remove TSqlCop's extension package by name
pushd '.\InstalledProducts\'
if (Test-Path -Path '.\SqlPackage') {
    Remove-Item 'SqlPackage'
}
popd

# remove TSqlCop's package by GUID
pushd '.\Packages\'
if (Test-Path -Path '.\{B94020EE-B348-4105-A362-A2ED89388CC8}') {
    Remove-Item '.\{B94020EE-B348-4105-A362-A2ED89388CC8}'
}
popd

# remove TSqlCop's autoload GUIDs
pushd '.\AutoLoadPackages\'
if (Test-Path -Path '.\{adfc4e64-0397-11d1-9f4e-00a0c911004f}') {
    Remove-Item '.\{adfc4e64-0397-11d1-9f4e-00a0c911004f}'
}
if (Test-Path -Path '.\{f1536ef8-92ec-443c-9ed7-fdadf150da82}') {
    Remove-Item '.\{f1536ef8-92ec-443c-9ed7-fdadf150da82}'
}
popd

# remove TSqlCop's menus
Remove-ItemProperty -Path '.\Menus' -Name '{B94020EE-B348-4105-A362-A2ED89388CC8}' -ErrorAction SilentlyContinue

# force the garbage collector to release any references to registry entries immediately
popd
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

# unload hive
reg unload HKLM\isolatedHive
popd
popd