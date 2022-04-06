Write-Host "How to output a result? Type 'file' or 'console'"
$outputMethod = Read-Host
Write-Host "`n"

$products = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"

foreach ($product in $products)
{
    $propertiesPath = "Registry::$($product)\InstallProperties"
    $isExists = Test-Path -Path $propertiesPath

    if ($isExists)
    {
        $properties = Get-ItemProperty -Path $propertiesPath

        if ($properties.DisplayName -eq "КриптоПро CSP")
        {
            $result = "$($properties.DisplayName) v.$($properties.DisplayVersion): $($properties.ProductId)"

            switch ($outputMethod) 
            {
                "file" { $result | Out-File -FilePath "$($PWD)\CryptoPro $($properties.DisplayVersion) ProductId.txt" }
                "console" { Write-Host $result }
                default { Write-Host $result }
            }
        }
    }
}

Write-Host "`n"
Write-Host "The task has been finished. Press Enter to exit"
Read-Host