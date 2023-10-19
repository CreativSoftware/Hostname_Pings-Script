
$hostnames = Import-Excel -Path .\hostname_list.xlsx

$pingable = New-Object System.Collections.ArrayList
$nonpingable = New-Object System.Collections.ArrayList

foreach ($name in $hostnames) {
    if (Test-Connection -TargetName $name.Name -Count 1 -Quiet) {
        $pingable.Add($name) | Out-Null
    } else {
        if (Test-Connection -TargetName $name.Address -Count 1 -Quiet) {
            $pingable.Add($name) | Out-Null
        } else {
            $nonpingable.Add($name) | Out-Null
        }
    }
}

$pingable | Export-Excel -Path .\Successful.xlsx
$nonpingable | Export-Excel -Path .\Unsuccessful.xlsx
