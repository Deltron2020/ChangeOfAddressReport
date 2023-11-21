$path = "\\network_path\Change_Of_Address"
$xlsFile = (Get-ChildItem -Path $path -force | Where-Object Extension -in ('.xls')).FullName
$xlsCount = (Get-ChildItem -Path $path -force | Where-Object Extension -in ('.xls') | Measure-Object).Count

if ($xlsCount -eq 1) {

     py "\\network_path\ChangeOfAddress\main.py" $xlsFile

 }

else {

    exit

 }