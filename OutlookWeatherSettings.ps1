### Set C' for Outlook Calendars  ###
#
# Designed to run on client machine on login, works for both 2013 & 2016/19 x64 Versions, not tested on x32 but should work
#
# V 1.0 13/05/2021 Martyn Smith

## Variables ##
# You will need to replace these paths with a shared directory that can be accessed by your users

$2013Config = "\\SERVERXXX\IT\Scripts\Outlook\Stream_Weather_2_B70BDE50A4BF6540BF14A69167770992-2013.dat"
$2016UpConfig = "\\SERVERXXX\IT\Scripts\Outlook\Stream_Weather_2_B70BDE50A4BF6540BF14A69167770992-2016+.dat"

## Office version detection borrowed from https://superuser.com/questions/1140114/how-to-detect-microsoft-office-version-name ##

$Keys = Get-Item -Path HKLM:\Software\RegisteredApplications | Select-Object -ExpandProperty property
$Product = $Keys | Where-Object {$_ -Match "Excel.Application."}
$OfficeVersion = ($Product.Replace("Excel.Application.","")+".0")

##

$ConfigLocation = "C:\Users\$($env:UserName)\AppData\Local\Microsoft\Outlook\RoamCache\"
cd $ConfigLocation
$ConfigName = Get-ChildItem -Filter *Stream_Weather_*

if ($OfficeVersion.Length -gt 1){  
    if ($OfficeVersion -eq "15.0") {
        copy-item -path $2013Config -Destination "$($ConfigLocation)\$($ConfigName)"
    } elseif($OfficeVersion -eq "16.0") {
         copy-item -path $2016UpConfig -Destination "$($ConfigLocation)\$($ConfigName)"
    } 
}


