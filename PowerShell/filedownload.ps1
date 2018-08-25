##cd to folder and clean old files

#$latest=Get-Content D:\trade\downloads\latest -First 1
##Write-Host $latest

#$DirPath = -join("D:\trade\downloads\", $latest)

##Write-Host $DirPath

#if((Test-Path -Path $DirPath )){
#   Write-Host "Downloaded already"
#   return (-7)
#}


if(!(Test-Path -Path "D:\trade\downloads\today" )){

    New-Item -ItemType directory -Path "D:\trade\downloads\today"

}
else
{
    Get-ChildItem D:\trade\downloads\today -Recurse | Remove-Item
}

##run python code
python D:\Users\shet\PycharmProjects\Trade1\Trade1.py

if(!(Test-Path -Path "D:\trade\downloads\today\date" ))
{

    retun (-2)

}

if(!(Test-Path -Path "D:\trade\downloads\today\data.csv" ))
{

    retun (-3)

}



##extract links from the file
$FILE = Get-Content D:\trade\downloads\today\date
$date, $cm_url, $fo_url, $mto_url = $FILE -split '\n'

##download files from URL
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($cm_url, "D:\trade\downloads\today\cm.csv.zip")
$wc.DownloadFile($fo_url, "D:\trade\downloads\today\fo.csv.zip")
$wc.DownloadFile($mto_url, "D:\trade\downloads\today\mto.csv")
$mto_file='D:\trade\downloads\today\mto.csv'
(Get-Content $mto_file | Select-Object -Skip 4) | Set-Content $mto_file
$a = Get-Content $mto_file
$b = 'RT,SN,SYMBOL,TYPE,DQG,DQ3,TP'
Set-Content $mto_file -value $b,$a


if(!(Test-Path -Path "D:\trade\downloads\today\cm.csv.zip" ))
{

    retun (-4)

}

if(!(Test-Path -Path "D:\trade\downloads\today\fo.csv.zip" ))
{

    retun (-5)

}

if(!(Test-Path -Path "D:\trade\downloads\today\mto.csv" ))
{

    retun (-6)

}


##unzip files
Expand-Archive -Path D:\trade\downloads\today\cm.csv.zip -DestinationPath D:\trade\downloads\today
Expand-Archive -Path D:\trade\downloads\today\fo.csv.zip -DestinationPath D:\trade\downloads\today
Remove-Item D:\trade\downloads\today\cm.csv.zip
Remove-Item D:\trade\downloads\today\fo.csv.zip
Get-ChildItem D:\trade\downloads\today | Where { $_.PSChildName -match "cm+"} | Rename-Item -newname cm.csv
Get-ChildItem D:\trade\downloads\today | Where { $_.PSChildName -match "fo+"} | Rename-Item -newname fo.csv
#archive
$date=Get-Content D:\trade\downloads\today\date -First 1
$folder_path=-join("D:\trade\downloads\",$date)
    if(!(Test-Path -Path $folder_path )){
        $dst=-join($folder_path,"\")
        Copy-Item -Path "D:\trade\downloads\today\" -Destination "$dst" -recurse
    }