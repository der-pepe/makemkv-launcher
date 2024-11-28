##
## MakeMKV Automatic updater
##

$topic=1053
$installPath = "C:\MEDIA\MakeMKV"

$url="https://forum.makemkv.com/forum/viewtopic.php?t=$topic"
$keyRegex = "\<code\>(.*)\<\/code\>"
$valRegex = "valid until end of (.*)\. "
$content = Invoke-WebRequest -Uri $url
$newKey  = [regex]::match($content, $keyRegex).Groups[1].Value
$validUntil = [regex]::match($content, $valRegex).Groups[1].Value

# typo fix
$validUntil = $validUntil.replace('Ocnjber','October')


$expiration = [datetime]::parseexact($validUntil, 'MMMM yyyy', $null).AddMonths(1)


New-ItemProperty -Path 'HKCU:\Software\MakeMKV' -Name 'app_Key' -Value $newKey -Force | out-null

Write-Host "License key:        $newKey"
Write-Host "License expiration: $expiration"

$OldVersion = (Get-Item $installPath\makemkv.exe).VersionInfo.ProductVersion
Write-Host "Installed version:  $OldVersion"

$url="https://makemkv.com/download/"
$pattern = "href\=\'(.*)\'\>MakeMKV (.*) for Windows"
$content = Invoke-WebRequest -Uri $url
$version = [regex]::match($content, $pattern)
$downloadUrl = "https://makemkv.com"+$version.Groups[1].Value
$NewVersion = "v{0}" -f $version.Groups[2].Value

Write-Host "Latest version:     $NewVersion"

if ($OldVersion -ne $NewVersion)
{
  Write-Host " (...Updating...) "
  Invoke-WebRequest -Uri $downloadUrl -OutFile "$env:Temp\Setup_MakeMKV_$NewVersion.exe"
  $installer = Start-Process -NoNewWindow -PassThru -FilePath "$env:Temp\Setup_MakeMKV_$NewVersion.exe" -ArgumentList "/S /NCRC /D=`"$installPath\`""
  Wait-Process -Id $installer.Id -Timeout 60
  Remove-Item -Path "$env:Temp\Setup_MakeMKV_$NewVersion.exe" -Force
}


If ($(Get-Date) -gt $expiration) {
  $dif = ($(Get-Date) - $expiration).Days +2
  Set-Date -Date (Get-Date).AddDays(0-$dif)
  Start-Process -FilePath "$installPath\MakeMKV.exe"
  Start-Sleep 5
  Set-Date -Date (Get-Date).AddDays($dif)
}
else
{
  Start-Process -FilePath "$installPath\MakeMKV.exe"
}

