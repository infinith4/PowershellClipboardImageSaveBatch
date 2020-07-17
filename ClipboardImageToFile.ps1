
Add-Type -AssemblyName System.Windows.Forms
$clipboardImage = [Windows.Forms.Clipboard]::GetImage()
if ($clipboardImage -ne $null)
{
  $dateTimeStr = Get-Date -Format "yyyyMMdd_HHmmss"
  $fileName = $dateTimeStr + "_clip.png"
  Write-Output $fileName
  $picturesDirPath = Join-Path $Env:UserProfile "Pictures"
  $clipboardImagesDirPath = Join-Path $picturesDirPath "ClipboardImages"
  Write-Output $clipboardImagesDirPath
  New-Item -ItemType Directory -Force -Path $clipboardImagesDirPath
  $outputFilePath = Join-Path $clipboardImagesDirPath $fileName
  Write-Output $outputFilePath
  $clipboardImage.Save($outputFilePath)
}