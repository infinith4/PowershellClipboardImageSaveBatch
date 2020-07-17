
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$clipboardImage = [Windows.Forms.Clipboard]::GetImage()
if ($clipboardImage -ne $null)
{
  $dateTimeStr = Get-Date -Format "yyyyMMdd_HHmmss"
  $fileName = $dateTimeStr + "_clip.jpg"
  #Write-Output $fileName
  $picturesDirPath = Join-Path $Env:UserProfile "Pictures"
  $clipboardImagesDirPath = Join-Path $picturesDirPath "ClipboardImages"
  #Write-Output $clipboardImagesDirPath
  New-Item -ItemType Directory -Force -Path $clipboardImagesDirPath
  $outputFilePath = Join-Path $clipboardImagesDirPath $fileName
  #Write-Output $outputFilePath
  #save as jpeg format
  #http://pc-diary.com/blog/2018/10/windows_powershell.html
  #initialize format and quality
  $quality = 100
  $encoderQuality = [System.Drawing.Imaging.Encoder]::Quality
  $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
  $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($encoderQuality, $quality)
  #codec
  $imageCodecInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|where {$_.MimeType -eq 'image/jpeg'}
  $clipboardImage.Save($outputFilePath, $imageCodecInfo, $encoderParams)
}