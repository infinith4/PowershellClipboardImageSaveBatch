
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
  # #保存形式と保存画質の初期設定
  # $encoderQuality = [System.Drawing.Imaging.Encoder]::Quality
  # $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
  # $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($myEncoder, $quality)
  $clipboardImage.Save($outputFilePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
}