[string]$filePath = "C:\test\pastel2.txt";

[string]$directory = [System.IO.Path]::GetDirectoryName($filePath);
[string]$strippedFileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath);
[string]$extension = [System.IO.Path]::GetExtension($filePath);
[string]$newFileName = $strippedFileName + "-" + [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + $extension;
[string]$newFilePath = [System.IO.Path]::Combine($directory, $newFileName);

Move-Item -LiteralPath $filePath -Destination $newFilePath;