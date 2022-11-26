# Seiichi Ariga <seiichi.ariga@gmail.com>
# .\output 内のExcelファイルを印刷する
# Windowsのみ対応

$currentDir = Get-Location
$workDir = $currentDir.Path + "\output\"

# Get-ChildItemはオブジェクトを返すので、ForEach-Objectが必要っぽい
$fileList = Get-ChildItem $workdir -Filter *.xlsx | ForEach-Object { "$_" }

Write-Output "以下のファイルを印刷します"
Write-Output $fileList

# 各ファイルに対してExcelを起動して印刷(Windowsのみ)
ForEach($file in $fileList) {
    Start-Process -FilePath $file -Verb print
}
