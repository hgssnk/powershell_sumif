#[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName System.Windows.Forms

# 出力先
$OUTPUT_CSV_PATH = ".\output.csv"

echo "Please select a file...."
Start-Sleep 2

# 入力ファイル選択
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'すべてのファイル|*'
    Title = '入力ファイルを選択'
}
if($FileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
    $INPUT_CSV_PATH = $FileBrowser.FileName
}

if ( Test-Path $INPUT_CSV_PATH )
{
  echo "Please wait...."
  echo ""

  # csv読み込み（タブ区切り）
  $in_csv = Import-Csv $INPUT_CSV_PATH -Delimiter "`t" -header @(1..2) 

  # 2列目でグループ化し、1列目の合計をとり、降順にソートs
  $sum = $in_csv | group "2" |select "Name",@{Name="Totals";Expression={($_.group | Measure-Object -sum "1").sum}} |Sort Totals -Desc

  # 結果をcsv出力
  $sum |Export-Csv -path $OUTPUT_CSV_PATH -Delimiter `t -Encoding UTF8

  # csvのヘッダを取り除く
  $out_csv = Get-Content $OUTPUT_CSV_PATH
  $out_csv[0] = $null
  $out_csv[1] = $null
  $out_csv | Out-File $OUTPUT_CSV_PATH

  echo 'SUCCESS!!'
  $CurrentDir = Split-Path $MyInvocation.MyCommand.Path
  $moge = $CurrentDir + $OUTPUT_CSV_PATH
  echo $moge
  $hoge = Read-Host
}
else
{
  echo 'no such file or directory'
  $hoge = Read-Host
}
