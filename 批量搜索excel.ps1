   $Excel = New-Object -ComObject Excel.Application 
$Excel.Visible = $true
foreach($file_excel in dir C:\Users\Administrator\Desktop\3\2)
{
$WorkBook = $Excel.Workbooks.Open("C:\Users\Administrator\Desktop\3\2\"+$file_excel)
$WorkSheet = $Workbook.Sheets.Item("Sheet1")
$SearchString = "你要搜索的字符串"
#括号内饰要搜索的列范围
$Range = $Worksheet.Range("A1:Z1").EntireColumn
$Search = $Range.find($SearchString)
#echo $Search
if($Search){
echo $file_excel
}
$WorkBook.Close()
}