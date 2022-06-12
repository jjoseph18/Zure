<#
.SYNOPSIS
    Takes invoice information from a spreadsheet

.DESCRIPTION
    Extracts the text from an excel spreadsheet and formats it onto a webpage

.EXAMPLE
    //

.OUTPUTS
    //

.NOTES
    //http://woshub.com/read-write-excel-files-powershell/
    //https://www.stevefenton.co.uk/2020/04/extract-an-excel-column-to-a-text-file-with-powershell/
    //https://www.zoho.com/invoice/templates/excel-invoice-template/
    //https://www.c-sharpcorner.com/blogs/insert-data-into-sql-server-table-using-powershell
#>

#establish SQL connection
$Ssms = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe"
$ServerName = "exceldbserver.database.windows.net"
$DataBaseName= "exceldb"
$ADUser = "aderant"
$ADPass = "Ad3rant!123"
$TableName = "Invoicing"

$Connection = New-Object System.Data.SqlClient.SqlConnection
$Connection.ConnectionString = "Server=$ServerName;database=$DataBaseName;User ID=$ADUser;Password=$ADPass;"
$Connection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $Connection



$SourceFile = "C:\Users\james\OneDrive\Documents\Aderant\Invoice.xlsx"

$ExcelAPP = New-Object -comobject Excel.Application
#Opens the whole excel workbook
$WorkBook = $ExcelAPP.Workbooks.Open($SourceFile)
#Selects the worksheet by name
$WorkSheet = $WorkBook.Sheets.Item("Sheet1")
#Count the total amount of rows used
$RowCount = $WorkSheet.UsedRange.Rows.Count
#Loop through all the rows and collect the data from each column
for($i=2;$i -le $RowCount;$i++){
    $BillTo = $WorkSheet.Range("A$i").Text
    $Description = $WorkSheet.Range("B$i").Text
    $Date = $WorkSheet.Range("C$i").Text
    $Hours = $WorkSheet.Range("D$i").Text
    $Amount = $WorkSheet.Range("E$i").Text

    #Output the variables onto a text file
    $insertquery="
    INSERT INTO $TableName
        ([BillTo],[Descriptions],[Dates],[HoursWorked],[Amount])
        VALUES
        ('$BillTo','$Description','$Date','$Hours','$Amount')"

        $Command.CommandText = $insertquery
        $Command.ExecuteNonQuery()
}
$Connection.Close();

