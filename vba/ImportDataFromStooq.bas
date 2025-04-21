Attribute VB_Name = "ImportDataFromStooq"
Option Explicit

Sub ImportStockDataFromStooq()
    Dim ticker As String
    Dim startDate As String
    Dim endDate As String
    Dim interval As String
    Dim urlCSV As String
    Dim folderPath As String
    Dim tempFile As String
    Dim xml As Object
    Dim wsDest As Worksheet
    Dim intervalCode As String
    Dim fileNum As Integer
    Dim sheetName As String

    ' Disable screen updating and automatic calculations to improve performance
    OptimizePerformance True

    ' Get the stock symbol and dates from the named cells
    ticker = GetCleanTicker
    startDate = GetFormattedDate("DLStartDate")
    endDate = GetFormattedDate("DLEndDate")
    interval = GetInterval

    ' Validate the interval
    If Not ValidateInterval(interval, intervalCode) Then
        MsgBox "Invalid interval selected. Please select a valid interval (Daily, Weekly, Monthly, Quarterly, Yearly).", vbCritical
        GoTo Cleanup
    End If

    ' Build the URL for downloading the CSV using the first letter of the interval
    urlCSV = BuildUrl(ticker, startDate, endDate, intervalCode)

    ' Temporary folder path for saving the CSV file
    folderPath = Environ$("TEMP") ' Using the temporary folder of Windows
    tempFile = folderPath & "\stooq_data.csv" ' Full path for the CSV file

    ' Download the CSV file using XMLHTTP
    Set xml = CreateObject("MSXML2.XMLHTTP")
    If Not DownloadCsv(xml, urlCSV, tempFile) Then
        MsgBox "Error downloading data. Please check the URL or try again later.", vbCritical
        GoTo Cleanup
    End If

    ' Create a new worksheet to import the data
    Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    ' Construct a compact sheet name
    sheetName = BuildSheetName(ticker, startDate, endDate, interval)

    ' Set the sheet name, handling potential errors
    On Error Resume Next
    wsDest.Name = Left(sheetName, 31) ' Truncate to 31 characters to fit Excel's sheet name limit
    On Error GoTo 0

    ' Import the data from the CSV file into the new worksheet
    ImportCsvData wsDest, tempFile

    ' Delete the temporary CSV file
    On Error GoTo FileError
    Kill tempFile

    MsgBox "Data successfully imported for ticker: " & UCase(ticker), vbInformation

Cleanup:
    ' Re-enable screen updating and automatic calculations
    OptimizePerformance False
    Set xml = Nothing
    Exit Sub

FileError:
    MsgBox "Error handling the temporary file. Please check the file path or permissions.", vbCritical
    Resume Cleanup
End Sub

Function GetCleanTicker() As String
    GetCleanTicker = LCase(Trim(Range("DLSymbol").Value)) ' User manually enters the symbol
End Function

Function GetFormattedDate(cellName As String) As String
    GetFormattedDate = Format(Range(cellName).Value, "yyyymmdd") ' Format the date
End Function

Function GetInterval() As String
    GetInterval = Trim(Range("DLInterval").Value) ' Get the interval value
End Function

Function ValidateInterval(ByVal interval As String, ByRef intervalCode As String) As Boolean
    If Len(interval) = 0 Then
        MsgBox "Interval is empty. Please select a valid interval (Daily, Weekly, Monthly, Quarterly, Yearly).", vbCritical
        ValidateInterval = False
        Exit Function
    End If

    intervalCode = LCase(Left(interval, 1))
    ValidateInterval = InStr("dwmqy", intervalCode) > 0
End Function

Function BuildUrl(ticker As String, startDate As String, endDate As String, intervalCode As String) As String
    BuildUrl = "https://stooq.com/q/d/l/?s=" & ticker & "&f=" & startDate & "&t=" & endDate & "&i=" & intervalCode
End Function

Function DownloadCsv(xml As Object, urlCSV As String, tempFile As String) As Boolean
    xml.Open "GET", urlCSV, False
    xml.Send

    If xml.Status = 200 Then
        Dim fileNum As Integer
        fileNum = FreeFile
        On Error GoTo FileError
        Open tempFile For Output As #fileNum
        Print #fileNum, xml.responseText
        Close #fileNum
        DownloadCsv = True
    Else
        DownloadCsv = False
    End If
    Exit Function

FileError:
    DownloadCsv = False
End Function

Sub ImportCsvData(wsDest As Worksheet, tempFile As String)
    With wsDest.QueryTables.Add(Connection:="TEXT;" & tempFile, Destination:=wsDest.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = xlWindows
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub OptimizePerformance(enable As Boolean)
    Application.ScreenUpdating = Not enable
    Application.Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
End Sub

Function BuildSheetName(ticker As String, startDate As String, endDate As String, interval As String) As String
    BuildSheetName = UCase(ticker) & "_" & Format(Range("DLStartDate").Value, "yyyy-mm") & "_" & Format(Range("DLEndDate").Value, "yyyy-mm") & "_" & Left(interval, 1)
End Function

