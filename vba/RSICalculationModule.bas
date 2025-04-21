Attribute VB_Name = "RSICalculationModule"
Option Explicit

Sub ImportAndCalculateRSI()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstEmptyCol As Long
    Dim period As Long
    Dim firstAvgRow As Long
    Dim changeCol As Long, gainCol As Long, lossCol As Long
    Dim avgGainCol As Long, avgLossCol As Long, rsCol As Long, rsiCol As Long
    Dim rsiRange As Range
    Dim ch As ChartObject
    Dim rsiStartRow As Long
    Dim sheetName As String

    ' Disable screen updating and automatic calculations for performance improvement
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Step 1: Import stock data
    On Error GoTo ErrorHandler
    Call ImportStockDataFromStooq

    ' Step 2: Use the newly created active sheet
    Set ws = ActiveSheet

    ' Step 3: Read RSI period from the named range "RSITimeFrame"
    On Error Resume Next
    period = CLng(ThisWorkbook.Names("RSITimeFrame").RefersToRange.Value)
    On Error GoTo ErrorHandler

    If period <= 0 Then
        MsgBox "Please set a valid RSI period (positive integer) in the named cell 'RSITimeFrame'.", vbExclamation
        GoTo Cleanup
    End If

    With ws
        ' Step 4: Remove columns Open, High, Low (B:D), then Volume (new C)
        .Range("B:D").Delete
        .Range("C:C").Delete

        ' Step 5: Find last row of data
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        ' Step 6: Replace "." with "," in the Close column (column B)
        .Range("B2:B" & lastRow).Replace What:=".", Replacement:=",", _
            LookAt:=xlPart, SearchOrder:=xlByRows

        ' Step 7: Find first empty column on row 1
        firstEmptyCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1

        ' Define column positions
        changeCol = firstEmptyCol
        gainCol = changeCol + 1
        lossCol = changeCol + 2
        avgGainCol = changeCol + 3
        avgLossCol = changeCol + 4
        rsCol = changeCol + 5
        rsiCol = changeCol + 6

        ' Step 8: Add headers
        .Cells(1, changeCol).Value = "Change"
        .Cells(1, gainCol).Value = "Gain"
        .Cells(1, lossCol).Value = "Loss"
        .Cells(1, avgGainCol).Value = "Average Gain"
        .Cells(1, avgLossCol).Value = "Average Loss"
        .Cells(1, rsCol).Value = "RS"
        .Cells(1, rsiCol).Value = "RSI"

        ' Step 9: Fill formulas
        ' Change = Close - Previous Close
        .Range(.Cells(3, changeCol), .Cells(lastRow, changeCol)).FormulaR1C1 = _
            "=RC2 - R[-1]C2"

        ' Gain = MAX(Change, 0)
        .Range(.Cells(3, gainCol), .Cells(lastRow, gainCol)).FormulaR1C1 = _
            "=MAX(RC[-1], 0)"

        ' Loss = MAX(-Change, 0)
        .Range(.Cells(3, lossCol), .Cells(lastRow, lossCol)).FormulaR1C1 = _
            "=MAX(-RC[-2], 0)"

        ' Start computing RSI from the row after the full period
        firstAvgRow = 2 + period

        ' Average Gain
        .Range(.Cells(firstAvgRow, avgGainCol), .Cells(lastRow, avgGainCol)).FormulaR1C1 = _
            "=AVERAGE(R[-" & (period - 1) & "]C[-2]:RC[-2])"

        ' Average Loss
        .Range(.Cells(firstAvgRow, avgLossCol), .Cells(lastRow, avgLossCol)).FormulaR1C1 = _
            "=AVERAGE(R[-" & (period - 1) & "]C[-2]:RC[-2])"

        ' RS = Avg Gain / Avg Loss
        .Range(.Cells(firstAvgRow, rsCol), .Cells(lastRow, rsCol)).FormulaR1C1 = _
            "=RC[-2]/RC[-1]"

        ' RSI = 100 - 100 / (1 + RS)
        .Range(.Cells(firstAvgRow, rsiCol), .Cells(lastRow, rsiCol)).FormulaR1C1 = _
            "=100 - 100 / (1 + RC[-1])"

        ' Step 10: Autofit for readability
        .Range(.Cells(1, firstEmptyCol), .Cells(1, rsiCol)).EntireColumn.AutoFit

        ' Format RSI and RS to 2 decimal places
        .Range(.Cells(2, rsiCol), .Cells(lastRow, rsiCol)).NumberFormat = "0.00"
        .Range(.Cells(2, rsCol), .Cells(lastRow, rsCol)).NumberFormat = "0.00"

        ' Set the start row for RSI (used in the chart later)
        rsiStartRow = firstAvgRow
    End With

    ' === APPLY CONDITIONAL FORMATTING TO RSI COLUMN ===
    With ws
        Set rsiRange = .Range(.Cells(rsiStartRow, rsiCol), .Cells(lastRow, rsiCol))

        ' Clear existing conditional formats
        rsiRange.FormatConditions.Delete

        ' RSI > 70 = GREEN
        With rsiRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="70")
            .Interior.Color = RGB(0, 255, 0)
        End With

        ' RSI < 30 = RED
        With rsiRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="30")
            .Interior.Color = RGB(255, 0, 0)
        End With

        ' 30 <= RSI <= 70 = YELLOW
        With rsiRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="30", Formula2:="70")
            .Interior.Color = RGB(255, 255, 0)
        End With
    End With

    ' Force the recalculation of all formulas (to avoid "Valeur périmée")
    Application.Calculate

    ' Step 11: Add RSI Thresholds legend and interpretation to the right (cells, top)
    With ws
        ' Positioning the threshold legend to the right of the data
        .Cells(1, rsiCol + 2).Value = "RSI Thresholds"
        .Cells(2, rsiCol + 2).Value = "RSI > 70 (Overbought)"
        .Cells(2, rsiCol + 2).Interior.Color = RGB(0, 255, 0) ' Green for 70
        .Cells(3, rsiCol + 2).Value = "RSI < 30 (Oversold)"
        .Cells(3, rsiCol + 2).Interior.Color = RGB(255, 0, 0) ' Red for 30
        .Cells(4, rsiCol + 2).Value = "30 <= RSI <= 70 (Neutral)"
        .Cells(4, rsiCol + 2).Interior.Color = RGB(255, 255, 0) ' Yellow for 30-70
    End With

    ' Step 12: Create the RSI chart to the right (in rsiCol + 2)
    Set ch = ws.ChartObjects.Add(Left:=ws.Cells(1, rsiCol + 2).Left, Width:=600, Top:=100, Height:=300)

    With ch.Chart
        .ChartType = xlLine
        .SetSourceData Source:=ws.Range(ws.Cells(rsiStartRow, rsiCol), ws.Cells(lastRow, rsiCol))
        .HasTitle = True
        .ChartTitle.Text = "RSI Over Time"

        ' Set the X-axis to use dates from column A
        .Axes(xlCategory).CategoryNames = ws.Range(ws.Cells(rsiStartRow, 1), ws.Cells(lastRow, 1))

        ' Set the minimum and maximum values for the Y-axis
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 100

        ' Add a trendline
        .SeriesCollection(1).Trendlines.Add Type:=xlLinear, Name:="RSI Trendline"
        .SeriesCollection(1).Trendlines(1).DisplayEquation = False
        .SeriesCollection(1).Trendlines(1).DisplayRSquared = False

        ' Add the trendline to the legend
        .SeriesCollection(1).Name = "RSI"
        .SeriesCollection(1).Trendlines(1).Name = "RSI Trendline"
        .HasLegend = True
    End With

    ' Construct a compact sheet name with "RSI"
    sheetName = UCase(ws.Name) & "_RSI"

    ' Set the sheet name, handling potential errors
    On Error Resume Next
    ws.Name = Left(sheetName, 31) ' Truncate to 31 characters to fit Excel's sheet name limit
    On Error GoTo 0

    ' Freeze the first row
    ws.Activate
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

    ' Resize columns A to K
    ws.Range("A:K").EntireColumn.AutoFit

    MsgBox "RSI calculation completed on sheet: " & ws.Name, vbInformation

Cleanup:
    ' Re-enable screen updating and automatic calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    GoTo Cleanup
End Sub


