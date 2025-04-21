Attribute VB_Name = "DataRefreshModule"
' Module-level public variables for managing the refresh cycle
Public nextDataRefreshTime As Date
Public isDataRefreshActive As Boolean

' Starts the automatic data refresh cycle by reading the interval from the named cell "RefreshInterval"
Sub StartDataRefreshCycle()
    Dim refreshIntervalSeconds As Double
    ' Read the refresh interval from the "Settings" sheet named cell "RefreshInterval"
    On Error Resume Next
    refreshIntervalSeconds = Sheets("Settings&Tools").Range("RefreshInterval").Value
    On Error GoTo 0
    
    If refreshIntervalSeconds <= 0 Then
        MsgBox "Please set a valid refresh interval (in seconds) in the named cell 'RefreshInterval'!", vbExclamation
        Exit Sub
    End If
    
    isDataRefreshActive = True
    Debug.Print "Data refresh cycle started at " & Now & " with an interval of " & refreshIntervalSeconds & " seconds."
    ScheduleNextDataRefresh refreshIntervalSeconds
End Sub

' Schedules the next data refresh based on the provided interval (in seconds)
Sub ScheduleNextDataRefresh(ByVal refreshIntervalSeconds As Double)
    nextDataRefreshTime = Now + TimeSerial(0, 0, refreshIntervalSeconds)
    Debug.Print "Next data refresh scheduled for " & nextDataRefreshTime
    Application.OnTime nextDataRefreshTime, "ExecuteDataRefresh"
End Sub

' Executes the data refresh (refreshing all connections) and then schedules the next refresh
Sub ExecuteDataRefresh()
    If isDataRefreshActive Then
        Debug.Print "Executing data refresh at " & Now
        ThisWorkbook.RefreshAll
        Dim refreshIntervalSeconds As Double
        refreshIntervalSeconds = Sheets("Settings&Tools").Range("RefreshInterval").Value
        ScheduleNextDataRefresh refreshIntervalSeconds
    End If
End Sub

' Stops the automatic data refresh cycle
Sub StopDataRefreshCycle()
    isDataRefreshActive = False
    On Error Resume Next
    Application.OnTime earliesttime:=nextDataRefreshTime, procedure:="ExecuteDataRefresh", schedule:=False
    On Error GoTo 0
    Debug.Print "Data refresh cycle stopped at " & Now
End Sub
