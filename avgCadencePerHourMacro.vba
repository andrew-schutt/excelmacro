Sub CadenceAvgPerHour()
    'cleanup values
    Range("J3:J60000").ClearContents
    Range("K3:K60000").ClearContents
    Range("L3:L60000").ClearContents
    
    'start columns (won't change)
    Dim cadenceCol, timeCol, avgCadPerHourCol, numRows As Single
    cadenceCol = 3: timeCol = 1: avgCadPerHourCol = 10
    'total columns holding time data
    numRows = Application.CountA(Range("A:A"))
    'variables needed for looping
    Dim i, cadCount, k, sum As Single
    Dim cadenceHour, nextCadenceHour As String
    
    'initalize loop variables
    cadCount = 0
    sum = 0
    
    For i = 3 To numRows
        k = i + 1
        cadenceHour = Mid(Cells(i, timeCol), 13, 2)
        nextCadenceHour = Mid(Cells(k, timeCol), 13, 2)
        cadenceVal = Cells(i, cadenceCol)
        
        If cadenceHour <> "" Then
            k = i + 1
            If cadenceHour < nextCadenceHour And cadCount <> 0 Then
                Cells(i, avgCadPerHourCol) = (sum / cadCount)
                Cells(i, 11) = sum
                Cells(i, 12) = cadCount
                sum = 0
                cadCount = 0
            ElseIf cadenceHour < nextCadenceHour And cadCount = 0 Then
                Cells(i, avgCadPerHourCol) = 0
                Cells(i, 11) = sum
                Cells(i, 12) = cadCount
            End If
            If cadenceVal <> "0" Then
                sum = sum + cadenceVal
                cadCount = cadCount + 1
            End If
        End If
    Next i
End Sub