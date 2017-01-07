Sub netsh_cmd()
    Dim selectedRange As Range
    Dim IPs(5) As String
    Dim count As Integer
    Dim command As String

    Set selectedRange = Application.Selection
    
    For Each cell In selectedRange.Cells
         IPs(count) = cell.Value
         count = count + 1
    Next cell
    
    If count = 5 Then
        command = "netsh interface ip set address " & Chr(34) & _
        "Local Area Connection" & Chr(34) & " static " & _
        IPs(0) & " " & IPs(1) & " " & IPs(2) & _
        " 1 && netsh interface ip set dns name=" & Chr(34) & "Local Area Connection" & Chr(34) & _
        " source=static addr=" & IPs(3) & " register=primary && netsh interface ip add dns name=" & _
        Chr(34) & "Local Area Connection" & Chr(34) & " addr=" & IPs(4) & " index=2"
    ElseIf count = 3 Then
        command = "netsh interface ip set address " & Chr(34) & _
        "Local Area Connection" & Chr(34) & " static " & _
        IPs(0) & " " & IPs(1) & " " & IPs(2) & " 1"
    End If
    CopyText command
    
End Sub


Sub CopyText(Text As String)
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub