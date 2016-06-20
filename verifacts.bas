 
Sub Stop_Button()
End
End Sub

Sub CLEAR()
Set aRange = Sheets("Sheet1").Range("A6.AL50000")
aRange.ClearContents
End Sub


Sub Add_Location()
  Dim NOTES As String
  NOTES = "CREDIT AR"
  Set HE = CreateObject("HostExplorer")
  Set CurrentHost = HE.CurrentHost
  Dim irow As Long
  irow = 6


Do
    If Range("A" & irow).Value = "" Then
    Application.StatusBar = "POE ADD COMPLETE"
    MsgBox "RUN COMPLETE"
    Exit Sub
    End If
    
    file = Range("A" & irow).Value
    POE = Range("B" & irow).Value
    CITY = Range("C" & irow).Value
    State = Range("D" & irow).Value
    ZIP = Range("E" & irow).Value
    PHONE = Range("F" & irow).Value
    NOTE = Range("G" & irow).Value
        
    icol = -1
      
    CurrentHost.pause 600
    CurrentHost.Keys (file)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("2^M")
    CurrentHost.pause 600
    CurrentHost.Keys (POE)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("/4^M")
    CurrentHost.pause 600
    CurrentHost.Keys (CITY)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys (State)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys (ZIP)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys (PHONE)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("//^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("4^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("16^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("17^M")
    CurrentHost.pause 600
    CurrentHost.Keys (NOTE)
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("5^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("17^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("//^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("/^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("16^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("17^M")
    CurrentHost.pause 600
    CurrentHost.Keys ("12^M")
    CurrentHost.pause 600

    Application.StatusBar = "Processing borrower " & file
   
    
irow = irow + 1

Loop
    
End Sub

