'
' VOC-Consultancy rapport tools
' 10-7-2014 Daniel van den Akker
'

Sub makeTXT(prefix As String, maxcount As Long)
   
    Dim myFile As String    ' Variable for file name in file loop
    Dim count As Long       ' Count up for file loop
    Dim aTexboxes()         ' Array with all cells
    Dim aField              ' Text of field to write, used in file loop
    
    ReDim aTexboxes(maxcount - 1)
      
    ' Fill array with data from excel sheet
    For i = 0 To maxcount - 1
        aTexboxes(i) = Cells(i + 1, 1)
    Next i
    
    ' Write array to files
    For Each aField In aTexboxes
        count = count + 1
        myFile = ".\txt\" & prefix & "\" & prefix & "_box" & count & ".txt" ' Formats as ./txt/Req X/Req X_box
        ' write
        Open myFile For Output As #1
        Print #1, aField
        Close #1
    Next aField
    ' informational respone
    MsgBox "There are " & count & " text boxes in this document"
End Sub

Sub tocompanytofile_Click()
    Call makeTXT("intro", 43)
End Sub
Sub tofile_Click()
    Call makeTXT("req_1", 84)
End Sub
