'
' VOC-Consultancy rapport tools
' 10-7-2014 Daniel van den Akker
'

Sub makeTXT(prefix As String)
   
    Dim myFile As String    ' Variable for file name in file loop
    Dim count As Long       ' Count up for file loop
    Dim aTextboxes()         ' Array with all cells
    Dim aField              ' Text of field to write, used in file loop
    Dim done As Boolean     ' for Dynamic array loop
    Dim i As Integer        ' counter for array loop
    
    ' initial value
    i = 0
    maxcount = 0
    done = False
      
    ' Fill array with data from excel sheet
    Do While Not done

        If Not IsEmpty(Cells(i + 1, 1)) Then
            maxcount = maxcount + 1
            ReDim Preserve aTextboxes(maxcount - 1)
            aTextboxes(i) = Cells(i + 1, 1)
        Else
            done = True
        End If
        i = i + 1
    Loop
    
    ' Write array to files
    For Each aField In aTextboxes
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
    Call makeTXT("intro")
End Sub
Sub tofile_Click()
    Call makeTXT("req_1")
End Sub
