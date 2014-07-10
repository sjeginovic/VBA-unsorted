Sub VOCTest()
'
' VOCTest Macro
'
'
    Call Intro
    Call Req1
End Sub

'
' VOC-Consultancy rapport tools
' 10-7-2014 Daniel van den Akker
'


' Function to generate in the word file, used for finding the fieldbox numbers
Private Sub writeNumbers_Click()
    Dim myFile As String
    Dim text As String
    Dim textline As String
    

    For Each aField In doc.FormFields
        If aField.Type = wdFieldFormTextInput Then
            count = count + 1
            text = ""
            
            myFile = ActiveDocument.Path & "\txt\req_1_box" & count & ".txt"
            Open myFile For Input As #1
    
            Do Until EOF(1)
                Line Input #1, textline
                text = text & textline
            Loop
            Close #1
            
            aField.Result = count & ": " & text
        End If
    Next aField
    MsgBox "There are " & count & " text boxes in this document"
End Sub

' Generic fuction to pull txt to a word doc
Private Sub txttodoc(prefix As String, start As Long)
    Dim myFile As String            ' File string used for read loop
    Dim text As String              ' Text from txt to import to document
    Dim textline As String          ' Textline for reading file loop
    Dim fileNr                      ' File itteration for read loop
    Dim count                       ' Total count itteration
    Dim done As Boolean             ' Bool for reading file loop
    Dim atextfield()                ' Dynmic array with vallues to import
    Dim amount As Long              ' Amount of text fields for this section

    ' Intial value
    fileNr = 0
    amount = 0
    ReDim atextfield(0)

    Do While Not done
        fileNr = fileNr + 1
        myFile = ActiveDocument.Path & "\txt\" & prefix & "\" & prefix & "_box" & fileNr & ".txt"
        If (Dir(myFile) > "") Then
            amount = amount + 1
            ReDim atextfield(amount)
            Open myFile For Input As #1     ' Open file
            Do Until EOF(1)                 ' Read all lines
                Line Input #1, textline     ' Read one line
                text = text & textline      ' Add line to text
            Loop
            Close #1                        ' Close file
            atextfield(amount - 1) = text
        Else
            done = True
        End If
    Loop

    ' Reset
    fileNr = 0
    done = False
    ' Go through all formtextinput fields
    For Each aField In ActiveDocument.FormFields
        If aField.Type = wdFieldFormTextInput Then
            count = count + 1
            ' Stop at formtextinput for this section
            If count >= start And count <= start + amount Then
                fileNr = fileNr + 1     ' Start counting the files
                text = ""               ' Reset value after last loop
                ' Set the file name to read
                myFile = ActiveDocument.Path & "\txt\" & prefix & "\" & prefix & "_box" & fileNr & ".txt"
                

                
                ' Set word file formtextinput field to what was in the file
                aField.Range.Fields(1).Result.text = text
            End If
            If count > start + amount Then
                Exit For
            End If
        End If
    Next aField
    ' informational result of function
    MsgBox "Completed " & prefix & " Changed " & fileNr & " of text boxes in this document"
End Sub

Private Sub Intro()
    Call txttodoc("intro", 3)
End Sub

Private Sub Req1()
    Call txttodoc("req_1", 348)
End Sub


