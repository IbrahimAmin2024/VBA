Attribute VB_Name = "subs"
Option Explicit

Sub LoadImage()

 On Error GoTo ErrorHandler
    
    ' Your code that might cause an error goes here
    
     Dim imgSourcePath As String ' To store the path of selected image
     
     Dim imgDestination As String 'To store the destination path to create the copy of selected image
     
     imgSourcePath = Trim(GetImagePath())
     
     Call CreateFolder
    
     imgDestination = ThisWorkbook.Path & "\Imgs\" & cpanel.id_txt & _
     "." & Split(imgSourcePath, ".")(UBound(Split(imgSourcePath, ".")))


    FileCopy imgSourcePath, imgDestination
    
    cpanel.img.PictureSizeMode = fmPictureSizeModeStretch
    
    cpanel.img.Picture = LoadPicture(imgDestination)
    
    cpanel.image_path.Value = imgDestination
    
    
    Exit Sub ' Exit the subroutine if no error occurs
    
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code

End Sub

Sub CreateFolder()

 On Error GoTo ErrorHandler
    Dim strFolder As String
    
    strFolder = ThisWorkbook.Path & Application.PathSeparator & "Imgs"
    
    If Dir(strFolder, vbDirectory) = "" Then
    
        MkDir strFolder
        
    End If
    
    Exit Sub ' Exit the subroutine if no error occurs
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code

End Sub


Sub Submit_Data()

 On Error GoTo ErrorHandler
 
    Dim shDatabase As Worksheet
    Dim rows_db As Long
    Dim counter As String
    
    Set shDatabase = ThisWorkbook.Sheets("Database") ' Change "Sheet1" to your sheet name


    If cpanel.Submit_btn.Caption = "Submit" Then
    counter = Replace(cpanel.Frame2.Caption, "Database: ", "")
    counter = Replace(counter, " |User's", "")
    
    rows_db = CLng(counter) + 1
    
    If rows_db = 1 Then ' Corrected the variable name to rows_db
        rows_db = shDatabase.Range("A" & shDatabase.rows.Count).End(xlUp).row + 1 ' Corrected rows.Count to shDatabase.Rows.Count
    Else
        rows_db = rows_db + 1
    End If
    
    ElseIf cpanel.Submit_btn.Caption = "Save" Then
        rows_db = cpanel.db.ListIndex + 2
    End If
    
    
    With shDatabase.Range("A" & rows_db)
        .Offset(0, 0).Value = "=Row()-1"
        .Offset(0, 1).Value = cpanel.em_txt.Value
        .Offset(0, 2).Value = IIf(cpanel.code_txt.Value = "", "0", cpanel.code_txt.Value)
        .Offset(0, 3).Value = cpanel.shift_combo.Value
        .Offset(0, 4).Value = cpanel.job_combo.Value
        .Offset(0, 5).Value = cpanel.activity_combo.Value
        .Offset(0, 6).Value = IIf(cpanel.notes_txt.Value = "", "Empty", cpanel.notes_txt.Value)
        .Offset(0, 7).Value = IIf(cpanel.image_path.Value = "", "Empty", cpanel.image_path.Value)
        .Offset(0, 8).Value = Format([Now()], "DD-MMM-YYYY HH:MM:SS")
        .Offset(0, 9).Value = Format([Now()], "DD-MMM-YYYY HH:MM:SS")
    End With
    
    Reset_Form
    
    Application.ScreenUpdating = True
    
    Dim lastIndex As Integer
    lastIndex = cpanel.db.ListCount - 1 ' Index of the last item
    ' Scrolling to the last item
    cpanel.db.TopIndex = lastIndex
    
    cpanel.em_txt.BackColor = vbWhite
    cpanel.code_txt.BackColor = vbWhite
    

    
    If cpanel.Submit_btn.Caption = "Save" Then
    cpanel.status.ForeColor = vbGreen
    cpanel.status.Caption = "Status : Data Edited Successfully!"
    cpanel.Submit_btn.Visible = False
    cpanel.Reset_btn.Visible = False
    Else
    cpanel.status.ForeColor = vbGreen
    cpanel.status.Caption = "Status : Data Submitted Successfully!"
    cpanel.Submit_btn.Visible = False
    cpanel.Reset_btn.Visible = False
    End If
    
    cpanel.img_status.Value = False
    Exit Sub ' Exit the subroutine if no error occurs
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code
End Sub

Sub Reset()
 On Error GoTo ErrorHandler
 
    Dim Delete_Confirm As VbMsgBoxResult
    Dim db As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set db = ThisWorkbook.Sheets("Database") 'Replace YourTableName with your Database Name

    ' Find the last used row in column A
    lastRow = db.Cells(db.rows.Count, "A").End(xlUp).row
    
    Delete_Confirm = MsgBox("Are you sure you want to reset the database?", vbQuestion + vbYesNo, "Confirmation")
    
    ' Check the user's response
    If Delete_Confirm = vbYes Then
    
    ' Loop through rows in reverse order and delete all rows except the first one
    For i = lastRow To 2 Step -1 ' Start from the last row and move upwards to the second row
        If i <> 1 Then ' Exclude the first row (assuming headers)
            db.rows(i).Delete
        End If
    Next i
    Reset_Form
    
    cpanel.status.ForeColor = vbRed
    cpanel.status.Caption = "Status : Database Account's Deleted Successfully!"
    
    cpanel.Submit_btn.Visible = False
    cpanel.Reset_btn.Visible = False
    
    cpanel.id_txt = id
    End If
        Exit Sub ' Exit the subroutine if no error occurs
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code

End Sub
