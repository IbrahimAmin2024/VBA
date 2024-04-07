Attribute VB_Name = "functions"
Option Explicit

' Notes: Functions must be top of any calls

Function ValidEntry() As Boolean

 On Error GoTo ErrorHandler
    ValidEntry = True
    
    With cpanel
    
        'Default Color
        
        '.em_txt.BackColor = vbBlack
        '.code_txt.BackColor = vbBlack
        
        'Validating Student's Name
        
        If Trim(.em_txt.Value) = "" Then
        
            .em_txt.BackColor = vbRed
            MsgBox "Please enter Employee's name.", vbOKOnly + vbInformation, "Employee's Name"
            
            .em_txt.SetFocus
            
            ValidEntry = False
            Exit Function
        
        
        End If
        
        'Validating Fathers's Name
        
        If Trim(.code_txt.Value) = "" Then
            
            .code_txt.BackColor = vbRed
            MsgBox "Please enter Code.", vbOKOnly + vbInformation, "Code"
            
            .code_txt.SetFocus
            
            ValidEntry = False
            Exit Function
        
        
        End If
        

        'Validating Image
        
        If .img.Picture Is Nothing Then
            If .img_status.Value = True Then
            .img.BorderColor = vbRed
            MsgBox "Please upload the PP Size Photo.", vbOKOnly + vbInformation, "Picture"
            .img.BorderColor = vbBlack
            ValidEntry = False
            Else
            ValidEntry = True
            End If
            
            Exit Function
            
        
        End If
        
    
    End With

    Exit Function ' Exit the subroutine if no error occurs
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code

End Function

Function GetImagePath() As String

 On Error GoTo ErrorHandler
 
    GetImagePath = ""
    
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .AllowMultiSelect = False
        
        .Filters.Add "Imgs", "*.gif;*.jpg;*.jpeg"
        
        If .Show <> 0 Then
        
        
            GetImagePath = .SelectedItems(1)
        
        End If
    
    End With
    
       Exit Function ' Exit the subroutine if no error occurs
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code

End Function


Function id() As Long

' Start error handler
On Error GoTo ErrorHandler
'Get Rows of DB
Dim rows_db As Long
Dim db As Worksheet

'Set A Refrence to "Database" sheet on thisworkbook
Set db = ThisWorkbook.Sheets("Database")

rows_db = db.Cells(db.rows.Count, "A").End(xlUp).row

If rows_db = 1 Then
id = 1
ElseIf rows_db > 1 Then
id = db.Cells(db.rows.Count, "A").End(xlUp).row
End If

ErrorHandler:
Resume Next:

End Function
