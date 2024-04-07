Attribute VB_Name = "db_controller"
Option Explicit
Public Sub Reset_Form()

On Error GoTo ErrorHandler
    Dim rows_db As Long
    Dim db As Worksheet

    ' Set a reference to the "DB" sheet in ThisWorkbook
    Set db = ThisWorkbook.Sheets("Database")


    With cpanel
    
        .em_txt.Value = ""
        .code_txt.Value = ""
        .shift_combo.Value = ""
        .job_combo.Value = ""
        .activity_combo.Value = ""
        .notes_txt.Value = ""
        .img.Picture = LoadPicture(vbNullString)
        
        .img_load_btn.Enabled = False ' Disable the button
        
        ' Assigning RowSource to the ListBox {Clear listbox}
        .db.RowSource = ""
        
        ' Assigning properties to the ListBox
        .db.ColumnCount = 10
        .db.ColumnHeads = True
        .db.ColumnWidths = "30,150,100,100,100,100,100,100,120,125"

        ' Identify the last non-blank row in column A of the "Database" sheet
        rows_db = db.Cells(db.rows.Count, "A").End(xlUp).row
        
        ' Show DB Row's Counter
    
        .Frame2.Caption = "Database: " & rows_db - 1 & " |User's"
 
        If rows_db > 1 Then
            ' If there's more than one row of data, set RowSource accordingly
            .db.RowSource = "Database!A2" & ":J" & rows_db
        ElseIf rows_db = 1 Then
 
        ' If there's only one row, include that in the RowSource
        .db.RowSource = "Database!A2:J2"
        
        Else
            ' If there's no data, clear the ListBox content
            .db.RowSource = ""
        End If
        
        ' Adding an item to a combo box
        .shift_combo.Clear
        .shift_combo.AddItem "1"
        .shift_combo.AddItem "2"
        .shift_combo.AddItem "3"
        .shift_combo.ListIndex = 0
        
        .job_combo.Clear
        .job_combo.AddItem "1"
        .job_combo.AddItem "2"
        .job_combo.AddItem "3"
        .job_combo.ListIndex = 0
        
        .activity_combo.Clear
        .activity_combo.AddItem "1"
        .activity_combo.AddItem "2"
        .activity_combo.AddItem "3"
        .activity_combo.ListIndex = 0
        
        .id_txt = id
    End With
    
        Exit Sub ' Exit the subroutine if no error occurs
ErrorHandler:
    Resume Next ' Resume execution after handling the error, 'Resume Next' continues execution at the next line of code

End Sub

