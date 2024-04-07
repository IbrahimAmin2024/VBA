VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cpanel 
   Caption         =   "Cpanel"
   ClientHeight    =   8550.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
   OleObjectBlob   =   "cpanel.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cpanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub new_acc_Click()
    Submit_btn.Visible = True
    Submit_btn.Caption = "Submit"
    Submit_btn.BackColor = vbGreen
    
    Reset_btn.Visible = True
    Reset_btn.Caption = "Reset"
    Reset_btn.BackColor = vbRed
    
    shift_combo.ListIndex = 0
    job_combo.ListIndex = 0
    activity_combo.ListIndex = 0
    
    'new_acc.Visible = False
    
    db.ListIndex = -1 ' Replace ListBox1 with the name of your ListBox control
    
    status.ForeColor = RGB(255, 165, 0) ' RGB values for orange
    status.Caption = "Status : Adding New User..!"
        
    em_txt.Value = ""
    code_txt.Value = ""
    notes_txt.Value = ""
    img.Picture = LoadPicture(vbNullString)
    
    id_txt = id
End Sub
Private Sub Reset_btn_Click()
If Reset_btn.Caption = "Reset" Then
    em_txt.Value = ""
    code_txt.Value = ""
    shift_combo.ListIndex = 0
    job_combo.ListIndex = 0
    activity_combo.ListIndex = 0
    img_status.Value = False
    img.Picture = LoadPicture(vbNullString)
    notes_txt.Value = ""
    
ElseIf Reset_btn.Caption = "Delete" Then

Dim i As VbMsgBoxResult
    
    i = MsgBox("Do you want to delete the selected record?", vbYesNo + vbQuestion, "Delete")
    
    If i = vbNo Then Exit Sub
    
    Dim row As Long
    
    row = db.List(db.ListIndex, 0) + 1
    
    ThisWorkbook.Sheets("Database").rows(row).Delete
    
    Call Reset_Form
    
    status.ForeColor = vbRed
    status.Caption = "Status : Deleted..!"
    
    Submit_btn.Visible = False
    Reset_btn.Visible = False
End If
End Sub

Private Sub Submit_btn_Click()
   If Submit_btn.Caption = "Submit" Then
    
    Dim i As VbMsgBoxResult
    
    i = MsgBox("Do you want to submit the data?", vbYesNo + vbQuestion, "Submit Data")
    
    If i = vbNo Then Exit Sub
    
    If ValidEntry = True Then
    
        Call Submit_Data
    
    End If
    
    ElseIf Submit_btn.Caption = "Save" Then
    
    i = MsgBox("Do you want to Save the data?", vbYesNo + vbQuestion, "Save Data")
    
    If i = vbNo Then Exit Sub
    
    If ValidEntry = True Then
    
        Call Submit_Data
    
    End If
    End If
End Sub

Private Sub UserForm_Initialize()
        id_txt = id
        status.ForeColor = RGB(255, 165, 0) ' RGB values for orange
        status.Caption = "Status : Welcome..!"
        Call Reset_Form
End Sub
Private Sub reset_db_Click()
        Call Reset
End Sub
Private Sub code_txt_Change()
em_txt.BackColor = vbWhite
End Sub
Private Sub em_txt_Change()
em_txt.BackColor = vbWhite
End Sub
Private Sub img_status_Change()
    If img_status.Value = True Then ' If the checkbox is checked
        img_load_btn.Enabled = True ' Enable the button
    Else
        img.Picture = LoadPicture(vbNullString)
        image_path.Value = ""
        img_load_btn.Enabled = False ' Disable the button
    End If
End Sub
Private Sub img_load_btn_Click()
    If Me.em_txt.Value = "" Then
    
        MsgBox "Please enter Employee's name first.", vbOKOnly + vbCritical, "Error"
        
    Else
    
        Call LoadImage
    
    End If
    
End Sub

Private Sub db_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim selectedIndex As Long
    selectedIndex = db.ListIndex
    
    If selectedIndex >= 0 Then ' Check if an item is selected
        Dim colValue As String
        colValue = db.List(selectedIndex, 0) ' Replace 0 with the column index (zero-based) you want
        
        'MsgBox "Selected index: " & selectedIndex & ", Name (Col 1 value): " & colValue
        
        
        new_acc.Visible = True
        
        Submit_btn.Visible = True
        Submit_btn.Caption = "Save"
        Submit_btn.BackColor = RGB(255, 165, 0) ' RGB values for orange
    
        Reset_btn.Visible = True
        Reset_btn.Caption = "Delete"
        Reset_btn.BackColor = vbRed
    
        id_txt.Value = db.List(selectedIndex, 0) 'Col 1 => employee id
        em_txt.Value = db.List(selectedIndex, 1)
        code_txt.Value = db.List(selectedIndex, 2)
        shift_combo.Value = db.List(selectedIndex, 3)
        job_combo.Value = db.List(selectedIndex, 4)
        activity_combo.Value = db.List(selectedIndex, 5)
        notes_txt.Value = db.List(selectedIndex, 6)
        
        img_status.Value = False
        
        If db.List(selectedIndex, 7) = "Empty" Then
        img.Picture = LoadPicture(vbNullString)
        image_path = ""
        Else
        image_path = db.List(selectedIndex, 7)
        img.Picture = LoadPicture(db.List(selectedIndex, 7))
        End If
        'img.Picture = IIf(db.List(selectedIndex, 7) = "Empty", LoadPicture(vbNullString), LoadPicture(db.List(selectedIndex, 7)))
        
        status.ForeColor = RGB(255, 165, 0) ' RGB values for orange
        status.Caption = "Status : View User ID [ " & db.List(selectedIndex, 0) & " ]"
    
    End If


End Sub
