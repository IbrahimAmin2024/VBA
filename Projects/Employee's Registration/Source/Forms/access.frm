VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} access 
   Caption         =   "Access"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   OleObjectBlob   =   "access.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub login_Click()
' user "Admin"
' pass "1234"
    If user = "Admin" And pass = "1234" Then
        Unload Me ' Close the current form
        cpanel.Show
        ' Perform actions if Yes is chosen
    Else
        user.BorderColor = vbRed
        pass.BorderColor = vbRed
        status.ForeColor = vbRed
        user.Value = ""
        pass.Value = ""
        status.Caption = "Status : Wrong Info..!"
        ' Perform actions if No is chosen
    End If
End Sub

Private Sub pass_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.status.ForeColor = &H80FF&
    status.Caption = "Status : Waiting..!"
    user.BorderColor = vbBlack
    pass.BorderColor = vbBlack
End Sub

Private Sub user_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    status.ForeColor = &H80FF&
    status.Caption = "Status : Waiting..!"
    user.BorderColor = vbBlack
    pass.BorderColor = vbBlack
End Sub
