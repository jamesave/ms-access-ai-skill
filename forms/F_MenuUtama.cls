Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd_BukaFormReferensi_Click()
   DoCmd.OpenForm "F_ReferensiHirarki"
End Sub

Private Sub cmdPilihMenu_Click()
    Dim sTarget As String

    If Me.lst_Menu.ListIndex < 0 Then
        MsgBox "Pilih Salah Satu", vbExclamation
        Exit Sub
    End If

    sTarget = Me.lst_Menu.Column(0)   ' TargetObject

    Select Case UCase$(Left(sTarget, 1))
        Case "F"
            DoCmd.OpenForm sTarget, acNormal
        Case "R"
            DoCmd.OpenReport sTarget, acViewPreview
        Case Else
            MsgBox "Unknown object type: " & sType, vbExclamation
    End Select

End Sub

Private Sub lbl_LinkLingkarRadius_Click()
        Application.FollowHyperlink PDFPath & LingkarBusurPDF
End Sub

Private Sub lbl_SkemaLining_Click()
        'Application.FollowHyperlink PDFPath & SkemaLiningPNG
        Dim sTarget As String
        sTarget = "F_ReferensiHirarki"
        DoCmd.OpenForm sTarget, acNormal

End Sub

Private Sub lst_Menu_AfterUpdate()
    ' Column(0) = TargetObject
    ' Column(1) = ItemDescription
    ' Column(2) = ItemDescription
    ' Column(3) = TargetObject
    ' Column(4) = ObjectType

    If Me.lst_Menu.ListIndex >= 0 Then
        Me.txt_ItemDesc = Me.lst_Menu.Column(1)
    Else
        Me.txt_ItemDesc = Null
    End If
End Sub
