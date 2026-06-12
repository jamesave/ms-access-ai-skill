Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboCariProduk_AfterUpdate()
        On Error GoTo Err_Handler

        Dim lngBlankoKepingAutoID    As Long

        lngBlankoKepingAutoID = Me.cboCariProduk
            Me.AllowEdits = True

        'Remove the filter because the recordset is searched as filtered
        If Me.FilterOn Then Me.FilterOn = False
        Me.DataEntry = False
        DoCmd.SearchForRecord acDataForm, Me.Name, acFirst, "[BlankoKepingAutoID] = " & lngBlankoKepingAutoID

        '============UI Option==============
        'Reset Product selection combo
        'Comment if the selected product should remain visible in the combo
        Me.cboCariProduk = ""
        '===================================

        'ControlStates

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "cboCariProduk_AferUpdate"
       Resume Exit_Handler
End Sub


Private Sub cmdCancel_Click()

      DoCmd.RunCommand (acCmdUndo)
      Me.cmdEditBlanko.SetFocus
      SetLockCtl (1)
End Sub

Private Sub cmdEditBlanko_Click()
      'On Error GoTo Err_Handler
      '
      '   If Me.Dirty = True Then
      '      DoCmd.RunCommand (acCmdUndo)
      '
      '   End If
      '      Me.cmdEditProfil.SetFocus
      '      SetLockCtl (1)
      '
      'Exit_Handler:
      '        Exit Sub
      '
      'Err_Handler:
      '        clsErrorHandler.HandleError Me.Name, "cmdCancel_Click"
      '        Resume Exit_Handler
         SetLockCtl (0)

End Sub


Private Sub cmdSimpanBlanko_Click()

   If Not Me.Dirty Then
      MsgBox ("Tidak ada perubahan Data. Data tidak disimpan")

   Else
      Me.Dirty = False
      MsgBox ("Data sudah disimpan")
   End If
   SetLockCtl (1)

End Sub



Private Sub cmdTambahBlanko_Click()
    On Error GoTo ErrorHandler

    ' Check if the form is already in a new record state
    If Me.NewRecord Then
        MsgBox "Already in a new record.", vbInformation
    Else
        ' Move to a new record
        DoCmd.GoToRecord , , acNewRec
    End If

    Exit Sub

ErrorHandler:

    clsErrorHandler.HandleError Me.Name, "cmdTambahBlanko_Click"
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   'Just in case the form is not locked and user change something.  Need to debug how the control is not working then.
   SetLockCtl (0)
   Me.lblWarning.Caption = "PERUBAHAN DATA"
   Me.lblWarning.ForeColor = RGB(237, 28, 36)
   Me.cmdCancel.Visible = True

End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim sqlstring As String
   'This form load from one single table.

        Me.RecordSource = "T_BlankoKeping"
    If (Nz(Me.OpenArgs) <> 0) Then
     '  Me.RecordSource = "T_BlankoKeping"
     '  sqlstring = "SELECT T_BlankoKeping.* FROM T_BlankoKeping WHERE (((T_BlankoKeping.BlankoKepingAutoID) = " & Me.OpenArgs & "));"
     '  MsgBox sqlstring
         Me.Filter = "BlankoKepingAutoID = " & Me.OpenArgs
         Me.FilterOn = True
    End If

    ' If allowedits is set to False, it will not allow the combobox to be selected.
   ' Me.AllowEdits = False
         SetLockCtl (1)
End Sub

Private Sub cmdKembaliKeDaftarBlanko_Click()
    Dim PilihBlankoKepingID As Long
    PilihBlankoKepingID = Me!BlankoKepingAutoID
    If Me.Dirty = True Then
        If MsgBox("Simpan Perubahan? ", vbYesNo) = vbYes Then
            Me.Dirty = False
        Else
            Me.Undo
        End If
    End If
    DoCmd.Close acForm, "F_BlankoKeping"
    DoCmd.OpenForm "F_DaftarBlankoKeping", OpenArgs:=PilihBlankoKepingAutoID
End Sub

Private Sub Tebal_TipisFlag_Click()
    If Tebal_TipisFlag.Value = False Then
        ' If checkbox is checked, block input to the field
        Tebal2.Enabled = False
    Else
        ' If checkbox is unchecked, allow input to the field
        Tebal2.Enabled = True
    End If
End Sub

Private Sub SetLockCtl(SetStatus As Integer)
        On Error GoTo Err_Handler
      Dim varPrefix       As Variant

         If SetStatus = 1 Then
            Me.txtBlankoKepingID.Locked = True
            Me.txtDeskripsiBlankoKeping.Locked = True
            Me.txtPanjangA.Locked = True
            Me.txtLebar.Locked = True
            Me.txtTebal1.Locked = True
            Me.txtTebal2.Locked = True
            Me.txtPanjangA.Locked = True
            Me.txtPanjangB.Locked = True
            Me.chkTebalTipisFlag.Locked = True
            Me.txtDiaBrakeShoe.Locked = True
            Me.txtDiaDrum.Locked = True
            Me.Detail.BackColor = 16777215 ' Background 1
            Me.lblWarning.Visible = False
            Me.lblWarning.Caption = "MODE DATA EDIT"
            Me.lblWarning.ForeColor = RGB(0, 0, 0)
            Me.lblWarning.BackColor = RGB(245, 157, 86)
            Me.cmdCancel.Visible = False ' This will give error when cancel cmd button is click since it's on focus

         Else  ' condition if the form is in edit mode

            Me.txtBlankoKepingID.Locked = False
            Me.txtDeskripsiBlankoKeping.Locked = False
            Me.txtPanjangA.Locked = False
            Me.txtLebar.Locked = False
            Me.txtTebal1.Locked = False
            Me.txtTebal2.Locked = False
            Me.txtPanjangA.Locked = False
            Me.txtPanjangB.Locked = False
            Me.chkTebalTipisFlag.Locked = False
            Me.txtDiaBrakeShoe.Locked = False
            Me.txtDiaDrum.Locked = False
            Me.Detail.BackColor = RGB(200, 200, 255)
            Me.lblWarning.Visible = True
            Me.cmdCancel.Visible = False
         End If

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "SetLockCtl"
        Resume Exit_Handler
End Sub
