Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdCancel_Click()
   On Error GoTo Err_Handler

   If Me.Dirty = True Then
      DoCmd.RunCommand (acCmdUndo)

   End If
      Me.cmdEditProfilKpngBtn.SetFocus
      SetLockCtl (1)

Exit_Handler:
      Exit Sub

Err_Handler:
      clsErrorHandler.HandleError Me.Name, "cmdCancel_Click"
      Resume Exit_Handler
End Sub

Private Sub cmdDaftarBlankoKeping_Click()

   DoCmd.OpenForm "F_DaftarBlankoKeping", acNormal
End Sub

Private Sub cmdEditProfilKpngBtn_Click()
      SetLockCtl (0)
End Sub

Private Sub cmdKembaliKeDaftarBaseProduk_Click()
   DoCmd.OpenForm "F_DaftarProfilKeping", acNormal
End Sub

Private Sub cmdSaveProfil_Click()
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
        SetLockCtl (1)
    End If
End Sub



Private Sub Form_Open(Cancel As Integer)
   'Set the code to lock the fields from edit
   SetLockCtl (1)
End Sub

Private Sub SetLockCtl(SetStatus As Integer)
   On Error GoTo Err_Handler
   Dim varPrefix       As Variant

   If SetStatus = 1 Then
      Me.txtKepingID.Locked = True
      Me.txtDeskripsiProfilKeping.Locked = True
      Me.txtKepingAutoID.Locked = True
      Me.cboBlankoKepingID.Locked = True
      Me.cboDeskripsiBlanko.Locked = True

      Me.sfrm_BlankoKeping.Locked = True
      Me.chkKepingRvt.Locked = True
      Me.cboPolaLubangID.Locked = True
      Me.txtPolaLubangID.Locked = True
      Me.cboUkuranLubangID.Locked = True
      Me.cboRivetID.Locked = True
      Me.chkLis.Locked = True
      Me.chkCoak.Locked = True
      Me.chkKoak.Locked = True

      Me.cmdHapusLbg.Visible = False

      Me.sfrm_PolaLubangImage.Locked = False

            Me.Detail.BackColor = RGB(245, 157, 86) ' Background 1
            Me.lblWarning.Visible = False
            Me.lblWarning.ForeColor = RGB(0, 0, 0)
            Me.lblWarning.BackColor = RGB(245, 157, 86)
            Me.cmdCancel.Visible = False ' This will give error when cancel cmd button is click since it's on focus

   Else  ' condition if the form is in edit mode

      Me.txtKepingID.Locked = False
      Me.txtDeskripsiProfilKeping.Locked = False
      Me.txtKepingAutoID.Locked = False
      Me.cboBlankoKepingID.Locked = False
      Me.cboDeskripsiBlanko.Locked = False

      Me.sfrm_BlankoKeping.Locked = False
      Me.chkKepingRvt.Locked = False
      Me.cboPolaLubangID.Locked = False
      Me.txtPolaLubangID.Locked = False
      Me.cboUkuranLubangID.Locked = False
      Me.cboRivetID.Locked = False
      Me.chkLis.Locked = False
      Me.chkCoak.Locked = False
      Me.chkKoak.Locked = False

      Me.cmdHapusLbg.Visible = True

      Me.sfrm_PolaLubangImage.Locked = False

      Me.Detail.BackColor = RGB(200, 200, 255)
      Me.lblWarning.ForeColor = RGB(255, 0, 0)
      Me.lblWarning.BackColor = RGB(245, 157, 86)
      Me.lblWarning.Caption = "MODE DATA EDIT"
      Me.lblWarning.Visible = True
      Me.cmdCancel.Visible = True

   End If

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "SetLockCtl"
        Resume Exit_Handler
End Sub
