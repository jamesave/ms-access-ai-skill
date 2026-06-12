Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cboBlankoKepingID_GotFocus()
   Me!cboBlankoKepingID.Requery
End Sub

Private Sub cboPolaLubangID_GotFocus()
   Me!cboPolaLubangID.Requery
End Sub

Private Sub cboRivetID_GotFocus()
   Me!cboRivetID.Requery
End Sub

Private Sub cboUkuranLubangID_GotFocus()
   Me!cboUkuranLubangID.Requery
End Sub

Private Sub cmd_HapusPolaLubang_Click()

 If MsgBox("Data lubang rivet akan terhapus. Setuju?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
      Me.PolaLubangID = Null
      Me.RivetID = Null
      Me.UkuranLubangID = Null
      Me.LubangFlag = Null

 End If
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler

   If Me.Dirty = True Then
      DoCmd.RunCommand (acCmdUndo)

   End If
      Me.cmdEditProfil.SetFocus
      SetLockCtl (1)

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "cmdCancel_Click"
        Resume Exit_Handler
End Sub

Private Sub cmdEditProfil_Click()
      SetLockCtl (0)
End Sub

Private Sub cmdKelolaFoto_Click()
   DoCmd.OpenForm "F_ImageProfilKeping", , , "[KepingAutoID]=" & Me.[KepingAutoID]
End Sub

Private Sub cmdRefreshcbobox_Click()
   Me.cboBlankoKepingID.Requery
   Me.cboPolaLubangID.Requery
   Me.cboRivetID.Requery

End Sub

Private Sub cmdSaveProfil_Click()
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
        SetLockCtl (1)
    End If
End Sub

Private Sub cmnOpenBlankoKeping_Click()
   'MsgBox BlankoKepingAutoID
   'DoCmd.OpenForm "F_BlankoKeping", , , "[BlankoKepingAutoID]=" & Me.[BlankoKepingAutoID]
   DoCmd.OpenForm "F_BlankoKeping", OpenArgs:=Me!BlankoKepingAutoID
End Sub

Private Sub cmdNewProfilKeping_Click()
   On Error GoTo Err_Handler

   ' Check if the form is already in a new record state
   If Me.NewRecord Then
       MsgBox "Already in a new record.", vbInformation
   Else
        ' Move to a new record
       DoCmd.GoToRecord , , acNewRec
   End If

   SetLockCtl (0)


   Exit Sub

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "cmdNewProfilKeping_Click"
        Resume Exit_Handler
End Sub




Private Sub Command76_Click()
   Debug.Print "FilterOn: "; Me.FilterOn
   Debug.Print "Filter: "; Me.Filter
   Debug.Print "RecordsetClone count: "; Me.RecordsetClone.RecordCount
End Sub

Private Sub Form_AfterUpdate()
   SetLockCtl (1)
End Sub

Private Sub Form_Dirty(Cancel As Integer)
   'Just in case the form is not locked and user change something.  Need to debug how the control is not working then.
   SetLockCtl (0)
   Me.lblWarning.Caption = "PERUBAHAN DATA"
   Me.lblWarning.ForeColor = RGB(237, 28, 36)
   Me.cmdCancel.Visible = True
End Sub

Private Sub Form_Load()
   'The Following block is to show form header if it's standalone, and hide if it's in a subform
   'This is not working
'   On Error GoTo Err_Handler
'    If Me.Parent Is Nothing Then
'        ' The form is opened as a standalone form
'        Me.Section("FormHeader").Visible = True
'    Else
'        ' The form is opened as a subform
'        Me.Section("FormHeader").Visible = False
'    End If
'
'Exit_Handler:
'        Exit Sub
'
'Err_Handler:
'        clsErrorHandler.HandleError Me.Name, "Form_Load"
'        Resume Exit_Handler
End Sub

Private Sub Form_Open(Cancel As Integer)


   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim sqlstring As String
   Dim QueryName As String
   Dim receivedArgs As Variant
   Dim Arraysize As Long


   QueryName = "Qry_ProfilKepingDetail_Primary"

   If Nz(Me.OpenArgs) = 0 Then
      Me.RecordSource = QueryName
   Else
      'Split the openArgs string by the delimiter ";"
      receivedArgs = Split(Me.OpenArgs, ";")
      MsgBox Nz(Me.OpenArgs)
      Arraysize = UBound(receivedArgs) - LBound(receivedArgs) + 1

      'Now do the SQL

      Set db = CurrentDb
      Set qdf = db.QueryDefs(QueryName)
      sqlstring = qdf.SQL
      sqlstring = Replace(sqlstring, ";", "") 'remove the last ; so the where statement can be inserted

   '   Me.RecordSource = "SELECT T_BaseProduk.BaseProdukID, T_BaseProduk.BaseProdukName, T_BaseProduk.Deskripsi, T_BaseProduk.Catatan, T_ListStatusBaseProduk.StatusProduk, T_KategoriProduk.KategoriProduk, T_ListAxlePos.AxlePos FROM T_ListStatusBaseProduk RIGHT JOIN (T_ListAxlePos RIGHT JOIN (T_KategoriProduk RIGHT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID) ON T_ListAxlePos.AxlePosID = T_BaseProduk.AxlePosID) ON T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID WHERE (((T_BaseProduk.BaseProdukID)=" & Me.OpenArgs & "));"

      For i = LBound(receivedArgs) To UBound(receivedArgs)
            If i = 0 Then
               sqlstring = sqlstring & " WHERE (((T_ProfilKeping.KepingID)=""" & receivedArgs(i) & """))" & vbCrLf

            Else
               sqlstring = sqlstring & " OR (((T_ProfilKeping.KepingID)=""" & receivedArgs(i) & """))"
            End If
      Next i
      sqlstring = sqlstring & ";"

      Me.RecordSource = sqlstring
      Set qdf = Nothing

   End If
   If Me.Filter <> "" Then
        Me.FilterOn = True
   End If
   'Set the code to lock the fields from edit
   SetLockCtl (1)
End Sub

Private Sub Tambah_Pola_Lubang_Click()

   ' DoCmd.OpenForm "F_PolaLubangDetail", , , , acFormAdd, acDialog

End Sub

Private Sub SetLockCtl(SetStatus As Integer)
        On Error GoTo Err_Handler
      Dim varPrefix       As Variant

         If SetStatus = 1 Then
            Me.txtKepingID.Locked = True
            Me.cboBlankoKepingID.Locked = True
            Me.cboDeskripsiBlanko.Locked = True
            Me.txtDeskripsiKeping.Locked = True
            'Me.imgGambarProfil.Locked = True
            Me.chkIndikator.Locked = True
            Me.chkCoak.Locked = True
            Me.chkKoak.Locked = True

            Me.Detail.BackColor = RGB(245, 157, 86) ' Background 1
            Me.lblWarning.Visible = False
            Me.lblWarning.ForeColor = RGB(0, 0, 0)
            Me.lblWarning.BackColor = RGB(245, 157, 86)
            Me.cmdCancel.Visible = False ' This will give error when cancel cmd button is click since it's on focus
            Me.cmd_HapusPolaLubang.Visible = False
'
         Else  ' condition if the form is in edit mode

            Me.txtKepingID.Locked = False
            Me.cboBlankoKepingID.Locked = False
            Me.cboDeskripsiBlanko.Locked = False
            Me.txtDeskripsiKeping.Locked = False
           ' Me.imgGambarProfil.Locked = False
            Me.chkIndikator.Locked = False
            Me.chkCoak.Locked = False
            Me.chkKoak.Locked = False

            Me.Detail.BackColor = RGB(200, 200, 255)
            Me.lblWarning.ForeColor = RGB(255, 0, 0)
            Me.lblWarning.BackColor = RGB(245, 157, 86)
            Me.lblWarning.Caption = "MODE DATA EDIT"
            Me.lblWarning.Visible = True
            Me.cmdCancel.Visible = True
            Me.cmd_HapusPolaLubang.Visible = True

         End If

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "SetLockCtl"
        Resume Exit_Handler
End Sub
