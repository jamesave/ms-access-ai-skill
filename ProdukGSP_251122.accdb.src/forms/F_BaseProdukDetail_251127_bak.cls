Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboFilterKategori_AfterUpdate()
   MsgBox "STOP_cboFilterKategori"
    If IsNull(Me.cboFilterKategori.Value) Then
        ' No filter
        Me.FilterOn = False
    Else
        ' Apply filter based on combo box value
        Me.Filter = "[KategoriProduk] = '" & Me.cboFilterKategori.Value & "'"
        Me.FilterOn = True
    End If
End Sub

Private Sub cmdCancel_Click()
      If Me.Dirty Then
         DoCmd.RunCommand (acCmdUndo)
      End If
      Me.cmdEditBaseProduk.SetFocus
      SetLockCtl (1)
End Sub


Private Sub cmdDelPKRecord_Click()
      Dim rs As DAO.Recordset
      Dim selectedID As Long

      selectedID = Me.txtSetBomID

      Set rs = Me.Sfr_PcsSetBOM.Form.RecordsetClone

      ' Filter recordset to the target record based on primary key
      MsgBox Me.txtSetBomID

      rs.Filter = "SetBOMID = " & selectedID
      Debug.Print "Filter applied: " & rs.Filter
      rs.MoveLast
      Debug.Print "Record count after filter: " & rs.RecordCount

      If rs.RecordCount > 0 Then
          rs.MoveFirst
          rs.Delete
          rs.Close
          Set rs = Nothing

          ' Requery the subform to refresh display after delete
          Me.Sfr_PcsSetBOM.Form.Requery
          Debug.Print selectedID
      Else
          MsgBox "Record not found."
      End If
End Sub

Private Sub cmdEditBaseProduk_Click()
      SetLockCtl (0)
End Sub

Private Sub cmdKembaliKeDaftarBaseProduk_Click()
    Dim PilihBaseProdukID As Long
    PilihBaseProdukID = Me!BaseProdukID
    MsgBox (PilihBaseProdukID)
    If Me.Dirty = True Then
        If MsgBox("Simpan Perubahan? ", vbYesNo) = vbYes Then
            Me.Dirty = False
        Else
            Me.Undo
        End If
    End If
    DoCmd.Close acForm, "F_BaseProdukDetail"
    DoCmd.OpenForm "F_DaftarBaseProduk", OpenArgs:=PilihBaseProdukID

End Sub

Private Sub cmdSimpanBaseProduk_Click()
   If Not Me.Dirty Then
      MsgBox ("Tidak ada perubahan Data. Data tidak disimpan")

   Else
      Me.Dirty = False
      MsgBox ("Data sudah disimpan")
   End If
   SetLockCtl (1)
End Sub

Private Sub cmdTambahBaseProduk_Click()
    On Error GoTo ErrorHandler

    ' Check if the form is already in a new record state
    If Me.NewRecord Then
        MsgBox "Already in a new record.", vbInformation
    Else
        ' Move to a new record
        DoCmd.GoToRecord , , acNewRec
    End If

    SetLockCtl (0)

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Private Sub cmdDebug_Click()
   MsgBox "See Immediate Window for Debug"
'   TmpData = RecCount(Me!BaseProdukID)
   Debug.Print Me.Sfr_PcsSetBOM.Form.RecordSource
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim i As Integer

    ' Open a recordset on the subform's recordsource
    Set rs = Me.Sfr_PcsSetBOM.Form.RecordsetClone

    If rs.EOF And rs.BOF Then
        Debug.Print "No records found."
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.MoveFirst
    Do While Not rs.EOF
        ' Print field names first time only (optional)
        If rs.AbsolutePosition = 0 Then
            For i = 0 To rs.Fields.Count - 1
                Debug.Print rs.Fields(i).Name & vbTab;
            Next i
            Debug.Print
        End If

        ' Print each field's value in the current record
        For i = 0 To rs.Fields.Count - 1
            Debug.Print Nz(rs.Fields(i).Value, "NULL") & vbTab;
        Next i

        Debug.Print ' New line for next record
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

End Sub
Public Sub UpdateCampurKeping()
    Dim RecCount As Long
    'recCount = DCount("*", "T_PcsSetBOM")
   RecCount = Me.Sfr_PcsSetBOM.Form.RecordsetClone.RecordCount

    If RecCount > 1 Then
      Me.txtCampurKeping.Value = RecCount & " : Campur Keping"
      Me.txtImageCode = "Pilih Salah Satu Keping"
      Me.txtselectedKepingID = ""
    ElseIf RecCount = 1 Then
      Dim val As Long
      val = Forms!F_BaseProdukDetail!Sfr_PcsSetBOM.Form!IDLink
      'TempLong = getKepingAutoID(Me!KepingID)
      'TempLong = getKepingAutoID(val)
      'MsgBox val & "-" & TempLong
      'TempStr = getPKImageName(TempLong)
      TempStr = GetPKImagePath(val)
       Me.txtCampurKeping.Value = ""
       'TempStr at the moment is not used. it is to store the image path
       Me.txtImageCode = Forms!F_BaseProdukDetail!Sfr_PcsSetBOM.Form!DeskripsiProfilKeping
       Me.txtselectedKepingID = Forms!F_BaseProdukDetail!Sfr_PcsSetBOM.Form!KepingID
       Me.txtSetBomID = Forms!F_BaseProdukDetail!Sfr_PcsSetBOM.Form!SetBOMID
    Else
      Me.txtCampurKeping.Value = ""
       Me.txtImageCode = ""
       Me.txtselectedKepingID = ""
       Me.txtSetBomID = ""
   End If
End Sub

Private Sub cmdUnfilter_Click()
   Me.FilterOn = False
   Me.cboFilterKategori = Null
End Sub

Private Sub cmnOpenPKDetail_Click()
    Dim filterValue As Variant
    filterValue = Me.txtselectedKepingID.Value

    If IsNull(filterValue) Then
        MsgBox "Please select a value first."
        Exit Sub
    End If
    MsgBox "KI? cmnOpenPKDetail_Click(): [T_ProfilKeping].[KepingID] = '" & filterValue & "'"
   DoCmd.OpenForm "F_ProfilKepingDetail", , , " [KepingID] = '" & filterValue & "'"
 '  DoCmd.OpenForm "F_ProfilKepingDetail", , , "  [T_ProfilKeping].[KepingAutoID] = " & filterValue
End Sub

Private Sub Form_Current()

   Me.Caption = "Produk Dasar: " & Nz(Me!BaseProdukName, "N/A")
   Me.Auto_Header0.Caption = "Produk Dasar: " & Nz(Me!BaseProdukName, "N/A")
  UpdateCampurKeping

End Sub



Private Sub Form_Dirty(Cancel As Integer)
   'Just in case the form is not locked and user change something.  Need to debug how the control is not working then.
   SetLockCtl (0)
   Me.lblWarning.Caption = "PERUBAHAN DATA"
   Me.lblWarning.ForeColor = RGB(237, 28, 36)
   Me.cmdCancel.Visible = True
End Sub

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Handler
      If Nz(Me.OpenArgs) = 0 Then
         Me.RecordSource = "Q_BaseProdukDetail"
      Else
         'MsgBox (nz(Me.OpenArgs)) 'The null zero
         Me.RecordSource = "SELECT T_BaseProduk.BaseProdukID, T_BaseProduk.BaseProdukName, T_BaseProduk.Deskripsi, T_BaseProduk.Catatan, T_ListStatusBaseProduk.StatusProduk, T_KategoriProduk.KategoriProduk, T_ListAxlePos.AxlePos FROM T_ListStatusBaseProduk RIGHT JOIN (T_ListAxlePos RIGHT JOIN (T_KategoriProduk RIGHT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID) ON T_ListAxlePos.AxlePosID = T_BaseProduk.AxlePosID) ON T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID WHERE (((T_BaseProduk.BaseProdukID)=" & Me.OpenArgs & "));"
      End If

   SetLockCtl (1)

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "Form_Open F_BaseProdukDetail"
        Resume Exit_Handler
End Sub

Private Function TotalQtyPcsPerSet(DataID As Long)

      TotalQtyPcsPerSet = Nz(DSum("Qty", "T_PcsSetBOM", "BaseProdukID = " & DataID))

End Function

Private Function RecCount(DataID As Long)

      RecCount = DCount("*", "T_PcsSetBOM", "BaseProdukID = " & DataID)

End Function

Private Sub cmdAddNewRecord_Click()
'    If Me.NewRecord Then
'        Me.Parent.Dirty = False ' Save the main record
'    End If

   ' Me.SetFocus ' Set focus to subform
    Me!Sfr_PcsSetBOM.SetFocus
    Debug.Print "HERE: " & Me!Sfr_PcsSetBOM.Form.RecordSource
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub SetLockCtl(SetStatus As Integer)
      On Error GoTo Err_Handler
      Dim varPrefix       As Variant

      If SetStatus = 1 Then  ' 1 means it's locked
         Me.txtBaseProdukName.Locked = True
         Me.cboStatProdFilt.Locked = True
         Me.cboKatProdFilt.Locked = True
         Me.txtDeskripsiBaseProduk.Locked = True
         Me.Sfr_PcsSetBOM.Locked = True
         Me.txtselectedKepingID.Locked = True
         Me.cboAxlePos.Locked = True
         Me.txtCatatan.Locked = True
         Me!Sfr_PcsSetBOM.Form.AllowAdditions = False
         Me!Sfr_PcsSetBOM.Form.AllowEdits = False
         'Me.SFR_ProfilKepingDetail.Locked = True

         Me.Detail.BackColor = 16777215 ' Background 1
         Me.lblWarning.Visible = False

         Me.lblWarning.ForeColor = RGB(0, 0, 0)
         Me.lblWarning.BackColor = RGB(245, 157, 86)
         Me.cmdCancel.Visible = False ' This will give error when cancel cmd button is click since it's on focus

      Else  ' condition if the form is in edit mode

         Me.txtBaseProdukName.Locked = False
         Me.cboStatProdFilt.Locked = False
         Me.cboKatProdFilt.Locked = False
         Me.txtDeskripsiBaseProduk.Locked = False
         Me.Sfr_PcsSetBOM.Locked = False
         Me.txtselectedKepingID.Locked = False
         Me.cboAxlePos.Locked = False
         Me.txtCatatan.Locked = False
         Me!Sfr_PcsSetBOM.Form.AllowAdditions = True
         Me!Sfr_PcsSetBOM.Form.AllowEdits = True
         'Me.SFR_ProfilKepingDetail.Locked = False

         Me.Detail.BackColor = RGB(200, 200, 255)
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
