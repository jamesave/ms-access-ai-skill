Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim FiltStr As String

Private Sub AxlePos_Label_Click()
If OrderBy = "AxlePos" Then
   OrderBy = "AxlePos desc"
Else
   OrderBy = "AxlePos"
End If
End Sub

Private Sub Catatan_Label_Click()
If OrderBy = "Catatan" Then
   OrderBy = "Catatan desc"
Else
   OrderBy = "Catatan"
End If
End Sub

Private Sub Deskripsi_Label_Click()
If OrderBy = "Deskripsi" Then
   OrderBy = "Deskripsi desc"
Else
   OrderBy = "Deskripsi"
End If
End Sub

' Sub berikut ini ditutup, digantikan oleh klik di SET ID (hyperlink)
' Sebelumnya cara kerjanya, harus click di row produk yang mau diubah baru klik edit
'Private Sub EditProdukDasarBtn_Click()
'   Dim BaseProdukID As Integer
'   BaseProdukID = Me!BaseProdukID
'   DoCmd.Close acForm, "F_DaftarBaseProduk"
'   DoCmd.OpenForm "F_BaseProdukDetail", OpenArgs:=BaseProdukID
'
'End Sub

Private Sub ExportExcelBtn_Click()
   On Error GoTo Err
      DoCmd.OutputTo acOutputForm, "F_DaftarBaseProduk", acFormatXLSX, "C:\Accurate\DaftarBaseProduk.xlsx", True

DoExit:
            Exit Sub

Err:
         MsgBox "Error Outputting Excel"
         Resume DoExit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ctl As Control
Dim fldName As String
Dim rst As Recordset
On Error GoTo ErrHandler
'MsgBox ("the keycode is " & KeyCode)
Select Case KeyCode
   Case 122
      If Shift <> 4 Then
         KeyCode = 0
      End If
   Case vbKeyEnd
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToLast
   Case vbKeyHome
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToFirst
   Case vbKeyUp
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToPrevious
   Case vbKeyDown
      KeyCode = 0
      DoCmd.RunCommand acCmdRecordsGoToNext
   Case vbKeyRight, vbKeyLeft
   Case 8  'backspace
   Case 9, 13  'tab or enter
   Case 48 To 57, 65 To 90  'any number or letter
      Set ctl = Screen.ActiveControl
      fldName = ctl.Name
      Select Case fldName
         Case "KatProdukFilt", "StatusBaseProdukFilt"
            Exit Sub
      End Select
      If fldName <> LastFld Then
         SrchVal = ""
      End If
      LastFld = fldName
      SrchVal = SrchVal & Chr(KeyCode)
      KeyCode = 0
      If fldName = "Address" Then
         SrchCrit = "[" & fldName & "] Like '*" & SrchVal & "*'"
      Else
         SrchCrit = "[" & fldName & "] Like '" & SrchVal & "*'"
      End If
      Set rst = Me.RecordsetClone
      rst.FindFirst SrchCrit
      If rst.NoMatch Then
         MsgBox (" Record not found! ")
      Else
         Me.Bookmark = rst.Bookmark
      End If
      rst.Close
   Case 107, 187
      If SrchVal = "" Then
         KeyCode = 0
         Exit Sub
      End If
      Set ctl = Screen.ActiveControl
      fldName = ctl.Name
      KeyCode = 0
      Set rst = Me.RecordsetClone
      rst.Bookmark = Me.Bookmark
      rst.FindNext SrchCrit
      If rst.NoMatch Then
         MsgBox (" Record not found! ")
      Else
         Me.Bookmark = rst.Bookmark
      End If
      rst.Close
   Case 109, 189
      If SrchVal = "" Then
         KeyCode = 0
         Exit Sub
      End If
      Set ctl = Screen.ActiveControl
      fldName = ctl.Name
      KeyCode = 0
      Set rst = Me.RecordsetClone
      rst.Bookmark = Me.Bookmark
      rst.FindPrevious SrchCrit
      If rst.NoMatch Then
         MsgBox (" Record not found! ")
      Else
         Me.Bookmark = rst.Bookmark
      End If
      rst.Close
   Case 27
      KeyCode = 0
      SrchVal = ""
   Case Else
      KeyCode = 0
End Select
Exit Sub
ErrHandler:
Select Case Err.Number
   Case 2046
   Case Else
      MsgBox Err.Number & " " & Err.Description
End Select
Resume Next
End Sub

Private Sub Form_Open(Cancel As Integer)
   Me.AllowAdditions = False
   Me.AllowEdits = False
   Me.AllowDeletions = False

   If Nz(Me.OpenArgs) <> 0 Then

   'The below lines are currently being ignored for now
   '  Me!BaseProdukID.SetFocus
   ' DoCmd.FindRecord Me.OpenArgs

   Else
      Me.OrderBy = "BaseProdukID"
      Me.OrderByOn = True
   End If
End Sub

Private Sub HapusKatFiltBtn_Click()
   Me.KatProdukFilt = ""
   BuildFiltStr
End Sub

Private Sub HapusSemuaFiltBtn_Click()
   Me.KatProdukFilt = ""
   Me.StatusBaseProdukFilt = ""
   BuildFiltStr
End Sub

Private Sub HapusStatusProdFiltBtn_Click()
   Me.StatusBaseProdukFilt = ""
   BuildFiltStr
End Sub

Private Sub KatProdukFilt_AfterUpdate()
   'Me.KatProdukFilt = ""
   BuildFiltStr
End Sub

Private Sub KatProdukFilt_GotFocus()
   Me.AllowEdits = True
   If FiltStr = "" Then
      Me.KatProdukFilt.RowSource = "SELECT DISTINCT T_BaseProduk.KategoriProdukID, T_KategoriProduk.KategoriProduk FROM T_KategoriProduk RIGHT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID;"
   Else
      Me.KatProdukFilt.RowSource = "SELECT DISTINCT T_BaseProduk.KategoriProdukID, T_KategoriProduk.KategoriProduk FROM T_KategoriProduk RIGHT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID WHERE (((T_KategoriProduk.KategoriProduk)=""" & FiltStr & """));"
   End If
      Me.KatProdukFilt.BoundColumn = 2
   Me.KatProdukFilt.Dropdown
End Sub

Private Sub KatProdukFilt_LostFocus()
   Me.AllowEdits = False
End Sub

Private Sub Kode_Label_Click()
If OrderBy = "BaseProdukID" Then
   OrderBy = "BaseProdukID desc"
Else
   OrderBy = "BaseProdukID"
End If
End Sub

Private Sub Produk_Label_Click()
If OrderBy = "BaseProdukName" Then
   OrderBy = "BaseProdukName desc"
Else
   OrderBy = "BaseProdukName"
End If
End Sub

Sub BuildFiltStr()
   FiltStr = ""
   If Me!KatProdukFilt <> "" Then
      FiltStr = "[KategoriProduk] = '" & Me!KatProdukFilt & "'"
   End If
   If Me!StatusBaseProdukFilt <> "" Then
      If FiltStr = "" Then
         FiltStr = "[StatusProduk] = '" & Me!StatusBaseProdukFilt & "'"
      Else
         FiltStr = FiltStr & " AND [StatusProduk] = '" & Me!StatusBaseProdukFilt & "'"
      End If
   End If
   If FiltStr = "" Then
      Me.Filter = ""
      Me.FilterOn = False
   Else
      Me.Filter = FiltStr
      Me.FilterOn = True
   End If
  ' MsgBox ("the Filtstr is " & FiltStr)
  ' KatProdukStrDONE = Nz(Me!KatProdukFilt, "")
  ' StatusBaseProdukStrDONE = Nz(Me!StatusBaseProdukFilt, "")
End Sub

Private Sub StatusBaseProduk_LostFocus()
   Me.AllowEdits = False
End Sub

Private Sub SetID_Click()
   Dim BaseProdukID As Long
   BaseProdukID = Me!BaseProdukID
   DoCmd.Close acForm, "F_DaftarBaseProduk"
   DoCmd.OpenForm "F_BaseProdukDetail", OpenArgs:=BaseProdukID
End Sub

Private Sub StatusBaseProdukFilt_GotFocus()
   Me.AllowEdits = True
   If FiltStr = "" Then
      Me.StatusBaseProdukFilt.RowSource = "SELECT DISTINCT T_BaseProduk.StatusProdukID, T_ListStatusBaseProduk.StatusProduk FROM T_ListStatusBaseProduk RIGHT JOIN T_BaseProduk ON T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID WHERE (((T_ListStatusBaseProduk.StatusProduk)=""Produksi""));"
   Else
      Me.StatusBaseProdukFilt.RowSource = "SELECT DISTINCT T_BaseProduk.StatusProdukID, T_ListStatusBaseProduk.StatusProduk FROM T_ListStatusBaseProduk RIGHT JOIN T_BaseProduk ON T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID WHERE (((T_ListStatusBaseProduk.StatusProduk)=""" & FiltStr & """));"
   End If
      Me.StatusBaseProdukFilt.BoundColumn = 2
   Me.StatusBaseProdukFilt.Dropdown
End Sub
