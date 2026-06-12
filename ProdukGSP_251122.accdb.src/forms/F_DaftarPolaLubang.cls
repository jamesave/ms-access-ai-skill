Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim FiltStr As String

'Catatan Form_F_DaftarPolaLubang
'Fungsi: Menunjukkan semua pola lubang yang ada
'Query: Qry_ListPolaLubang_byHoleQty
'Temp place holder for folder location:
' C:\Users\gunai\OneDrive\Galangsa_GD\DBProduk\Images\

'from T_List_PolaLubang

Private Sub ClearFilter_Click()
   Me.JmlLubangFilt = ""
   BuildFiltStr
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.OrderBy = "T_ListPolaLubang.JumlahLubang"
    Me.OrderByOn = True
    Me.AllowAdditions = False
    Me.AllowEdits = False
    Me.AllowDeletions = False
    'Below here is copy from form open F_DaftarBlankoKeping
    'If (Nz(Me.OpenArgs) <> 0) Then
     '   Me!T_ListPolaLubang.JumlahLubang.SetFocus
      '  DoCmd.FindRecord Me.OpenArgs
    'End If
    '__copy until here.. not sure if the setfocus is proper?

End Sub

Private Sub JmlLubangFilt_AfterUpdate()
   BuildFiltStr
End Sub

Private Sub JmlLubangFilt_GotFocus()
   Me.AllowEdits = True
   If Nz(FiltStr, "") = "" Then
      Me.JmlLubangFilt.RowSource = "SELECT DISTINCT JumlahLubang FROM T_ListPolaLubang ORDER BY JumlahLubang;"
   Else
      Me.JmlLubangFilt.RowSource = "SELECT DISTINCT JumlahLubang FROM T_ListPolaLubang ORDER BY JumlahLubang WHERE JumlahLubang = " & FiltStr & ";"
   End If
   Me.JmlLubangFilt.Dropdown
End Sub

Private Sub JmlLubangFilt_LostFocus()
   Me.AllowEdits = False
End Sub

Private Sub JumlahLubang_Label_Click()
    If Me.OrderBy = "JumlahLubang" Then
       Me.OrderBy = "JumlahLubang DESC"
    Else
       Me.OrderBy = "JumlahLubang"
    End If
End Sub

Private Sub LokasiImage_Label_Click()
    If Me.OrderBy = "LokasiImage" Then
       Me.OrderBy = "LokasiImage DESC"
    Else
       Me.OrderBy = "LokasiImage"
    End If
End Sub

Private Sub PolaLubang_Click()
   'Buka F_DaftarPolaLubang_Detail dengan Pola Lubang yang dipilih
   DoCmd.OpenForm "F_DaftarPolaLubang_Detail", , , "PolaLubang = '" & Me!PolaLubang & "'"

End Sub

Private Sub PolaLubang_Label_Click()
    If Me.OrderBy = "PolaLubang" Then
       Me.OrderBy = "PolaLubang DESC"
    Else
       Me.OrderBy = "PolaLubang"
    End If
End Sub

Private Sub Tambah_Click()
    Me.AllowAdditions = True
    DoCmd.RunCommand acCmdRecordsGoToNew ' untuk saat ini tidak bisa records go to new, karena record source distinct
    Me!PolaLubang = "BARU"
   ' Me.AllowAdditions = False   ' Perhaps can change this in later time.  removed ffor now 12/8/25
   'EditBlankoBtn_Click
End Sub

Sub BuildFiltStr()
   FiltStr = ""
   If Me!JmlLubangFilt <> "" Then
      FiltStr = "[JumlahLubang] = " & Me!JmlLubangFilt
   End If
   If FiltStr = "" Then
      Me.Filter = ""
      Me.FilterOn = False
   Else
      Me.Filter = FiltStr
      Me.FilterOn = True
      'MsgBox FiltStr
   End If

End Sub
