Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
   If Nz(Me.OpenArgs) = 0 Then
      Me.RecordSource = "Q_BaseProdukDetail"
   Else
   'MsgBox (nz(Me.OpenArgs)) 'The null zero
   Me.RecordSource = "SELECT T_BaseProduk.BaseProdukID, T_BaseProduk.BaseProdukName, T_BaseProduk.Deskripsi, T_BaseProduk.Catatan, T_ListStatusBaseProduk.StatusProduk, T_KategoriProduk.KategoriProduk, T_ListAxlePos.AxlePos FROM T_ListStatusBaseProduk RIGHT JOIN (T_ListAxlePos RIGHT JOIN (T_KategoriProduk RIGHT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID) ON T_ListAxlePos.AxlePosID = T_BaseProduk.AxlePosID) ON T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID WHERE (((T_BaseProduk.BaseProdukID)=" & Me.OpenArgs & "));"

   End If
End Sub
