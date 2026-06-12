Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub BlankoKepingID_AfterUpdate()

   Me.DeskripsiBlanko.Value = Me.BlankoKepingID.Value
   Me.AllowEdits = False
End Sub

Private Sub BlankoKepingID_BeforeUpdate(Cancel As Integer)
   Me.AllowEdits = True
End Sub

Private Sub EditProfilBtn_Click()
     Me.AllowEdits = True
End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim sqlstring As String
   Dim QueryName As String
   Dim receivedArgs As Variant
   Dim Arraysize As Long


   QueryName = "Qry_ProfilKepingDetail_Primary"

   Me.AllowEdits = False
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

End Sub

Private Sub Tambah_Pola_Lubang_Click()

   ' DoCmd.OpenForm "F_PolaLubangDetail", , , , acFormAdd, acDialog

End Sub
