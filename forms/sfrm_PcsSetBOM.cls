Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Form_Open(Cancel As Integer)
   Me.AllowEdits = False
   Me.AllowAdditions = False
End Sub

Private Sub LihatProfilBtn_Click()
   Dim ProfilKeping As String
   Dim recordresult As Integer
   Dim rs As DAO.Recordset
   Dim allrecords As String

   Set rs = Me.Controls("T_PcsSetBOM subform2").Form.RecordsetClone

   ' Move to the last record to get an accurate record count
   If Not rs.EOF Then rs.MoveLast

   If rs.RecordCount > 1 Then
      rs.MoveFirst
      Do Until rs.EOF
         ProfilKeping = ProfilKeping & rs!KepingID & ";"

         rs.MoveNext
      Loop
      'remove the last ","

      ProfilKeping = Left(ProfilKeping, Len(ProfilKeping) - 1)
         MsgBox ProfilKeping
      DoCmd.OpenForm "F_ProfilKepingDetail", OpenArgs:=ProfilKeping
   ElseIf rs.RecordCount = 1 Then
      ProfilKeping = Me!KepingID
      DoCmd.OpenForm "F_ProfilKepingDetail", OpenArgs:=ProfilKeping
   Else
       MsgBox "No entries found in the subform."
   End If

  ' recordresult = me


End Sub
