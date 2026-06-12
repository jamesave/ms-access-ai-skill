Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub L_Label_Click()
    If Me.OrderBy = "L" Then
       Me.OrderBy = "L DESC"
    Else
       Me.OrderBy = "L"
    End If
End Sub

Private Sub PolaLubang_Label_Click()
   If Me.OrderBy = "PolaLubangID" Then
      Me.OrderBy = "PolaLubangID DESC"
   Else
      Me.OrderBy = "PolaLubangID"
   End If
End Sub

Private Sub PolaLubangID_Click()
   DoCmd.OpenForm "F_PolaLubangDetail", , , "[PolaLubangID] =" & Me.PolaLubangID
End Sub

Private Sub X_Label_Click()
   If Me.OrderBy = "X" Then
      Me.OrderBy = "X DESC"
   Else
      Me.OrderBy = "X"
   End If
End Sub
