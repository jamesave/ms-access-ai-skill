Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Tambah_Pola_Lubang_Click()

    DoCmd.OpenForm "F_PolaLubangDetail", , , , acFormAdd, acDialog

End Sub
