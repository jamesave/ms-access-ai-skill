Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddBlankoKeping_Click()
    DoCmd.OpenForm "F_BlankoKeping", , , , acFormAdd, acDialog
End Sub

Private Sub KepingID_Click()
    On Error Resume Next
    If Not Me.NewRecord Then
    ' Reference the main form and update the text field
      'Me.Parent!selectedKepingID = Me!KepingID
      Forms!F_BaseProdukDetail.Form!selectedKepingID = Me!KepingID
      MsgBox Me!KepingID
    Else
      Me.Parent!selectedKepingID = Null
    End If

   On Error GoTo 0
End Sub

Private Sub Qty_Click()
    ' Make sure the main form is open and reference the control in the main form
    If Not IsNull(Forms!F_BaseProdukDetail!SF_PcsSetBOM.Form!KepingID) Then ' Replace FieldName with the name of your field
        Forms!F_BaseProdukDetail.Form!selectedKepingID = Forms!F_BaseProdukDetail!SF_PcsSetBOM.Form!KepingID
        '=[Forms]![MainFormName]![SubformControlName].Form![txtSelectedProductName]
'T_PcsSetBOM_sfrm2
        MsgBox Forms!F_BaseProdukDetail.Form!selectedKepingID
        '=[SF_PcsSetBOM].[Form]![KepingId]
    End If

End Sub
