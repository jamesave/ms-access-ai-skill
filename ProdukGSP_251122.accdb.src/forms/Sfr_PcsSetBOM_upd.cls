Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddBlankoKeping_Click()
    DoCmd.OpenForm "F_BlankoKeping", , , , acFormAdd, acDialog
End Sub

Private Sub Form_AfterInsert()
   UpdateCampurKeping
End Sub

Private Sub Form_AfterUpdate()
   UpdateCampurKeping
End Sub

Private Sub Form_Current()
   ' Dim recCount As Long
    'recCount = DCount("*", "T_PcsSetBOM")
  ' recCount = Me.RecordsetClone.RecordCount

   ' If recCount > 1 Then
  '      Me.txtCampurKepingIndicator.Value = recCount & " : Campur Keping"
   ' End If
   UpdateCampurKeping
End Sub




Private Sub IDLink_Click()
   Dim TempStr As String
   Dim TempLong As Long

    On Error Resume Next
    If Not Me.NewRecord Then
    ' Reference the main form and update the text field
      'Me.Parent!txtselectedKepingID = Me!KepingID
      Forms!F_BaseProdukDetail.Form!txtselectedKepingID = Me!KepingID


      TempLong = getKepingAutoID(Me!KepingID)
     ' TempStr = getPKImageName(TempLong)
      TempStr = GetPKImagePath(TempLong)

      'MsgBox Me!KepingID & " :here: " & TempLong & " :here: " & TempStr
      'Forms!F_BaseProdukDetail.Form!txtImageCode = TempStr
      Forms!F_BaseProdukDetail.Form!txtImageCode = Me.DeskripsiProfilKeping
      Forms!F_BaseProdukDetail.Form!PKViewer.Picture = TempStr
      Forms!F_BaseProdukDetail.Form!txtSetBomID = Me.SetBOMID

    Else
      Me.Parent!txtselectedKepingID = Null
    End If

   On Error GoTo 0
End Sub

Private Sub KepingID_Click()
'    On Error Resume Next
'    If Not Me.NewRecord Then
'    ' Reference the main form and update the text field
'      Forms!F_BaseProdukDetail.Form!txtselectedKepingID = Me!KepingID
'     ' MsgBox Me!KepingID
'    Else
'      Me.Parent!txtselectedKepingID = Null
'    End If
'
'   On Error GoTo 0

   Dim TempStr As String
   Dim TempLong As Long

    On Error Resume Next
    If Not Me.NewRecord Then
    ' Reference the main form and update the text field
      'Me.Parent!txtselectedKepingID = Me!KepingID
      Forms!F_BaseProdukDetail.Form!txtselectedKepingID = Me!KepingID


      TempLong = getKepingAutoID(Me!KepingID)
     ' TempStr = getPKImageName(TempLong)
      TempStr = GetPKImagePath(TempLong)

      'MsgBox Me!KepingID & " :here: " & TempLong & " :here: " & TempStr
      Forms!F_BaseProdukDetail.Form!txtImageCode = Me!DeskripsiProfilKeping
      Forms!F_BaseProdukDetail.Form!PKViewer.Picture = TempStr
    Else
      Me.Parent!txtselectedKepingID = Null
    End If

   On Error GoTo 0

End Sub

Private Sub Qty_Click()
    ' Make sure the main form is open and reference the control in the main form
    If Not IsNull(Forms!F_BaseProdukDetail!Sfr_PcsSetBOM.Form!KepingID) Then ' Replace FieldName with the name of your field
        Forms!F_BaseProdukDetail.Form!txtselectedKepingID = Forms!F_BaseProdukDetail!Sfr_PcsSetBOM.Form!KepingID
        '=[Forms]![MainFormName]![SubformControlName].Form![txtSelectedProductName]
       ' MsgBox Forms!F_BaseProdukDetail.Form!txtselectedKepingID
        '=[sfr_PcsSetBOM].[Form]![KepingId]
    End If

End Sub

Private Sub UpdateCampurKeping()
    Dim RecCount As Long
    'recCount = DCount("*", "T_PcsSetBOM")
   RecCount = Me.RecordsetClone.RecordCount

    If RecCount > 1 Then
        Me.txtCampurKepingIndicator.Value = RecCount & " : Campur Keping"
    Else
         Me.txtCampurKepingIndicator.Value = ""
   End If
End Sub

Private Sub QtyPcsPerSet()
   QtyPcsPerSet = DSum("Quantity", "YourQueryName", "ProductID = 25")
   'QtyTotal = Sum([Qty])
End Sub
