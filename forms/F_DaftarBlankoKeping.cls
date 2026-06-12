Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim SrchVal As String
Dim SrchCrt As String

Private Sub BlankoKepingID_Click()
   On Error GoTo Err_Handler

          '=====================Equivalent Syntax=========================
          ' DoCmd.OpenForm "frmProductDetail", , , "ProductID = " & Me.ProductID
          ' DoCmd.OpenForm FormName:="frmProductDetail", WhereCondition:="ProductID = " & Me.ProductID
          '===============================================================

   ' MsgBox (Me!BlankoKepingAutoID)
    DoCmd.OpenForm "F_BlankoKeping", OpenArgs:=Me!BlankoKepingAutoID
    'DoCmd.Close acForm, "F_DaftarBlankoKeping"

Exit_Handler:
    Exit Sub

Err_Handler:
   'The line below is causing compile error.. need to debug
   'clsErrorHandler.HandleError Me.Name, "BlankoKepingID_Click", True
    Resume Exit_Handler
End Sub

Private Sub lblBlankoKepingID_Click()
    If Me.OrderBy = "BlankoKepingID" Then
       Me.OrderBy = "BlankoKepingID DESC"
    Else
       Me.OrderBy = "BlankoKepingID"
    End If
End Sub


Private Sub cmdEditBlankoBtn_Click()
    Dim PilihBlankoKepingID As String
    MsgBox (Me!BlankoKepingAutoID)

    PilihBlankoKepingID = Me!BlankoKepingAutoID
    DoCmd.Close acForm, "F_DaftarBlankoKeping"
    DoCmd.OpenForm "F_BlankoKeping", OpenArgs:=PilihBlankoKepingID



'    PilihBlankoKepingID = Me!BlankoKepingID
'    DoCmd.Close acForm, "F_DaftarBlankoKeping"
'    DoCmd.OpenForm "F_BlankoKeping", OpenArgs:=PilihBlankoKepingID
End Sub

Private Sub cmdExportExcelBtn_Click()

  ' On Error GoTo Err
   On Error GoTo Err_Handler
      DoCmd.OutputTo acOutputForm, "F_Daftarblankokeping", acFormatXLSX, "C:\Accurate\DaftarBlankoKeping.xlsx", True

ExitHere:
    Exit Sub

Exit_Handler:
         Exit Sub

Err_Handler:
      MsgBox "Error " & Err.Number & " : " & Err.Description, vbExclamation, "Error Outputting Excel"
    ' Optional: handle specific errors
    'If Err.Number = 2302 Then
    '    MsgBox "Cek path atau file Excel yang sedang terbuka.", vbInformation
    'End If
       Resume ExitHere
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.OrderBy = "BlankoKepingID"
    Me.OrderByOn = True
    Me.AllowAdditions = False
    Me.AllowEdits = False
    Me.AllowDeletions = False
    If (Nz(Me.OpenArgs) <> 0) Then
        Me!BlankoKepingID.SetFocus
        DoCmd.FindRecord Me.OpenArgs
    End If
End Sub





Private Sub cmdTambahKepingBlankoBtn_Click()
    Me.AllowAdditions = True
    DoCmd.RunCommand acCmdRecordsGoToNew
    Me!BlankoKepingID = ""
    Me.AllowAdditions = False
   ' EditBlankoBtn_Click
End Sub


Private Sub lblDeskripsiBlankoKeping_Click()
    If Me.OrderBy = "DeskripsiBlankoKeping" Then
       Me.OrderBy = "DeskripsiBlankoKeping DESC"
    Else
       Me.OrderBy = "DeskripsiBlankoKeping"
    End If
End Sub



Private Sub lblPanjang_Click()
   If Me.OrderBy = "Panjang" Then
      Me.OrderBy = "Panjang DESC"
   Else
      Me.OrderBy = "Panjang"
   End If
End Sub

Private Sub lblLebar_Click()
    If Me.OrderBy = "Lebar" Then
       Me.OrderBy = "Lebar DESC"
    Else
       Me.OrderBy = "Lebar"
    End If
End Sub

Private Sub lblTebal_Click()
    If Me.OrderBy = "Tebal" Then
       Me.OrderBy = "Tebal DESC"
    Else
       Me.OrderBy = "Tebal"
    End If
End Sub

Private Sub lblTebal2_Click()
    If Me.OrderBy = "Tebal2" Then
       Me.OrderBy = "Tebal2 DESC"
    Else
        Me.OrderBy = "Tebal2"
    End If
End Sub



Private Sub lblDiameterDrum_Click()
    If Me.OrderBy = "DiameterDrum" Then
       Me.OrderBy = "DiameterDrum DESC"
    Else
       Me.OrderBy = "DiameterDrum"
    End If
End Sub

Private Sub lblDiaBrakeShoe_Click()
    If Me.OrderBy = "DiameterBrakeShoe" Then
       Me.OrderBy = "DiameterBrakeShoe DESC"
    Else
       Me.OrderBy = "DiameterBrakeShoe"
    End If
End Sub


Private Sub Tebal_TipisFlag_Label_Click()
    If Me.OrderBy = "Tebal_TipisFlag" Then
       Me.OrderBy = "Tebal_TipisFlag DESC"
    Else
       Me.OrderBy = "Tebal_TipisFlag"
    End If
End Sub





Private Sub CatatanBlankoKeping_Label_Click()
    If Me.OrderBy = "CatatanBlankoKeping" Then
       Me.OrderBy = "CatatanBlankoKeping DESC"
    Else
       Me.OrderBy = "CatatanBlankoKeping"
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandler
Dim ctl As Control
Dim fldName As String
Dim rst As Recordset
'MsgBox ("the keycode is " & KeyCode)
Select Case KeyCode
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
   Case 48 To 57, 65 To 90 'all numbers and letters
      Set ctl = Screen.ActiveControl 'find out what is the activecontrol
      fldName = ctl.Name
      SrchVal = SrchVal & Chr(KeyCode) 'allow to compound each stroke to be added
      KeyCode = 0
      SrchCrt = "[" & fldName & "] Like '*" & SrchVal & "*'"
      Set rst = Me.RecordsetClone  'rst is a recordset dim. copy the recordset to this variable
      rst.FindFirst SrchCrt
      If rst.NoMatch Then
        MsgBox ("Record not found: " & SrchCrt)
      Else
        Me.Bookmark = rst.Bookmark
      End If
      rst.Close
   Case 27 'Escape Key
      SrchVal = ""
      KeyCode = 0

   Case 122 'F11 Function key
      'If alternate key is pressed
      'Shift =1, Ctrl =2, Alt =4
      If Shift <> 4 Then
        KeyCode = 0
      End If
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
