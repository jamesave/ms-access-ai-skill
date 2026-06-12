Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
   'Me!NameFilt = Nz(NameStrDONE, "")
  ' Me!CityFilt = Nz(CityStrDONE, "")
   'BuildFiltStr
   'If Nz(Me.OpenArgs) <> 0 Then
   '   Me!CustID.SetFocus
   '   DoCmd.FindRecord Me.OpenArgs
   'End If
End Sub

Sub BuildFiltStr()
   'FiltStr = ""
   'If Me!NameFilt <> "" Then
    '  FiltStr = "[LastName] = '" & Me!NameFilt & "'"
  ' End If
   'If Me!CityFilt <> "" Then
    '  If FiltStr = "" Then
   '      FiltStr = "[City] = '" & Me!CityFilt & "'"
   '   Else
   '      FiltStr = FiltStr & " AND [City] = '" & Me!CityFilt & "'"
    '  End If
 '  End If
  ' If Me!StateFilt <> "" Then
  '    If FiltStr = "" Then
   '      FiltStr = "[ST] = '" & Me!StateFilt & "'"
   '   Else
   '      FiltStr = FiltStr & " AND [ST] = '" & Me!StateFilt & "'"
   '   End If
  ' End If
  ' If FiltStr = "" Then
   '   Me.Filter = ""
   '   Me.FilterOn = False
  ' Else
   '   Me.Filter = FiltStr
   '   Me.FilterOn = True
  ' End If
   'NameStrDONE = Nz(Me!NameFilt, "")
  ' CityStrDONE = Nz(Me!CityFilt, "")
End Sub
