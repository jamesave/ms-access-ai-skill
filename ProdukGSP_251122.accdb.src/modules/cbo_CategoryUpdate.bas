Attribute VB_Name = "cbo_CategoryUpdate"
Option Compare Database

Private Sub cboCategory_AfterUpdate()
    Me.subTable.Form.Filter = "CategoryField = '" & Me.cboCategory.Value & "'"
    Me.subTable.Form.FilterOn = True
End Sub
