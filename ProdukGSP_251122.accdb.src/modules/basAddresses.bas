Attribute VB_Name = "basAddresses"
Option Compare Database

Option Explicit

Function GetToken(stLn, stDelim)

    Dim iDelim As Integer, stToken As String
    iDelim = InStr(1, stLn, stDelim)
    If (iDelim <> 0) Then
        stToken = LTrim$(RTrim$(Mid$(stLn, 1, iDelim - 1)))
        stLn = Mid$(stLn, iDelim + 1)
    Else
        stToken = LTrim$(RTrim$(Mid$(stLn, 1)))
        stLn = ""
    End If
    GetToken = stToken

End Function
Function InParam(fld, Param)

    Dim stToken As String
    If IsNull(fld) Then fld = ""
    Do While (Len(Param) > 0)
        stToken = GetToken(Param, ",")
        If stToken = LTrim$(RTrim$(fld)) Then
            InParam = -1
            Exit Function
        Else
            InParam = 0
        End If
    Loop

End Function

Function GetFileTypes() As String

    ' returns list of file extensions for use in
    ' common dialogue opened from Images form

    Dim dbs As DAO.Database, rst As DAO.Recordset

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("qryFileTypes")

    ' loop through recordset and build string of file extensions
    ' delimited by semi colon and space
    With rst
        If Not .BOF And Not .EOF Then
            .MoveFirst
            Do Until .EOF
                GetFileTypes = GetFileTypes & "; " & !FileExtension
                .MoveNext
            Loop
            ' remove leading semi colon and space
            GetFileTypes = Mid$(GetFileTypes, 3)
        End If
        .Close
    End With

    Set rst = Nothing
    Set dbs = Nothing

End Function

Public Sub PropertySet(strValue As String, strProp As String, strForm As String)

    ' sets custom property of form

    Dim dbs As DAO.Database, ctr As DAO.Container, doc As DAO.Document
    Dim blnReturn As Boolean

    ' Return reference to current database.
    Set dbs = CurrentDb
    ' Return reference to Forms container.
    Set ctr = dbs.Containers!Forms
    ' Return reference to document.
    Set doc = ctr.Documents(strForm)
    blnReturn = SetAccessProperty(doc, _
        strProp, dbText, strValue)
    ' Evaluate return value.
    If Not blnReturn Then
        MsgBox "Property not set", vbExclamation, "Error"
    End If

End Sub
Function SetAccessProperty(obj As Object, strName As String, _
        intType As Integer, varSetting As Variant) As Boolean

    ' general purpose function from Access 97 Help
    ' for setting custom properties of objects

    Dim prp As Property
    Const conPropNotFound As Integer = 3270

    On Error GoTo ErrorSetAccessProperty
    ' Explicitly refer to Properties collection.
    obj.Properties(strName) = varSetting
    obj.Properties.Refresh
    SetAccessProperty = True

ExitSetAccessProperty:
    Exit Function

ErrorSetAccessProperty:
    If Err = conPropNotFound Then

        ' Create property, denote type, and set initial value.
        Set prp = obj.CreateProperty(strName, intType, varSetting)
        ' Append Property object to Properties collection.
        obj.Properties.Append prp
        obj.Properties.Refresh
        SetAccessProperty = True
        Resume ExitSetAccessProperty
    Else
        MsgBox Err & ": " & vbCrLf & Err.Description
        SetAccessProperty = False
        Resume ExitSetAccessProperty
    End If

End Function
Public Function GetDefaultPath()

    ' gets value of frmImages form's custom DefaultPath property

    Dim dbs As Database, doc As Document

    Set dbs = CurrentDb
    Set doc = dbs.Containers("Forms")!frmImages

    On Error Resume Next
    GetDefaultPath = doc.Properties("DefaultPath")
    If Err > 0 Then
        GetDefaultPath = "C:\"
    End If

End Function

Public Function GetFile(strPath As String) As String

    ' gets file name from path

    Dim n As Integer

    ' step back through path until backslash character
    ' encountered and return all charcters to right
    ' of backslash
    For n = Len(strPath) To 0 Step -1
        If Mid$(strPath, n, 1) = "\" Then
            GetFile = Mid$(strPath, n + 1)
            Exit For
        End If
    Next n

End Function

Public Function ShowImage()

    Dim frm As Form

    ' hide Web toolbar
    CommandBars("Web").Visible = False

    ' if frmImages has current record then
    ' set Picture property of image control on
    ' frmImages form to path to current image file,
    ' show image control and hide hyperlink label
    Set frm = Forms!frmImages
    With frm
        If Not IsNull(!ImageID) Then
            !Image0.Picture = !txtImagePath
            !Image0.Visible = True
            !Label1.Visible = False
            .SetFocus
        End If
    End With

End Function
