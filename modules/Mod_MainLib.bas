Attribute VB_Name = "Mod_MainLib"
Option Compare Database

' Main Library. Contains Paths to all the directories.
' Also contain commons command or tasks that can be used multiple forms

'MainPath is the main folder for all related images and files for thie project
Public Const MainPath As String = "C:\Users\gunai\OneDrive\Galangsa_GD\DBProduk\"

'PDFPath is the folder where the PDF files for info is stored
Public Const PDFPath As String = MainPath & "PDF\"
Public Const LingkarBusurPDF As String = "LingkarBusurArc125-410mm.pdf"
Public Const SkemaLiningPNG As String = "SkemaLining.png"


'PcsPicPath is folder to store picture of a piece of the product
Public Const PcsPicPath As String = "C:\Users\gunai\OneDrive\Galangsa_GD\DBProduk\Images\ProfilKeping\"

'HolePatternPath is folder to store the hole pattern images
Public Const HolePatternPath As String = "C:\Users\gunai\OneDrive\Galangsa_GD\DBProduk\Images\PolaLubang\"

'ImageDefaultFile is the file to show that the image is not available
Public Const ImageDefaultFile As String = "ImageDefault.png"
Public Const ImageDefaultPath As String = "C:\Users\gunai\OneDrive\Galangsa_GD\DBProduk\Images\"

'Return a string with a current time stamp with the format below.
Function GetTimeStamp() As String
  GetTimeStamp = Format(Now, "yymmdd_hhnnss")
End Function

'Here is the original blocks
'GetImagePath([id]) as ImagePath is called from the sql part in the form

' Function GetPKImagePath is to return the image path.
' Known Use: ProfilKeping to get ImageCode
' value 0 (default) for product images
' value 1 for hole pattern images
' DataID (is the file selected by the user).  The function returns path file where the selected file will be stored.  The idea is supposedly to check the "selected file"

Public Function GetPKImagePath(ByVal DataID As Long, Optional AlertBadName As Boolean = False) As String
  Dim ImageFolder As String
  Dim ImageName As String
  Dim ImagePath As String

  ImageName = getPKImageName(DataID)

   If ImageName = "" Then
        ImageName = ImageDefaultFile
        ImageFolder = ImageDefaultPath
   Else
        ImageName = ImageName
        ImageFolder = PcsPicPath
   End If

'  MsgBox "ImageName : " & ImageName & " DataID :" & DataID

   ImagePath = ImageFolder & ImageName
   GetPKImagePath = ImagePath
' FileExists = Dir(ImagePath)
'  If (FileExists = "") Then
'    If AlertBadName Then MsgBox "There is no file with name " & ImagePath, vbInformation
'    Debug.Print "There is no file with name " & ImagePath, vbInformation
'    GetPKImagePath = ImageFolder & BadImageName
'  Else
'    GetPKImagePath = ImagePath
'  End If

End Function

Public Function getPKImageName(DataID As Long) As String
   'this is hard coded for T_ProfilKeping at the moment
   getPKImageName = Nz(DLookup("ImageCode", "T_ProfilKeping", "KepingAutoID = " & DataID), "")
End Function

Public Function getKepingAutoID(DataID As String) As Long
   'this is hard coded for T_ProfilKeping at the moment
   getKepingAutoID = Nz(DLookup("KepingAutoID", "T_ProfilKeping", "KepingID = '" & DataID & "'"), "")

End Function

Public Function GetHPImagePath(ByVal DataID As String, Optional AlertBadName As Boolean = False) As String
  Dim ImageFolder As String
  Dim ImageName As String
  Dim ImagePath As String

  ImageName = getHPImageName(DataID)

   If ImageName = "" Then
        ImageName = ImageDefaultFile
        ImageFolder = ImageDefaultPath
   Else
        ImageName = ImageName
        ImageFolder = HolePatternPath
   End If

 'The following debug msgbox is not good idea to turn on especially if the data has too many rows
 'MsgBox "GetHPImagePath ImageName : " & ImageName & " DataID :" & DataID

   ImagePath = ImageFolder & ImageName
   GetHPImagePath = ImagePath
' FileExists = Dir(ImagePath)
'  If (FileExists = "") Then
'    If AlertBadName Then MsgBox "There is no file with name " & ImagePath, vbInformation
'    Debug.Print "There is no file with name " & ImagePath, vbInformation
'    GetPKImagePath = ImageFolder & BadImageName
'  Else
'    GetPKImagePath = ImagePath
'  End If

End Function

Public Function getHPImageName(DataID As String) As String
   'this is hard coded for T_ProfilKeping at the moment
   getHPImageName = Nz(DLookup("LokasiImage", "T_ListPolaLubang", "PolaLubang = '" & DataID & "'"), "")
End Function

Public Function ExtractMiddleID(ByVal strInput As String) As Long
'This require Tools > Reference > Microsoft VBScript Regular Expression 5.5
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    With regEx
        .Global = False
        .IgnoreCase = True
        .Pattern = "\D*(\d+)\D*"  ' Captures first sequence of digits as middle ID
    End With

    If regEx.Test(strInput) Then
        ExtractMiddleID = regEx.Execute(strInput)(0).SubMatches(0)
    Else
        ExtractMiddleID = 0
    End If
End Function
