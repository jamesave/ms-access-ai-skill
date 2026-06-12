Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'Cara kerja form di sini
'Form di sini hanya untuk mengubah tampilan foto profil keping
'User akan memilih gambar; di mana form akan menampilkan foto gambar tersebut.
'Apabila gambar sudah tepat, maka user akan menggunakan tombol simpan untuk mengcopy gambar.
'Gambar dicopy dengan pembubuhan timestamp. file name ini akan disimpan di dalam ImageCode
'Untuk saat ini, gambar harus dalam jpg,  jpeg atau png
'Gambar yang disimpan dengan nama file "PK_<profilkepingAutoID>_timestamp"


'Default value untuk ImageCode adalah ImageDefaultFile (dari Mod_MainLib)


Dim jpgNm As String 'File name with timestamp
Dim DestPath As String 'DestPath = PcsPicPath from Mod_MainLib
Dim SelectedFile As String
Dim myfileExt As String
Dim OGFileName As String
Dim OGFilePath As String

Private Sub ResetImageCodetoDefault()
'     Place your reset code here
'     For example, clear all textboxes:
'        Dim ctrl As Control
'           For Each ctrl In Me.Controls
'           If TypeName(ctrl) = "TextBox" Then ctrl.Value = ""
'     Next ctrl
      Me.ImageCode = ""
      GetPKImagePath (Me.KepingAutoID)
'         SelectedFile = ImageDefaultFile
'         OGFileName = Mid(SelectedFile, InStrRev(SelectedFile, "\") + 1)
'         Me.ImageCode = OGFileName  'At this point, the db is dirty and OGfilename is ImageDefault.png
'         Me.ImgViewer.Picture = OGFileName 'Display Image


End Sub


Private Sub cmdCancel_Click()
On Error GoTo Err_Handler

   If Me.Dirty = True Then
      DoCmd.RunCommand (acCmdUndo)

   End If
    '  Me.cmdEditFotoProfil.SetFocus
      Me.KepingAutoID.SetFocus
      SetLockCtl (1)

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "cmdCancel_Click"
        Resume Exit_Handler
End Sub

Private Sub cmdDeleteImg_Click()
   If MsgBox("Yakin untuk menghapus gambar?", vbQuestion + vbYesNo) = vbYes Then
      ResetImageCodetoDefault
      Me.Dirty = False
      Me.KepingAutoID.SetFocus
      SetLockCtl (1)
   End If
End Sub

Private Sub cmdEditFotoProfil_Click()
'First of all you need to enable some Database references related to Access Objects
'Tools > Reference > Microsoft Office 16.0 Object Library


'SelectedFile is origin full path and
'OGfilename is the original file name
'jpgNm is the file name to be saved as
'destination path is saved in main lib


   Dim FD As FileDialog

   Set FD = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
    'FD.AllowMultiSelect = False
    FD.Title = "Select an Image"
    FD.Filters.Clear
    FD.Filters.Add "Image Files", "*.jpg; *.jpeg; *.png; *.bmp; *.gif"

       SetLockCtl (0)
   With FD
      .AllowMultiSelect = False
      .Title = "Pilih hanya satu file gambar."

      If .Show = True Then
         SelectedFile = FD.SelectedItems(1)
         OGFileName = Mid(SelectedFile, InStrRev(SelectedFile, "\") + 1)
         OGFilePath = Left(SelectedFile, InStrRev(SelectedFile, "\"))
        ' Me.ImageCode = OGfilename  ' Directly going to recordsource

         jpgNm = "PK_" & Me.KepingAutoID & "_" & GetTimeStamp & ".jpg"
         Me.ImageCode = jpgNm 'Inserting the address or the path as text into the field. At this point, the db is dirty.
         DestPath = PcsPicPath & jpgNm
         MsgBox "Copying from " & SelectedFile & "Copying to " & DestPath
         FileCopy SelectedFile, DestPath
         Me.Requery    ' Refreshes data
         Me.Repaint    ' Redraws the form

         'Me.ImgViewer.Picture = SelectedFile 'Display Image
'
'      Else
'         MsgBox "Wrong File, , , only JPG image file allowed.", vbExc1amation, "l Wrong File Format l"
'
       End If
    End With

      SetLockCtl (1)
End Sub

Private Sub cmdSimpanFotoProfil_Click()

      'This should work only if the Me.ImageCode = jpgNm is already dirty
      If Me.Dirty Then
         DestPath = PcsPicPath & jpgNm
         MsgBox "Copying from " & SelectedFile & "Copying to " & DestPath
         FileCopy SelectedFile, DestPath
         Me.ImgViewer.Picture = ""
         Me.ImgViewer.Picture = DestPath
      End If
      SetLockCtl (1)

End Sub


Private Sub Form_Dirty(Cancel As Integer)
   'Just in case the form is not locked and user change something.  Need to debug how the control is not working then.
   SetLockCtl (0)
   Me.lblWarning.Caption = "PERUBAHAN DATA"
   Me.lblWarning.ForeColor = RGB(237, 28, 36)
   Me.cmdCancel.Visible = True

End Sub

Private Sub SetLockCtl(SetStatus As Integer)
   On Error GoTo Err_Handler

   ' Status = 1 ; the form is locked
   ' Status is else when condition if the form is in edit mode

         If SetStatus = 1 Then
            Me.txtKepingID.Locked = True
            Me.txtDeskripsiProfilKeping.Locked = True
            Me.txtNoteGambar.Locked = True
            Me.txtImageCode.Locked = True

            Me.Detail.BackColor = 16777215 ' Background 1
            Me.lblWarning.Visible = False
            Me.lblWarning.Caption = "MODE DATA EDIT"
            Me.lblWarning.ForeColor = RGB(0, 0, 0)
            Me.lblWarning.BackColor = RGB(245, 157, 86)
            Me.cmdCancel.Visible = False ' This will give error when cancel cmd button is click since it's on focus
            Me.cmdDeleteImg.Visible = False
            Me.cmdSimpanFotoProfil.Visible = False

         Else

            Me.txtKepingID.Locked = False
            Me.txtDeskripsiProfilKeping.Locked = False
            Me.txtNoteGambar.Locked = False
            Me.txtImageCode.Locked = False
            Me.Detail.BackColor = RGB(200, 200, 255)
            Me.lblWarning.Visible = True
            Me.cmdCancel.Visible = True
            Me.cmdDeleteImg.Visible = True
            Me.cmdSimpanFotoProfil.Visible = True
         End If

Exit_Handler:
        Exit Sub

Err_Handler:
        clsErrorHandler.HandleError Me.Name, "SetLockCtl"
        Resume Exit_Handler
End Sub


Private Sub Form_Open(Cancel As Integer)
   SetLockCtl (1)
End Sub
