Attribute VB_Name = "Mod_DragDropFile"
Option Compare Database
Option Explicit

' --- Windows API declarations ---
Public Declare PtrSafe Function DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long) As Long
Public Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" _
    (ByVal hDrop As LongPtr, ByVal iFile As Long, ByVal lpszFile As String, ByVal cch As Long) As Long
Public Declare PtrSafe Function DragFinish Lib "shell32.dll" (ByVal hDrop As LongPtr) As Long
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
    (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, _
     ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
    (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Public Const GWL_WNDPROC = -4
Public Const WM_DROPFILES = &H233

Public lpPrevWndProc As LongPtr
Public frmHwnd As LongPtr

' Custom window procedure
Public Function WindowProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    If uMsg = WM_DROPFILES Then
        Dim filePath As String
        Dim length As Long

        filePath = String(260, 0)
        length = DragQueryFile(wParam, 0, filePath, 260)
        filePath = Left(filePath, length)

        ' Set textbox and update image
        Forms!F_KelolaImg!ImagePath = filePath
        Forms!F_KelolaImg!imgPreview.Picture = filePath

        DragFinish wParam
    End If

    WindowProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
End Function

' Hook the window
Public Sub HookForm()
    Dim hWndForm As LongPtr

    hWndForm = FindWindow("OMain", vbNullString)
    If hWndForm <> 0 Then
        DragAcceptFiles hWndForm, True
        lpPrevWndProc = SetWindowLongPtr(hWndForm, GWL_WNDPROC, AddressOf WindowProc)
        frmHwnd = hWndForm
    End If
End Sub
