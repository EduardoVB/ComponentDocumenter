Attribute VB_Name = "mBrowseForFolder"
Option Explicit

Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_VALIDATEFAILED = 3

Private Const MAX_PATH = 260

Public gWindowTitle As String
Public gCommonDialogEx_ShowFolder_StartFolder As String

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
   On Error Resume Next
   Dim ret As Long
   Dim sBuffer As String
   Dim iEh As Long
   
   Select Case uMsg
       Case BFFM_INITIALIZED
           Call SendMessageString(hWnd, BFFM_SETSELECTION, 1, gCommonDialogEx_ShowFolder_StartFolder)
           SetWindowText hWnd, gWindowTitle
            iEh = FindWindowEx(hWnd, 0, "Edit", "")
            SetWindowText iEh, gCommonDialogEx_ShowFolder_StartFolder
       Case BFFM_SELCHANGED
           sBuffer = Space(MAX_PATH)
           ret = SHGetPathFromIDList(lp, sBuffer)
           If ret = 1 Then
               Call SendMessageString(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
               iEh = FindWindowEx(hWnd, 0, "Edit", "")
               SetWindowText iEh, sBuffer
           End If
        Case BFFM_VALIDATEFAILED
            Call SendMessageString(hWnd, BFFM_SETSELECTION, 1, "")
   End Select
   BrowseCallbackProc = 0
End Function
