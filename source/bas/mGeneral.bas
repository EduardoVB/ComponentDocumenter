Attribute VB_Name = "mGeneral"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const DT_CALCRECT As Long = &H400
Private Const SM_CXEDGE As Long = 45
Private Const SM_CXVSCROLL As Long = 2
Private Const CB_GETMINVISIBLE As Long = &H1702&
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Const gstrSEP_DIR$ = "\"                         ' Directory separator character
'Private Const gstrAT$ = "@"
Private Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Private Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
'Private Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Private Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long
Private Declare Function GetACP Lib "kernel32" () As Long

Private Const CP_UTF8 As Long = 65001

Public gIcon As StdPicture

Public Function ConvertToUTF8(ByRef Source As String) As Byte()
  Dim length As Long
  Dim Pointer As Long
  Dim Size As Long
  Dim Buffer() As Byte
  Const CP_ACP As Long = 0
  
  If Len(Source) > 0 Then
    length = Len(Source)
    Pointer = StrPtr(Source)
    Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, length, 0, 0, 0, 0)
    If Size > 0 Then
        ReDim Buffer(0 To Size - 1)
        
        WideCharToMultiByte CP_UTF8, 0, Pointer, length, VarPtr(Buffer(0)), Size, 0, 0
        ConvertToUTF8 = Buffer
    Else
        Size = WideCharToMultiByte(CP_ACP, 0, Pointer, length, 0, 0, 0, 0)
        If Size > 0 Then
            ReDim Buffer(0 To Size - 1)
            
            WideCharToMultiByte CP_ACP, 0, Pointer, length, VarPtr(Buffer(0)), Size, 0, 0
            ConvertToUTF8 = Buffer
        End If
    End If
  End If
End Function

Public Function GetTempDir() As String
    Dim lChar As Long
    
    GetTempDir = String$(255, 0)
    lChar = GetTempPath(255, GetTempDir)
    GetTempDir = Left$(GetTempDir, lChar)
    AddDirSep GetTempDir
End Function

Public Sub AddDirSep(strPathName As String)
    strPathName = RTrim$(strPathName)
    If Right$(strPathName, Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR Then
        If Right$(strPathName, Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
            strPathName = strPathName & gstrSEP_DIR
        End If
    End If
End Sub


Public Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    
'    Debug.Print Err.Number, Err.Description
    FileExists = (Err.Number = 0) Or (Err.Number = 70)
    
    Close intFileNum

    Err.Clear
End Function

Private Sub SeparatePathAndFileName(FullPath As String, _
    Optional ByRef Path As String, _
    Optional ByRef FileName As String)

    Dim nSepPos As Long
    Dim nSepPos2 As Long
    Dim fUsingDriveSep As Boolean

    nSepPos = InStrRev(FullPath, gstrSEP_DIR)
    nSepPos2 = InStrRev(FullPath, gstrSEP_DIRALT)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
    End If
    nSepPos2 = InStrRev(FullPath, gstrSEP_DRIVE)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
        fUsingDriveSep = True
    End If

    If nSepPos = 0 Then
        'Separator was not found.
        Path = CurDir$
        FileName = FullPath
    Else
        If fUsingDriveSep Then
            Path = Left$(FullPath, nSepPos)
        Else
            Path = Left$(FullPath, nSepPos - 1)
        End If
        FileName = Mid$(FullPath, nSepPos + 1)
    End If
End Sub

Public Function GetFileName(nFileFullPath As String) As String
    Dim iFileName As String
    
    SeparatePathAndFileName nFileFullPath, , iFileName
    GetFileName = iFileName
End Function

Public Function GetFolder(nFileFullPath As String) As String
    Dim iFolderPath As String
    
    SeparatePathAndFileName nFileFullPath, iFolderPath
    GetFolder = iFolderPath
    AddDirSep GetFolder
End Function

Public Function FolderExists(ByVal nFolderPath As String) As Boolean
    On Error Resume Next

    FolderExists = (GetAttr(nFolderPath) And vbDirectory) = vbDirectory

    Err.Clear
End Function

Public Function RTFBold(nText As String) As String
    RTFBold = "\b " & nText & "\b0 "
End Function

Public Function RTFUnderline(nText As String) As String
    RTFUnderline = "\ul " & nText & "\ul0 "
End Function

Public Function RTFItalic(nText As String) As String
    RTFItalic = "\i " & nText & "\i0 "
End Function

Public Function AddToList(nList As Variant, nValue As Variant, Optional nOnlyIfMissing As Boolean, Optional nFirstElement As Long = 0) As Boolean
    Dim i As Long
    Dim iAdd As Boolean
    
    If Not nOnlyIfMissing Then
        iAdd = True
    Else
        iAdd = Not IsInList(nList, nValue, nFirstElement)
    End If
    If iAdd Then
        i = UBound(nList) + 1
        ReDim Preserve nList(LBound(nList) To i)
        nList(i) = nValue
        AddToList = True
    End If
End Function

Public Function IsInList(nList As Variant, nValue As Variant, Optional nFirstElement As Long = 0, Optional nLastElement As Long = -1) As Boolean
    Dim c As Long
    
    If nLastElement = -1 Then
        nLastElement = UBound(nList)
    Else
        If nLastElement > UBound(nList) Then
            nLastElement = UBound(nList)
        End If
    End If
    
    For c = nFirstElement To nLastElement
        If nList(c) = nValue Then
            IsInList = True
            Exit For
        End If
    Next c
End Function

Public Function IndexInList(nList As Variant, nValue As Variant) As Long
    Dim c As Long
    
    IndexInList = LBound(nList) - 1
    For c = LBound(nList) To UBound(nList)
        If nList(c) = nValue Then
            IndexInList = c
            Exit For
        End If
    Next c
End Function

Public Function Trim2(nText As String) As String
    Dim iChar As String
    
    Trim2 = nText
    iChar = Left$(Trim2, 1)
    Do While (iChar = " ") Or (iChar = vbTab) Or (iChar = vbCr) Or (iChar = vbLf) Or (iChar = Chr(160))
        Trim2 = Mid$(Trim2, 2)
        iChar = Left$(Trim2, 1)
    Loop
    iChar = Right$(Trim2, 1)
    Do While (iChar = " ") Or (iChar = vbTab) Or (iChar = vbCr) Or (iChar = vbLf) Or (iChar = Chr(160))
        Trim2 = Left$(Trim2, Len(Trim2) - 1)
        iChar = Right$(Trim2, 1)
    Loop
End Function

Public Sub SaveBinaryFile(nFilePath As String, nBytes() As Byte)
    Dim iFreeFile As Long
    
    iFreeFile = FreeFile
    Open nFilePath For Binary Access Write As #iFreeFile
    Put #iFreeFile, , nBytes
    Close #iFreeFile
End Sub

Public Function AppPath4Reg() As String
    Static sValue As String
    
    If sValue = "" Then
        sValue = Replace(App_Path, "\", "_")
    End If
    
    AppPath4Reg = sValue
End Function

Public Function AutoSizeDropDownWidth(Combo As Object) As Long
    '**************************************************************
    'PURPOSE: Automatically size the combo box drop down width
    '         based on the width of the longest item in the combo box
    
    'PARAMETERS: Combo - ComboBox to size
    
    'RETURNS: True if successful, false otherwise
    
    'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
    '                conversion from twips to pixels are made.
    '                API functions require units in pixels
    '
    '             2. Combo Box's parent is a form or other
    '                container that support the hDC property
    
    'EXAMPLE: AutoSizeDropDownWidth Combo1
    '****************************************************************
    Dim LRet As Long
    Dim lCurrentWidth As Single
    Dim rectCboText As RECT
    Dim lParentHDC As Long
    Dim lListCount As Long
    Dim lCtr As Long
    Dim lTempWidth As Long
    Dim lWidth As Long
    Dim sSavedFont As String
    Dim sngSavedSize As Single
    Dim bSavedBold As Boolean
    Dim bSavedItalic As Boolean
    Dim bSavedUnderline As Boolean
    Dim bFontSaved As Boolean
    Dim iRc As RECT
    Dim iMaxItemsWithoutScrollBar As Long
    
    On Error GoTo errorHandler
    
    If Not TypeOf Combo Is ComboBox Then Exit Function
    
    lParentHDC = Combo.Parent.hDC
    If lParentHDC = 0 Then Exit Function
    lListCount = Combo.ListCount
    If lListCount = 0 Then Exit Function
    
    'Change font of parent to combo box's font
    'Save first so it can be reverted when finished
    'this is necessary for drawtext API Function
    'which is used to determine longest string in combo box
    With Combo.Parent
        sSavedFont = .FontName
        sngSavedSize = .FontSize
        bSavedBold = .FontBold
        bSavedItalic = .FontItalic
        bSavedUnderline = .FontUnderLine
        
        .FontName = Combo.FontName
        .FontSize = Combo.FontSize
        .FontBold = Combo.FontBold
        .FontItalic = Combo.FontItalic
        .FontUnderLine = Combo.FontItalic
    End With
    
    bFontSaved = True
    
    'Get the width of the largest item
    For lCtr = 0 To lListCount
       DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, DT_CALCRECT
       'adjust the number added (20 in this case to
       'achieve desired right margin
       lTempWidth = rectCboText.Right - rectCboText.Left + GetSystemMetrics(SM_CXEDGE) * 2
    
       If (lTempWidth > lWidth) Then
          lWidth = lTempWidth
       End If
    Next
     
    iMaxItemsWithoutScrollBar = SendMessageLong(Combo.hWnd, CB_GETMINVISIBLE, 0&, 0&)
    
    If Combo.ListCount > iMaxItemsWithoutScrollBar Then
         lTempWidth = lTempWidth + GetSystemMetrics(SM_CXVSCROLL)
    End If
     
     
    GetWindowRect Combo.hWnd, iRc
    LRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, iRc.Right - iRc.Left, 0)
    
    lCurrentWidth = SendMessageLong(Combo.hWnd, CB_GETDROPPEDWIDTH, 0, 0)
    
    If lCurrentWidth > lWidth Then 'current drop-down width is
    '                               sufficient
'        AutoSizeDropDownWidth = True
        AutoSizeDropDownWidth = lCurrentWidth
        GoTo errorHandler
        Exit Function
    End If
     
    'don't allow drop-down width to
    'exceed screen.width
    If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20
    
    LRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)
    AutoSizeDropDownWidth = lWidth
'    AutoSizeDropDownWidth = LRet > 0

errorHandler:
    On Error Resume Next
    If bFontSaved Then
    'restore parent's font settings
      With Combo.Parent
        .FontName = sSavedFont
        .FontSize = sngSavedSize
        .FontUnderLine = bSavedUnderline
        .FontBold = bSavedBold
        .FontItalic = bSavedItalic
     End With
    End If
End Function

Public Property Get App_Path() As String
    Static sValue As String
    
    If sValue = "" Then
        sValue = App.Path
        If Right$(sValue, 1) = "\" Then
            sValue = Left$(sValue, Len(sValue) - 1)
        End If
    End If
    App_Path = sValue
End Property
