VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Component Documenter"
   ClientHeight    =   7692
   ClientLeft      =   2220
   ClientTop       =   2400
   ClientWidth     =   11004
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7692
   ScaleWidth      =   11004
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrCheckcboReListVisible 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7620
      Top             =   2220
   End
   Begin VB.ComboBox cboRef 
      BackColor       =   &H00F4FFFE&
      Height          =   336
      Left            =   8040
      TabIndex        =   18
      Text            =   "cboRef"
      Top             =   2220
      Visible         =   0   'False
      Width           =   1872
   End
   Begin VB.CommandButton cmdReference 
      Caption         =   "R"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "(Ctrl+J) Reference (make a link to) a Control, Class, Enum, Property, Method or Event"
      Top             =   2220
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   10.2
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":048A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Add link tag to reference something"
      Top             =   2220
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdBold 
      Caption         =   "B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Add Bold tag"
      Top             =   2220
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Timer tmrSetFocus 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   6240
   End
   Begin VB.CommandButton cmdAppliesTo 
      Caption         =   "···"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10320
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.CommandButton cmdLongDescriptionMenu 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   10380
      TabIndex        =   8
      Top             =   2220
      Visible         =   0   'False
      Width           =   372
   End
   Begin ComctlLib.TreeView trv1 
      Height          =   5868
      Left            =   144
      TabIndex        =   0
      Top             =   120
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   10351
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   2
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin RichTextLib.RichTextBox rtbAux 
      Height          =   900
      Left            =   9360
      TabIndex        =   13
      Top             =   7872
      Visible         =   0   'False
      Width           =   804
      _ExtentX        =   1418
      _ExtentY        =   1588
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":07CC
   End
   Begin VB.TextBox txtParamsInfo 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   576
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   6372
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4440
      TabIndex        =   12
      Top             =   4968
      Visible         =   0   'False
      Width           =   6372
   End
   Begin VB.Timer tmrNodeRightClick 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   6240
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   480
      Top             =   6240
   End
   Begin VB.TextBox txtShortDescription 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4868
      Visible         =   0   'False
      Width           =   6372
   End
   Begin VB.TextBox txtLongDescription 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2580
      Visible         =   0   'False
      Width           =   6372
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   4440
      TabIndex        =   3
      Top             =   808
      Visible         =   0   'False
      Width           =   6372
   End
   Begin VB.Label lblAppliesTo 
      Alignment       =   2  'Center
      Caption         =   "This definition applies to n objects"
      Height          =   252
      Left            =   7380
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   3036
   End
   Begin VB.Label lblParamsInfo 
      Caption         =   "Parameters:"
      Height          =   252
      Left            =   4440
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   2004
   End
   Begin VB.Label lblCurrentAction 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F32732&
      Height          =   360
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   4572
   End
   Begin VB.Label lblShortDescription 
      Caption         =   "Short description:"
      Height          =   252
      Left            =   4440
      TabIndex        =   10
      Top             =   4556
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label lblLongDescription 
      AutoSize        =   -1  'True
      Caption         =   "Long description:"
      Height          =   240
      Left            =   4440
      TabIndex        =   7
      Top             =   2256
      Visible         =   0   'False
      Width           =   1368
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   252
      Left            =   4440
      TabIndex        =   2
      Top             =   520
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Menu mnuComponent 
      Caption         =   "&Component"
      Begin VB.Menu mnuNewComponentDB 
         Caption         =   "&New Component database"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenComponentDB 
         Caption         =   "&Open Component database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnyDeleteComponentDB 
         Caption         =   "Delete component database"
      End
      Begin VB.Menu mnuComponentProperties 
         Caption         =   "Component properties"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Enabled         =   0   'False
      Begin VB.Menu mnuImport 
         Caption         =   "&Import data from OCX/DLL file"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDataDelete 
         Caption         =   "Delete current item"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLoadFromOrphanMember2 
         Caption         =   "Load long description from orphan member"
      End
      Begin VB.Menu mnuSetMethodToExistentDefinition2 
         Caption         =   "Set member to an existent definition"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListBlankFields 
         Caption         =   "List blank fields"
      End
      Begin VB.Menu mnuListMarkupLinksErrors 
         Caption         =   "List markup link errors"
      End
      Begin VB.Menu mnuListOrphanMembers 
         Caption         =   "List orphan members"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteOrphanMembers 
         Caption         =   "Delete orphan members"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete all data"
      End
   End
   Begin VB.Menu mnuPopupGeneral 
      Caption         =   "mnuPopupGeneral"
      Visible         =   0   'False
      Begin VB.Menu mnuNewClass 
         Caption         =   "New Class"
      End
      Begin VB.Menu mnuNewControl 
         Caption         =   "New Control"
      End
      Begin VB.Menu mnuNewEnum 
         Caption         =   "New Enum"
      End
   End
   Begin VB.Menu mnuPopupClassOrControl 
      Caption         =   "mnuPopupClassOrControl"
      Visible         =   0   'False
      Begin VB.Menu mnuNewProperty 
         Caption         =   "New Property"
      End
      Begin VB.Menu mnuNewMethod 
         Caption         =   "New Method"
      End
      Begin VB.Menu mnuNewEvent 
         Caption         =   "New Event"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteObject 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyName 
         Caption         =   "Copy Name"
      End
   End
   Begin VB.Menu mnuPopupClasses 
      Caption         =   "mnuPopupClasses"
      Visible         =   0   'False
      Begin VB.Menu mnuNewClass2 
         Caption         =   "New Class"
      End
   End
   Begin VB.Menu mnuPopupControls 
      Caption         =   "mnuPopupControls"
      Visible         =   0   'False
      Begin VB.Menu mnuNewControl2 
         Caption         =   "New Control"
      End
   End
   Begin VB.Menu mnuPopupEnums 
      Caption         =   "mnuPopupEnums"
      Visible         =   0   'False
      Begin VB.Menu mnuNewEnum2 
         Caption         =   "New Enum"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnumsOrderAllByName 
         Caption         =   "Order all by name"
      End
      Begin VB.Menu mnuEnumsOrderAllByValue 
         Caption         =   "Order all by value"
      End
   End
   Begin VB.Menu mnuPopupMembersParent 
      Caption         =   "mnuPopupMembersParent"
      Visible         =   0   'False
      Begin VB.Menu mnuNewMember 
         Caption         =   "New"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyList 
         Caption         =   "Copy list"
      End
   End
   Begin VB.Menu mnuPopupMember 
      Caption         =   "mnuPopupMember"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteMember 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyName2 
         Caption         =   "Copy name"
      End
   End
   Begin VB.Menu mnuPopupEnum 
      Caption         =   "mnuPopupEnum"
      Visible         =   0   'False
      Begin VB.Menu mnuNewEnumConstant 
         Caption         =   "New constant"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteObject2 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyName3 
         Caption         =   "Copy Name"
      End
      Begin VB.Menu mnuCopyEnumList 
         Caption         =   "Copy list"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConstantsOrderedByName 
         Caption         =   "Ordered by Name"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuConstantsOrderedByValue 
         Caption         =   "Ordered by Value"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Enabled         =   0   'False
      Begin VB.Menu mnuReportHTML 
         Caption         =   "HTML"
      End
      Begin VB.Menu mnuReportRTF 
         Caption         =   "RTF"
      End
      Begin VB.Menu mnuReportPrint 
         Caption         =   "Print (PDF generation)"
      End
      Begin VB.Menu mnuReportText 
         Caption         =   "Plain text"
      End
      Begin VB.Menu Sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportingOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuPopupLongDescription 
      Caption         =   "mnuPopupLongDescription"
      Visible         =   0   'False
      Begin VB.Menu mnuLoadFromOrphanMember 
         Caption         =   "Load long description from orphan member"
      End
      Begin VB.Menu mnuSetMethodToExistentDefinition 
         Caption         =   "Set to an existent definition"
      End
      Begin VB.Menu mnuCopyFromSameNameMember 
         Caption         =   "mnuCopyFromSameNameMember"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPreferences 
      Caption         =   "Preferences"
   End
   Begin VB.Menu mnuLinkType 
      Caption         =   "mnuLinkType"
      Visible         =   0   'False
      Begin VB.Menu mnuLinkToControl 
         Caption         =   "Link to Control"
      End
      Begin VB.Menu mnuLinkToClass 
         Caption         =   "Link to Class"
      End
      Begin VB.Menu mnuLinkToEnum 
         Caption         =   "Link to Enum"
      End
      Begin VB.Menu mnuLinkToProperty 
         Caption         =   "Link to Property"
      End
      Begin VB.Menu mnuLinkToMethod 
         Caption         =   "Link to Method"
      End
      Begin VB.Menu mnuLinkToEvent 
         Caption         =   "Link to Event"
      End
   End
   Begin VB.Menu mnuRefType 
      Caption         =   "mnuRefType"
      Visible         =   0   'False
      Begin VB.Menu mnuRefTypeList 
         Caption         =   "Reference a Control"
         Index           =   0
      End
      Begin VB.Menu mnuRefTypeList 
         Caption         =   "Reference a Class"
         Index           =   1
      End
      Begin VB.Menu mnuRefTypeList 
         Caption         =   "Reference an Enum"
         Index           =   2
      End
      Begin VB.Menu mnuRefTypeList 
         Caption         =   "Reference a Property"
         Index           =   3
      End
      Begin VB.Menu mnuRefTypeList 
         Caption         =   "Reference a Method"
         Index           =   4
      End
      Begin VB.Menu mnuRefTypeList 
         Caption         =   "Reference an Event"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum CDPrintMode
    cdPrintModeContiguous = 0
    cdSeparatePages = 1
End Enum

Private Enum CDHTMLMode
    cdHTMLOnePage = 0
    cdHTMLPerObject = 1
    cdHTMLPerMethod = 2
End Enum

Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton As Long
   hwndCombo As Long
   hwndEdit As Long
   hWndList As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type
 
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Const SW_SHOW = 5
Private Const CB_SHOWDROPDOWN = &H14F

Private Type Charrange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hDC As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As Charrange ' Range of text to draw (see above declaration)
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const TV_FIRST = &H1100
Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_SELECTITEM = (TV_FIRST + 11)
Private Const TVGN_CARET = 9
Private Const TVGN_FIRSTVISIBLE = &H5

Private Const PARAMFLAG_NONE As Long = &H0&
Private Const PARAMFLAG_FIN As Long = &H1&
Private Const PARAMFLAG_FOUT As Long = &H2&
Private Const PARAMFLAG_FLCID As Long = &H4&
Private Const PARAMFLAG_FRETVAL As Long = &H8&
Private Const PARAMFLAG_FOPT As Long = &H10&
Private Const PARAMFLAG_FHASDEFAULT As Long = &H20&
Private Const PARAMFLAG_FHASCUSTDATA As Long = &H40&
Private mCurrentDBPath As String

Private Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long

Private mControlsEditZone As Collection
Private mControlsEditZoneVisible As Boolean

Private Enum ECurrentAction
    ecaDefault
    ecaAddClass
    ecaAddControl
    ecaAddEnum
    ecaEditClass
    ecaEditControl
    ecaEditEnum
    ecaAddProperty
    ecaAddMethod
    ecaAddEvent
    ecaEditProperty
    ecaEditMethod
    ecaEditEvent
    ecaAddConstant
    ecaEditConstant
    ecaEditIntroduction
    ecaEditEndNotes
End Enum

Private Enum ENodeType
    entNone
    entClassesParent
    entControlsParent
    entEnumsParent
    entClass
    entControl
    entEnum
    entPropertiesParent
    entMethodsParent
    entEventsParent
    entProperty
    entMethod
    entEvent
    entConstant
    enIntroduction
    entEndNotes
End Enum

Private mObjectType_s(2) As String
Private mObjectType_s2(2) As String
Private mObjectType_p(2) As String
Private mMemberType_s(3) As String
Private mMemberType_p(3) As String
Private mMemberTypeRec(3) As Recordset

Private mCurrentAction As ECurrentAction
Private mDatabase As Database

Private mClasses As Recordset
Private mControls As Recordset
Private mEnums As Recordset
Private mProperties As Recordset
Private mMethods As Recordset
Private mEvents As Recordset
Private mConstants As Recordset
Private mGeneral_Information As Recordset

Private mSelectedType As ENodeType
Private mSelectedID As Long
Private mSelectedSecondaryID As Long
Private mShowingTree As Boolean
Private mDeletingNode As Boolean

Private Const cHTMLDefaultHeadSection As String = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"" >" & _
    vbCrLf & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & _
    vbCrLf & "<head>" & vbCrLf & _
    "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf & _
    "<meta charset=""utf-8"" />" & vbCrLf & _
    "<!--[PAGE_TITLE]-->" & vbCrLf & _
    "<!--[PAGE_DESCRIPTION]-->" & vbCrLf & _
    "<!--[STYLESHEET_INFO]-->" & vbCrLf & _
    "</head>"

Private Const cHTMLDefaultStyleSheet1 As String = "body {" & vbCrLf & _
    "  background-color: white;" & vbCrLf & _
    "  font:.875em/1.35 'Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;" & vbCrLf & _
    "  margin-top:0;" & vbCrLf & _
    "  margin-left:20;" & vbCrLf & _
    "  margin-bottom:0;" & vbCrLf & _
    "  padding-bottom:15px;" & vbCrLf & _
    "  padding-left:15px;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf & _
    "h1 {" & vbCrLf & _
    "  color: black;" & vbCrLf & _
    "  font-family:'Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;" & vbCrLf & _
    "  font-weight:normal;" & vbCrLf & _
    "  font-size: 200%;" & vbCrLf & _
    "  margin-top:0;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf & _
    "p  {" & vbCrLf & _
    "  color:#2a2a2a;" & vbCrLf & _
    "  font-size: 100%;" & vbCrLf & _
    "  line-height:18px;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf
Private Const cHTMLDefaultStyleSheet2 As String = cHTMLDefaultStyleSheet1 & "a {color:#069;" & vbCrLf & _
    "  text-decoration:none;" & vbCrLf & _
    "  font-weight: 600;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf & _
    "a:visited {" & vbCrLf & _
    "  color:#069;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf & _
    "pre {background: #f4f4f4;" & vbCrLf & _
    "    padding: 5px;" & vbCrLf & _
    "    margin:0;" & vbCrLf & _
    "    font-family: Consolas,Courier,monospace!important;" & vbCrLf & _
    "    font-style: normal;" & vbCrLf & _
    "    font-weight: normal;" & vbCrLf & _
    "    overflow: auto;" & vbCrLf & _
    "    word-wrap: normal;" & vbCrLf & _
    "    border: 1px solid #ddd;" & vbCrLf & _
    "    line-height: 1.6;" & vbCrLf & _
    "    display: block;" & vbCrLf & _
    "   border: 3px solid #eee;" & vbCrLf & _
    "   border-left: 3px solid #aca;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf
Private Const cHTMLDefaultStyleSheet As String = cHTMLDefaultStyleSheet2 & "header {" & vbCrLf & _
    "  background-color: white;" & vbCrLf & _
    "  font:.875em/1.35 'Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;" & vbCrLf & _
    "  margin-top:0;" & vbCrLf & _
    "  margin-left:0;" & vbCrLf & _
    "  margin-bottom:0;" & vbCrLf & _
    "  font-size: 130%;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf & _
    "header a {color:#066;" & vbCrLf & _
    "  text-decoration:none;" & vbCrLf & _
    "  font-weight: 400;" & vbCrLf & _
    "  font-size: 100%;" & vbCrLf & _
    " }" & vbCrLf & _
    "" & vbCrLf & _
    "header a:visited {" & vbCrLf & _
    "  color:#066;" & vbCrLf & _
    " }" & vbCrLf

Private Const cHTMLDefaultPageHeaderMP As String = "<header>" & vbCrLf & _
    "<br>" & vbCrLf & _
    "> [COMPONENT_NAME] > <a href=""index.html""> Reference</a> > [CURRENT_ITEM]" & vbCrLf & _
    "<br><br>" & vbCrLf & _
    "</header>"
    
Private Const cHTMLDefaultPageHeaderOP As String = ""
Private Const cHTMLDefaultPageFooter As String = ""

Private Const cRTFHeaders As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang11274{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 Consolas;}}{\colortbl ;\red0\green0\blue0;\red14\green22\blue190;\red0\green128\blue0;}\viewkind4\uc1\cf1\fs24"

Private mTOCItems() As String
Private mTOCItems_Level() As Long
Private mTOCItems_Page() As Long

Private mTOC_Ub As Long
Private mTOC_Index As Long

Private mLinkErrors() As String
Private mComponentName As String
Private mComponentVersion As String
Private mComponentReleaseDate As Date
Private mHTML_Mode As CDHTMLMode
Private mPrint_Mode As CDPrintMode
Private mExternalCSS As Boolean
Private mReplaceCSSFile As Boolean
Private mMargin As Long
Private mHTML_HeadSection As String
Private mHTML_StyleSheet As String
Private mHTML_PageHeaderMP_Template As String
Private mHTML_PageHeaderOP_Template As String
Private mHTML_PageHeaderMP As String
Private mHTML_PageHeaderOP As String
Private mHTML_PageFooter_Template As String
Private mHTML_PageFooter As String
Private mPages() As String
Private mPages_FileNames() As String
Private mNonUniqueMemberNamePages() As String
Private mNewEnumsOrderedByValue As Boolean
Public PrinterIndex As Long
Private mFileImported As Boolean
Private mfrmFieldsModifAlertShowed As Boolean
Private mAppFont As StdFont
Private mFontPropertion As Single
Private mLinkTypeSelected As String
Private mRefTables() As String
Private mRefTableIndex As Long
Private mcboRefhWndList As Long

Public Function UnfoldRelativePath(ByVal sPath As String) As String
    Dim sBuff As String
    
    sBuff = Space$(261)
    If PathCanonicalize(sBuff, sPath) Then
        UnfoldRelativePath = Left$(sBuff, InStr(sBuff, vbNullChar) - 1)
    Else
        UnfoldRelativePath = sPath
    End If
End Function

Private Sub cboRef_Click()
    Dim iStr As String
    
    Select Case mRefTableIndex
        Case 0 ' controls
            iStr = "[c["
        Case 1 ' classes
            iStr = "[o["
        Case 2 ' enums
            iStr = "[["
        Case 3 ' properties
            iStr = "[p["
        Case 4 ' methods
            iStr = "[m["
        Case 5 ' events
            iStr = "[e["
    End Select
    iStr = iStr & cboRef.Text & "]]"
    
    txtLongDescription.SelText = iStr
    txtLongDescription.SetFocus
    cboRef.Visible = False
End Sub

Private Sub cboRef_LostFocus()
    txtLongDescription.SetFocus
    cboRef.Visible = False
End Sub

Private Sub cmdAppliesTo_Click()
    MsgBox cmdAppliesTo.ToolTipText & "."
End Sub

Private Sub cmdBold_Click()
    Dim iSS As Long
    Dim iSL As Long
    Dim txt As TextBox
    
    If cmdBold.Tag <> "1" Then Exit Sub
    
    Set txt = txtLongDescription
    iSS = txt.SelStart
    iSL = txt.SelLength
    
    txt.SelLength = 0
    txt.SelText = "<b>"
    txt.SelStart = iSS + iSL + 3
    txt.SelText = "</b>"
    txt.SelStart = iSS
    txt.SelLength = iSL + 7
    txt.SetFocus
    cmdBold.Tag = ""
End Sub

Private Sub cmdLink_Click()
    Dim iSS As Long
    Dim iSL As Long
    Dim txt As TextBox
    Dim iStart As String
    Dim iRec As Recordset
    Dim iST As String
    Dim iControls As Boolean
    Dim iClasses As Boolean
    Dim iEnums As Boolean
    Dim iProperties As Boolean
    Dim iMethods As Boolean
    Dim iEvents As Boolean
    Dim c As Long
    
    If cmdLink.Tag <> "1" Then Exit Sub
    
    Set txt = txtLongDescription
    iSS = txt.SelStart
    iSL = txt.SelLength
    
    If iSL > 0 Then
        iST = txt.SelText
        c = 0
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls WHERE (Name = '" & iST & "')")
        If (iRec.RecordCount > 0) Then iControls = True: c = c + 1
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes WHERE (Name = '" & iST & "')")
        If (iRec.RecordCount > 0) Then iClasses = True: c = c + 1
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Enums WHERE (Name = '" & iST & "')")
        If (iRec.RecordCount > 0) Then iEnums = True: c = c + 1
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Properties WHERE (Name = '" & iST & "')")
        If (iRec.RecordCount > 0) Then iProperties = True: c = c + 1
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Methods WHERE (Name = '" & iST & "')")
        If (iRec.RecordCount > 0) Then iMethods = True: c = c + 1
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Events WHERE (Name = '" & iST & "')")
        If (iRec.RecordCount > 0) Then iEvents = True: c = c + 1
        If c = 0 Then
            MsgBox "Nothing found with that name", vbExclamation
            Exit Sub
        ElseIf c = 1 Then
            Select Case True
                Case iControls
                    iStart = "[c["
                Case iClasses
                    iStart = "[o["
                Case iEnums
                    iStart = "[["
                Case iProperties
                    iStart = "[p["
                Case iMethods
                    iStart = "[m["
                Case iEvents
                    iStart = "[e["
            End Select
        Else
            mnuLinkToControl.Visible = iControls
            mnuLinkToClass.Visible = iClasses
            mnuLinkToEnum.Visible = iEnums
            mnuLinkToProperty.Visible = iProperties
            mnuLinkToMethod.Visible = iMethods
            mnuLinkToEvent.Visible = iEvents
            mLinkTypeSelected = ""
            PopupMenu mnuLinkType
            If mLinkTypeSelected = "" Then Exit Sub
            iStart = mLinkTypeSelected
        End If
    Else
        iStart = "[["
    End If
    
    txt.SelLength = 0
    txt.SelText = iStart
    txt.SelStart = iSS + iSL + Len(iStart)
    txt.SelText = "]]"
    txt.SelStart = iSS
    txt.SelLength = iSL + 2 + Len(iStart)
    txt.SetFocus
    cmdLink.Tag = ""
End Sub

Private Sub cmdLongDescriptionMenu_Click()
    Dim iCurrentType As Long
    Dim iCurrentID As Long
    Dim iRec As Recordset
    Dim c As Long
    
    iCurrentType = CurrentType
    iCurrentID = GetCurrentMemberID
    
    mnuLoadFromOrphanMember.Enabled = ThereAreOrphanMembers And ((mCurrentAction = ecaEditProperty) Or (mCurrentAction = ecaEditMethod) Or (mCurrentAction = ecaEditEvent))
    If (iCurrentType > 0) And (iCurrentID <> 0) Then
        mnuSetMethodToExistentDefinition.Enabled = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(iCurrentType) & " WHERE (Name = '" & txtName.Text & "') AND (" & mMemberType_s(iCurrentType) & "_ID <> " & CStr(iCurrentID) & ") AND (Auxiliary_Field = 1)").RecordCount > 0
    Else
        mnuSetMethodToExistentDefinition.Enabled = False
    End If
    
    For c = 1 To mnuCopyFromSameNameMember.UBound
        Unload mnuCopyFromSameNameMember(c)
    Next
    mnuCopyFromSameNameMember(0).Visible = False
    
    If (iCurrentType > 0) And (iCurrentID <> 0) Then
        Debug.Print
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(iCurrentType) & " WHERE Name = ('" & txtName.Text & "') AND (" & mMemberType_s(iCurrentType) & "_ID <> " & iCurrentID & ") AND (Long_Description <> '') AND (Long_Description <> '" & Replace(txtLongDescription.Text, "'", "''") & "')")
        If iRec.RecordCount > 0 Then
            iRec.MoveLast
            If iRec.RecordCount = 1 Then
                mnuCopyFromSameNameMember(0).Caption = "Load from " & LCase$(mMemberType_s(iCurrentType)) & " of the same name"
                mnuCopyFromSameNameMember(0).Tag = iRec(mMemberType_s(iCurrentType) & "_ID").Value
                mnuCopyFromSameNameMember(0).Visible = True
            Else
                iRec.MoveFirst
                c = 0
                Do Until iRec.EOF
                    If c > 0 Then Load mnuCopyFromSameNameMember(c)
                    mnuCopyFromSameNameMember(c).Caption = "Load from " & LCase$(mMemberType_s(iCurrentType)) & " of the same name " & c + 1
                    mnuCopyFromSameNameMember(c).Tag = iRec(mMemberType_s(iCurrentType) & "_ID").Value
                    mnuCopyFromSameNameMember(c).Visible = True
                    c = c + 1
                    iRec.MoveNext
                Loop
            End If
        End If
    End If
    PopupMenu mnuPopupLongDescription
End Sub

Private Sub cmdReference_Click()
    Dim iRec As Recordset
    Dim c As Long
    Dim iRCList As RECT
    Dim iRCCombo As RECT
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iPt As POINTAPI
    
    If cmdReference.Tag <> "1" Then Exit Sub
    
    cboRef.Left = -10000
    cboRef.Clear
    Set cboRef.Font = txtLongDescription.Font
    
    For c = 0 To UBound(mRefTables)
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mRefTables(c))
        mnuRefTypeList(c).Visible = (iRec.RecordCount > 0)
    Next
    
    GetCaretPos iPt
    ClientToScreen txtLongDescription.hWnd, iPt
    ScreenToClient Me.hWnd, iPt
    
    mRefTableIndex = -1
    PopupMenu mnuRefType, , iPt.X * Screen.TwipsPerPixelX, iPt.Y * Screen.TwipsPerPixelY
    If mRefTableIndex = -1 Then Exit Sub
    
    cboRef.Left = -Me.Left - 10000
    cboRef.Clear
    Set iRec = mDatabase.OpenRecordset("SELECT DISTINCT Name FROM " & mRefTables(mRefTableIndex) & " ORDER BY Name")
    iRec.MoveFirst
    Do Until iRec.EOF
        cboRef.AddItem iRec!Name
        iRec.MoveNext
    Loop
    
    AutoSizeDropDownWidth cboRef
    cboRef.Visible = True
    Call SendMessage(cboRef.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)
    
    mcboRefhWndList = GetComboListHwnd(cboRef)
    
    GetWindowRect mcboRefhWndList, iRCList
    iWidth = iRCList.Right - iRCList.Left
    iHeight = iRCList.Bottom - iRCList.Top
    
    txtLongDescription.SetFocus
    
    GetCaretPos iPt
    ClientToScreen txtLongDescription.hWnd, iPt
    If (iPt.Y + iHeight) > Screen.Height / Screen.TwipsPerPixelY Then
        iPt.Y = Screen.Height / Screen.TwipsPerPixelY - iHeight
    End If
    iRCList.Left = iPt.X
    iRCList.Top = iPt.Y
    iRCList.Right = iRCList.Left + iWidth
    iRCList.Bottom = iRCList.Top + iHeight
    
    cboRef.SetFocus
    
    SetWindowPos mcboRefhWndList, 0, iRCList.Left, iRCList.Top, iWidth, iRCList.Bottom - iRCList.Top, 0
    ShowWindow mcboRefhWndList, SW_SHOW
    tmrCheckcboReListVisible.Enabled = True
    
    cmdReference.Tag = ""
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    trv1_MouseUp Button, Shift, X, Y
End Sub

Private Sub mnuCopyFromSameNameMember_Click(Index As Integer)
    Dim iRec As Recordset
    Dim iCurrentType As Long
    
    iCurrentType = CurrentType
    If iCurrentType > 0 Then
        If Trim(txtLongDescription.Text) <> "" Then
            If MsgBox("This will replace the current text and can't be undone, continue?", vbYesNo) = vbNo Then Exit Sub
        End If
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(iCurrentType))
        iRec.FindFirst mMemberType_s(iCurrentType) & "_ID = " & mnuCopyFromSameNameMember(Index).Tag
        If Not iRec.NoMatch Then
            txtLongDescription.Text = iRec!Long_Description
        End If
    End If
End Sub

Private Sub mnuData_Click()
    Dim iCurrentType As Long
    Dim iCurrentID As Long
    
    UpdateData
    
    iCurrentType = CurrentType
    iCurrentID = GetCurrentMemberID
    
    mnuLoadFromOrphanMember2.Enabled = ThereAreOrphanMembers And ((mCurrentAction = ecaEditProperty) Or (mCurrentAction = ecaEditMethod) Or (mCurrentAction = ecaEditEvent))
    If (iCurrentType > 0) And (iCurrentID <> 0) Then
        mnuSetMethodToExistentDefinition2.Enabled = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(iCurrentType) & " WHERE (Name = '" & txtName.Text & "') AND (" & mMemberType_s(iCurrentType) & "_ID <> " & CStr(iCurrentID) & ") AND (Auxiliary_Field = 1)").RecordCount > 0
    Else
        mnuSetMethodToExistentDefinition2.Enabled = False
    End If
End Sub

Private Sub mnuDataAdd_Click()
    If (mSelectedType = entPropertiesParent) Or (mSelectedType = entMethodsParent) Or (mSelectedType = entEventsParent) Then
        mnuNewMember_Click
    ElseIf (mSelectedType = entEnum) Then
        mnuNewEnumConstant_Click
    Else
        PopupMenu mnuPopupGeneral, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub mnuDataDelete_Click()
    If (mSelectedType = entClass) Or (mSelectedType = entControl) Or (mSelectedType = entEnum) Then
        mnuDeleteObject_Click
    ElseIf (mSelectedType = entProperty) Or (mSelectedType = entMethod) Or (mSelectedType = entEvent) Then
        mnuDeleteMember_Click
    ElseIf (mSelectedType = entConstant) Then
        mnuDeleteMember_Click
    End If
End Sub

Private Sub mnuEnumsOrderAllByName_Click()
    mEnums.MoveFirst
    Do Until mEnums.EOF
        mEnums.Edit
        mEnums!Ordered_By_Value = False
        mEnums.Update
        mEnums.Bookmark = mEnums.LastModified
        mEnums.MoveNext
    Loop
    mNewEnumsOrderedByValue = False
    SaveSettingBase "General", "NewEnumsOrderedByValue", CStr(Abs(CLng(mNewEnumsOrderedByValue)))
    mCurrentAction = ecaDefault
    ShowTree
End Sub

Private Sub mnuEnumsOrderAllByValue_Click()
    mEnums.MoveFirst
    Do Until mEnums.EOF
        mEnums.Edit
        mEnums!Ordered_By_Value = True
        mEnums.Update
        mEnums.Bookmark = mEnums.LastModified
        mEnums.MoveNext
    Loop
    mNewEnumsOrderedByValue = True
    SaveSettingBase "General", "NewEnumsOrderedByValue", CStr(Abs(CLng(mNewEnumsOrderedByValue)))
    mCurrentAction = ecaDefault
    ShowTree
End Sub

Private Sub mnuLinkToClass_Click()
    mLinkTypeSelected = "[o["
End Sub

Private Sub mnuLinkToControl_Click()
    mLinkTypeSelected = "[c["
End Sub

Private Sub mnuLinkToEnum_Click()
    mLinkTypeSelected = "[["
End Sub

Private Sub mnuLinkToEvent_Click()
    mLinkTypeSelected = "[e["
End Sub

Private Sub mnuLinkToMethod_Click()
    mLinkTypeSelected = "[m["
End Sub

Private Sub mnuLinkToProperty_Click()
    mLinkTypeSelected = "[p["
End Sub

Private Sub mnuListBlankFields_Click()
    Dim iStr As String
    Dim iRec As Recordset
    Dim txt As TextBox
    Dim iAT As String
    Dim iEnums As Recordset
    
    Set iEnums = mEnums.Clone
    iEnums.Index = "PrimaryKey"
    iStr = GetSettingBase("General", "BlankFieldsSelection", "")
    If iStr <> "" Then frmBlankFieldsOptions.FieldsString = iStr
    frmBlankFieldsOptions.Show vbModal
    If frmBlankFieldsOptions.OKPressed Then
        If InStr(frmBlankFieldsOptions.FieldsString, "1") = 0 Then
            MsgBox "Nothing selected.", vbExclamation
            Exit Sub
        End If
        SaveSettingBase "General", "BlankFieldsSelection", frmBlankFieldsOptions.FieldsString
        Set txt = frmMessage.txtMessage
        
        If frmBlankFieldsOptions.Field(0) Then ' controls/classes long descriptions
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls WHERE Long_Description = ''")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    txt.SelText = iRec!Name & " control long description" & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes WHERE Long_Description = ''")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    txt.SelText = iRec!Name & " class long description" & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmBlankFieldsOptions.Field(1) Then ' controls/classes short descriptions
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls WHERE Short_Description = ''")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    txt.SelText = iRec!Name & " control short description" & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes WHERE Short_Description = ''")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    txt.SelText = iRec!Name & " class short description" & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmBlankFieldsOptions.Field(2) Then ' members params info
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Properties WHERE (Params_Info = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(1, iRec!Property_ID, True)
                    txt.SelText = iRec!Name & " property parameters info" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Methods WHERE (Params_Info = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(2, iRec!Method_ID, True)
                    txt.SelText = iRec!Name & " method parameters info" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Events WHERE (Params_Info = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(3, iRec!Event_ID, True)
                    txt.SelText = iRec!Name & " event parameters info" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmBlankFieldsOptions.Field(3) Then ' members long description
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Properties WHERE (Long_Description = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(1, iRec!Property_ID, True)
                    txt.SelText = iRec!Name & " property long description" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Methods WHERE (Long_Description = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(2, iRec!Method_ID, True)
                    txt.SelText = iRec!Name & " method long description" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Events WHERE (Long_Description = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(3, iRec!Event_ID, True)
                    txt.SelText = iRec!Name & " event long description" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmBlankFieldsOptions.Field(4) Then ' members short description
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Properties WHERE (Short_Description = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(1, iRec!Property_ID, True)
                    txt.SelText = iRec!Name & " property short description" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Methods WHERE (Short_Description = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(2, iRec!Method_ID, True)
                    txt.SelText = iRec!Name & " method short description" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Events WHERE (Short_Description = '') AND (Auxiliary_Field = 1)")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iAT = GetAppliesTo(3, iRec!Event_ID, True)
                    txt.SelText = iRec!Name & " event short description" & IIf(iAT <> "", " (" & iAT & ")", "") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmBlankFieldsOptions.Field(5) Then ' enums description
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Enums WHERE (Description = '')")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    txt.SelText = iRec!Name & " enum description" & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmBlankFieldsOptions.Field(6) Then ' constants description
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Description = '')")
            If Not iRec.EOF Then
                iRec.MoveFirst
                Do Until iRec.EOF
                    iEnums.Seek "=", iRec!Enum_ID
                    
                    txt.SelText = iRec!Name & " constant description" & IIf(iEnums.NoMatch, "", " in " & iEnums!Name & " enum") & vbCrLf
                    iRec.MoveNext
                Loop
            End If
        End If
        If frmMessage.txtMessage.Text = "" Then
            Unload frmMessage
            MsgBox "No blank fields found.", vbInformation
        Else
            frmMessage.Caption = "List of blank fields"
            frmMessage.Show vbModal
        End If
        Set frmMessage = Nothing
    End If
    Set frmBlankFieldsOptions = Nothing
End Sub

Private Sub mnuLoadFromOrphanMember_Click()
    If CurrentType > 0 Then
        frmSelectOrpahnMember.LoadList mDatabase, mMemberType_p(CurrentType)
        frmSelectOrpahnMember.Show vbModal
        If frmSelectOrpahnMember.OKPressed Then
            If Trim2(txtLongDescription.Text) <> "" Then
                If MsgBox("This will replace the current text and can't be undone, continue?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
            End If
            txtLongDescription.Text = frmSelectOrpahnMember.Text
            UpdateData
        End If
        Set frmSelectOrpahnMember = Nothing
    End If
End Sub

Private Sub mnuLoadFromOrphanMember2_Click()
    mnuLoadFromOrphanMember_Click
End Sub

Private Sub mnuPreferences_Click()
    Set frmPreferences.AppFont = mAppFont
    frmPreferences.Show vbModal
    mFontPropertion = mAppFont.Size / GetDesiredFontSize
    PlaceDataControls
End Sub

Private Sub mnuRefTypeList_Click(Index As Integer)
    mRefTableIndex = Index
End Sub

Private Sub mnuReport_Click()
    UpdateData
End Sub

Private Sub mnuReportHTML_Click()
    Dim iSc As SmartConcat
    Dim iDlg As cDlg
    Dim iOutFolderPath As String
    Dim iOutFolderPath2 As String
    Dim iDo As Boolean
    Dim c As Long
    
    If Not ShowfrmReportSelection("HTML") Then Exit Sub
    
    ReDim mLinkErrors(0)
    
    Set iDlg = New cDlg
    
    iDlg.DialogTitle = "Select the output folder"
    iDlg.FolderName = GetSettingBase("General", "OutputPathHTML", "")
    iDlg.ShowFolder
    If iDlg.Canceled Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    iOutFolderPath = iDlg.FolderName
    If Right$(iOutFolderPath, 1) <> "\" Then iOutFolderPath = iOutFolderPath & "\"
    SaveSettingBase "General", "OutputPathHTML", iOutFolderPath
    Set iDlg = Nothing
    
    ReDim mPages(0)
    ReDim mPages_FileNames(0)
    
    If mHTML_Mode = cdHTMLPerMethod Then
        iOutFolderPath2 = iOutFolderPath & "html_files\"
        DoHTMLPerMethod
    ElseIf mHTML_Mode = cdHTMLPerObject Then
        iOutFolderPath2 = iOutFolderPath & "html_files\"
        DoHTMLPerObject
    ElseIf mHTML_Mode = cdHTMLOnePage Then
        iOutFolderPath2 = iOutFolderPath
        DoHTMLOnePage
    End If
    
    If mExternalCSS Then
        If mReplaceCSSFile Then
            iDo = True
        Else
            If Not FileExists(iOutFolderPath2 & "styles.css") Then
                iDo = True
            End If
        End If
        If iDo Then
            AddPage "styles.css", mHTML_StyleSheet
        End If
    End If
        
    If Not FolderExists(iOutFolderPath) Then
        MkDir iOutFolderPath
        If Not FolderExists(iOutFolderPath) Then
            MsgBox "Could not create folder.", vbCritical
            Exit Sub
        End If
    End If
    
    If Not FolderExists(iOutFolderPath2) Then
        MkDir iOutFolderPath2
        If Not FolderExists(iOutFolderPath2) Then
            MsgBox "Could not create folder.", vbCritical
            Exit Sub
        End If
    End If
    
    For c = 1 To UBound(mPages)
        If FileExists(iOutFolderPath2 & mPages_FileNames(c)) Then
            On Error Resume Next
            Kill iOutFolderPath2 & mPages_FileNames(c)
            On Error GoTo 0
        End If
        SaveBinaryFile iOutFolderPath2 & mPages_FileNames(c), ConvertToUTF8(mPages(c))
    Next c
    
    If (mHTML_Mode = cdHTMLPerMethod) Or (mHTML_Mode = cdHTMLPerObject) Then
        If FileExists(iOutFolderPath & mComponentName & "_reference.html") Then
            On Error Resume Next
            Kill iOutFolderPath & mComponentName & "_reference.html"
            On Error GoTo 0
        End If
        SaveBinaryFile iOutFolderPath & mComponentName & "_reference.html", ConvertToUTF8(Replace(mPages(1), "href=""", "href=""html_files/"))
    End If
    
    Erase mPages
    Erase mPages_FileNames
    Unload frmReportSelection
    Set frmReportSelection = Nothing
    
    Screen.MousePointer = vbDefault
    
    If UBound(mLinkErrors) > 0 Then
        MsgBox "Found " & UBound(mLinkErrors) & " link errors."
    End If
End Sub

Private Sub DoHTMLPerMethod()
    Dim iSc As SmartConcat
    Dim iRec As Recordset
    Dim c As Long
    Dim iDesc As String
    Dim iSc2 As SmartConcat
    Dim iSc3 As SmartConcat
    Dim iLinkCounter As Long
    Dim t As Long
    Dim m As Long
    Dim iTTypes As Recordset
    Dim iTmpSections() As String
    Dim iFileName As String
    Dim iAux_MemberList_Names() As String
    Dim iAux_MemberList_FileNames() As String
    Dim iMembers_Names(4) As Variant
    Dim iMembers_FileNames(4) As Variant
    Dim iName As String
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTEnums As Recordset
    Dim iStrs() As String
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    Set iSc = New SmartConcat
    ReDim mNonUniqueMemberNamePages(0)
    
    ' HTML
    ' Head
    iSc.AddString GetHTMLHeadSection(mComponentName & " Reference", mComponentName & " component's reference - Documentation.")
    iSc.AddString "<body>" & vbCrLf
    
    ' Body
    If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace$(Replace$(mHTML_PageHeaderMP, "> [CURRENT_ITEM]", ""), "[CURRENT_ITEM]", "") & vbCrLf
    iSc.AddString "<h1>" & mComponentName & IIf((frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoVersion.Value = 1) And (mComponentVersion <> ""), " " & mComponentVersion, "") & " Reference</h1>" & vbCrLf
    If (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoReleaseDate.Value = 1) And (mComponentReleaseDate <> 0) Then
        iSc.AddString "<p>Release date: " & FormatDateTime(mComponentReleaseDate, vbShortDate) & "</p><br>" & vbCrLf
    End If
    
    iDesc = GetGeneralInfo("Introduction")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        iSc.AddString "<p>" & TxtToHTML(iDesc, "index") & "</p><br>" & vbCrLf
    End If

    If (iTControls.RecordCount > 0) Or (iTClasses.RecordCount > 0) Then
        iSc.AddString "<a href=""objects.html"">Objects</a><br>" & vbCrLf
    End If
    If mProperties.RecordCount > 0 Then
        iSc.AddString "<a href=""properties.html"">Properties</a><br>" & vbCrLf
    End If
    If mMethods.RecordCount > 0 Then
        iSc.AddString "<a href=""methods.html"">Methods</a><br>" & vbCrLf
    End If
    If mEvents.RecordCount > 0 Then
        iSc.AddString "<a href=""events.html"">Events</a><br>" & vbCrLf
    End If
    If mConstants.RecordCount > 0 Then
        iSc.AddString "<a href=""constants.html"">Constants</a><br>" & vbCrLf
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoEndNotes.Value = 1) Then
        iSc.AddString "<br>"
        iSc.AddString "<p>" & TxtToHTML(iDesc, "index") & "</p><br>" & vbCrLf
    End If
    
    If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
    iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
    AddPage "index.html", iSc.GenerateCurrentString
    Set iSc = Nothing

    ' Objects
    Set iSc = New SmartConcat
    ' Head
    iSc.AddString GetHTMLHeadSection(mComponentName & " Objects", mComponentName & " Objects Documentation.")
    iSc.AddString "<body>" & vbCrLf
    
    ' Body
    If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", "Objects") & vbCrLf
    iSc.AddString "<h1>" & mComponentName & " Objects</h1>" & vbCrLf
    
    If iTControls.RecordCount > 0 Then
        iTControls.MoveFirst
        Do Until iTControls.EOF
            If frmReportSelection.IsItemSelected(1, iTControls!Name) Then
                iSc.AddString "<a href=""" & LCase(iTControls!Name) & "_control.html"">" & iTControls!Name & "</a> control<br>" & vbCrLf
            End If
            iTControls.MoveNext
        Loop
    End If

    If iTClasses.RecordCount > 0 Then
        If iTControls.RecordCount > 0 Then
            iSc.AddString "<br>"
        End If
        iTClasses.MoveFirst
        Do Until iTClasses.EOF
            If frmReportSelection.IsItemSelected(2, iTClasses!Name) Then
                iSc.AddString "<a href=""" & LCase(iTClasses!Name) & "_object.html"">" & iTClasses!Name & "</a> object<br>" & vbCrLf
            End If
            iTClasses.MoveNext
        Loop
    End If

    If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
    iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
    AddPage "objects.html", iSc.GenerateCurrentString
    Set iSc = Nothing
    
    For c = 1 To 3
        ReDim iAux_MemberList_Names(0)
        ReDim iAux_MemberList_FileNames(0)
        iMembers_Names(c) = iAux_MemberList_Names
        iMembers_FileNames(c) = iAux_MemberList_FileNames
    Next c
    
    For t = 1 To 2 ' controls and classes
        If frmReportSelection.chkType(t).Value Then
            If t = 1 Then
                Set iTTypes = iTControls
            ElseIf t = 2 Then
                Set iTTypes = iTClasses
            End If
        
            If iTTypes.RecordCount > 0 Then 'if there are controls (or classes)
                iTTypes.MoveFirst
                Do Until iTTypes.EOF ' for each control or class
                    If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                        Set iSc = New SmartConcat
                        
                        ' Head
                        iSc.AddString GetHTMLHeadSection(iTTypes!Name & " " & LCase(mObjectType_s2(t)), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & " Documentation.")
                        iSc.AddString "<body>" & vbCrLf
                        ' Body
                        If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", iTTypes!Name) & vbCrLf
                        iSc.AddString "<h1>" & iTTypes!Name & " " & LCase(mObjectType_s2(t)) & "</h1>" & vbCrLf
                        ' Description
                        If iTTypes!Long_Description <> "" Then
                            iDesc = iTTypes!Long_Description
                        Else
                            iDesc = iTTypes!Short_Description
                        End If
                        If iDesc <> "" Then
                            iSc.AddString "<p>" & TxtToHTML(iDesc, iTTypes!Name & " " & LCase(mObjectType_s2(t))) & "</p><br>" & vbCrLf
                        End If
                        
                        iLinkCounter = 0
                        ReDim iTmpSections(0)
                        For m = 1 To 3 ' properties, methods or events
                            iAux_MemberList_Names = iMembers_Names(m)
                            iAux_MemberList_FileNames = iMembers_FileNames(m)
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then ' if there are properties (methods or events)
                                iLinkCounter = iLinkCounter + 1
                                iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & mMemberType_p(m) & "</a></li>"
                                Set iSc2 = New SmartConcat
                                iSc2.AddString "<br>"
                                iSc2.AddString "<h2><a name=""ilink" & CStr(iLinkCounter) & """ /><p class=""section"">" & mMemberType_p(m) & "</p></a></h2>" & vbCrLf
                                iRec.MoveFirst
                                Do Until iRec.EOF ' for each property, method or event
                                    iFileName = GetPageFileName(m, iRec)
                                    iSc2.AddString "<a href=""" & iFileName & """>" & iRec!Name & "</a><br>" & vbCrLf
                                    If Not IsInList(mPages_FileNames, iFileName) Then ' if the property, method or event has not be done already
                                        
                                        Set iSc3 = New SmartConcat
                                        ' Head
                                        iSc3.AddString GetHTMLHeadSection(iRec!Name & " " & mMemberType_s(m), iRec!Name & " " & mMemberType_s(m) & ", parameters and information.")
                                        iSc3.AddString "<body>" & vbCrLf
                                        ' Body
                                        If mHTML_PageHeaderMP <> "" Then iSc3.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", iRec!Name) & vbCrLf
                                        iSc3.AddString "<h1>" & iRec!Name & " " & mMemberType_s(m) & "</h1>" & vbCrLf
                                        iSc3.AddString "<p>Applies to: " & TxtToHTML(GetAppliesTo(m, iRec.Fields(mMemberType_p(m) & "." & mMemberType_s(m) & "_ID")), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p>" & vbCrLf
                                        ' Description
                                        If iRec!Params_Info <> "" Then
                                            iSc3.AddString "<p>" & TxtToHTML(HTMLFormatParameters(iRec!Params_Info), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                        End If
                                        If iRec!Long_Description <> "" Then
                                            iSc3.AddString "<h2>Description:</h2>" & vbCrLf
                                            iSc3.AddString "<p>" & TxtToHTML(iRec!Long_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                        ElseIf iRec!Short_Description <> "" Then
                                            iSc3.AddString "<h2>Description:</h2>" & vbCrLf
                                            iSc3.AddString "<p>" & TxtToHTML(iRec!Short_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                        End If
                                        If mHTML_PageFooter <> "" Then iSc3.AddString mHTML_PageFooter & vbCrLf
                                        iSc3.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
                                        
                                        AddPage iFileName, iSc3.GenerateCurrentString ' add the page for the property, method or event
                                    
                                        ' this is used to make the global list of properties, methods and events
                                        AddToList iAux_MemberList_Names, IIf(IsMemberNameUnique(m, iRec!Name), iRec!Name, iRec!Name & " (" & TxtToHTML(GetAppliesTo(m, iRec.Fields(mMemberType_p(m) & "." & mMemberType_s(m) & "_ID")), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & ")")
                                        AddToList iAux_MemberList_FileNames, iFileName
                                    End If
                                    iRec.MoveNext
                                Loop
                                ' add to a temporary list one whole section of list of properties (with links to each individual file), methods or events
                                AddToList iTmpSections, iSc2.GenerateCurrentString
                            End If
                            iMembers_Names(m) = iAux_MemberList_Names
                            iMembers_FileNames(m) = iAux_MemberList_FileNames
                        Next m
                        
                        ' add each section of properties, methos and events
                        For c = 1 To UBound(iTmpSections)
                            iSc.AddString iTmpSections(c)
                        Next c
                        
                        If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
                        iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
                        
                        AddPage LCase(iTTypes!Name) & "_" & LCase(mObjectType_s2(t)) & ".html", iSc.GenerateCurrentString
                    End If
                    iTTypes.MoveNext
                Loop
            End If
        End If
    Next t
    
    ' Properties, Methods and Events
    For m = 1 To 3
        iAux_MemberList_Names = iMembers_Names(m)
        iAux_MemberList_FileNames = iMembers_FileNames(m)
        OrderVector iAux_MemberList_Names, iAux_MemberList_FileNames
        If UBound(iAux_MemberList_Names) > 0 Then
            Set iSc = New SmartConcat
            
            ' Head
            iSc.AddString GetHTMLHeadSection(mComponentName & " " & mMemberType_p(m), mComponentName & " " & mMemberType_p(m) & " List.")
            iSc.AddString "<body>" & vbCrLf
            ' Body
            If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", mMemberType_p(m)) & vbCrLf
            iSc.AddString "<h1>" & mComponentName & " " & mMemberType_p(m) & "</h1>" & vbCrLf
            
            For c = 1 To UBound(iAux_MemberList_Names)
                iFileName = iAux_MemberList_FileNames(c)
                iName = iAux_MemberList_Names(c)
                iSc.AddString "<a href=""" & iFileName & """>" & iName & "</a><br>" & vbCrLf
            Next c
            
            If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
            iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
            
            AddPage LCase(mMemberType_p(m)) & ".html", iSc.GenerateCurrentString
        End If
    Next m
    
    ' Constants
    If frmReportSelection.chkType(3).Value Then
        Set iTEnums = mEnums.Clone
        iTEnums.Index = "Name"
        If iTEnums.RecordCount > 0 Then
            Set iSc = New SmartConcat
            
            ' Head
            iSc.AddString GetHTMLHeadSection(mComponentName & " Enumerations", mComponentName & " Enumeration List.")
            iSc.AddString "<body>" & vbCrLf
            ' Body
            If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", "Constants") & vbCrLf
            iSc.AddString "<h1>" & mComponentName & " Enumerations</h1>" & vbCrLf
            
            iTEnums.MoveFirst
            Do Until iTEnums.EOF
                If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
                    If iRec.RecordCount > 0 Then
                        iSc.AddString "<a href=""" & LCase(iTEnums!Name) & "_enumeration.html"">" & iTEnums!Name & "</a><br>" & vbCrLf
                        
                        Set iSc2 = New SmartConcat
                        
                        ' Head
                        iSc2.AddString GetHTMLHeadSection(iTEnums!Name & " Enumeration", iTEnums!Name & " Enumeration Documentation.")
                        iSc2.AddString "<style>.table2, .table2 th, .table2 td {border: 2px solid #e0e0e0;border-collapse: collapse;} .table2 th, .table2 td {padding: 8px;text-align: left;}</style>"
                        iSc2.AddString "<body>" & vbCrLf
                        ' Body
                        If mHTML_PageHeaderMP <> "" Then iSc2.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", iTEnums!Name) & vbCrLf
                        iSc2.AddString "<h1>" & iTEnums!Name & " Enumeration</h1>" & vbCrLf
                        iSc2.AddString "<p>" & TxtToHTML(iTEnums!Description, iTEnums!Name & " Enumeration") & "</p><br>" & vbCrLf
                        
                        iSc2.AddString "<table class=""table table2"">" & vbCrLf
                        iRec.MoveFirst
                        Do Until iRec.EOF
                            iSc2.AddString "  <tr>" & vbCrLf
                            iSc2.AddString "    <td>"
                            iSc2.AddString "<b>" & iRec!Name & "</b> = "
                            iSc2.AddString "    </td>" & vbCrLf
                            iSc2.AddString "    <td>"
                            iSc2.AddString iRec!Value
                            iSc2.AddString "    </td>" & vbCrLf
                            If (iRec!Description <> "") Then
                                iSc2.AddString "    <td>"
                                iSc2.AddString "<a style=""color:#0040A0""> ' " & TxtToHTML(Replace(iRec!Description, vbCrLf, vbTab), iRec!Name & " constant of " & iTEnums!Name & " Enumeration") & "</a>"
                                iSc2.AddString "    </td>" & vbCrLf
                            End If
                            iRec.MoveNext
                            iSc2.AddString "  </tr>" & vbCrLf
                        Loop
                        iSc2.AddString "</table>" & vbCrLf
                        
                        If mHTML_PageFooter <> "" Then iSc2.AddString mHTML_PageFooter & vbCrLf
                        iSc2.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
                        
                        AddPage LCase(iTEnums!Name) & "_enumeration.html", iSc2.GenerateCurrentString
                    End If
                End If
                iTEnums.MoveNext
            Loop
            
            If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
            iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
            
            AddPage "constants.html", iSc.GenerateCurrentString
        End If
    End If
    
    ' add an "index" page for each member that has different definitions with the same name (used in different controls/classes)
    For c = 1 To UBound(mNonUniqueMemberNamePages)
        iStrs = Split(mNonUniqueMemberNamePages(c), "|")
        m = Val(iStrs(0))
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(m) & " WHERE (Name = '" & iStrs(1) & "') AND (Auxiliary_Field = 1)")
        iRec.MoveFirst
        Set iSc = New SmartConcat
        
        ' Head
        iSc.AddString GetHTMLHeadSection(iRec!Name & " " & mMemberType_s(m), iStrs(1) & " " & mMemberType_s(m) & " (list)")
        iSc.AddString "<body>" & vbCrLf
        ' Body
        If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", iStrs(1)) & vbCrLf
        iSc3.AddString "<h1>" & iStrs(1) & " " & mMemberType_s(m) & " list</h1>" & vbCrLf
        
        iRec.MoveFirst
        Do Until iRec.EOF
            iFileName = GetPageFileName(m, iRec)
            iSc.AddString "<a href=""" & iFileName & """>" & iRec!Name & "</a> (" & TxtToRTF(GetAppliesTo(m, iRec.Fields(mMemberType_s(m) & "_ID"), True)) & ")" & "<br>" & vbCrLf
            rtbAux.Text = ""
            iRec.MoveNext
        Loop
    
        If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
        iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
        
        AddPage LCase$(iStrs(1)) & "_" & LCase$(mMemberType_s(m)) & ".html", iSc.GenerateCurrentString
    Next
End Sub

Private Sub DoHTMLPerObject()
    Dim iSc As SmartConcat
    Dim iRec As Recordset
    Dim c As Long
    Dim iDesc As String
    Dim iSc2 As SmartConcat
    Dim iSc3 As SmartConcat
    Dim iLinkCounter As Long
    Dim t As Long
    Dim m As Long
    Dim iTTypes As Recordset
    Dim iTmpSections_PME_List() As String
    Dim iTmpSections_PME_Details() As String
    Dim iTmpSections_PME_Details_Joined() As String
    Dim iTmpSections_Constants() As String
    Dim iName As String
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTEnums As Recordset
    Dim iLinks_ToReplace() As String
    Dim iLinks_Replacement() As String
    Dim iLinks_ToReplace_General() As String
    Dim iLinks_Replacement_General() As String
    Dim iStr As String
    Dim p As Long
    Dim iPutPMELinks As Boolean
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    Set iSc = New SmartConcat
    
    ' HTML
    ' Head
    iSc.AddString GetHTMLHeadSection(mComponentName & " Reference", mComponentName & " component's reference - Documentation.")
    iSc.AddString "<body>" & vbCrLf
    
    ' Body
    If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace$(Replace$(mHTML_PageHeaderMP, "> [CURRENT_ITEM]", ""), "[CURRENT_ITEM]", "") & vbCrLf
    iSc.AddString "<h1>" & mComponentName & IIf((frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoVersion.Value = 1) And (mComponentVersion <> ""), " " & mComponentVersion, "") & " Reference</h1>" & vbCrLf
    If (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoReleaseDate.Value = 1) And (mComponentReleaseDate <> 0) Then
        iSc.AddString "<p>Release date: " & FormatDateTime(mComponentReleaseDate, vbShortDate) & "</p><br>" & vbCrLf
    End If

    iDesc = GetGeneralInfo("Introduction")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        iSc.AddString "<p>" & TxtToHTML(iDesc, "index") & "</p><br>" & vbCrLf
    End If
    
    If iTControls.RecordCount > 0 Then
        iTControls.MoveFirst
        Do Until iTControls.EOF
            If frmReportSelection.IsItemSelected(1, iTControls!Name) Then
                iSc.AddString "<a href=""" & LCase(iTControls!Name) & "_control.html"">" & iTControls!Name & "</a> control<br>" & vbCrLf
            End If
            iTControls.MoveNext
        Loop
    End If
    If iTClasses.RecordCount > 0 Then
        If iTControls.RecordCount > 0 Then
            iSc.AddString "<br>"
        End If
        iTClasses.MoveFirst
        Do Until iTClasses.EOF
            If frmReportSelection.IsItemSelected(2, iTClasses!Name) Then
                iSc.AddString "<a href=""" & LCase(iTClasses!Name) & "_object.html"">" & iTClasses!Name & "</a> object<br>" & vbCrLf
            End If
            iTClasses.MoveNext
        Loop
    End If
    If mConstants.RecordCount > 0 Then
        iSc.AddString "<a href=""constants.html"">Constants</a><br>" & vbCrLf
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoEndNotes.Value = 1) Then
        iSc.AddString "<br>"
        iSc.AddString "<p>" & TxtToHTML(iDesc, "index") & "</p><br>" & vbCrLf
    End If
    
    If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
    iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
    AddPage "index.html", iSc.GenerateCurrentString
    Set iSc = Nothing

    ' Objects
    ReDim iLinks_ToReplace_General(0)
    ReDim iLinks_Replacement_General(0)
    For t = 1 To 2 ' controls and classes
        If frmReportSelection.chkType(t).Value Then
            If t = 1 Then
                Set iTTypes = iTControls
            ElseIf t = 2 Then
                Set iTTypes = iTClasses
            End If
        
            If iTTypes.RecordCount > 0 Then 'if there are controls (or classes)
                iTTypes.MoveFirst
                Do Until iTTypes.EOF ' for each control or class
                    If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                        Set iSc = New SmartConcat
                        
                        ' Head
                        iSc.AddString GetHTMLHeadSection(iTTypes!Name & " " & LCase(mObjectType_s2(t)), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & " Documentation.")
                        iSc.AddString "<body>" & vbCrLf
                        ' Body
                        If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", iTTypes!Name) & vbCrLf
                        iSc.AddString "<h1>" & iTTypes!Name & " " & LCase(mObjectType_s2(t)) & "</h1>" & vbCrLf
                        ' Description
                        If iTTypes!Long_Description <> "" Then
                            iDesc = iTTypes!Long_Description
                        Else
                            iDesc = iTTypes!Short_Description
                        End If
                        If iDesc <> "" Then
                            iSc.AddString "<p>" & TxtToHTML(iDesc, iTTypes!Name & " " & LCase(mObjectType_s2(t))) & "</p><br>" & vbCrLf
                        End If
                        
                        iLinkCounter = 0
                        ReDim iLinks_ToReplace(0)
                        ReDim iLinks_Replacement(0)
                        ReDim iTmpSections_PME_List(0)
                        ReDim iTmpSections_PME_Details_Joined(0)
                        c = 0
                        For m = 1 To 3
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then c = c + 1
                            If c > 1 Then Exit For
                        Next
                        iPutPMELinks = (c > 1)
                        For m = 1 To 3 ' properties, methods or events
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then ' if there are properties (methods or events)
                                iLinkCounter = iLinkCounter + 1
                                Set iSc2 = New SmartConcat
                                If iPutPMELinks Then
                                    iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & mMemberType_p(m) & "</a></li>"
                                    iSc2.AddString "<br>"
                                    iSc2.AddString "<a name=""ilink" & CStr(iLinkCounter) & """ /><h2><p class=""section"">" & mMemberType_p(m) & "</p></h2></a>" & vbCrLf
                                Else
                                    iSc2.AddString "<br><br><br><h1>" & mMemberType_p(m) & "</h1><br>"
                                End If
                                
                                Set iSc3 = New SmartConcat
                                If iPutPMELinks Then
                                    iSc3.AddString "<br><br><br><h1>" & mMemberType_p(m) & "</h1><br>"
                                End If
                                
                                ReDim iTmpSections_PME_Details(0)
                                iRec.MoveFirst
                                Do Until iRec.EOF ' for each property, method or event
                                    iLinkCounter = iLinkCounter + 1
                                    iSc2.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & iRec!Name & "</a><br>" & vbCrLf
                                    AddToList iLinks_ToReplace, LCase(iRec!Name) & "_" & LCase(mMemberType_s(m)) & ".html"
                                    AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
                                    
                                    iSc3.AddString "<a name=""ilink" & CStr(iLinkCounter) & """ /><h2><p class""section"">" & iRec!Name & " " & mMemberType_s(m) & "</p></h2></a>" & vbCrLf
                                    ' Description
                                    If iRec!Params_Info <> "" Then
                                        iSc3.AddString "<p>" & TxtToHTML(HTMLFormatParameters(iRec!Params_Info), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p>" & vbCrLf
                                    End If
                                    If iRec!Long_Description <> "" Then
                                        iSc3.AddString "<h3>Description:</h3>" & vbCrLf
                                        iSc3.AddString "<p>" & TxtToHTML(iRec!Long_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                    ElseIf iRec!Short_Description <> "" Then
                                        iSc3.AddString "<h3>Description:</h3>" & vbCrLf
                                        iSc3.AddString "<p>" & TxtToHTML(iRec!Short_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                    End If
                                    
                                    iRec.MoveNext
                                Loop
                                If Not iPutPMELinks Then
                                    iSc2.AddString "<br><br><br>"
                                End If
                                AddToList iTmpSections_PME_Details, iSc3.GenerateCurrentString
                                ' add to a temporary list one whole section of list of properties (with links to each individual file), methods or events
                                AddToList iTmpSections_PME_List, iSc2.GenerateCurrentString
                                AddToList iTmpSections_PME_Details_Joined, Join(iTmpSections_PME_Details, vbCrLf)
                            End If
                        Next m
                        
                        ' add each section of properties, methos and events
                        For c = 1 To UBound(iTmpSections_PME_List)
                            iSc.AddString iTmpSections_PME_List(c)
                        Next c
                        ' and after that the details
                        For c = 1 To UBound(iTmpSections_PME_Details_Joined)
                            iSc.AddString iTmpSections_PME_Details_Joined(c)
                        Next c
                        
                        If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
                        iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
                        
                        iStr = iSc.GenerateCurrentString
                        For c = 1 To UBound(iLinks_ToReplace)
                            AddToList iLinks_ToReplace_General, iLinks_ToReplace(c)
                            AddToList iLinks_Replacement_General, LCase(iTTypes!Name) & "_" & LCase(mObjectType_s2(t)) & ".html" & iLinks_Replacement(c)
                            If InStr(iStr, "<a href=""" & iLinks_ToReplace(c) & """>") Then
                                iStr = Replace$(iStr, "<a href=""" & iLinks_ToReplace(c) & """>", "<a href=""" & iLinks_Replacement(c) & """>")
                            End If
                        Next
                        AddPage LCase(iTTypes!Name) & "_" & LCase(mObjectType_s2(t)) & ".html", iStr
                    End If
                    iTTypes.MoveNext
                Loop
            End If
        End If
    Next t
    
    If frmReportSelection.chkType(3).Value Then
        ' Constants
        ReDim iTmpSections_Constants(0)
        Set iTEnums = mEnums.Clone
        iTEnums.Index = "Name"
        If iTEnums.RecordCount > 0 Then
            Set iSc = New SmartConcat
            
            ' Header
            iSc.AddString GetHTMLHeadSection(mComponentName & " Enumerations", mComponentName & " Enumeration List.")
            iSc.AddString "<style>.table2, .table2 th, .table2 td {border: 2px solid #e0e0e0;border-collapse: collapse;} .table2 th, .table2 td {padding: 8px;text-align: left;}</style>"
            iSc.AddString "<body>" & vbCrLf
            ' Body
            If mHTML_PageHeaderMP <> "" Then iSc.AddString Replace(mHTML_PageHeaderMP, "[CURRENT_ITEM]", "Constants") & vbCrLf
            iSc.AddString "<h1>" & mComponentName & " Enumerations</h1><br>" & vbCrLf
            
            iLinkCounter = 0
            ReDim iLinks_ToReplace(0)
            ReDim iLinks_Replacement(0)
            iTEnums.MoveFirst
            Do Until iTEnums.EOF
                If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
                    If iRec.RecordCount > 0 Then
                        iLinkCounter = iLinkCounter + 1
                        iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & iTEnums!Name & "</a></li>"
                        
                        AddToList iLinks_ToReplace, LCase(iTEnums!Name) & "_enumeration.html"
                        AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
                        
                        Set iSc2 = New SmartConcat
                        iSc2.AddString "<br>"
                        iSc2.AddString "<a name=""ilink" & CStr(iLinkCounter) & """ /><h2><p class""section"">" & iTEnums!Name & "</h2>" & "</p></h2></a>" & vbCrLf
                        iSc2.AddString "<p>" & TxtToHTML(iTEnums!Description, iTEnums!Name & " Enumeration") & "</p><br>" & vbCrLf
                        
                        iSc2.AddString "<table class=""table table2"">" & vbCrLf
                        iRec.MoveFirst
                        Do Until iRec.EOF
                            iSc2.AddString "  <tr>" & vbCrLf
                            iSc2.AddString "    <td>"
                            iSc2.AddString "<b>" & iRec!Name & "</b> = "
                            iSc2.AddString "    </td>" & vbCrLf
                            iSc2.AddString "    <td>"
                            iSc2.AddString iRec!Value
                            iSc2.AddString "    </td>" & vbCrLf
                            If (iRec!Description <> "") Then
                                iSc2.AddString "    <td>"
                                iSc2.AddString "<a style=""color:#0040A0""> ' " & TxtToHTML(Replace(iRec!Description, vbCrLf, vbTab), iRec!Name & " constant of " & iTEnums!Name & " Enumeration") & "</a>"
                                iSc2.AddString "    </td>" & vbCrLf
                            End If
                            iRec.MoveNext
                            iSc2.AddString "  </tr>" & vbCrLf
                        Loop
                        iSc2.AddString "</table>" & vbCrLf
                        
                        AddToList iTmpSections_Constants, iSc2.GenerateCurrentString
                    End If
                End If
                iTEnums.MoveNext
            Loop
            
            ' add each section of enums
            For c = 1 To UBound(iTmpSections_Constants)
                iSc.AddString iTmpSections_Constants(c)
            Next c
            If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
            iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
            
            iStr = iSc.GenerateCurrentString
            For c = 1 To UBound(iLinks_ToReplace)
                AddToList iLinks_ToReplace_General, iLinks_ToReplace(c)
                AddToList iLinks_Replacement_General, "constants.html" & iLinks_Replacement(c)
                If InStr(iStr, "<a href=""" & iLinks_ToReplace(c) & """>") Then
                    iStr = Replace$(iStr, "<a href=""" & iLinks_ToReplace(c) & """>", "<a href=""" & iLinks_Replacement(c) & """>")
                End If
            Next
            AddPage "constants.html", iStr
        End If
    End If
    
    For p = 1 To UBound(mPages)
        For c = 1 To UBound(iLinks_ToReplace_General)
            If InStr(mPages(p), "<a href=""" & iLinks_ToReplace_General(c) & """>") Then
                mPages(p) = Replace$(mPages(p), "<a href=""" & iLinks_ToReplace_General(c) & """>", "<a href=""" & iLinks_Replacement_General(c) & """>")
            End If
        Next
    Next
End Sub

Private Sub DoHTMLOnePage()
    Dim iSc As SmartConcat
    Dim iSc2 As SmartConcat
    Dim iSc3 As SmartConcat
    Dim iSc4 As SmartConcat
    Dim iRec As Recordset
    Dim c As Long
    Dim iDesc As String
    Dim iLinkCounter As Long
    Dim t As Long
    Dim m As Long
    Dim iTTypes As Recordset
    Dim iTmpSections_PME_List() As String
    Dim iTmpSections_PME_Details() As String
    Dim iTmpSections_PME_Details_Joined() As String
    Dim iTmpSections_Constants() As String
    Dim iName As String
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTEnums As Recordset
    Dim iLinks_ToReplace() As String
    Dim iLinks_Replacement() As String
    Dim iStr As String
    Dim p As Long
    Dim iLinks As New Collection
    Dim iPutPMELinks As Boolean
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    Set iSc = New SmartConcat
    iLinkCounter = 0
    ReDim iLinks_ToReplace(0)
    ReDim iLinks_Replacement(0)
    
    ' HTML
    ' Head
    iSc.AddString GetHTMLHeadSection(mComponentName & " Reference", mComponentName & " component's reference - Documentation.", "<style>.table2, .table2 th, .table2 td {border: 2px solid #e0e0e0;border-collapse: collapse;} .table2 th, .table2 td {padding: 8px;text-align: left;}</style>")
    iSc.AddString "<body>" & vbCrLf
    
    ' Body
    If mHTML_PageHeaderOP <> "" Then iSc.AddString mHTML_PageHeaderOP & vbCrLf
    iSc.AddString "<h1>" & mComponentName & IIf((frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoVersion.Value = 1) And (mComponentVersion <> ""), " " & mComponentVersion, "") & " Reference</h1>" & vbCrLf
    If (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoReleaseDate.Value = 1) And (mComponentReleaseDate <> 0) Then
        iSc.AddString "<p>Release date: " & FormatDateTime(mComponentReleaseDate, vbShortDate) & "</p><br>" & vbCrLf
    End If

    iDesc = GetGeneralInfo("Introduction")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        iSc.AddString "<p>" & TxtToHTML(iDesc, "index") & "</p><br>" & vbCrLf
    End If
    
    If iTControls.RecordCount > 0 Then
        iTControls.MoveFirst
        Do Until iTControls.EOF
            If frmReportSelection.IsItemSelected(1, iTControls!Name) Then
                iLinkCounter = iLinkCounter + 1
                iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & iTControls!Name & "</a> control<br></li>" & vbCrLf
                AddToList iLinks_ToReplace, LCase(iTControls!Name) & "_control.html"
                AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
                iLinks.Add "ilink" & CStr(iLinkCounter), LCase(iTControls!Name) & "_control"
            End If
            iTControls.MoveNext
        Loop
    End If
    If iTClasses.RecordCount > 0 Then
        If iTControls.RecordCount > 0 Then
            iSc.AddString "<br>"
        End If
        iTClasses.MoveFirst
        Do Until iTClasses.EOF
            If frmReportSelection.IsItemSelected(2, iTClasses!Name) Then
                iLinkCounter = iLinkCounter + 1
                iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & iTClasses!Name & "</a> object<br></li>" & vbCrLf
                AddToList iLinks_ToReplace, LCase(iTClasses!Name) & "_object.html"
                AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
                iLinks.Add "ilink" & CStr(iLinkCounter), LCase(iTClasses!Name) & "_object"
            End If
            iTClasses.MoveNext
        Loop
    End If
    If mConstants.RecordCount > 0 Then
        iLinkCounter = iLinkCounter + 1
        iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>Constants<br></li>" & vbCrLf
        AddToList iLinks_ToReplace, "constants.html"
        AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
        iLinks.Add "ilink" & CStr(iLinkCounter), "constants"
    End If
    
    ' Objects
    For t = 1 To 2 ' controls and classes
        If frmReportSelection.chkType(t).Value Then
            If t = 1 Then
                Set iTTypes = iTControls
            ElseIf t = 2 Then
                Set iTTypes = iTClasses
            End If
            
            If iTTypes.RecordCount > 0 Then 'if there are controls (or classes)
                iTTypes.MoveFirst
                Do Until iTTypes.EOF ' for each control or class
                    If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                        Set iSc2 = New SmartConcat
                        
                        iSc2.AddString "<a name=""" & iLinks(LCase(iTTypes!Name) & "_" & LCase(mObjectType_s2(t))) & """><h1><p class=""section"">" & iTTypes!Name & " " & LCase(mObjectType_s2(t)) & "</p></h1></a>" & vbCrLf
                        
                        ' Description
                        If iTTypes!Long_Description <> "" Then
                            iDesc = iTTypes!Long_Description
                        Else
                            iDesc = iTTypes!Short_Description
                        End If
                        If iDesc <> "" Then
                            iSc2.AddString "<p>" & TxtToHTML(iDesc, iTTypes!Name & " " & LCase(mObjectType_s2(t))) & "</p><br>" & vbCrLf
                        End If
                        
                        ReDim iTmpSections_PME_List(0)
                        ReDim iTmpSections_PME_Details_Joined(0)
                        c = 0
                        For m = 1 To 3
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then c = c + 1
                            If c > 1 Then Exit For
                        Next
                        iPutPMELinks = (c > 1)
                        For m = 1 To 3 ' properties, methods or events
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then ' if there are properties (methods or events)
                                iLinkCounter = iLinkCounter + 1
                                Set iSc3 = New SmartConcat
                                If iPutPMELinks Then
                                    iSc2.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & mMemberType_p(m) & "</a></li>"
                                    iSc3.AddString "<br>"
                                    iSc3.AddString "<a name=""ilink" & CStr(iLinkCounter) & """ /><h2><p class=""section"">" & mMemberType_p(m) & "</p></h2></a>" & vbCrLf
                                Else
                                    iSc3.AddString "<br><br><br><h1>" & mMemberType_p(m) & "</h1><br>"
                                End If
                                
                                Set iSc4 = New SmartConcat
                                If iPutPMELinks Then
                                    iSc4.AddString "<br><br><br><h1>" & mMemberType_p(m) & "</h1><br>"
                                End If
                                
                                ReDim iTmpSections_PME_Details(0)
                                iRec.MoveFirst
                                Do Until iRec.EOF ' for each property, method or event
                                    iLinkCounter = iLinkCounter + 1
                                    iSc3.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & iRec!Name & "</a><br>" & vbCrLf
                                    AddToList iLinks_ToReplace, LCase(iRec!Name) & "_" & LCase(mMemberType_s(m)) & ".html"
                                    AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
                                    
                                    iSc4.AddString "<a name=""ilink" & CStr(iLinkCounter) & """ /><h2><p class""section"">" & iRec!Name & " " & mMemberType_s(m) & "</p></h2></a>" & vbCrLf
                                    ' Description
                                    If iRec!Params_Info <> "" Then
                                        iSc4.AddString "<p>" & TxtToHTML(HTMLFormatParameters(iRec!Params_Info), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p>" & vbCrLf
                                    End If
                                    If iRec!Long_Description <> "" Then
                                        iSc4.AddString "<h3>Description:</h3>" & vbCrLf
                                        iSc4.AddString "<p>" & TxtToHTML(iRec!Long_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                    ElseIf iRec!Short_Description <> "" Then
                                        iSc4.AddString "<h3>Description:</h3>" & vbCrLf
                                        iSc4.AddString "<p>" & TxtToHTML(iRec!Short_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m)) & "</p><br>" & vbCrLf
                                    End If
                                    
                                    iRec.MoveNext
                                Loop
                                If Not iPutPMELinks Then
                                    iSc3.AddString "<br><br><br>"
                                End If
                                AddToList iTmpSections_PME_Details, iSc4.GenerateCurrentString
                                ' add to a temporary list one whole section of list of properties (with links to each individual file), methods or events
                                AddToList iTmpSections_PME_List, iSc3.GenerateCurrentString
                                AddToList iTmpSections_PME_Details_Joined, Join(iTmpSections_PME_Details, vbCrLf)
                            End If
                        Next m
                        
                        ' add each section of properties, methos and events
                        For c = 1 To UBound(iTmpSections_PME_List)
                            iSc2.AddString iTmpSections_PME_List(c)
                        Next c
                        ' and after that the details
                        For c = 1 To UBound(iTmpSections_PME_Details_Joined)
                            iSc2.AddString iTmpSections_PME_Details_Joined(c)
                        Next c
                        
                        If mHTML_PageFooter <> "" Then iSc2.AddString mHTML_PageFooter & vbCrLf
                        iSc2.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
                        
                        iSc.AddString iSc2.GenerateCurrentString
                    End If
                    iTTypes.MoveNext
                Loop
            End If
        End If
    Next t
    
    If frmReportSelection.chkType(3).Value Then
        ' Constants
        ReDim iTmpSections_Constants(0)
        Set iTEnums = mEnums.Clone
        iTEnums.Index = "Name"
        If iTEnums.RecordCount > 0 Then
            iSc.AddString "<a name=""" & iLinks("constants") & """><h1><p class=""section"">" & mComponentName & " Enumerations</p></h1></a>" & vbCrLf
            
            iTEnums.MoveFirst
            Do Until iTEnums.EOF
                If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
                    If iRec.RecordCount > 0 Then
                        iLinkCounter = iLinkCounter + 1
                        iSc.AddString "<li class=""toc1""><a href=""#ilink" & CStr(iLinkCounter) & """>" & iTEnums!Name & "</a></li>"
                        
                        AddToList iLinks_ToReplace, LCase(iTEnums!Name) & "_enumeration.html"
                        AddToList iLinks_Replacement, "#ilink" & CStr(iLinkCounter)
                        
                        Set iSc3 = New SmartConcat
                        iSc3.AddString "<br>"
                        iSc3.AddString "<a name=""ilink" & CStr(iLinkCounter) & """ /><h2><p class""section"">" & iTEnums!Name & "</h2>" & "</p></h2></a>" & vbCrLf
                        iSc3.AddString "<p>" & TxtToHTML(iTEnums!Description, iTEnums!Name & " Enumeration") & "</p><br>" & vbCrLf
                        
                        iSc3.AddString "<table class=""table table2"">" & vbCrLf
                        iRec.MoveFirst
                        Do Until iRec.EOF
                            iSc3.AddString "  <tr>" & vbCrLf
                            iSc3.AddString "    <td>"
                            iSc3.AddString "<b>" & iRec!Name & "</b> = "
                            iSc3.AddString "    </td>" & vbCrLf
                            iSc3.AddString "    <td>"
                            iSc3.AddString iRec!Value
                            iSc3.AddString "    </td>" & vbCrLf
                            If (iRec!Description <> "") Then
                                iSc3.AddString "    <td>"
                                iSc3.AddString "<a style=""color:#0040A0""> ' " & TxtToHTML(Replace(iRec!Description, vbCrLf, vbTab), iRec!Name & " constant of " & iTEnums!Name & " Enumeration") & "</a>"
                                iSc3.AddString "    </td>" & vbCrLf
                            End If
                            iRec.MoveNext
                            iSc3.AddString "  </tr>" & vbCrLf
                        Loop
                        iSc3.AddString "</table>" & vbCrLf
                        
                        AddToList iTmpSections_Constants, iSc3.GenerateCurrentString
                    End If
                End If
                iTEnums.MoveNext
            Loop
            
            ' add each section of enums
            For c = 1 To UBound(iTmpSections_Constants)
                iSc.AddString iTmpSections_Constants(c)
            Next c
        End If
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoEndNotes.Value = 1) Then
        iSc.AddString "<br>"
        iSc.AddString "<p>" & TxtToHTML(iDesc, "index") & "</p><br>" & vbCrLf
    End If
    
    If mHTML_PageFooter <> "" Then iSc.AddString mHTML_PageFooter & vbCrLf
    iSc.AddString "</body>" & vbCrLf & "</html>" & vbCrLf
    iStr = iSc.GenerateCurrentString
    Set iSc = Nothing
    For c = 1 To UBound(iLinks_ToReplace)
        If InStr(iStr, "<a href=""" & iLinks_ToReplace(c) & """>") Then
            iStr = Replace$(iStr, "<a href=""" & iLinks_ToReplace(c) & """>", "<a href=""" & iLinks_Replacement(c) & """>")
        End If
    Next
    AddPage LCase$(mComponentName) & "_reference.html", iStr
    
End Sub

Private Sub mnuReportingOptions_Click()
    Dim iLng As Long
    
    frmReportingOptions.optHTML(mHTML_Mode).Value = True
    frmReportingOptions.chkExternalCSS.Value = IIf(mExternalCSS, 1, 0)
    frmReportingOptions.chkReplaceCSSFile.Value = IIf(mReplaceCSSFile, 1, 0)
    frmReportingOptions.optPrint(mPrint_Mode).Value = True
    
    frmReportingOptions.HTML_HeadSection = mHTML_HeadSection
    frmReportingOptions.HTML_PageHeaderMP = mHTML_PageHeaderMP_Template
    frmReportingOptions.HTML_PageHeaderOP = mHTML_PageHeaderOP_Template
    frmReportingOptions.HTML_StyleSheet = mHTML_StyleSheet
    frmReportingOptions.HTML_PageFooter = mHTML_PageFooter_Template
    
    frmReportingOptions.Show vbModal
    
    SaveSettingBase "ReportingOptions", "HTML_Mode", frmReportingOptions.HTML_Mode
    SaveSettingBase "ReportingOptions", "HTML_ExternalCSS", frmReportingOptions.ExternalCSS
    If frmReportingOptions.ExternalCSS Then
        SaveSettingBase "ReportingOptions", "HTML_ReplaceCSSFile", frmReportingOptions.ReplaceCSSFile
    End If
    SaveSettingBase "ReportingOptions", "Print_Mode", frmReportingOptions.Print_Mode
    
    If Trim2(frmReportingOptions.HTML_HeadSection) = Trim2(cHTMLDefaultHeadSection) Then
        SaveSettingBase "ReportingOptions", "HTML_HeadSection", ""
    Else
        SaveSettingBase "ReportingOptions", "HTML_HeadSection", Trim2(frmReportingOptions.HTML_HeadSection)
    End If
    If Trim2(frmReportingOptions.HTML_HeadSection) = Trim2(cHTMLDefaultStyleSheet) Then
        SaveSettingBase "ReportingOptions", "HTML_StyleSheet", ""
    Else
        SaveSettingBase "ReportingOptions", "HTML_StyleSheet", Trim2(frmReportingOptions.HTML_StyleSheet)
    End If
    If Trim2(frmReportingOptions.HTML_PageHeaderMP) = Trim2(cHTMLDefaultPageHeaderMP) Then
        SaveSettingBase "ReportingOptions", "HTML_PageHeaderMP", ""
    Else
        SaveSettingBase "ReportingOptions", "HTML_PageHeaderMP", Trim2(frmReportingOptions.HTML_PageHeaderMP)
    End If
    If Trim2(frmReportingOptions.HTML_PageHeaderOP) = Trim2(cHTMLDefaultPageHeaderOP) Then
        SaveSettingBase "ReportingOptions", "HTML_PageHeaderOP", ""
    Else
        SaveSettingBase "ReportingOptions", "HTML_PageHeaderOP", Trim2(frmReportingOptions.HTML_PageHeaderOP)
    End If
    If Trim2(frmReportingOptions.HTML_PageFooter) = Trim2(cHTMLDefaultPageFooter) Then
        SaveSettingBase "ReportingOptions", "HTML_PageFooter", ""
    Else
        SaveSettingBase "ReportingOptions", "HTML_PageFooter", Trim2(frmReportingOptions.HTML_PageFooter)
    End If
    
    LoadReportingOptions
    
    Set frmReportingOptions = Nothing
End Sub

Private Sub mnuReportPrint_Click()
    If Not ShowfrmReportSelection("Print") Then Exit Sub
    
    frmSelectPrinter.Show vbModal
    If frmSelectPrinter.OKPressed Then
        DoPrint
    End If
    Set frmSelectPrinter = Nothing
    
    Unload frmReportSelection
    Set frmReportSelection = Nothing
End Sub

Private Sub mnuReportRTF_Click()
    Dim iDlg As New cDlg
    Dim iFilePath As String
    Dim iRec As Recordset
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTEnums As Recordset
    Dim t As Long
    Dim m As Long
    Dim iTTypes As Recordset
    Dim iDesc As String
    
    If Not ShowfrmReportSelection("RTF") Then Exit Sub
    
    Set iDlg = New cDlg
    iDlg.Filter = "Rich Text Format (*.rtf)|*.rtf"
    iDlg.FileName = GetSetting(App.Title, AppPath4Reg, "OutputPathRTF_" & mComponentName, mComponentName & "_reference" & ".RTF")
    iDlg.ShowSave
    If iDlg.Canceled Then Exit Sub
    iFilePath = iDlg.FileName
    SaveSetting App.Title, AppPath4Reg, "OutputPathRTF_" & mComponentName, iFilePath
    Set iDlg = Nothing
    
    Screen.MousePointer = vbHourglass
    
    rtbAux.Text = ""
    rtbAux.Font.Name = "Arial"
    rtbAux.Font.Size = 12
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    iDesc = GetGeneralInfo("Introduction")
    If (((mControls.RecordCount * frmReportSelection.chkType(1).Value) + (mClasses.RecordCount * frmReportSelection.chkType(2).Value) + (mEnums.RecordCount * frmReportSelection.chkType(3).Value)) > 1) Or (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        ' title
        rtbAux.SelAlignment = rtfCenter
        rtbAux.SelBold = True
        rtbAux.SelFontSize = 28
        rtbAux.SelText = vbCrLf & vbCrLf & vbCrLf & mComponentName & IIf((frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoVersion.Value = 1) And (mComponentVersion <> ""), " " & mComponentVersion, "") & " Reference" & vbCrLf & vbCrLf
        rtbAux.SelFontSize = 12
        rtbAux.SelBold = False
        rtbAux.SelAlignment = rtfLeft
        If (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoReleaseDate.Value = 1) And (mComponentReleaseDate <> 0) Then
            rtbAux.SelText = "Release date: " & FormatDateTime(mComponentReleaseDate, vbShortDate) & "" & vbCrLf & vbCrLf & vbCrLf
        End If
        
        If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
            rtbAux.SelFontSize = 22
            rtbAux.SelIndent = 0
            rtbAux.SelBold = True
            rtbAux.SelColor = &H808080 ' &HE189A
            rtbAux.SelText = "Introduction" & vbCrLf & vbCrLf
            rtbAux.SelColor = vbBlack
            rtbAux.SelBold = False
            
            AddRTF iDesc
            rtbAux.SelText = vbCrLf & vbCrLf
        End If
    End If
    rtbAux.SelAlignment = rtfLeft
    
    For t = 1 To 2 ' controls and classes
        If frmReportSelection.chkType(t).Value Then
            If t = 1 Then
                Set iTTypes = iTControls
            ElseIf t = 2 Then
                Set iTTypes = iTClasses
            End If
            If iTTypes.RecordCount > 0 Then
                iTTypes.MoveFirst
                Do Until iTTypes.EOF
                    If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                        ' title
                        rtbAux.SelFontSize = 22
                        rtbAux.SelIndent = 0
                        rtbAux.SelBold = True
                        rtbAux.SelColor = &H808080 ' &HE189A
                        rtbAux.SelText = vbCrLf & vbCrLf & iTTypes!Name & " " & LCase(mObjectType_s2(t)) & vbCrLf & vbCrLf
                        rtbAux.SelColor = vbBlack
                        rtbAux.SelBold = False
                        
                        ' Description
                        rtbAux.SelFontSize = 12
                        If iTTypes!Long_Description <> "" Then
                            iDesc = iTTypes!Long_Description
                        Else
                            iDesc = iTTypes!Short_Description
                        End If
                        If iDesc <> "" Then
                            AddRTF iDesc
                            rtbAux.SelText = vbCrLf & vbCrLf & vbCrLf & vbCrLf
                        End If
                        
                        For m = 1 To 3 ' properties, methods or events
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then
                                rtbAux.SelFontSize = 14
                                rtbAux.SelBold = True
                                rtbAux.SelColor = &HDB7920 'vbBlu
                                rtbAux.SelFontSize = 18
                                rtbAux.SelText = mMemberType_p(m)
                                rtbAux.SelColor = vbBlack
                                rtbAux.SelFontSize = 14
                                rtbAux.SelText = ": " & vbCrLf & vbCrLf & vbCrLf
                                rtbAux.SelBold = False
                                rtbAux.SelFontSize = 12
                                
                                iRec.MoveFirst
                                Do Until iRec.EOF
                                    ' Property, method or event name
                                    rtbAux.SelBold = True
                                    rtbAux.SelText = iRec!Name
                                    rtbAux.SelBold = False
                                    rtbAux.SelItalic = True
                                    If mMemberType_s(m) = "Method" Then
                                        If Left$(iRec!Params_Info, 17) = "Return Type:" & vbTab & "None" Then
                                            rtbAux.SelText = " " & LCase$(mMemberType_s(m))
                                        Else
                                            rtbAux.SelText = " function"
                                        End If
                                    Else
                                        rtbAux.SelText = " " & LCase$(mMemberType_s(m))
                                    End If
                                    rtbAux.SelItalic = False
                                    rtbAux.SelText = ":" & vbCrLf & vbCrLf
                                    rtbAux.SelIndent = rtbAux.SelIndent + 500
        
                                    ' Description
                                    If iRec!Params_Info <> "" Then
        '                                rtbAux.SelText = "Parameters information:" & vbCrLf
                                        If Left$(iRec!Params_Info, 20) = "Return Type:" & vbTab & "None." & vbCrLf Then
                                            AddRTF Mid$(iRec!Params_Info, 21)
                                        Else
                                            AddRTF iRec!Params_Info
                                        End If
                                        rtbAux.SelFontName = rtbAux.Font.Name
                                        rtbAux.SelText = vbCrLf & vbCrLf
                                    End If
                                    iDesc = ""
                                    If iRec!Long_Description <> "" Then
                                        iDesc = iRec!Long_Description
                                    ElseIf iRec!Short_Description <> "" Then
                                        iDesc = iRec!Short_Description
                                    End If
                                    If iDesc <> "" Then
                                        'rtbAux.SelText = vbCrLf & vbCrLf & "Description:" & vbCrLf
                                        AddRTF iDesc
                                        rtbAux.SelText = vbCrLf & vbCrLf
                                    End If
                                    
                                    rtbAux.SelIndent = rtbAux.SelIndent - 500
                                    iRec.MoveNext
                                Loop
                                rtbAux.SelText = vbCrLf & vbCrLf & vbCrLf & vbCrLf ' leave some space
                            End If
                        Next m
                    End If
                    iTTypes.MoveNext
                Loop
            End If
        End If
    Next
    
    ' Constants
    If frmReportSelection.chkType(3).Value Then
        Set iTEnums = mEnums.Clone
        iTEnums.Index = "Name"
        If iTEnums.RecordCount > 0 Then
            
            ' title
            rtbAux.SelFontSize = 14
            rtbAux.SelBold = True
            rtbAux.SelColor = &HDB7920 'vbBlu
            rtbAux.SelFontSize = 18
            rtbAux.SelText = vbCrLf & vbCrLf & "Constants" & vbCrLf & vbCrLf
            rtbAux.SelColor = vbBlack
            rtbAux.SelFontSize = 14
            rtbAux.SelBold = False
            rtbAux.SelFontSize = 12
            
            iTEnums.MoveFirst
            Do Until iTEnums.EOF
                If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
                    If iRec.RecordCount > 0 Then
                        
                        rtbAux.SelFontSize = 14
                        rtbAux.SelIndent = 0
                        rtbAux.SelBold = True
                        rtbAux.SelColor = &HE189A
                        rtbAux.SelText = vbCrLf & iTEnums!Name & " enumeration" & vbCrLf & vbCrLf
                        rtbAux.SelColor = vbWindowText
                        rtbAux.SelFontSize = 12
                        If iTEnums!Description <> "" Then
                            AddRTF iTEnums!Description
                            rtbAux.SelText = vbCrLf & vbCrLf
                        End If
                        rtbAux.SelBold = False
                        
                        iRec.MoveFirst
                        Do Until iRec.EOF
                            rtbAux.SelIndent = 0
                            rtbAux.SelBold = True
                            rtbAux.SelText = iRec!Name & " = " & iRec!Value & vbCrLf
                            rtbAux.SelBold = False
                            rtbAux.SelIndent = 500
                            If iRec!Description <> "" Then
                                AddRTF iRec!Description
                                rtbAux.SelRTF = vbCrLf & vbCrLf
                            End If
                            iRec.MoveNext
                        Loop
                    End If
                End If
                iTEnums.MoveNext
            Loop
        End If
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoEndNotes.Value = 1) Then
        rtbAux.SelFontSize = 22
        rtbAux.SelIndent = 0
        rtbAux.SelBold = True
        rtbAux.SelColor = &H808080 ' &HE189A
        rtbAux.SelText = vbCrLf & vbCrLf & "End Notes" & vbCrLf & vbCrLf
        rtbAux.SelColor = vbBlack
        rtbAux.SelBold = False
        
        AddRTF iDesc
        rtbAux.SelText = vbCrLf & vbCrLf
    End If
    
    Unload frmReportSelection
    Set frmReportSelection = Nothing
    
    rtbAux.SaveFile iFilePath
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl Is trv1 Then
        If KeyCode = vbKeyDelete Then
            If mnuDataDelete.Enabled Then
                If MsgBox("Delete member " & trv1.SelectedItem.Text & "?", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
                mnuDataDelete_Click
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Set gIcon = Me.Icon
    Dim iStr As String
    Dim iStrs() As String
    
    mFontPropertion = 1
    Set mAppFont = txtLongDescription.Font
    iStr = GetSetting(App.Title, "Preferences", "FontAttr", "")
    If iStr <> "" Then
        iStrs = Split(iStr, "|")
        If UBound(iStrs) = 5 Then
            mAppFont.Name = iStrs(0)
            mAppFont.Size = Val(iStrs(1)) / 100
            mAppFont.Bold = CBool(iStrs(2))
            mAppFont.Charset = Val(iStrs(3))
            mAppFont.Italic = CBool(iStrs(4))
            mAppFont.Weight = Val(iStrs(5))
        End If
    End If
    mFontPropertion = mAppFont.Size / 12
    SetControlsFont
    
    mObjectType_s(1) = "Control"
    mObjectType_s(2) = "Class"
    mObjectType_s2(1) = "Control"
    mObjectType_s2(2) = "Object"
    mObjectType_p(1) = "Controls"
    mObjectType_p(2) = "Classes"
    mMemberType_s(1) = "Property"
    mMemberType_s(2) = "Method"
    mMemberType_s(3) = "Event"
    mMemberType_p(1) = "Properties"
    mMemberType_p(2) = "Methods"
    mMemberType_p(3) = "Events"
    
    Set mControlsEditZone = New Collection
    mControlsEditZone.Add lblName
    mControlsEditZone.Add txtName
    mControlsEditZone.Add lblLongDescription
    mControlsEditZone.Add lblShortDescription
    mControlsEditZone.Add txtLongDescription
    mControlsEditZone.Add txtShortDescription
    mControlsEditZone.Add txtValue
    mControlsEditZone.Add lblParamsInfo
    mControlsEditZone.Add txtParamsInfo
    
    Set mMemberTypeRec(1) = mProperties
    Set mMemberTypeRec(2) = mMethods
    Set mMemberTypeRec(3) = mEvents
    
    ReDim mRefTables(5)
    mRefTables(0) = "Controls"
    mRefTables(1) = "Classes"
    mRefTables(2) = "Enums"
    mRefTables(3) = "Properties"
    mRefTables(4) = "Methods"
    mRefTables(5) = "Events"
    
    mCurrentDBPath = GetSetting(App.Title, AppPath4Reg, "CurrentDBPath", "")
    If mCurrentDBPath <> "" Then
        If FileExists(mCurrentDBPath) Then
            OpenTheDatabase
            mSelectedType = Val(GetSetting(App.Title, AppPath4Reg, "SelectedType", "0"))
            mSelectedID = Val(GetSetting(App.Title, AppPath4Reg, "SelectedID", "0"))
            mSelectedSecondaryID = Val(GetSetting(App.Title, AppPath4Reg, "SelectedSecondaryID", "0"))
        End If
    End If
    
    ShowTree
    trv1_Click
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UpdateData
End Sub

Private Sub Form_Resize()
    If Me.Width < 9700 Then Me.Width = 9700
    If Me.Height < 8000 Then Me.Height = 8000
    
    If Me.WindowState <> vbMinimized Then
        trv1.Height = Me.ScaleHeight - (trv1.Top * 2)
        If Me.ScaleWidth < 10000 Then
            trv1.Width = 3000
        ElseIf Me.ScaleWidth < 12000 Then
            trv1.Width = 4200
        ElseIf Me.ScaleWidth < 14000 Then
            trv1.Width = 5000
        Else
            trv1.Width = 5800
        End If
        SetFontsTo GetDesiredFontSize
        PlaceDataControls
    End If
End Sub

Private Function GetDesiredFontSize() As Single
    If Me.ScaleWidth < 10000 Then
        GetDesiredFontSize = 9
    ElseIf Me.ScaleWidth < 12000 Then
        GetDesiredFontSize = 10
    ElseIf Me.ScaleWidth < 14000 Then
        GetDesiredFontSize = 11
    Else
        GetDesiredFontSize = 12
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, AppPath4Reg, "CurrentDBPath", mCurrentDBPath
    SaveSetting App.Title, AppPath4Reg, "SelectedType", mSelectedType
    SaveSetting App.Title, AppPath4Reg, "SelectedID", mSelectedID
    SaveSetting App.Title, AppPath4Reg, "SelectedSecondaryID", mSelectedSecondaryID
    If Not mDatabase Is Nothing Then
        mDatabase.Close
        Set mDatabase = Nothing
    End If
End Sub

Private Sub mnuConstantsOrderedByName_Click()
    If mSelectedType <> entEnum Then Err.Raise 1235
    mEnums.Index = "PrimaryKey"
    mEnums.Seek "=", mSelectedID
    If mEnums.NoMatch Then Err.Raise 1234
    mEnums.Edit
    mEnums!Ordered_By_Value = False
    mEnums.Update
    mCurrentAction = ecaDefault
    ShowTree
End Sub

Private Sub mnuConstantsOrderedByValue_Click()
    If mSelectedType <> entEnum Then Err.Raise 1235
    mEnums.Index = "PrimaryKey"
    mEnums.Seek "=", mSelectedID
    If mEnums.NoMatch Then Err.Raise 1234
    mEnums.Edit
    mEnums!Ordered_By_Value = True
    mEnums.Update
    mCurrentAction = ecaDefault
    ShowTree
End Sub

Private Sub mnuCopyEnumList_Click()
    Dim iMembers As Recordset
    Dim iStr As String
    
    If mSelectedType <> entEnum Then Err.Raise 1234
    mEnums.Index = "PrimaryKey"
    mEnums.Seek "=", mSelectedID
    If mEnums.NoMatch Then Err.Raise 1234
    
    Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Constants.Enum_ID = " & mEnums!Enum_ID & ") ORDER BY Constants.Value")
    iStr = ""
    If iMembers.RecordCount > 0 Then
        iMembers.MoveFirst
        Do Until iMembers.EOF
            If iStr <> "" Then iStr = iStr & vbCrLf
            iStr = iStr & "    " & iMembers!Name & " = " & iMembers!Value
            iMembers.MoveNext
        Loop
    End If
    Clipboard.Clear
    Clipboard.SetText iStr
End Sub

Private Sub mnuCopyList_Click()
    Dim iStr As String
    Dim iNode As Node
    
    If trv1.SelectedItem.Children Then
        Set iNode = trv1.SelectedItem.Child
        Do Until iNode Is Nothing
            iStr = iStr & iNode.Text & vbCrLf
            Set iNode = iNode.Next
        Loop
    End If
    Clipboard.Clear
    Clipboard.SetText iStr
End Sub

Private Sub mnuCopyName_Click()
    Clipboard.Clear
    Clipboard.SetText trv1.SelectedItem.Text
End Sub

Private Sub mnuCopyName2_Click()
    mnuCopyName_Click
End Sub

Private Sub mnuCopyName3_Click()
    mnuCopyName_Click
End Sub

Private Sub mnuDeleteAll_Click()
    If MsgBox("This will make the whole database blank, continue?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    ControlsEditZoneVisible = False
    CurrentAction = ecaDefault
    mSelectedType = entNone
    mSelectedID = 0
    
    If mClasses.RecordCount > 0 Then
        mClasses.MoveLast
        Do Until mClasses.RecordCount = 0
            mClasses.Delete
            mClasses.MoveLast
        Loop
    End If
    If mControls.RecordCount > 0 Then
        mControls.MoveLast
        Do Until mControls.RecordCount = 0
            mControls.Delete
            mControls.MoveLast
        Loop
    End If
    If mEnums.RecordCount > 0 Then
        mEnums.MoveLast
        Do Until mEnums.RecordCount = 0
            mEnums.Delete
            mEnums.MoveLast
        Loop
    End If
    If mProperties.RecordCount > 0 Then
        mProperties.MoveLast
        Do Until mProperties.RecordCount = 0
            mProperties.Delete
            mProperties.MoveLast
        Loop
    End If
    If mMethods.RecordCount > 0 Then
        mMethods.MoveLast
        Do Until mMethods.RecordCount = 0
            mMethods.Delete
            mMethods.MoveLast
        Loop
    End If
    If mEvents.RecordCount > 0 Then
        mEvents.MoveLast
        Do Until mEvents.RecordCount = 0
            mEvents.Delete
            mEvents.MoveLast
        Loop
    End If
    SaveSettingBase "General", "FileImported", "0"
    mFileImported = False
    ShowTree
End Sub

Private Sub mnuDeleteMember_Click()
    Dim iNode As Node
    Dim iCurrentObjectTypePluralStr As String
    Dim iMembers As Recordset
    
    mDeletingNode = True
    If mSelectedType = entProperty Then
        iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
        Set iMembers = mDatabase.OpenRecordset("SELECT * FROM " & iCurrentObjectTypePluralStr & "_Properties WHERE (Property_ID = " & mSelectedID & ") AND (" & GetCurrentObjectTypeSingularStr & "_ID = " & mSelectedSecondaryID & ")")
        If iMembers.RecordCount <> 1 Then Err.Raise 1234
        iMembers.MoveFirst
        If iMembers.RecordCount <> 1 Then Err.Raise 1234
        iMembers.Delete
        Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties WHERE (Property_ID = " & mSelectedID & ")")
        If iMembers.RecordCount = 0 Then
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties WHERE (Property_ID = " & mSelectedID & ")")
            If iMembers.RecordCount = 0 Then
                mProperties.Index = "PrimaryKey"
                mProperties.Seek "=", mSelectedID
                If mProperties.NoMatch Then Err.Raise 1234
                mProperties.Delete
            End If
        End If
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    ElseIf mSelectedType = entMethod Then
        iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
        Set iMembers = mDatabase.OpenRecordset("SELECT * FROM " & iCurrentObjectTypePluralStr & "_Methods WHERE (Method_ID = " & mSelectedID & ") AND (" & GetCurrentObjectTypeSingularStr & "_ID = " & mSelectedSecondaryID & ")")
        If iMembers.RecordCount <> 1 Then Err.Raise 1234
        iMembers.MoveFirst
        If iMembers.RecordCount <> 1 Then Err.Raise 1234
        iMembers.Delete
        Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods WHERE (Method_ID = " & mSelectedID & ")")
        If iMembers.RecordCount = 0 Then
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods WHERE (Method_ID = " & mSelectedID & ")")
            If iMembers.RecordCount = 0 Then
                mMethods.Index = "PrimaryKey"
                mMethods.Seek "=", mSelectedID
                If mMethods.NoMatch Then Err.Raise 1234
                mMethods.Delete
            End If
        End If
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    ElseIf mSelectedType = entEvent Then
        iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
        Set iMembers = mDatabase.OpenRecordset("SELECT * FROM " & iCurrentObjectTypePluralStr & "_Events WHERE (Event_ID = " & mSelectedID & ") AND (" & GetCurrentObjectTypeSingularStr & "_ID = " & mSelectedSecondaryID & ")")
        If iMembers.RecordCount <> 1 Then Err.Raise 1234
        iMembers.MoveFirst
        If iMembers.RecordCount <> 1 Then Err.Raise 1234
        iMembers.Delete
        Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Events WHERE (Event_ID = " & mSelectedID & ")")
        If iMembers.RecordCount = 0 Then
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Events WHERE (Event_ID = " & mSelectedID & ")")
            If iMembers.RecordCount = 0 Then
                mEvents.Index = "PrimaryKey"
                mEvents.Seek "=", mSelectedID
                If mEvents.NoMatch Then Err.Raise 1234
                mEvents.Delete
            End If
        End If
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    ElseIf mSelectedType = entConstant Then
        mConstants.Index = "PrimaryKey"
        mConstants.Seek "=", mSelectedID
        If mConstants.NoMatch Then Err.Raise 1234
        mConstants.Delete
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    End If
    mDeletingNode = False
End Sub

Private Sub NavigateNearNode()
    If Not trv1.SelectedItem.Previous Is Nothing Then
        trv1.SelectedItem.Previous.EnsureVisible
        trv1.SelectedItem.Previous.Selected = True
    Else
        If Not trv1.SelectedItem.Next Is Nothing Then
            trv1.SelectedItem.Next.EnsureVisible
            trv1.SelectedItem.Next.Selected = True
        Else
            If Not trv1.SelectedItem.FirstSibling Is Nothing Then
                trv1.SelectedItem.FirstSibling.EnsureVisible
                trv1.SelectedItem.FirstSibling.Selected = True
            Else
                If Not trv1.SelectedItem.Parent Is Nothing Then
                    trv1.SelectedItem.Parent.Selected = True
                End If
            End If
        End If
    End If
    UpdateCurrentSelected
End Sub

Private Sub mnuDeleteObject_Click()
    Dim iNode As Node
    Dim iRec As Recordset
    Dim iUsed As Boolean
    Dim iRec2 As Recordset
    Dim iProperties As Recordset
    Dim iMethods As Recordset
    Dim iEvents As Recordset
    
    If mSelectedType = entClass Then
        Set iProperties = mProperties.Clone
        Set iMethods = mMethods.Clone
        Set iEvents = mEvents.Clone
        iProperties.Index = "PrimaryKey"
        iMethods.Index = "PrimaryKey"
        iEvents.Index = "PrimaryKey"
        mClasses.Index = "PrimaryKey"
        mClasses.Seek "=", mSelectedID
        If mClasses.NoMatch Then Err.Raise 1234
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties WHERE (Class_ID = " & mClasses!Class_ID & ")")
        If Not iRec.EOF Then
            iRec.MoveFirst
            Do Until iRec.EOF
                iUsed = False
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties WHERE (Property_ID = " & iRec!Property_ID & ") AND (Class_ID <> " & mClasses!Class_ID & ")")
                If iRec2.RecordCount > 0 Then iUsed = True
                If Not iUsed Then
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties WHERE (Property_ID = " & iRec!Property_ID & ")")
                    If iRec2.RecordCount > 0 Then iUsed = True
                End If
                If Not iUsed Then
                    iProperties.Seek "=", iRec!Property_ID
                    If iProperties.NoMatch Then Err.Raise 12367
                    iProperties.Edit
                    iProperties!Auxiliary_Field = 0
                    iProperties.Update
                End If
                iRec.MoveNext
            Loop
        End If
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods WHERE (Class_ID = " & mClasses!Class_ID & ")")
        If Not iRec.EOF Then
            iRec.MoveFirst
            Do Until iRec.EOF
                iUsed = False
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods WHERE (Method_ID = " & iRec!Method_ID & ") AND (Class_ID <> " & mClasses!Class_ID & ")")
                If iRec2.RecordCount > 0 Then iUsed = True
                If Not iUsed Then
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods WHERE (Method_ID = " & iRec!Method_ID & ")")
                    If iRec2.RecordCount > 0 Then iUsed = True
                End If
                If Not iUsed Then
                    iMethods.Seek "=", iRec!Method_ID
                    If iMethods.NoMatch Then Err.Raise 12367
                    iMethods.Edit
                    iMethods!Auxiliary_Field = 0
                    iMethods.Update
                End If
                iRec.MoveNext
            Loop
        End If
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes_Events WHERE (Class_ID = " & mClasses!Class_ID & ")")
        If Not iRec.EOF Then
            iRec.MoveFirst
            Do Until iRec.EOF
                iUsed = False
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Events WHERE (Event_ID = " & iRec!Event_ID & ") AND (Class_ID <> " & mClasses!Class_ID & ")")
                If iRec2.RecordCount > 0 Then iUsed = True
                If Not iUsed Then
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Events WHERE (Event_ID = " & iRec!Event_ID & ")")
                    If iRec2.RecordCount > 0 Then iUsed = True
                End If
                If Not iUsed Then
                    iEvents.Seek "=", iRec!Event_ID
                    If iEvents.NoMatch Then Err.Raise 12367
                    iEvents.Edit
                    iEvents!Auxiliary_Field = 0
                    iEvents.Update
                End If
                iRec.MoveNext
            Loop
        End If
        mClasses.Delete
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    ElseIf mSelectedType = entControl Then
        mControls.Index = "PrimaryKey"
        mControls.Seek "=", mSelectedID
        If mControls.NoMatch Then Err.Raise 1234
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties WHERE (Control_ID = " & mControls!control_ID & ")")
        If Not iRec.EOF Then
            iRec.MoveFirst
            Do Until iRec.EOF
                iUsed = False
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties WHERE (Property_ID = " & iRec!Property_ID & ") AND (Control_ID <> " & mControls!control_ID & ")")
                If iRec2.RecordCount > 0 Then iUsed = True
                If Not iUsed Then
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties WHERE (Property_ID = " & iRec!Property_ID & ")")
                    If iRec2.RecordCount > 0 Then iUsed = True
                End If
                If Not iUsed Then
                    iProperties.Seek "=", iRec!Property_ID
                    If iProperties.NoMatch Then Err.Raise 12367
                    iProperties.Edit
                    iProperties!Auxiliary_Field = 0
                    iProperties.Update
                End If
                iRec.MoveNext
            Loop
        End If
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods WHERE (Control_ID = " & mControls!control_ID & ")")
        If Not iRec.EOF Then
            iRec.MoveFirst
            Do Until iRec.EOF
                iUsed = False
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods WHERE (Method_ID = " & iRec!Method_ID & ") AND (Control_ID <> " & mControls!control_ID & ")")
                If iRec2.RecordCount > 0 Then iUsed = True
                If Not iUsed Then
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods WHERE (Method_ID = " & iRec!Method_ID & ")")
                    If iRec2.RecordCount > 0 Then iUsed = True
                End If
                If Not iUsed Then
                    iMethods.Seek "=", iRec!Method_ID
                    If iMethods.NoMatch Then Err.Raise 12367
                    iMethods.Edit
                    iMethods!Auxiliary_Field = 0
                    iMethods.Update
                End If
                iRec.MoveNext
            Loop
        End If
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls_Events WHERE (Control_ID = " & mControls!control_ID & ")")
        If Not iRec.EOF Then
            iRec.MoveFirst
            Do Until iRec.EOF
                iUsed = False
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Events WHERE (Event_ID = " & iRec!Event_ID & ") AND (Control_ID <> " & mControls!control_ID & ")")
                If iRec2.RecordCount > 0 Then iUsed = True
                If Not iUsed Then
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Events WHERE (Event_ID = " & iRec!Event_ID & ")")
                    If iRec2.RecordCount > 0 Then iUsed = True
                End If
                If Not iUsed Then
                    iEvents.Seek "=", iRec!Event_ID
                    If iEvents.NoMatch Then Err.Raise 12367
                    iEvents.Edit
                    iEvents!Auxiliary_Field = 0
                    iEvents.Update
                End If
                iRec.MoveNext
            Loop
        End If
        mControls.Delete
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    ElseIf mSelectedType = entEnum Then
        mEnums.Index = "PrimaryKey"
        mEnums.Seek "=", mSelectedID
        If mEnums.NoMatch Then Err.Raise 1234
        mEnums.Delete
        Set iNode = trv1.SelectedItem
        NavigateNearNode
        trv1.Nodes.Remove (iNode.Key)
    End If
    
End Sub

Private Sub mnuDeleteObject2_Click()
    mnuDeleteObject_Click
End Sub

Private Sub mnuDeleteOrphanMembers_Click()
    Dim c As Long
    Dim iStr As String
    Dim iRec As Recordset
    Dim iPos As Long
    Dim iCount As Long
    Dim iList() As String
    
    ReDim iList(0)
    For c = 1 To 3
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(c) & " WHERE (Auxiliary_Field = 0)")
        If iRec.RecordCount > 0 Then
            iRec.MoveLast
            If iStr = "" Then
                iStr = "There are "
            Else
                iStr = iStr & ", "
            End If
            iStr = iStr & iRec.RecordCount & " " & IIf(iRec.RecordCount = 1, mMemberType_s(c), mMemberType_p(c))
            iCount = iCount + iRec.RecordCount
            
            iRec.MoveFirst
            Do Until iRec.EOF
                AddToList iList, iRec!Name
                iRec.MoveNext
            Loop
        End If
    Next c
    iPos = InStrRev(iStr, ", ")
    If iPos > 0 Then
        iStr = Left$(iStr, iPos - 1) & " and " & Mid$(iStr, iPos + 2)
    End If
    If iStr = "" Then
        MsgBox "No orphan members found.", vbInformation
    Else
        If iCount = 1 Then iStr = Replace$(iStr, "There are ", "There is ")
        If MsgBox(iStr & "." & vbCrLf & "Delete " & IIf(iCount = 1, "it", "them") & "?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
            Exit Sub
        End If
    
        For c = 1 To 3
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(c) & " WHERE (Auxiliary_Field = 0)")
            If iRec.RecordCount > 0 Then
                iRec.MoveLast
                Do Until iRec.RecordCount = 0
                    iRec.Delete
                    iRec.MoveLast
                Loop
            End If
        Next c
    End If
    
    mSelectedType = 0
    mSelectedID = 0
    mSelectedSecondaryID = 0
    mCurrentAction = ecaDefault
    ControlsEditZoneVisible = False
    ClearControlsEditZone
    
    ShowTree
End Sub

Private Sub mnuReportText_Click()
    Dim iDlg As New cDlg
    Dim iFilePath As String
    Dim iRec As Recordset
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTEnums As Recordset
    Dim t As Long
    Dim m As Long
    Dim iTTypes As Recordset
    Dim iDesc As String
    Dim iSc As SmartConcat
    
    If Not ShowfrmReportSelection("TXT") Then Exit Sub
    
    Set iDlg = New cDlg
    iDlg.Filter = "Plain Text (*.txt)|*.txt"
    iDlg.FileName = GetSetting(App.Title, AppPath4Reg, "OutputPathTXT_" & mComponentName, mComponentName & "_reference" & ".txt")
    iDlg.ShowSave
    If iDlg.Canceled Then Exit Sub
    iFilePath = iDlg.FileName
    SaveSetting App.Title, AppPath4Reg, "OutputPathTXT_" & mComponentName, iFilePath
    Set iDlg = Nothing
    
    Screen.MousePointer = vbHourglass
    
    Set iSc = New SmartConcat
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    iDesc = GetGeneralInfo("Introduction")
    If (((mControls.RecordCount * frmReportSelection.chkType(1).Value) + (mClasses.RecordCount * frmReportSelection.chkType(2).Value) + (mEnums.RecordCount * frmReportSelection.chkType(3).Value)) > 1) Or (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        ' title
        
        iSc.AddString UCase$(mComponentName & " Reference") & vbCrLf & vbCrLf
        If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
            iSc.AddString "Introduction:" & vbCrLf
            rtbAux.Text = ""
            AddRTF iDesc
            iSc.AddString rtbAux.Text & vbCrLf & vbCrLf
        End If
    End If
    
    For t = 1 To 2 ' controls and classes
        If frmReportSelection.chkType(t).Value Then
            If t = 1 Then
                Set iTTypes = iTControls
            ElseIf t = 2 Then
                Set iTTypes = iTClasses
            End If
            If iTTypes.RecordCount > 0 Then
                iTTypes.MoveFirst
                Do Until iTTypes.EOF
                    If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                        ' title
                        iSc.AddString "*" & iTTypes!Name & " " & LCase(mObjectType_s2(t)) & "*" & vbCrLf
                        
                        ' Description
                        If iTTypes!Long_Description <> "" Then
                            iDesc = iTTypes!Long_Description
                        Else
                            iDesc = iTTypes!Short_Description
                        End If
                        If iDesc <> "" Then
                            rtbAux.Text = ""
                            AddRTF iDesc
                            iSc.AddString "- " & rtbAux.Text & vbCrLf & vbCrLf
                        End If
                        
                        For m = 1 To 3 ' properties, methods or events
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then
                                iSc.AddString mMemberType_p(m)
                                iSc.AddString ": " & vbCrLf & vbCrLf
                                iRec.MoveFirst
                                Do Until iRec.EOF
                                    ' Property, method or event name
                                    iSc.AddString iRec!Name
                                    If mMemberType_s(m) = "Method" Then
                                        If Left$(iRec!Params_Info, 17) = "Return Type:" & vbTab & "None" Then
                                            iSc.AddString " " & LCase$(mMemberType_s(m))
                                        Else
                                            iSc.AddString " function"
                                        End If
                                    Else
                                        iSc.AddString " " & LCase$(mMemberType_s(m))
                                    End If
                                    iSc.AddString ":" & vbCrLf
        
                                    ' Description
                                    If iRec!Params_Info <> "" Then
                                        rtbAux.Text = ""
                                        If Left$(iRec!Params_Info, 20) = "Return Type:" & vbTab & "None." & vbCrLf Then
                                            AddRTF Mid$(iRec!Params_Info, 21)
                                        Else
                                            AddRTF iRec!Params_Info
                                        End If
                                        iSc.AddString TabLines(rtbAux.Text) & vbCrLf
                                    End If
                                    iDesc = ""
                                    If iRec!Long_Description <> "" Then
                                        iDesc = iRec!Long_Description
                                    ElseIf iRec!Short_Description <> "" Then
                                        iDesc = iRec!Short_Description
                                    End If
                                    If iDesc <> "" Then
                                        rtbAux.Text = ""
                                        AddRTF iDesc
                                        iSc.AddString IIf(m = 2, vbCrLf, "") & TabLines(rtbAux.Text) & vbCrLf & vbCrLf
                                    End If
                                    iSc.AddString vbCrLf
                                    
                                    iRec.MoveNext
                                Loop
                                If Not iRec.EOF Then
                                    iSc.AddString vbCrLf & vbCrLf  ' leave some space
                                End If
                            End If
                        Next m
                    End If
                    iTTypes.MoveNext
                Loop
            End If
        End If
    Next
    
    ' Constants
    If frmReportSelection.chkType(3).Value Then
        Set iTEnums = mEnums.Clone
        iTEnums.Index = "Name"
        If iTEnums.RecordCount > 0 Then
            
            ' title
            iSc.AddString "Constants" & vbCrLf & vbCrLf
            
            iTEnums.MoveFirst
            Do Until iTEnums.EOF
                If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
                    If iRec.RecordCount > 0 Then
                        iSc.AddString iTEnums!Name & " ENUMERATION" & vbCrLf & vbCrLf
                        If iTEnums!Description <> "" Then
                            rtbAux.Text = ""
                            AddRTF iTEnums!Description
                            iSc.AddString rtbAux.Text & vbCrLf & vbCrLf
                        End If
                        
                        iRec.MoveFirst
                        Do Until iRec.EOF
                            iSc.AddString vbTab & iRec!Name & " = " & iRec!Value & vbCrLf
                            If iRec!Description <> "" Then
                                rtbAux.Text = ""
                                AddRTF iRec!Description
                                iSc.AddString TabLines(Replace(rtbAux.Text, vbCrLf & vbCrLf, vbCrLf)) & vbCrLf & vbCrLf
                            End If
                            iRec.MoveNext
                        Loop
                        iSc.AddString vbCrLf
                    End If
                End If
                iTEnums.MoveNext
            Loop
        End If
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoEndNotes.Value = 1) Then
        iSc.AddString "End Notes" & vbCrLf & vbCrLf
        rtbAux.Text = ""
        AddRTF iDesc
        iSc.AddString rtbAux.Text
    End If
    
    Unload frmReportSelection
    Set frmReportSelection = Nothing
    
    rtbAux.Text = ""
    rtbAux.Text = iSc.GenerateCurrentString
    rtbAux.SaveFile iFilePath, rtfText
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuListMarkupLinksErrors_Click()
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTTypes As Recordset
    Dim iTEnums As Recordset
    Dim iRec As Recordset
    Dim iSc As New SmartConcat
    Dim iDesc As String
    Dim t As Long
    Dim m As Long
    Dim c As Long
    
    ReDim mLinkErrors(0)
    ReDim mNonUniqueMemberNamePages(0)
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    iDesc = GetGeneralInfo("Introduction")
    If iDesc <> "" Then
        Call TxtToHTML(iDesc, "index")
    End If
    
    ' Objects
    For t = 1 To 2 ' controls and classes
        If t = 1 Then
            Set iTTypes = iTControls
        ElseIf t = 2 Then
            Set iTTypes = iTClasses
        End If
        
        If iTTypes.RecordCount > 0 Then 'if there are controls (or classes)
            iTTypes.MoveFirst
            Do Until iTTypes.EOF ' for each control or class
                ' Description
                If iTTypes!Long_Description <> "" Then
                    iDesc = iTTypes!Long_Description
                Else
                    iDesc = iTTypes!Short_Description
                End If
                If iDesc <> "" Then
                    Call TxtToHTML(iDesc, iTTypes!Name & " " & LCase(mObjectType_s2(t)))
                End If
                For m = 1 To 3 ' properties, methods or events
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                    If iRec.RecordCount > 0 Then ' if there are properties (methods or events)
                        iRec.MoveFirst
                        Do Until iRec.EOF ' for each property, method or event
                            ' Description
                            If iRec!Params_Info <> "" Then
                                Call TxtToHTML(HTMLFormatParameters(iRec!Params_Info), iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m))
                            End If
                            If iRec!Long_Description <> "" Then
                                Call TxtToHTML(iRec!Long_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m))
                            ElseIf iRec!Short_Description <> "" Then
                                Call TxtToHTML(iRec!Short_Description, iTTypes!Name & " " & LCase(mObjectType_s2(t)) & ", " & iRec!Name & " " & mMemberType_s(m))
                            End If
                            iRec.MoveNext
                        Loop
                    End If
                Next m
                iTTypes.MoveNext
            Loop
        End If
    Next t
    
    ' Constants
    Set iTEnums = mEnums.Clone
    iTEnums.Index = "Name"
    If iTEnums.RecordCount > 0 Then
        iTEnums.MoveFirst
        Do Until iTEnums.EOF
            Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
            If iRec.RecordCount > 0 Then
                Call TxtToHTML(iTEnums!Description, iTEnums!Name & " Enumeration")
                iRec.MoveFirst
                Do Until iRec.EOF
                    If (iRec!Description <> "") Then
                        Call TxtToHTML(Replace(iRec!Description, vbCrLf, vbTab), iRec!Name & " constant of " & iTEnums!Name & " Enumeration")
                    End If
                    iRec.MoveNext
                Loop
            End If
            iTEnums.MoveNext
        Loop
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If iDesc <> "" Then
        Call TxtToHTML(iDesc, "index")
    End If
    
    If UBound(mLinkErrors) > 0 Then
        Set iSc = New SmartConcat
        For c = 1 To UBound(mLinkErrors)
            iSc.AddString mLinkErrors(c) & vbCrLf
        Next c
        frmMessage.Message = "Found " & UBound(mLinkErrors) & " link error" & IIf(UBound(mLinkErrors) > 1, "s", "") & ":" & vbCrLf & iSc.GenerateCurrentString
        frmMessage.Show vbModal
    Else
        MsgBox "No link errors found.", vbInformation
    End If
End Sub

Private Sub mnuComponent_Click()
    UpdateData
End Sub

Private Sub mnuImport_Click()
    Dim iDlg As New cDlg
    Dim iTLI As TypeLibInfo
    Dim t As Long
    Dim iName As String
    Dim iClasses() As Long
    Dim iClasses_Names() As String
    Dim iClasses_DefaultInterface() As Long
    Dim iClasses_DefaultEventInterface() As Long
    Dim iControls() As Long
    Dim iControls_Names() As String
    Dim iControls_DefaultInterface() As Long
    Dim iControls_DefaultEventInterface() As Long
    Dim iEnums() As Long
    Dim iEnums_Names() As String
    Dim c As Long
    Dim iTypeInfo As TypeInfo
    Dim m As Long
    Dim iTMembers As Recordset
    Dim iStr As String
    Dim iProperties() As Long
    Dim iProperties_Names() As String
    Dim iProperties_ReadOnly() As Boolean
    Dim iProperties_HelpString() As String
    Dim iMethods() As Long
    Dim iEvents() As Long
    Dim iIk As InvokeKinds
    Dim i As Long
    Dim c2 As Long
    Dim iOptional As Boolean
    Dim iDefaultValue As String
    Dim iLng As Long
    Dim iMemberInfo As MemberInfo
    Dim iByRef As Boolean
    Dim iTypeName As String
    Dim iParamsInfo As String
    Dim iPos As Long
    Dim iRec As Recordset
    Dim iParamName As String
    Dim iRec2 As Recordset
    Dim iEditCurrent As Boolean
    Dim iRec3 As Recordset
    Dim iIDToFind As Long
    Dim iDoNotUpdateParamsInfo As Boolean
    Dim iDoNotUpdateShortDescription As Boolean
    
    If mClasses Is Nothing Then Exit Sub
    If (mClasses.RecordCount > 0) Or (mControls.RecordCount > 0) Or (mEnums.RecordCount > 0) Then
        If Not mClasses Is Nothing Then
            If (mControls.RecordCount > 0) Or (mClasses.RecordCount > 0) Or (mEnums.RecordCount > 0) Then
                If MsgBox("This will replace existent information, all short descriptions defined in the component will overwrite the ones of the database. Controls, classes and members (properties, etc) not found on the components will be hided, OK?" & vbCrLf & vbCrLf & "Note: enums not found are left, since some enums may not be exposed by the component, so if you changed enums names or removed some, you'll have to delete or rename them manually.", vbOKCancel + vbDefaultButton2) = vbCancel Then Exit Sub
            End If
        End If
    End If
    
    ReDim iClasses(0)
    ReDim iClasses_Names(0)
    ReDim iControls(0)
    ReDim iControls_Names(0)
    ReDim iEnums(0)
    ReDim iEnums_Names(0)
    
    iDlg.Filter = "ActiveX Controls (*.OCX)|*.OCX|ActiveX DLLs (*.DLL)|*.DLL|Type Libraries (*.TLB;*.OLB)|*.TLB;*.OLB|ActiveX Executables (*.EXE)|*.EXE|Any ActiveX File (*.OCX;*.DLL;*.TLB;*.OLB;*.EXE)|*.OCX;*.DLL;*.TLB;*.OLB;*.EXE|All Files (*.*)|*."
    iDlg.FilterIndex = 5
    iDlg.ShowOpen
    If iDlg.Canceled Then Exit Sub
        
    mSelectedType = entNone
    CurrentAction = ecaDefault
    ControlsEditZoneVisible = False
    On Error Resume Next
    Set iTLI = TLI.TypeLibInfoFromFile(iDlg.FileName)
    If Err.Number Then
        MsgBox "Error: " & Err.Description, vbExclamation
        Err.Clear
        Exit Sub
    End If
    
    mComponentVersion = iTLI.MajorVersion & "." & iTLI.MinorVersion
    mComponentReleaseDate = CDate(CLng(CDate(FileDateTime(iDlg.FileName))))
    
    mGeneral_Information.Seek "=", "ComponentVersion"
    If mGeneral_Information.NoMatch Then
        mGeneral_Information.AddNew
        mGeneral_Information!Name = "ComponentVersion"
    Else
        mGeneral_Information.Edit
    End If
    mGeneral_Information!Value = mComponentVersion
    mGeneral_Information.Update
    
    mGeneral_Information.Seek "=", "ComponentReleaseDate"
    If mGeneral_Information.NoMatch Then
        mGeneral_Information.AddNew
        mGeneral_Information!Name = "ComponentReleaseDate"
    Else
        mGeneral_Information.Edit
    End If
    mGeneral_Information!Value = CLng(mComponentReleaseDate)
    mGeneral_Information.Update
    
    On Error GoTo 0
    For t = 1 To iTLI.TypeInfos.Count
        If Not ((iTLI.TypeInfos(t).AttributeMask And 16) = 16) Then
            iName = iTLI.TypeInfos(t).Name
            If LCase(iTLI.TypeInfos(t).TypeKindString) = "enum" Then
                AddToList iEnums, t
                AddToList iEnums_Names, iName
            ElseIf LCase(iTLI.TypeInfos(t).TypeKindString) = "coclass" Then
                If (iTLI.TypeInfos(t).AttributeMask And 32) = 32 Then
                    AddToList iControls, t
                    AddToList iControls_Names, iName
                ElseIf Not ((iTLI.TypeInfos(t).DefaultInterface.Members.Count = 7) And (iTLI.TypeInfos(t).DefaultEventInterface Is Nothing)) Then ' Without members, without events, most probably a property bag
                    AddToList iClasses, t
                    AddToList iClasses_Names, iName
                End If
            End If
        End If
    Next t
    
    ' General information
    If iTLI.Name <> mComponentName Then
        If MsgBox("Replace Component name '" & mComponentName & "' with component name from file '" & iTLI.Name & "'?", vbYesNo) = vbYes Then
            mGeneral_Information.Seek "=", "ComponentName"
            If mGeneral_Information.NoMatch Then
                mGeneral_Information.AddNew
                mGeneral_Information!Name = "ComponentName"
            Else
                mGeneral_Information.Edit
            End If
            mGeneral_Information!Value = iTLI.Name
            mGeneral_Information.Update
            mComponentName = iTLI.Name
            Me.Caption = App.Title & " - " & mComponentName & " (" & GetFileName(mCurrentDBPath) & ")"
            LoadReportingOptions
        End If
    End If
    
    ReDim iClasses_DefaultInterface(UBound(iClasses))
    ReDim iClasses_DefaultEventInterface(UBound(iClasses))
    ReDim iControls_DefaultInterface(UBound(iControls))
    ReDim iControls_DefaultEventInterface(UBound(iControls))
    
    For c = 1 To UBound(iClasses)
        iClasses_DefaultInterface(c) = -1
        iClasses_DefaultEventInterface(c) = -1
    Next
    For c = 1 To UBound(iControls)
        iControls_DefaultInterface(c) = -1
        iControls_DefaultEventInterface(c) = -1
    Next
    
    On Error Resume Next
    For c = 1 To UBound(iClasses)
        Set iTypeInfo = iTLI.TypeInfos(iClasses(c))
        iClasses_DefaultInterface(c) = iTypeInfo.DefaultInterface.TypeInfoNumber + 1
        iClasses_DefaultEventInterface(c) = iTypeInfo.DefaultEventInterface.TypeInfoNumber + 1
    Next c
    For c = 1 To UBound(iControls)
        Set iTypeInfo = iTLI.TypeInfos(iControls(c))
        iControls_DefaultInterface(c) = iTypeInfo.DefaultInterface.TypeInfoNumber + 1
        iControls_DefaultEventInterface(c) = iTypeInfo.DefaultEventInterface.TypeInfoNumber + 1
    Next c
    On Error GoTo 0
    
    OrderVector iClasses_Names, iClasses, iClasses_DefaultInterface, iClasses_DefaultEventInterface
    OrderVector iControls_Names, iControls, iControls_DefaultInterface, iControls_DefaultEventInterface
    OrderVector iEnums_Names, iEnums
    
    ' Enums
    mEnums.Index = "Name"
    For c = 1 To UBound(iEnums)
        Set iTypeInfo = iTLI.TypeInfos(iEnums(c))
        iName = iTypeInfo.Name
        mEnums.Seek "=", iName
        If mEnums.NoMatch Then
            mEnums.AddNew
            If mNewEnumsOrderedByValue Then
                mEnums!Ordered_By_Value = True
            End If
            mEnums!Name = iName
        Else
            mEnums.Edit
        End If
        If mEnums!Description = "" Then
            If Not ((UBound(Split(iTypeInfo.HelpString, " ")) = 1) And (Right$(iTypeInfo.HelpString, 6) = "class.")) Then ' sometimes it seems that VB6 puts the name of the class where it is & " class."
                mEnums!Description = iTypeInfo.HelpString
            End If
        End If
        mEnums.Update
        mEnums.Bookmark = mEnums.LastModified
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & mEnums!Enum_ID & ") ORDER BY Name")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        For m = 1 To iTypeInfo.Members.Count
            iStr = iTypeInfo.Members(m).Name
            iTMembers.FindFirst "Name = '" & iStr & "'"
            If iTMembers.NoMatch Then
                iTMembers.AddNew
                iTMembers!Enum_ID = mEnums!Enum_ID
                iTMembers!Name = iStr
            Else
                iTMembers.Edit
            End If
            iTMembers!Auxiliary_Field = 0
            iTMembers!Value = iTypeInfo.Members(m).Value
            If iTMembers!Description = "" Then
                If iTypeInfo.Members(m).HelpString <> iTypeInfo.Members(m).Name Then
                    iTMembers!Description = iTypeInfo.Members(m).HelpString
                End If
            End If
            iTMembers.Update
        Next
        iTMembers.FindFirst "Auxiliary_Field = 1"
        Do Until iTMembers.NoMatch
            iTMembers.Delete
            iTMembers.FindFirst "Auxiliary_Field = 1"
        Loop
    Next c
    mEnums.Index = "PrimaryKey"

    mClasses.Index = "Name"
    mControls.Index = "Name"
    mProperties.Index = "Name"
    mMethods.Index = "Name"
    mEvents.Index = "Name"
    
    If mProperties.RecordCount > 0 Then
        mProperties.MoveFirst
        Do Until mProperties.EOF
            If mProperties!Auxiliary_Field <> 0 Then
                mProperties.Edit
                mProperties!Auxiliary_Field = 0
                mProperties.Update
                mProperties.Bookmark = mProperties.LastModified
            End If
            mProperties.MoveNext
        Loop
    End If
    If mMethods.RecordCount > 0 Then
        mMethods.MoveFirst
        Do Until mMethods.EOF
            If mMethods!Auxiliary_Field <> 0 Then
                mMethods.Edit
                mMethods!Auxiliary_Field = 0
                mMethods.Update
                mMethods.Bookmark = mMethods.LastModified
            End If
            mMethods.MoveNext
        Loop
    End If
    If mEvents.RecordCount > 0 Then
        mEvents.MoveFirst
        Do Until mEvents.EOF
            If mEvents!Auxiliary_Field <> 0 Then
                mEvents.Edit
                mEvents!Auxiliary_Field = 0
                mEvents.Update
                mEvents.Bookmark = mEvents.LastModified
            End If
            mEvents.MoveNext
        Loop
    End If
    
    
    ' Controls
    For c = 1 To UBound(iControls)
        Set iTypeInfo = iTLI.TypeInfos(iControls(c))
        iName = iTypeInfo.Name
        mControls.Seek "=", iName
        If mControls.NoMatch Then
            mControls.AddNew
            mControls!Name = iName
        Else
            mControls.Edit
        End If
        If iTypeInfo.HelpString <> "" Then
            mControls!Short_Description = iTypeInfo.HelpString
        End If
        mControls.Update
        mControls.Bookmark = mControls.LastModified
        
        ' Properties
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties WHERE (Control_ID = " & mControls!control_ID & ")")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        ' Methods
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods WHERE (Control_ID = " & mControls!control_ID & ")")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        ' Events
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Events WHERE (Control_ID = " & mControls!control_ID & ")")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        
        ReDim iProperties(0)
        ReDim iProperties_Names(0)
        ReDim iProperties_ReadOnly(0)
        ReDim iProperties_HelpString(0)
        ReDim iMethods(0)
        ReDim iEvents(0)
        
        If iControls_DefaultInterface(c) > -1 Then
            Set iTypeInfo = iTLI.TypeInfos(iControls_DefaultInterface(c))
        End If
        
        For m = 1 To iTypeInfo.Members.Count
            If ((iTypeInfo.Members(m).AttributeMask And FUNCFLAG_FHIDDEN) = 0) And ((iTypeInfo.Members(m).AttributeMask And FUNCFLAG_FRESTRICTED) = 0) Then ' Not hidden
                iIk = iTypeInfo.Members(m).InvokeKind
'                    If iIk = INVOKE_EVENTFUNC Then
                    'AddToList iEvents, m
                If iIk = INVOKE_FUNC Then
                 '   If (iTypeInfo.Members(m).ReturnType.VarType = VT_VOID) Or (iTypeInfo.Members(m).ReturnType.VarType = VT_HRESULT) Then
                    AddToList iMethods, m
                  '  Else
                   '     AddToList iFunctions, m
                    'End If
                ElseIf (iIk = INVOKE_PROPERTYGET) Or (iIk = INVOKE_PROPERTYPUT) Or (iIk = INVOKE_PROPERTYPUTREF) Then
                    i = IndexInList(iProperties_Names, iTypeInfo.Members(m).Name)
                    If i = -1 Then
                        AddToList iProperties_Names, iTypeInfo.Members(m).Name
                        i = UBound(iProperties_Names)
                        ReDim Preserve iProperties(i)
                        iProperties(i) = m
                        ReDim Preserve iProperties_ReadOnly(i)
                        ReDim Preserve iProperties_HelpString(i)
                        iProperties_ReadOnly(i) = True
                    End If
                    If iProperties_HelpString(i) = "" Then
                        iProperties_HelpString(i) = iTypeInfo.Members(m).HelpString
                    End If
                    If (iIk = INVOKE_PROPERTYPUT) Or (iIk = INVOKE_PROPERTYPUTREF) Then
                        iProperties_ReadOnly(i) = False
                    End If
                End If
            End If
        Next m
        
        If iControls_DefaultEventInterface(c) > -1 Then
            Set iTypeInfo = iTLI.TypeInfos(iControls_DefaultEventInterface(c))
            For m = 1 To iTypeInfo.Members.Count
                If ((iTypeInfo.Members(m).AttributeMask And FUNCFLAG_FHIDDEN) = 0) And ((iTypeInfo.Members(m).AttributeMask And FUNCFLAG_FRESTRICTED) = 0) Then ' Not hidden
                    iIk = iTypeInfo.Members(m).InvokeKind
                    If iIk = INVOKE_FUNC Then
                        AddToList iEvents, m
                    End If
                End If
            Next m
        End If
        
        If iControls_DefaultInterface(c) <> -1 Then
            Set iTypeInfo = iTLI.TypeInfos(iControls_DefaultInterface(c))
        Else
            Set iTypeInfo = iTLI.TypeInfos(iControls(c))
        End If
        
        ' Properties
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties WHERE (Control_ID = " & mControls!control_ID & ")")
        
        For m = 1 To UBound(iProperties)
            Set iMemberInfo = iTypeInfo.Members(iProperties(m))
            
            ' Params info:
            iStr = ""
            On Error Resume Next
            iStr = GetTypeName(iMemberInfo.ReturnType)
            On Error GoTo 0
            If iStr = "" Then iStr = "[unknown]"
            iStr = "Type: " & "<b>" & iStr & "</b>"
            If iProperties_ReadOnly(m) Then
                If iMemberInfo.MemberID = 0 Then
                    iStr = iStr & " (Default property, read only)"
                Else
                    iStr = iStr & " (Read only)"
                End If
            ElseIf iMemberInfo.MemberID = 0 Then
                iStr = iStr & " (Default property)"
            End If
            If (iMemberInfo.AttributeMask And 1024) <> 0 Then
                iPos = InStr(iStr, ")")
                If iPos > 0 Then
                    iStr = Left(iStr, iPos - 1) & ", not available at design time)"
                Else
                    iStr = iStr & " (Not available at design time)"
                End If
            End If
            iParamsInfo = iStr
            
            iStr = ""
            If iMemberInfo.Parameters.Count > 0 Then
                iStr = vbCrLf & "Additional parameter(s):" & vbCrLf
                For c2 = 1 To iMemberInfo.Parameters.Count
                    iParamName = "<b>" & iMemberInfo.Parameters(c2).Name & "</b>"
                    iLng = 28 - Len(iParamName)
                    If iLng < 1 Then iLng = 1
                    On Error Resume Next
                    iTypeName = GetTypeName(iMemberInfo.Parameters(c2).VarTypeInfo)
                    On Error GoTo 0
                    If iTypeName = "" Then iTypeName = "[unknown]"
                    iOptional = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOPT) <> 0
                    If iOptional Then
                        On Error Resume Next
                        iDefaultValue = ""
                        iDefaultValue = CStr(iMemberInfo.Parameters(c2).DefaultValue)
                        On Error GoTo 0
                        iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                        iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                        On Error Resume Next
                        If (iMemberInfo.Parameters(c2).VarTypeInfo.VarType = 0) Then
                            If (iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name <> "") Then
                                iDefaultValue = GetConstantName2(iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name, Val(iDefaultValue))
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    iByRef = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOUT) <> 0
                    iStr = iStr & "    " & IIf(iByRef, "In/Out", "In") & vbTab & IIf(iOptional, "Optional", "Required") & vbTab & iParamName & vbTab & "As" & vbTab & iTypeName & IIf(iOptional And (iDefaultValue <> "") And (Not iByRef), vbTab & "Default value: " & iDefaultValue, "") & vbCrLf
                Next c2
            End If
            If iStr <> "" Then
                iParamsInfo = iParamsInfo & vbCrLf & iStr
            End If
            ' End of Params info
            
            iStr = iMemberInfo.Name
            mProperties.Seek "=", iStr
            If mProperties.NoMatch Then
                mProperties.AddNew
                mProperties!Name = iStr
            Else
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties, Properties WHERE (Properties.Property_ID = Controls_Properties.Property_ID) AND (Control_ID = " & mControls!control_ID & ") AND (Properties.Name = '" & iStr & "')")
                iEditCurrent = False
                If Not iRec2.EOF Then
                    iRec2.MoveLast
                    If (iRec2.RecordCount = 1) And iRec2.Fields("Properties.Auxiliary_Field").Value = 0 Then
                        If iRec2!ID_Params_Info_Replaced <> 0 Then
                            iDoNotUpdateParamsInfo = True
                        End If
                        If iRec2!ID_Short_Description_Replaced <> 0 Then
                            iDoNotUpdateShortDescription = True
                        End If
                        mProperties.Index = "PrimaryKey"
                        mProperties.Seek "=", iRec2.Fields("Properties.Property_ID").Value
                        If mProperties.NoMatch Then Err.Raise 1236
                        iEditCurrent = True
                    End If
                End If
                If iEditCurrent Then
                    mProperties.Edit
                ElseIf ParamsInfoAreTheSame(mProperties!Params_Info, iParamsInfo) And ((mProperties!Short_Description = "") Or (mProperties!Short_Description = iProperties_HelpString(m)) Or (iProperties_HelpString(m) = "")) Then
                    mProperties.Edit
                Else
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Properties WHERE (Name = '" & iStr & "')")
                    iRec2.MoveFirst
                    Do Until iRec2.EOF
                        If (RemoveNAADT(iRec2!Params_Info) = RemoveNAADT(iParamsInfo)) And ((mProperties!Short_Description = "") Or (iRec2!Short_Description = iProperties_HelpString(m)) Or (iProperties_HelpString(m) = "")) Then
                            iIDToFind = iRec2!Property_ID
                            Set iRec3 = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties, Properties WHERE (Properties.Property_ID = Controls_Properties.Property_ID) AND (Control_ID = " & mControls!control_ID & ") AND (Properties.Name = '" & iStr & "')")
                            If iRec3.RecordCount = 1 Then
                                iRec3.MoveFirst
                                If iRec3!ID_Params_Info_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Properties.Property_ID").Value
                                    iDoNotUpdateParamsInfo = True
                                End If
                                If iRec3!ID_Short_Description_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Properties.Property_ID").Value
                                    iDoNotUpdateShortDescription = True
                                End If
                            End If
                            mProperties.Index = "PrimaryKey"
                            mProperties.Seek "=", iIDToFind
                            If mProperties.NoMatch Then Err.Raise 1236
                            mProperties.Edit
                            Exit Do
                        End If
                        iRec2.MoveNext
                    Loop
                    If iRec2.EOF Then
                        mProperties.AddNew
                        mProperties!Name = iStr
                    End If
                End If
            End If
            mProperties!Auxiliary_Field = 1
            If (iProperties_HelpString(m) <> "") And (Not iDoNotUpdateShortDescription) Then
                mProperties!Short_Description = iProperties_HelpString(m)
            End If
            
            If Not iDoNotUpdateParamsInfo Then
                mProperties!Params_Info = iParamsInfo
            End If
            
            mProperties.Update
            mProperties.Index = "Name"
            mProperties.Bookmark = mProperties.LastModified
            
            iTMembers.FindFirst "Property_ID = " & mProperties!Property_ID
            If iTMembers.NoMatch Then
                iTMembers.AddNew
                iTMembers!Property_ID = mProperties!Property_ID
                iTMembers!control_ID = mControls!control_ID
            Else
                iTMembers.Edit
            End If
            iTMembers!Auxiliary_Field = 0
            iTMembers.Update
        Next m
        
        iTMembers.FindFirst "Auxiliary_Field = 1"
        Do Until iTMembers.NoMatch
            iTMembers.Delete
            iTMembers.FindFirst "Auxiliary_Field = 1"
        Loop
        
        ' Methods
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods WHERE (Control_ID = " & mControls!control_ID & ")")
        
        For m = 1 To UBound(iMethods)
            Set iMemberInfo = iTypeInfo.Members(iMethods(m))
            
            ' Params info:
            iStr = ""
            On Error Resume Next
            iStr = GetTypeName(iMemberInfo.ReturnType)
            On Error GoTo 0
            If iStr = "" Then
                iParamsInfo = "Return Type unknown"
            Else
                If (iMemberInfo.ReturnType.VarType = VT_VOID) Or (iMemberInfo.ReturnType.VarType = VT_HRESULT) Then
                    iParamsInfo = "Return Type:" & vbTab & "None."
                Else
                    iStr = GetTypeName(iMemberInfo.ReturnType)
                    iStr = "Return Type:" & vbTab & "<b>" & iStr & "</b>"
                    iParamsInfo = iStr
                End If
            End If
            
            iStr = ""
            If iMemberInfo.Parameters.Count > 0 Then
                iStr = "Parameter(s):" & vbCrLf
                For c2 = 1 To iMemberInfo.Parameters.Count
                    iParamName = "<b>" & iMemberInfo.Parameters(c2).Name & "</b>"
                    iLng = 28 - Len(iParamName)
                    If iLng < 1 Then iLng = 1
                    On Error Resume Next
                    iTypeName = GetTypeName(iMemberInfo.Parameters(c2).VarTypeInfo)
                    On Error GoTo 0
                    If iTypeName = "" Then iTypeName = "[unknown]"
                    iOptional = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOPT) <> 0
                    If iOptional Then
                        On Error Resume Next
                        iDefaultValue = ""
                        iDefaultValue = CStr(iMemberInfo.Parameters(c2).DefaultValue)
                        On Error GoTo 0
                        iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                        iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                        On Error Resume Next
                        If (iMemberInfo.Parameters(c2).VarTypeInfo.VarType = 0) Then
                            If (iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name <> "") Then
                                iDefaultValue = GetConstantName2(iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name, Val(iDefaultValue))
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    
                    iByRef = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOUT) <> 0
                    iStr = iStr & "    " & IIf(iByRef, "In/Out", "In") & vbTab & IIf(iOptional, "Optional", "Required") & vbTab & iParamName & vbTab & "As" & vbTab & iTypeName & IIf(iOptional And (iDefaultValue <> "") And (Not iByRef), vbTab & "Default value: " & iDefaultValue, "") & vbCrLf
                Next c2
            Else
                iStr = "No parameters."
            End If
            If iStr <> "" Then
                iParamsInfo = iParamsInfo & vbCrLf & iStr
            Else
                iParamsInfo = iParamsInfo & vbCrLf & "This method has no parameters."
            End If
            ' End of Params Info
            
            iStr = iMemberInfo.Name
            mMethods.Seek "=", iStr
            If mMethods.NoMatch Then
                mMethods.AddNew
                mMethods!Name = iStr
            Else
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods, Methods WHERE (Methods.Method_ID = Controls_Methods.Method_ID) AND (Control_ID = " & mControls!control_ID & ") AND (Methods.Name = '" & iStr & "')")
                iEditCurrent = False
                If Not iRec2.EOF Then
                    iRec2.MoveLast
                    If (iRec2.RecordCount = 1) And iRec2.Fields("Methods.Auxiliary_Field").Value = 0 Then
                        If iRec2!ID_Params_Info_Replaced <> 0 Then
                            iDoNotUpdateParamsInfo = True
                        End If
                        If iRec2!ID_Short_Description_Replaced <> 0 Then
                            iDoNotUpdateShortDescription = True
                        End If
                        mMethods.Index = "PrimaryKey"
                        mMethods.Seek "=", iRec2.Fields("Methods.Method_ID").Value
                        If mMethods.NoMatch Then Err.Raise 1236
                        iEditCurrent = True
                    End If
                End If
                If iEditCurrent Then
                    mMethods.Edit
                ElseIf ParamsInfoAreTheSame(mMethods!Params_Info, iParamsInfo) And ((mMethods!Short_Description = "") Or (mMethods!Short_Description = iTypeInfo.HelpString) Or (iTypeInfo.HelpString = "")) Then
                    mMethods.Edit
                Else
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Methods WHERE (Name = '" & iStr & "')")
                    iRec2.MoveFirst
                    Do Until iRec2.EOF
                        If (iRec2!Params_Info = iParamsInfo) And ((iRec2!Short_Description = "") Or (iRec2!Short_Description = iMemberInfo.HelpString) Or (iMemberInfo.HelpString = "")) Then
                            iIDToFind = iRec2!Method_ID
                            Set iRec3 = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods, Methods WHERE (Methods.Method_ID = Controls_Methods.Method_ID) AND (Control_ID = " & mControls!control_ID & ") AND (Methods.Name = '" & iStr & "')")
                            If iRec3.RecordCount = 1 Then
                                iRec3.MoveFirst
                                If iRec3!ID_Params_Info_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Methods.Method_ID").Value
                                    iDoNotUpdateParamsInfo = True
                                End If
                                If iRec3!ID_Short_Description_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Methods.Method_ID").Value
                                    iDoNotUpdateShortDescription = True
                                End If
                            End If
                            mMethods.Index = "PrimaryKey"
                            mMethods.Seek "=", iIDToFind
                            If mMethods.NoMatch Then Err.Raise 1236
                            mMethods.Edit
                            Exit Do
                        End If
                        iRec2.MoveNext
                    Loop
                    If iRec2.EOF Then
                        mMethods.AddNew
                        mMethods!Name = iStr
                    End If
                End If
            End If
            mMethods!Auxiliary_Field = 1
            If (iMemberInfo.HelpString <> "") And (Not iDoNotUpdateShortDescription) Then
                mMethods!Short_Description = iMemberInfo.HelpString
            End If
            
            If Not iDoNotUpdateParamsInfo Then
                mMethods!Params_Info = iParamsInfo
            End If
            
            mMethods.Update
            mMethods.Index = "Name"
            mMethods.Bookmark = mMethods.LastModified
            
            iTMembers.FindFirst "Method_ID = " & mMethods!Method_ID
            If iTMembers.NoMatch Then
                iTMembers.AddNew
                iTMembers!Method_ID = mMethods!Method_ID
                iTMembers!control_ID = mControls!control_ID
            Else
                iTMembers.Edit
            End If
            iTMembers!Auxiliary_Field = 0
            iTMembers.Update
        Next m
        
        iTMembers.FindFirst "Auxiliary_Field = 1"
        Do Until iTMembers.NoMatch
            iTMembers.Delete
            iTMembers.FindFirst "Auxiliary_Field = 1"
        Loop
        
        ' Events
        If iControls_DefaultEventInterface(c) <> -1 Then
            Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Events WHERE (Control_ID = " & mControls!control_ID & ")")
            Set iTypeInfo = iTLI.TypeInfos(iControls_DefaultEventInterface(c))
            
            For m = 1 To UBound(iEvents)
                Set iMemberInfo = iTypeInfo.Members(iEvents(m))
                
                ' Params info:
                iStr = ""
                If iMemberInfo.Parameters.Count > 0 Then
                    iStr = "Parameter(s):" & vbCrLf
                    For c2 = 1 To iMemberInfo.Parameters.Count
                        iParamName = "<b>" & iMemberInfo.Parameters(c2).Name & "</b>"
                        iLng = 28 - Len(iParamName)
                        If iLng < 1 Then iLng = 1
                        On Error Resume Next
                        iTypeName = GetTypeName(iMemberInfo.Parameters(c2).VarTypeInfo)
                        On Error GoTo 0
                        If iTypeName = "" Then iTypeName = "[unknown]"
                        iOptional = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOPT) <> 0
                        If iOptional Then
                            On Error Resume Next
                            iDefaultValue = ""
                            iDefaultValue = CStr(iMemberInfo.Parameters(c2).DefaultValue)
                            On Error GoTo 0
                            iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                            iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                            On Error Resume Next
                            If (iMemberInfo.Parameters(c2).VarTypeInfo.VarType = 0) Then
                                If (iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name <> "") Then
                                    iDefaultValue = GetConstantName2(iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name, Val(iDefaultValue))
                                End If
                            End If
                            On Error GoTo 0
                        End If
                        iByRef = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOUT) <> 0
                        iStr = iStr & "    " & IIf(iByRef, "Returns value", "Info") & vbTab & iParamName & vbTab & "As" & vbTab & iTypeName & IIf(iOptional And (iDefaultValue <> "") And (Not iByRef), vbTab & "Default value: " & iDefaultValue, "") & vbCrLf
                    Next c2
                Else
                    iStr = "No parameters."
                End If
                If iStr <> "" Then
                    iParamsInfo = iStr
                Else
                    iParamsInfo = "This Event has no parameters."
                End If
                ' End of Params Info
                
                iStr = iMemberInfo.Name
                mEvents.Seek "=", iStr
                If mEvents.NoMatch Then
                    mEvents.AddNew
                    mEvents!Name = iStr
                Else
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Controls_Events, Events WHERE (Events.Event_ID = Controls_Events.Event_ID) AND (Control_ID = " & mControls!control_ID & ") AND (Events.Name = '" & iStr & "')")
                    iEditCurrent = False
                    If Not iRec2.EOF Then
                        iRec2.MoveLast
                        If (iRec2.RecordCount = 1) And iRec2.Fields("Events.Auxiliary_Field").Value = 0 Then
                            If iRec2!ID_Params_Info_Replaced <> 0 Then
                                iDoNotUpdateParamsInfo = True
                            End If
                            If iRec2!ID_Short_Description_Replaced <> 0 Then
                                iDoNotUpdateShortDescription = True
                            End If
                            mEvents.Index = "PrimaryKey"
                            mEvents.Seek "=", iRec2.Fields("Events.Event_ID").Value
                            If mEvents.NoMatch Then Err.Raise 1236
                            iEditCurrent = True
                        End If
                    End If
                    If iEditCurrent Then
                        mEvents.Edit
                    ElseIf ParamsInfoAreTheSame(mEvents!Params_Info, iParamsInfo) And ((mEvents!Short_Description = "") Or (mEvents!Short_Description = iTypeInfo.HelpString) Or (iTypeInfo.HelpString = "")) Then
                        mEvents.Edit
                    Else
                        Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Events WHERE (Name = '" & iStr & "')")
                        iRec2.MoveFirst
                        Do Until iRec2.EOF
                            If (iRec2!Params_Info = iParamsInfo) And ((iRec2!Short_Description = "") Or (iRec2!Short_Description = iMemberInfo.HelpString) Or (iMemberInfo.HelpString = "")) Then
                                iIDToFind = iRec2!Event_ID
                                Set iRec3 = mDatabase.OpenRecordset("SELECT * FROM Controls_Events, Events WHERE (Events.Event_ID = Controls_Events.Event_ID) AND (Control_ID = " & mControls!control_ID & ") AND (Events.Name = '" & iStr & "')")
                                If iRec3.RecordCount = 1 Then
                                    iRec3.MoveFirst
                                    If iRec3!ID_Params_Info_Replaced <> 0 Then
                                        iIDToFind = iRec3.Fields("Events.Event_ID").Value
                                        iDoNotUpdateParamsInfo = True
                                    End If
                                    If iRec3!ID_Short_Description_Replaced <> 0 Then
                                        iIDToFind = iRec3.Fields("Events.Event_ID").Value
                                        iDoNotUpdateShortDescription = True
                                    End If
                                End If
                                mEvents.Index = "PrimaryKey"
                                mEvents.Seek "=", iIDToFind
                                If mEvents.NoMatch Then Err.Raise 1236
                                mEvents.Edit
                                Exit Do
                            End If
                            iRec2.MoveNext
                        Loop
                        If iRec2.EOF Then
                            mEvents.AddNew
                            mEvents!Name = iStr
                        End If
                    End If
                End If
                mEvents!Auxiliary_Field = 1
                If (iMemberInfo.HelpString <> "") And (Not iDoNotUpdateShortDescription) Then
                    mEvents!Short_Description = iMemberInfo.HelpString
                End If
                
                If Not iDoNotUpdateParamsInfo Then
                    mEvents!Params_Info = iParamsInfo
                End If
                
                mEvents.Update
                mEvents.Index = "Name"
                mEvents.Bookmark = mEvents.LastModified
                
                iTMembers.FindFirst "Event_ID = " & mEvents!Event_ID
                If iTMembers.NoMatch Then
                    iTMembers.AddNew
                    iTMembers!Event_ID = mEvents!Event_ID
                    iTMembers!control_ID = mControls!control_ID
                Else
                    iTMembers.Edit
                End If
                iTMembers!Auxiliary_Field = 0
                iTMembers.Update
            Next m
            
            iTMembers.FindFirst "Auxiliary_Field = 1"
            Do Until iTMembers.NoMatch
                iTMembers.Delete
                iTMembers.FindFirst "Auxiliary_Field = 1"
            Loop
        End If
        
    Next c
    
    ' Classes
    For c = 1 To UBound(iClasses)
        Set iTypeInfo = iTLI.TypeInfos(iClasses(c))
        iName = iTypeInfo.Name
        mClasses.Seek "=", iName
        If mClasses.NoMatch Then
            mClasses.AddNew
            mClasses!Name = iName
        Else
            mClasses.Edit
        End If
        If iTypeInfo.HelpString <> "" Then
            mClasses!Short_Description = iTypeInfo.HelpString
        End If
        mClasses.Update
        mClasses.Bookmark = mClasses.LastModified
        
        ' Properties
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties WHERE (Class_ID = " & mClasses!Class_ID & ")")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        ' Methods
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods WHERE (Class_ID = " & mClasses!Class_ID & ")")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        ' Events
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Events WHERE (Class_ID = " & mClasses!Class_ID & ")")
        If iTMembers.RecordCount > 0 Then
            iTMembers.MoveFirst
            Do Until iTMembers.EOF
                If iTMembers!Auxiliary_Field <> 1 Then
                    iTMembers.Edit
                    iTMembers!Auxiliary_Field = 1
                    iTMembers.Update
                    iTMembers.Bookmark = iTMembers.LastModified
                End If
                iTMembers.MoveNext
            Loop
        End If
        
        ReDim iProperties(0)
        ReDim iProperties_Names(0)
        ReDim iProperties_ReadOnly(0)
        ReDim iProperties_HelpString(0)
        ReDim iMethods(0)
        ReDim iEvents(0)
        
        If iClasses_DefaultInterface(c) > -1 Then
            Set iTypeInfo = iTLI.TypeInfos(iClasses_DefaultInterface(c))
        End If
        
        For m = 1 To iTypeInfo.Members.Count
            If iTypeInfo.Members(m).AttributeMask = 0 Then  ' Not hidden
                iIk = iTypeInfo.Members(m).InvokeKind
                If iIk = INVOKE_FUNC Then
                    AddToList iMethods, m
                ElseIf (iIk = INVOKE_PROPERTYGET) Or (iIk = INVOKE_PROPERTYPUT) Or (iIk = INVOKE_PROPERTYPUTREF) Then
                    i = IndexInList(iProperties_Names, iTypeInfo.Members(m).Name)
                    If i = -1 Then
                        AddToList iProperties_Names, iTypeInfo.Members(m).Name
                        i = UBound(iProperties_Names)
                        ReDim Preserve iProperties(i)
                        iProperties(i) = m
                        ReDim Preserve iProperties_ReadOnly(i)
                        ReDim Preserve iProperties_HelpString(i)
                        iProperties_ReadOnly(i) = True
                    End If
                    If iProperties_HelpString(i) = "" Then
                        iProperties_HelpString(i) = iTypeInfo.Members(m).HelpString
                    End If
                    If (iIk = INVOKE_PROPERTYPUT) Or (iIk = INVOKE_PROPERTYPUTREF) Then
                        iProperties_ReadOnly(i) = False
                    End If
                End If
            End If
        Next m
        
        If iClasses_DefaultEventInterface(c) > -1 Then
            Set iTypeInfo = iTLI.TypeInfos(iClasses_DefaultEventInterface(c))
            For m = 1 To iTypeInfo.Members.Count
                If ((iTypeInfo.Members(m).AttributeMask And FUNCFLAG_FHIDDEN) = 0) And ((iTypeInfo.Members(m).AttributeMask And FUNCFLAG_FRESTRICTED) = 0) Then ' Not hidden
                    iIk = iTypeInfo.Members(m).InvokeKind
                    If iIk = INVOKE_FUNC Then
                        AddToList iEvents, m
                    End If
                End If
            Next m
        End If
        
        If iClasses_DefaultInterface(c) <> -1 Then
            Set iTypeInfo = iTLI.TypeInfos(iClasses_DefaultInterface(c))
        Else
            Set iTypeInfo = iTLI.TypeInfos(iClasses(c))
        End If
        
        ' Properties
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties WHERE (Class_ID = " & mClasses!Class_ID & ")")
        
        For m = 1 To UBound(iProperties)
            Set iMemberInfo = iTypeInfo.Members(iProperties(m))
            
            ' Params info:
            iStr = ""
            On Error Resume Next
            iStr = GetTypeName(iMemberInfo.ReturnType)
            On Error GoTo 0
            If iStr = "" Then iStr = "[unknown]"
            iStr = "Type: " & "<b>" & iStr & "</b>"
            If iProperties_ReadOnly(m) Then
                If iMemberInfo.MemberID = 0 Then
                    iStr = iStr & " (Default property, read only)"
                Else
                    iStr = iStr & " (Read only)"
                End If
            ElseIf iMemberInfo.MemberID = 0 Then
                iStr = iStr & " (Default property)"
            End If
            iParamsInfo = iStr
            
            iStr = ""
            If iMemberInfo.Parameters.Count > 0 Then
                iStr = vbCrLf & "Additional parameter(s):" & vbCrLf
                For c2 = 1 To iMemberInfo.Parameters.Count
                    iParamName = "<b>" & iMemberInfo.Parameters(c2).Name & "</b>"
                    iLng = 28 - Len(iParamName)
                    If iLng < 1 Then iLng = 1
                    On Error Resume Next
                    iTypeName = GetTypeName(iMemberInfo.Parameters(c2).VarTypeInfo)
                    On Error GoTo 0
                    If iTypeName = "" Then iTypeName = "[unknown]"
                    iOptional = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOPT) <> 0
                    If iOptional Then
                        On Error Resume Next
                        iDefaultValue = ""
                        iDefaultValue = CStr(iMemberInfo.Parameters(c2).DefaultValue)
                        On Error GoTo 0
                        iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                        iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                        On Error Resume Next
                        If (iMemberInfo.Parameters(c2).VarTypeInfo.VarType = 0) Then
                            If (iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name <> "") Then
                                iDefaultValue = GetConstantName2(iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name, Val(iDefaultValue))
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    iByRef = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOUT) <> 0
                    iStr = iStr & "    " & IIf(iByRef, "In/Out", "In") & vbTab & IIf(iOptional, "Optional", "Required") & vbTab & iParamName & vbTab & "As" & vbTab & iTypeName & IIf(iOptional And (iDefaultValue <> "") And (Not iByRef), vbTab & "Default value: " & iDefaultValue, "") & vbCrLf
                Next c2
            End If
            If iStr <> "" Then
                iParamsInfo = iParamsInfo & vbCrLf & iStr
            End If
            ' End of Params info
            
            iStr = iMemberInfo.Name
            mProperties.Seek "=", iStr
            If mProperties.NoMatch Then
                mProperties.AddNew
                mProperties!Name = iStr
            Else
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties, Properties WHERE (Properties.Property_ID = Classes_Properties.Property_ID) AND (Class_ID = " & mClasses!Class_ID & ") AND (Properties.Name = '" & iStr & "')")
                iEditCurrent = False
                If Not iRec2.EOF Then
                    iRec2.MoveLast
                    If (iRec2.RecordCount = 1) And iRec2.Fields("Properties.Auxiliary_Field").Value = 0 Then
                        If iRec2!ID_Params_Info_Replaced <> 0 Then
                            iDoNotUpdateParamsInfo = True
                        End If
                        If iRec2!ID_Short_Description_Replaced <> 0 Then
                            iDoNotUpdateShortDescription = True
                        End If
                        mProperties.Index = "PrimaryKey"
                        mProperties.Seek "=", iRec2.Fields("Properties.Property_ID").Value
                        If mProperties.NoMatch Then Err.Raise 1236
                        iEditCurrent = True
                    End If
                End If
                If iEditCurrent Then
                    mProperties.Edit
                ElseIf ParamsInfoAreTheSame(mProperties!Params_Info, iParamsInfo) And ((mProperties!Short_Description = "") Or (mProperties!Short_Description = iProperties_HelpString(m)) Or (iProperties_HelpString(m) = "")) Then
                    mProperties.Edit
                Else
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Properties WHERE (Name = '" & iStr & "')")
                    iRec2.MoveFirst
                    Do Until iRec2.EOF
                        If (RemoveNAADT(iRec2!Params_Info) = RemoveNAADT(iParamsInfo)) And ((mProperties!Short_Description = "") Or (iRec2!Short_Description = iProperties_HelpString(m)) Or (iProperties_HelpString(m) = "")) Then
                            iIDToFind = iRec2!Property_ID
                            Set iRec3 = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties, Properties WHERE (Properties.Property_ID = Classes_Properties.Property_ID) AND (Class_ID = " & mClasses!Class_ID & ") AND (Properties.Name = '" & iStr & "')")
                            If iRec3.RecordCount = 1 Then
                                iRec3.MoveFirst
                                If iRec3!ID_Params_Info_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Properties.Property_ID").Value
                                    iDoNotUpdateParamsInfo = True
                                End If
                                If iRec3!ID_Short_Description_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Properties.Property_ID").Value
                                    iDoNotUpdateShortDescription = True
                                End If
                            End If
                            mProperties.Index = "PrimaryKey"
                            mProperties.Seek "=", iIDToFind
                            If mProperties.NoMatch Then Err.Raise 1236
                            mProperties.Edit
                            Exit Do
                        End If
                        iRec2.MoveNext
                    Loop
                    If iRec2.EOF Then
                        mProperties.AddNew
                        mProperties!Name = iStr
                    End If
                End If
            End If
            mProperties!Auxiliary_Field = 1
            If (iProperties_HelpString(m) <> "") And (Not iDoNotUpdateShortDescription) Then
                mProperties!Short_Description = iProperties_HelpString(m)
            End If
            
            If Not iDoNotUpdateParamsInfo Then
                mProperties!Params_Info = iParamsInfo
            End If
            
            mProperties.Update
            mProperties.Index = "Name"
            mProperties.Bookmark = mProperties.LastModified
            
            iTMembers.FindFirst "Property_ID = " & mProperties!Property_ID
            If iTMembers.NoMatch Then
                iTMembers.AddNew
                iTMembers!Property_ID = mProperties!Property_ID
                iTMembers!Class_ID = mClasses!Class_ID
            Else
                iTMembers.Edit
            End If
            iTMembers!Auxiliary_Field = 0
            iTMembers.Update
        Next m
        
        iTMembers.FindFirst "Auxiliary_Field = 1"
        Do Until iTMembers.NoMatch
            iTMembers.Delete
            iTMembers.FindFirst "Auxiliary_Field = 1"
        Loop
        
        ' Methods
        Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods WHERE (Class_ID = " & mClasses!Class_ID & ")")
        
        For m = 1 To UBound(iMethods)
            Set iMemberInfo = iTypeInfo.Members(iMethods(m))
            
            ' Params info:
            iStr = ""
            On Error Resume Next
            iStr = GetTypeName(iMemberInfo.ReturnType)
            On Error GoTo 0
            If iStr = "" Then
                iParamsInfo = "Return Type unknown"
            Else
                If (iMemberInfo.ReturnType.VarType = VT_VOID) Or (iMemberInfo.ReturnType.VarType = VT_HRESULT) Then
                    iParamsInfo = "Return Type:" & vbTab & "None."
                Else
                    iStr = GetTypeName(iMemberInfo.ReturnType)
                    iStr = "Return Type:" & vbTab & "<b>" & iStr & "</b>"
                    iParamsInfo = iStr
                End If
            End If
            
            iStr = ""
            If iMemberInfo.Parameters.Count > 0 Then
                iStr = "Parameter(s):" & vbCrLf
                For c2 = 1 To iMemberInfo.Parameters.Count
                    iParamName = "<b>" & iMemberInfo.Parameters(c2).Name & "</b>"
                    iLng = 28 - Len(iParamName)
                    If iLng < 1 Then iLng = 1
                    iTypeName = ""
                    On Error Resume Next
                    iTypeName = GetTypeName(iMemberInfo.Parameters(c2).VarTypeInfo)
                    On Error GoTo 0
                    If iTypeName = "" Then iTypeName = "[unknown]"
                    iOptional = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOPT) <> 0
                    If iOptional Then
                        On Error Resume Next
                        iDefaultValue = ""
                        iDefaultValue = CStr(iMemberInfo.Parameters(c2).DefaultValue)
                        On Error GoTo 0
                        iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                        iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                        On Error Resume Next
                        If (iMemberInfo.Parameters(c2).VarTypeInfo.VarType = 0) Then
                            If (iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name <> "") Then
                                iDefaultValue = GetConstantName2(iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name, Val(iDefaultValue))
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    iByRef = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOUT) <> 0
                    iStr = iStr & "    " & IIf(iByRef, "In/Out", "In") & vbTab & IIf(iOptional, "Optional", "Required") & vbTab & iParamName & vbTab & "As" & vbTab & iTypeName & IIf(iOptional And (iDefaultValue <> "") And (Not iByRef), vbTab & "Default value: " & iDefaultValue, "") & vbCrLf
                Next c2
            Else
                iStr = "No parameters."
            End If
            If iStr <> "" Then
                iParamsInfo = iParamsInfo & vbCrLf & iStr
            Else
                iParamsInfo = iParamsInfo & vbCrLf & "This method has no parameters."
            End If
            ' End of Params Info
            
            iStr = iMemberInfo.Name
            mMethods.Seek "=", iStr
            If mMethods.NoMatch Then
                mMethods.AddNew
                mMethods!Name = iStr
            Else
                Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods, Methods WHERE (Methods.Method_ID = Classes_Methods.Method_ID) AND (Class_ID = " & mClasses!Class_ID & ") AND (Methods.Name = '" & iStr & "')")
                iEditCurrent = False
                If Not iRec2.EOF Then
                    iRec2.MoveLast
                    If (iRec2.RecordCount = 1) And iRec2.Fields("Methods.Auxiliary_Field").Value = 0 Then
                        If iRec2!ID_Params_Info_Replaced <> 0 Then
                            iDoNotUpdateParamsInfo = True
                        End If
                        If iRec2!ID_Short_Description_Replaced <> 0 Then
                            iDoNotUpdateShortDescription = True
                        End If
                        mMethods.Index = "PrimaryKey"
                        mMethods.Seek "=", iRec2.Fields("Methods.Method_ID").Value
                        If mMethods.NoMatch Then Err.Raise 1236
                        iEditCurrent = True
                    End If
                End If
                If iEditCurrent Then
                    mMethods.Edit
                ElseIf ParamsInfoAreTheSame(mMethods!Params_Info, iParamsInfo) And ((mMethods!Short_Description = "") Or (mMethods!Short_Description = iTypeInfo.HelpString) Or (iTypeInfo.HelpString = "")) Then
                    mMethods.Edit
                Else
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Methods WHERE (Name = '" & iStr & "')")
                    iRec2.MoveFirst
                    Do Until iRec2.EOF
                        If (iRec2!Params_Info = iParamsInfo) And ((iRec2!Short_Description = "") Or (iRec2!Short_Description = iMemberInfo.HelpString) Or (iMemberInfo.HelpString = "")) Then
                            iIDToFind = iRec2!Method_ID
                            Set iRec3 = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods, Methods WHERE (Methods.Method_ID = Classes_Methods.Method_ID) AND (Class_ID = " & mClasses!Class_ID & ") AND (Methods.Name = '" & iStr & "')")
                            If iRec3.RecordCount = 1 Then
                                iRec3.MoveFirst
                                If iRec3!ID_Params_Info_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Methods.Method_ID").Value
                                    iDoNotUpdateParamsInfo = True
                                End If
                                If iRec3!ID_Short_Description_Replaced <> 0 Then
                                    iIDToFind = iRec3.Fields("Methods.Method_ID").Value
                                    iDoNotUpdateShortDescription = True
                                End If
                            End If
                            mMethods.Index = "PrimaryKey"
                            mMethods.Seek "=", iIDToFind
                            If mMethods.NoMatch Then Err.Raise 1236
                            mMethods.Edit
                            Exit Do
                        End If
                        iRec2.MoveNext
                    Loop
                    If iRec2.EOF Then
                        mMethods.AddNew
                        mMethods!Name = iStr
                    End If
                End If
            End If
            mMethods!Auxiliary_Field = 1
            If (iMemberInfo.HelpString <> "") And (Not iDoNotUpdateShortDescription) Then
                mMethods!Short_Description = iMemberInfo.HelpString
            End If
            
            If Not iDoNotUpdateParamsInfo Then
                mMethods!Params_Info = iParamsInfo
            End If
            
            mMethods.Update
            mMethods.Index = "Name"
            mMethods.Bookmark = mMethods.LastModified
            
            iTMembers.FindFirst "Method_ID = " & mMethods!Method_ID
            If iTMembers.NoMatch Then
                iTMembers.AddNew
                iTMembers!Method_ID = mMethods!Method_ID
                iTMembers!Class_ID = mClasses!Class_ID
            Else
                iTMembers.Edit
            End If
            iTMembers!Auxiliary_Field = 0
            iTMembers.Update
        Next m
        
        iTMembers.FindFirst "Auxiliary_Field = 1"
        Do Until iTMembers.NoMatch
            iTMembers.Delete
            iTMembers.FindFirst "Auxiliary_Field = 1"
        Loop
        
        ' Events
        If iClasses_DefaultEventInterface(c) <> -1 Then
            Set iTMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Events WHERE (Class_ID = " & mClasses!Class_ID & ")")
            Set iTypeInfo = iTLI.TypeInfos(iClasses_DefaultEventInterface(c))
            
            For m = 1 To UBound(iEvents)
                Set iMemberInfo = iTypeInfo.Members(iEvents(m))
                
                ' Params info:
                iStr = ""
                If iMemberInfo.Parameters.Count > 0 Then
                    iStr = "Parameter(s):" & vbCrLf
                    For c2 = 1 To iMemberInfo.Parameters.Count
                        iParamName = "<b>" & iMemberInfo.Parameters(c2).Name & "</b>"
                        iLng = 28 - Len(iParamName)
                        If iLng < 1 Then iLng = 1
                        iTypeName = ""
                        On Error Resume Next
                        iTypeName = GetTypeName(iMemberInfo.Parameters(c2).VarTypeInfo)
                        On Error GoTo 0
                        If iTypeName = "" Then iTypeName = "[unknown]"
                        iOptional = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOPT) <> 0
                        If iOptional Then
                            On Error Resume Next
                            iDefaultValue = ""
                            iDefaultValue = CStr(iMemberInfo.Parameters(c2).DefaultValue)
                            On Error GoTo 0
                            iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                            iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                            On Error Resume Next
                            If (iMemberInfo.Parameters(c2).VarTypeInfo.VarType = 0) Then
                                If (iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name <> "") Then
                                    iDefaultValue = GetConstantName2(iMemberInfo.Parameters(c2).VarTypeInfo.TypeInfo.Name, Val(iDefaultValue))
                                End If
                            End If
                            On Error GoTo 0
                        End If
                        iByRef = (iMemberInfo.Parameters(c2).Flags And PARAMFLAG_FOUT) <> 0
                        iStr = iStr & "    " & IIf(iByRef, "Returns value", "Info") & vbTab & iParamName & vbTab & "As" & vbTab & iTypeName & IIf(iOptional And (iDefaultValue <> "") And (Not iByRef), vbTab & "Default value: " & iDefaultValue, "") & vbCrLf
                    Next c2
                Else
                    iStr = "No parameters."
                End If
                If iStr <> "" Then
                    iParamsInfo = iStr
                Else
                    iParamsInfo = "This Event has no parameters."
                End If
                ' End of Params Info
                
                iStr = iMemberInfo.Name
                mEvents.Seek "=", iStr
                iDoNotUpdateParamsInfo = False
                iDoNotUpdateShortDescription = False
                If mEvents.NoMatch Then
                    mEvents.AddNew
                    mEvents!Name = iStr
                Else
                    Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Classes_Events, Events WHERE (Events.Event_ID = Classes_Events.Event_ID) AND (Class_ID = " & mClasses!Class_ID & ") AND (Events.Name = '" & iStr & "')")
                    iEditCurrent = False
                    If Not iRec2.EOF Then
                        iRec2.MoveLast
                        If (iRec2.RecordCount = 1) And iRec2.Fields("Events.Auxiliary_Field").Value = 0 Then
                            If iRec2!ID_Params_Info_Replaced <> 0 Then
                                iDoNotUpdateParamsInfo = True
                            End If
                            If iRec2!ID_Short_Description_Replaced <> 0 Then
                                iDoNotUpdateShortDescription = True
                            End If
                            mEvents.Index = "PrimaryKey"
                            mEvents.Seek "=", iRec2.Fields("Events.Event_ID").Value
                            If mEvents.NoMatch Then Err.Raise 1236
                            iEditCurrent = True
                        End If
                    End If
                    If iEditCurrent Then
                        mEvents.Edit
                    ElseIf ParamsInfoAreTheSame(mEvents!Params_Info, iParamsInfo) And ((mEvents!Short_Description = "") Or (mEvents!Short_Description = iTypeInfo.HelpString) Or (iTypeInfo.HelpString = "")) Then
                        mEvents.Edit
                    Else
                        Set iRec2 = mDatabase.OpenRecordset("SELECT * FROM Events WHERE (Name = '" & iStr & "')")
                        iRec2.MoveFirst
                        Do Until iRec2.EOF
                            If (iRec2!Params_Info = iParamsInfo) And ((iRec2!Short_Description = "") Or (iRec2!Short_Description = iMemberInfo.HelpString) Or (iMemberInfo.HelpString = "")) Then
                                iIDToFind = iRec2!Event_ID
                                Set iRec3 = mDatabase.OpenRecordset("SELECT * FROM Classes_Events, Events WHERE (Events.Event_ID = Classes_Events.Event_ID) AND (Class_ID = " & mClasses!Class_ID & ") AND (Events.Name = '" & iStr & "')")
                                If iRec3.RecordCount = 1 Then
                                    iRec3.MoveFirst
                                    If iRec3!ID_Params_Info_Replaced <> 0 Then
                                        iIDToFind = iRec3.Fields("Events.Event_ID").Value
                                        iDoNotUpdateParamsInfo = True
                                    End If
                                    If iRec3!ID_Short_Description_Replaced <> 0 Then
                                        iIDToFind = iRec3.Fields("Events.Event_ID").Value
                                        iDoNotUpdateShortDescription = True
                                    End If
                                End If
                                mEvents.Index = "PrimaryKey"
                                mEvents.Seek "=", iIDToFind
                                If mEvents.NoMatch Then Err.Raise 1236
                                mEvents.Edit
                                Exit Do
                            End If
                            iRec2.MoveNext
                        Loop
                        If iRec2.EOF Then
                            mEvents.AddNew
                            mEvents!Name = iStr
                        End If
                    End If
                End If
                mEvents!Auxiliary_Field = 1
                If (iMemberInfo.HelpString <> "") And (Not iDoNotUpdateShortDescription) Then
                    mEvents!Short_Description = iMemberInfo.HelpString
                End If
                
                If Not iDoNotUpdateParamsInfo Then
                    mEvents!Params_Info = iParamsInfo
                End If
                
                mEvents.Update
                mEvents.Index = "Name"
                mEvents.Bookmark = mEvents.LastModified
                
                iTMembers.FindFirst "Event_ID = " & mEvents!Event_ID
                If iTMembers.NoMatch Then
                    iTMembers.AddNew
                    iTMembers!Event_ID = mEvents!Event_ID
                    iTMembers!Class_ID = mClasses!Class_ID
                Else
                    iTMembers.Edit
                End If
                iTMembers!Auxiliary_Field = 0
                iTMembers.Update
            Next m
            
            iTMembers.FindFirst "Auxiliary_Field = 1"
            Do Until iTMembers.NoMatch
                iTMembers.Delete
                iTMembers.FindFirst "Auxiliary_Field = 1"
            Loop
        End If
        
    Next c
    
    mClasses.Index = "PrimaryKey"
    mControls.Index = "PrimaryKey"
    mProperties.Index = "PrimaryKey"
    mMethods.Index = "PrimaryKey"
    mEvents.Index = "PrimaryKey"

    SaveSettingBase "General", "FileImported", "1"
    mFileImported = True
    ShowTree False
    SelectCurrentNode
    
End Sub

Private Sub mnuListOrphanMembers_Click()
    Dim c As Long
    Dim iStr As String
    Dim iRec As Recordset
    Dim iPos As Long
    Dim iCount As Long
    Dim iList() As String
    
    ReDim iList(0)
    For c = 1 To 3
        Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(c) & " WHERE (Auxiliary_Field = 0)")
        If iRec.RecordCount > 0 Then
            iRec.MoveLast
            If iStr = "" Then
                iStr = "There are "
            Else
                iStr = iStr & ", "
            End If
            iStr = iStr & iRec.RecordCount & " " & IIf(iRec.RecordCount = 1, mMemberType_s(c), mMemberType_p(c))
            iCount = iCount + iRec.RecordCount
            
            iRec.MoveFirst
            Do Until iRec.EOF
                AddToList iList, iRec!Name & " " & LCase$(mMemberType_s(c))
                iRec.MoveNext
            Loop
        End If
    Next c
    iPos = InStrRev(iStr, ", ")
    If iPos > 0 Then
        iStr = Left$(iStr, iPos - 1) & " and " & Mid$(iStr, iPos + 2)
    End If
    If iStr = "" Then
        MsgBox "No orphan members found.", vbInformation
    Else
        If iCount = 1 Then iStr = Replace$(iStr, "There are ", "There is ")
        iStr = iStr & ":" & vbCrLf & vbCrLf
        For c = 1 To UBound(iList)
            iStr = iStr & iList(c) & vbCrLf
        Next c
        
        frmMessage.Message = iStr
        frmMessage.lblNote.Caption = "Orphan members are properties, methods or events that there were before a component file was imported but that is not currently present, they are not used by any control or class. It could had been an old property or if you renamed a property without renaming it first here it is left orphan."
        frmMessage.Show vbModal
    End If
End Sub

Private Sub mnuNewEnumConstant_Click()
    CurrentAction = ecaAddConstant
End Sub

Private Sub mnuNewClass_Click()
    CurrentAction = ecaAddClass
End Sub

Private Sub mnuNewClass2_Click()
    mnuNewClass_Click
End Sub

Private Sub mnuNewControl_Click()
    CurrentAction = ecaAddControl
End Sub

Private Sub mnuNewControl2_Click()
    mnuNewControl_Click
End Sub

Private Sub mnuNewEnum_Click()
    CurrentAction = ecaAddEnum
End Sub

Private Sub mnuNewEnum2_Click()
    mnuNewEnum_Click
End Sub

Private Property Let ControlsEditZoneVisible(nValue As Boolean)
    Dim ctl As Control
    
    If nValue <> mControlsEditZoneVisible Then
        mControlsEditZoneVisible = nValue
        For Each ctl In mControlsEditZone
            ctl.Visible = mControlsEditZoneVisible
        Next
        lblShortDescription.Caption = "Short description:"
        txtValue.Visible = False
        lblParamsInfo.Visible = False
        txtParamsInfo.Visible = False
        If (mCurrentAction = ecaAddProperty) Or (mCurrentAction = ecaAddMethod) Or (mCurrentAction = ecaAddEvent) Or (mCurrentAction = ecaEditProperty) Or (mCurrentAction = ecaEditMethod) Or (mCurrentAction = ecaEditEvent) Then
            lblParamsInfo.Visible = mControlsEditZoneVisible
            txtParamsInfo.Visible = mControlsEditZoneVisible
        End If
        If (mCurrentAction = ecaAddEnum) Or (mCurrentAction = ecaEditEnum) Then
            lblShortDescription.Visible = False
            txtShortDescription.Visible = False
            lblLongDescription.Caption = "Description:"
        ElseIf (mCurrentAction = ecaAddConstant) Or (mCurrentAction = ecaEditConstant) Then
            txtValue.Visible = mControlsEditZoneVisible
            txtShortDescription.Visible = False
            lblShortDescription.Caption = "Value:"
        ElseIf (mCurrentAction = ecaEditIntroduction) Or (mCurrentAction = ecaEditEndNotes) Then
            lblName.Visible = False
            txtName.Visible = False
            lblLongDescription.Caption = "(Optional section, leave blank to ignore)"
            lblShortDescription.Visible = False
            txtShortDescription.Visible = False
        Else
            lblLongDescription.Caption = "Long description:"
        End If
        PlaceDataControls
        cmdLongDescriptionMenu.Visible = nValue And (CurrentType > 0) And mnuLoadFromOrphanMember2.Visible
        cmdBold.Visible = nValue
        cmdLink.Visible = nValue
        cmdReference.Visible = nValue
        ShowAppliesTo
    End If
End Property

Private Sub ShowAppliesTo()
    Dim iStr As String
    Dim iCurrentType As Long
    Dim iCurrentID As Long
    
    lblAppliesTo.Visible = False
    cmdAppliesTo.Visible = False
    
    iCurrentType = CurrentType
    iCurrentID = GetCurrentMemberID
    
    If (iCurrentType > 0) And (iCurrentID <> 0) Then
        iStr = GetAppliesTo(CurrentType, iCurrentID, True)
        If InStr(iStr, ", ") > 0 Then
            lblAppliesTo.Visible = True
            lblAppliesTo.Caption = "This definition applies to " & UBound(Split(iStr, ",")) + 1 & " objects"
            lblAppliesTo.ToolTipText = iStr
            cmdAppliesTo.Visible = True
            cmdAppliesTo.ToolTipText = iStr
        End If
    End If

End Sub

Private Property Get ControlsEditZoneVisible() As Boolean
    ControlsEditZoneVisible = mControlsEditZoneVisible
End Property


Private Sub ClearControlsEditZone()
    Dim ctl As Control
    
    For Each ctl In mControlsEditZone
        If TypeName(ctl) = "TextBox" Then
            ctl.Text = ""
        End If
    Next
    
End Sub

Private Sub mnuNewEvent_Click()
    CurrentAction = ecaAddEvent
End Sub

Private Sub mnuNewMember_Click()
    If mSelectedType = entPropertiesParent Then
        mnuNewProperty_Click
    ElseIf mSelectedType = entMethodsParent Then
        mnuNewMethod_Click
    ElseIf mSelectedType = entEventsParent Then
        mnuNewEvent_Click
    End If
End Sub

Private Sub mnuNewMethod_Click()
    CurrentAction = ecaAddMethod
End Sub

Private Sub mnuNewComponentDB_Click()
    Dim iComponentName As String
    Dim c As Long
    
    iComponentName = Trim$(InputBox("Enter the component name", "New component database"))
    If iComponentName <> "" Then
        If FileExists(App.Path & "\databases\" & iComponentName & ".mdb") Then
            c = 1
            Do Until Not FileExists(App.Path & "\databases\" & iComponentName & "_" & CStr(c) & ".mdb")
                c = c + 1
            Loop
            mCurrentDBPath = App.Path & "\databases\" & iComponentName & "_" & CStr(c) & ".mdb"
        Else
            mCurrentDBPath = App.Path & "\databases\" & iComponentName & ".mdb"
        End If
        
        If Not FileExists(App.Path & "\BlankDB.mdb") Then
            MsgBox "Error, " & App.Path & "\BlankDB.mdb" & " not found.", vbCritical
            Exit Sub
        End If
        If Not mDatabase Is Nothing Then
            mDatabase.Close
            Set mDatabase = Nothing
        End If
        
        If Not FolderExists(GetFolder(mCurrentDBPath)) Then
            MsgBox "Error: folder " & GetFolder(mCurrentDBPath) & " does not exist.", vbCritical
            Exit Sub
        End If
        FileCopy App.Path & "\BlankDB.mdb", mCurrentDBPath
        
        mSelectedType = 0
        mSelectedID = 0
        mSelectedSecondaryID = 0
        mCurrentAction = ecaDefault
        ControlsEditZoneVisible = False
        ClearControlsEditZone
        
        OpenTheDatabase
        mGeneral_Information.Seek "=", "ComponentName"
        If mGeneral_Information.NoMatch Then
            mGeneral_Information.AddNew
            mGeneral_Information!Name = "ComponentName"
        Else
            mGeneral_Information.Edit
        End If
        mGeneral_Information!Value = iComponentName
        mGeneral_Information.Update
        
        If Not mGeneral_Information Is Nothing Then
            mGeneral_Information.Seek "=", "Introduction"
            If mGeneral_Information.NoMatch Then
                mGeneral_Information.AddNew
                mGeneral_Information!Name = "Introduction"
                mGeneral_Information.Update
            End If
            
            mGeneral_Information.Seek "=", "EndNotes"
            If mGeneral_Information.NoMatch Then
                mGeneral_Information.AddNew
                mGeneral_Information!Name = "EndNotes"
                mGeneral_Information.Update
            End If
        End If
        
        mComponentName = iComponentName
        Me.Caption = App.Title & " - " & mComponentName & " (" & GetFileName(mCurrentDBPath) & ")"
        LoadReportingOptions
        
        ShowTree
        trv1.Nodes(1).EnsureVisible
        trv1.SelectedItem = trv1.Nodes(2)
        trv1_Click
        trv1.SelectedItem = trv1.Nodes(1)
        trv1_Click
    
    End If

End Sub

Private Sub mnuNewProperty_Click()
    CurrentAction = ecaAddProperty
End Sub

Private Sub mnuOpenComponentDB_Click()
    On Error Resume Next
    frmSelectComponentDB.Show vbModal
    If Err.Number Then Exit Sub
    On Error GoTo 0
    If (frmSelectComponentDB.DBPath <> "") And (frmSelectComponentDB.DBPath <> mCurrentDBPath) Then
        mSelectedType = 0
        mSelectedID = 0
        mSelectedSecondaryID = 0
        mCurrentAction = ecaDefault
        If Not mDatabase Is Nothing Then
            mDatabase.Close
            Set mDatabase = Nothing
        End If
        ControlsEditZoneVisible = False
        ClearControlsEditZone
        
        mCurrentDBPath = frmSelectComponentDB.DBPath
        OpenTheDatabase
        LoadReportingOptions
        
        ShowTree
        trv1.Nodes(1).EnsureVisible
        trv1.SelectedItem = trv1.Nodes(2)
        trv1_Click
        trv1.SelectedItem = trv1.Nodes(1)
        trv1_Click
    End If
    Set frmSelectComponentDB = Nothing
End Sub

Public Sub DoPrint()
    Dim c As Long
    Dim iDotsLeft As Long
    Dim iDotsRight As Long
    Dim iPagesAdded As Long
    Dim iStr As String
    Dim iPagePrev As Long
    Dim iDesc As String
    
    Screen.MousePointer = vbHourglass
    mMargin = Printer.ScaleY(20, vbMillimeters, vbTwips)
    
    On Error Resume Next
    Print "";
    If Err.Number Then
        Err.Clear
        ' printer error, probably a dialog was canceled
        ' exit silently
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    On Error GoTo 0
    
    PrintComponentTitle
    Print1
    
    iPagePrev = Printer.Page
    Printer_NewPage
    
    Printer.FontSize = 30
    PrintCenteredTop "Table of contents" & vbCrLf
    Printer.FontSize = 12
    For c = 1 To mTOC_Index
        Printer.CurrentX = mMargin
        CheckForNewPage
        iDotsLeft = Printer.TextWidth(String$(mTOCItems_Level(c) * 4, " ") & mTOCItems(c) & "  ")
        iDotsRight = Printer.ScaleWidth - mMargin - Printer.TextWidth(mTOCItems_Page(c) + iPagesAdded & " ")
        Printer.Print String$(mTOCItems_Level(c) * 4, " ") & mTOCItems(c);
        Printer.CurrentX = iDotsRight
        Printer.Print "  " & mTOCItems_Page(c) + iPagesAdded
        DrawTOCDots iDotsLeft, iDotsRight
    Next
    iPagesAdded = Printer.Page - iPagePrev
    Printer.KillDoc
    Set Printer = Printers(PrinterIndex)

    PrintComponentTitle
    
    Printer.FontSize = 30
    PrintCenteredTop "Table of contents" & vbCrLf
    Printer.FontSize = 12
    For c = 1 To mTOC_Index
        Printer.CurrentX = mMargin
        CheckForNewPage
        iDotsLeft = Printer.TextWidth(String$(mTOCItems_Level(c) * 4, " ") & mTOCItems(c) & "  ")
        iDotsRight = Printer.ScaleWidth - mMargin - Printer.TextWidth(mTOCItems_Page(c) + iPagesAdded & " ")
        Printer.Print String$(mTOCItems_Level(c) * 4, " ") & mTOCItems(c);
        Printer.CurrentX = iDotsRight
        Printer.Print "  " & mTOCItems_Page(c) + iPagesAdded
        DrawTOCDots mMargin + iDotsLeft, iDotsRight
    Next
    Printer_NewPage
    
    Print1
    Printer.EndDoc
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintComponentTitle()
    Dim iDesc As String
    Dim iStr As String
    
    ' component title
    iDesc = GetGeneralInfo("Introduction")
    If (((mControls.RecordCount * frmReportSelection.chkType(1).Value) + (mClasses.RecordCount * frmReportSelection.chkType(2).Value) + (mEnums.RecordCount * frmReportSelection.chkType(3).Value)) > 1) Or (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        Printer.FontSize = 82
        Printer.FontBold = True
        Printer.ForeColor = &HC0C0C0
        PrintCentered mComponentName & IIf((frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoVersion.Value = 1) And (mComponentVersion <> ""), " " & mComponentVersion, "") & vbCrLf & vbCrLf & "Reference", 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
        If (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoReleaseDate.Value = 1) And (mComponentReleaseDate <> 0) Then
            Printer.FontSize = 14
            Printer.Print vbCrLf & vbCrLf
            iStr = "Release date: " & FormatDateTime(mComponentReleaseDate, vbShortDate)
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr)) / 2
            Printer.Print iStr
            Printer.Print
        End If
        Printer_NewPage
        Printer.FontSize = 12
        Printer.FontBold = False
        Printer.ForeColor = vbBlack
    End If

End Sub

Private Sub Print1()
    Dim iRec As Recordset
    Dim iTControls As Recordset
    Dim iTClasses As Recordset
    Dim iTEnums As Recordset
    Dim t As Long
    Dim m As Long
    Dim iTTypes As Recordset
    Dim iDesc As String
    Dim iStr As String
    Dim c As Long
    Dim iLng As Long
    Dim iLng2 As Long
    
    mTOC_Ub = 100
    ReDim mTOCItems(mTOC_Ub)
    ReDim mTOCItems_Level(mTOC_Ub)
    ReDim mTOCItems_Page(mTOC_Ub)
    mTOC_Index = 0
    
    Set iTControls = mControls.Clone
    Set iTClasses = mClasses.Clone
    iTControls.Index = "Name"
    iTClasses.Index = "Name"
    
    iDesc = GetGeneralInfo("Introduction")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoIntroduction.Value = 1) Then
        AddToTOC "Introduction", 0
        rtbAux.Text = ""
        rtbAux.SelFontSize = 36
        rtbAux.SelIndent = 0
        rtbAux.SelBold = True
        rtbAux.SelColor = &H808080 ' &HE189A
        rtbAux.SelText = "Introduction" & vbCrLf & vbCrLf
        rtbAux.SelColor = vbBlack
        rtbAux.SelBold = False
        
        AddRTF iDesc
        rtbAux.SelText = vbCrLf & vbCrLf
        PrintRTB rtbAux
        Printer_NewPage
    End If
    
    For t = 1 To 2 ' 1: controls, 2: classes
        If frmReportSelection.chkType(t).Value Then
            If t = 1 Then
                Set iTTypes = iTControls
            ElseIf t = 2 Then
                Set iTTypes = iTClasses
            End If
            
            If iTTypes.RecordCount > 0 Then
                PrintSeparation 5
                iTTypes.MoveFirst
                Do Until iTTypes.EOF
                    If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                        SeparatePrintedItems
                        ' title
                        Printer.FontSize = 36
                        Printer.CurrentX = mMargin
                        Printer.FontBold = True
                        'Printer.ForeColor = &HE189A
                        Printer.ForeColor = &H808080
    
                        Do Until Printer.TextWidth(iTTypes!Name & " " & LCase(mObjectType_s2(t))) <= (Printer.ScaleWidth - mMargin * 2)
                            Printer.FontSize = Printer.FontSize - 1
                        Loop
                        PrintCenteredTop iTTypes!Name & " " & LCase(mObjectType_s2(t)) & vbCrLf, 0, 0, Printer.ScaleWidth
                        Printer.ForeColor = vbWindowText
                        Printer.FontBold = False
                        
                        AddToTOC iTTypes!Name & " " & LCase(mObjectType_s2(t)), 0
                        
                        ' Description
                        rtbAux.Text = ""
                        rtbAux.SelFontSize = 12
                        If iTTypes!Long_Description <> "" Then
                            iDesc = iTTypes!Long_Description
                        Else
                            iDesc = iTTypes!Short_Description
                        End If
                        If iDesc <> "" Then
                            AddRTF iDesc
                            rtbAux.SelText = vbCrLf & vbCrLf
                            PrintRTB rtbAux
                            rtbAux.Text = ""
                        End If
                        SeparatePrintedItems
                        
                        For m = 1 To 3
                            Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mObjectType_p(t) & "_" & mMemberType_p(m) & ", " & mMemberType_p(m) & " WHERE (" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID = " & mObjectType_p(t) & "_" & mMemberType_p(m) & "." & mMemberType_s(m) & "_ID) AND (" & mObjectType_s(t) & "_ID = " & iTTypes.Fields(mObjectType_s(t) & "_ID") & ") ORDER BY " & mMemberType_p(m) & ".Name")
                            If iRec.RecordCount > 0 Then
                                SeparatePrintedItems
                                rtbAux.SelFontSize = 26
                               ' rtbAux.SelUnderline = True
                               ' rtbAux.SelBold = True
                                rtbAux.SelColor = vbBlue
                                rtbAux.SelText = mMemberType_p(m)
                                rtbAux.SelColor = vbWindowText
                             '   rtbAux.SelUnderline = False
                                rtbAux.SelText = ": " & vbCrLf & vbCrLf
                               ' rtbAux.SelBold = False
                                rtbAux.SelFontSize = 12
                                
                                AddToTOC mMemberType_p(m), 1
                                
                                iRec.MoveFirst
                                Do Until iRec.EOF
                                    Printer.FontSize = 50
                                    If (Printer.ScaleHeight - mMargin - Printer.CurrentY - Printer.TextHeight("tq") * 2) < 0 Then Printer_NewPage
                                    ' Property, method or event name
                                    rtbAux.SelFontSize = 18
                                    rtbAux.SelBold = True
                                    rtbAux.SelColor = &HC51212
                                    rtbAux.SelText = iRec!Name
                                    rtbAux.SelBold = False
                                    rtbAux.SelColor = vbBlack
                                    rtbAux.SelItalic = True
                                    If mMemberType_s(m) = "Method" Then
                                        If Left$(iRec!Params_Info, 17) = "Return Type:" & vbTab & "None" Then
                                            rtbAux.SelText = " " & LCase$(mMemberType_s(m))
                                        Else
                                            rtbAux.SelText = " function"
                                        End If
                                    Else
                                        rtbAux.SelText = " " & LCase$(mMemberType_s(m))
                                    End If
                                    rtbAux.SelItalic = False
                                    rtbAux.SelFontSize = 12
                                    rtbAux.SelText = ":" & vbCrLf & vbCrLf
                                    AddToTOC iRec!Name, 2
                                    
                                    ' Description
                                    If iRec!Params_Info <> "" Then
        '                                rtbAux.SelText = "Parameters information:" & vbCrLf
                                        rtbAux.SelIndent = rtbAux.SelIndent + 500
                                        AddRTF iRec!Params_Info
                                        rtbAux.SelFontSize = 12
                                        rtbAux.SelText = vbCrLf & vbCrLf
                                        rtbAux.SelIndent = rtbAux.SelIndent - 500
                                    End If
                                    iDesc = ""
                                    If iRec!Long_Description <> "" Then
                                        iDesc = iRec!Long_Description
                                    ElseIf iRec!Short_Description <> "" Then
                                        iDesc = iRec!Short_Description
                                    End If
                                    If iDesc <> "" Then
                                        rtbAux.SelFontSize = 12
                                        AddRTF iDesc
                                    End If
                                    
                                    'rtbAux.SelIndent = rtbAux.SelIndent - 500
                                    iRec.MoveNext
                                    
                                    PrintRTB rtbAux
                                    rtbAux.Text = ""
                                    Printer.FontSize = 12 ' 24
                                    
                                    If Not iRec.EOF Then
                                        SeparatePrintedItems 3
                                    End If
                                Loop
                                PrintSeparation 5
                            End If
                        Next m
                    End If
                    iTTypes.MoveNext
                    If Not iTTypes.EOF Then
                        If frmReportSelection.IsItemSelected(t, iTTypes!Name) Then
                            SeparatePrintedItems
                        End If
                    End If
                Loop
            End If
        End If
    Next
    
    ' Constants
    If frmReportSelection.chkType(3).Value Then
        ' title
        If mPrint_Mode = cdSeparatePages Then
            Printer.FontSize = 70
        Else
            Printer.FontSize = 40
        End If
        Printer.CurrentX = mMargin
        Printer.FontBold = True
        Printer.ForeColor = &HC0C0C0
        Do Until Printer.TextWidth("Constants") <= (Printer.ScaleWidth - mMargin * 2)
            Printer.FontSize = Printer.FontSize - 1
        Loop
        If mPrint_Mode = cdSeparatePages Then
            Printer_NewPage
        Else
            PrintSeparation 3
            If (Printer.ScaleHeight - mMargin - Printer.CurrentY - Printer.TextHeight("tq") * 3) < 0 Then
                Printer_NewPage
            End If
        End If
        PrintCenteredTop "Constants" & vbCrLf, 0, 0, Printer.ScaleWidth
        Printer.ForeColor = vbWindowText
        Printer.FontBold = False
        AddToTOC "Constants", 0
        If mPrint_Mode = cdSeparatePages Then Printer_NewPage
        
        Set iTEnums = mEnums.Clone
        iTEnums.Index = "Name"
        If iTEnums.RecordCount > 0 Then
            iTEnums.MoveFirst
            Do Until iTEnums.EOF
                If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Enum_ID = " & iTEnums!Enum_ID & ") ORDER BY " & IIf(iTEnums!Ordered_By_Value, "Value", "Name"))
                    If iRec.RecordCount > 0 Then
                        
                        ' title
                        Printer.FontSize = 20
                        Printer.CurrentX = mMargin
                        Printer.FontBold = True
                        Printer.ForeColor = &H808080
                        Do Until Printer.TextWidth(iTEnums!Name & " enumeration") <= (Printer.ScaleWidth - mMargin * 2)
                            Printer.FontSize = Printer.FontSize - 1
                        Loop
                        If (Printer.ScaleHeight - mMargin - Printer.CurrentY - Printer.TextHeight("tq") * 3) < 0 Then Printer_NewPage
                        'PrintCenteredTop iTEnums!Name & " enumeration" & vbCrLf, 0, 0, Printer.ScaleWidth
                        Printer.Print iTEnums!Name & " enumeration" & vbCrLf
                        Printer.ForeColor = vbWindowText
                        Printer.FontBold = False
                        AddToTOC iTEnums!Name, 1
                        
                        If iTEnums!Description <> "" Then
                            Printer.CurrentY = Printer.CurrentY + Printer.TwipsPerPixelY * 50
                            Printer.DrawWidth = 20
                            iLng2 = Printer.CurrentY
                            iLng = Printer.CurrentY - Printer.TwipsPerPixelY * 50
                            Printer.Line (mMargin, iLng)-(Printer.ScaleWidth - mMargin, iLng), &HC0C0C0
                            Printer.CurrentY = iLng2
                            AddRTF iTEnums!Description
                            PrintRTB rtbAux
                            rtbAux.Text = ""
                            iLng2 = Printer.CurrentY + Printer.TwipsPerPixelY * 50
                            Printer.Line (mMargin, iLng2)-(Printer.ScaleWidth - mMargin, iLng2), &HC0C0C0
                            Printer.CurrentX = mMargin
                            Printer.Print
                        End If
                        
                        iRec.MoveFirst
                        Do Until iRec.EOF
                            rtbAux.SelFontSize = 12
                            rtbAux.SelIndent = 500
                            rtbAux.SelBold = True
                            rtbAux.SelText = iRec!Name & " = " & iRec!Value & vbCrLf
                            rtbAux.SelBold = False
                            If iRec!Description <> "" Then
                                rtbAux.SelIndent = 1000
                                AddRTF iRec!Description
                                rtbAux.SelFontSize = 12
                                rtbAux.SelText = vbCrLf & vbCrLf
                            End If
                            rtbAux.SelIndent = 0
                            iRec.MoveNext
                        Loop
                        PrintRTB rtbAux
                        Printer.FontSize = 12
                        Printer.Print
                    End If
                    rtbAux.Text = ""
                End If
                iTEnums.MoveNext
                If Not iTEnums.EOF Then
                    If frmReportSelection.IsItemSelected(3, iTEnums!Name) Then
                        SeparatePrintedItems
                    End If
                End If
            Loop
        End If
    End If
    
    iDesc = GetGeneralInfo("EndNotes")
    If (iDesc <> "") And (frmReportSelection.chkInfo.Value = 1) And (frmReportSelection.chkInfoEndNotes.Value = 1) Then
        Printer_NewPage
        AddToTOC "End Notes", 0
        rtbAux.Text = ""
        rtbAux.SelFontSize = 36
        rtbAux.SelIndent = 0
        rtbAux.SelBold = True
        rtbAux.SelColor = &H808080 ' &HE189A
        rtbAux.SelText = "End Notes" & vbCrLf & vbCrLf
        rtbAux.SelColor = vbBlack
        rtbAux.SelBold = False
        
        AddRTF iDesc
        rtbAux.SelText = vbCrLf & vbCrLf
        PrintRTB rtbAux
    End If
    
    
    ' Print the table of contents
    Printer.CurrentX = mMargin
    Printer.CurrentY = mMargin
End Sub

Private Sub mnuComponentProperties_Click()
    If mGeneral_Information Is Nothing Then Exit Sub
    
    frmComponentProperties.txtName.Text = mComponentName
    frmComponentProperties.txtVersion.Text = mComponentVersion
    If mComponentReleaseDate <> 0 Then
        frmComponentProperties.txtReleaseDate.Text = FormatDateTime(mComponentReleaseDate, vbShortDate)
    End If
    frmComponentProperties.Show vbModal
    If frmComponentProperties.OKPressed Then
        mGeneral_Information.Seek "=", "ComponentName"
        If Not mGeneral_Information.NoMatch Then
            mComponentName = frmComponentProperties.ComponentName
            mGeneral_Information.Edit
            mGeneral_Information!Value = mComponentName
            mGeneral_Information.Update
            Me.Caption = App.Title & " - " & mComponentName & " (" & GetFileName(mCurrentDBPath) & ")"
            LoadReportingOptions
        End If
        
        mGeneral_Information.Seek "=", "ComponentVersion"
        mComponentVersion = frmComponentProperties.ComponentVersion
        If mGeneral_Information.NoMatch Then
            mGeneral_Information.AddNew
            mGeneral_Information!Name = "ComponentVersion"
        Else
            mGeneral_Information.Edit
        End If
        mGeneral_Information!Value = mComponentVersion
        mGeneral_Information.Update
        
        mGeneral_Information.Seek "=", "ComponentReleaseDate"
        If IsDate(frmComponentProperties.ComponentReleaseDate) Then
            mComponentReleaseDate = CLng(CDate(frmComponentProperties.ComponentReleaseDate))
        Else
            mComponentReleaseDate = 0
        End If
        If mGeneral_Information.NoMatch Then
            mGeneral_Information.AddNew
            mGeneral_Information!Name = "ComponentReleaseDate"
        Else
            mGeneral_Information.Edit
        End If
        mGeneral_Information!Value = CLng(mComponentReleaseDate)
        mGeneral_Information.Update
    End If
    Set frmComponentProperties = Nothing
End Sub

Private Function GetCurrentMemberID() As Long
    On Error Resume Next
    Select Case CurrentType
        Case 1 ' property
            GetCurrentMemberID = mProperties!Property_ID
        Case 2 ' method
            GetCurrentMemberID = mMethods!Method_ID
        Case 3 ' event
            GetCurrentMemberID = mEvents!Event_ID
    End Select
End Function

Private Function GetCurrentMemeberTypeRecordSet() As Recordset
    On Error Resume Next
    Select Case CurrentType
        Case 1 ' property
            Set GetCurrentMemeberTypeRecordSet = mProperties
        Case 2 ' method
            Set GetCurrentMemeberTypeRecordSet = mMethods
        Case 3 ' event
            Set GetCurrentMemeberTypeRecordSet = mEvents
    End Select
End Function

Private Sub mnuSetMethodToExistentDefinition_Click()
    Dim iCurrentType As Long
    Dim iCurrentID As Long
    Dim iCurrentAction As Long
    Dim iCurrentObjectTypePluralStr As String
    Dim iCurrentObjectID As Long
    Dim iMembers As Recordset
    Dim iCurrentMemberTypeRecordSet As Recordset
    Dim iRec As Recordset
    
    iCurrentType = CurrentType
    iCurrentID = GetCurrentMemberID
    
    If mSelectedID = 0 Then
        MsgBox "Error 12351, no selected ID", vbExclamation
        Exit Sub
    End If
    If (iCurrentType > 0) And (iCurrentID <> 0) Then
        frmSelectMemberDefinition.LoadList mDatabase, mMemberType_p(iCurrentType), txtName.Text, mMemberType_s(iCurrentType), iCurrentID, txtParamsInfo.Text, txtShortDescription.Text
        frmSelectMemberDefinition.Show vbModal
        If frmSelectMemberDefinition.OKPressed Then
            
            Set iCurrentMemberTypeRecordSet = GetCurrentMemeberTypeRecordSet
            
            If Not iCurrentMemberTypeRecordSet Is Nothing Then
                iCurrentAction = CurrentAction
                
                iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
                If iCurrentObjectTypePluralStr = "" Then
                    MsgBox "Error 12349 could not get CurrentObjectType", vbExclamation
                    Exit Sub
                End If
                iCurrentObjectID = GetCurrentObjectID
                If iCurrentObjectID = 0 Then
                    MsgBox "Error 12350 could not get CurrentObjectID", vbExclamation
                    Exit Sub
                End If
                Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_" & mMemberType_p(iCurrentType))

                ' new table item
                iMembers.AddNew
                iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
                'iMembers!Property_ID = mProperties!Property_ID
                iMembers.Fields(mMemberType_s(iCurrentType) & "_ID") = frmSelectMemberDefinition.MemberID
                If frmSelectMemberDefinition.ParamsInfo <> txtParamsInfo.Text Then
                    Set iRec = mDatabase.OpenRecordset("Params_Info_Replaced")
                    iRec.AddNew
                    iRec!Params_Info = txtParamsInfo.Text
                    iRec.Update
                    iRec.Bookmark = iRec.LastModified
                    iMembers!ID_Params_Info_Replaced = iRec!ID_Params_Info_Replaced
                End If
                If frmSelectMemberDefinition.ShortDescription <> txtShortDescription.Text Then
                    Set iRec = mDatabase.OpenRecordset("Short_Description_Replaced")
                    iRec.AddNew
                    iRec!Short_Description = txtShortDescription.Text
                    iRec.Update
                    iRec.Bookmark = iRec.LastModified
                    iMembers!ID_Short_Description_Replaced = iRec!ID_Short_Description_Replaced
                End If
                iMembers.Update
                
                ' set as orphan if the old member definition is not used anywhere else
                Set iMembers = mDatabase.OpenRecordset("SELECT * FROM " & iCurrentObjectTypePluralStr & "_" & mMemberType_p(iCurrentType) & " WHERE (" & mMemberType_s(iCurrentType) & "_ID = " & mSelectedID & ")")
                If iMembers.RecordCount = 0 Then
                    MsgBox "Error 12352 could not update the current ID", vbExclamation
                Else
                    iMembers.MoveLast
                    If iMembers.RecordCount = 1 Then
                        iCurrentMemberTypeRecordSet.Seek "=", iCurrentID
                        If iCurrentMemberTypeRecordSet.NoMatch Then
                            MsgBox "Error 12353 could not set as orphan member", vbExclamation
                        End If
                        iCurrentMemberTypeRecordSet.Edit
                        iCurrentMemberTypeRecordSet!Auxiliary_Field = 0
                        iCurrentMemberTypeRecordSet.Update
                    End If
                End If
                
                ' remove old table item
                iMembers.FindFirst "(" & GetCurrentObjectTypeSingularStr & "_ID = " & iCurrentObjectID & ") AND (" & mMemberType_s(iCurrentType) & "_ID = " & iCurrentID & ")"
                If iMembers.NoMatch Then
                    MsgBox "Error 12348 removing current ID from recordset", vbExclamation
                    Exit Sub
                Else
                    iMembers.Delete
                End If
                
                mSelectedID = frmSelectMemberDefinition.MemberID
                iCurrentMemberTypeRecordSet.Seek "=", mSelectedID
                If iCurrentMemberTypeRecordSet.NoMatch Then
                    MsgBox "Error 12354 could not find new member record", vbExclamation
                End If
                mSelectedSecondaryID = iCurrentObjectID
                
                ShowTree
                ShowAppliesTo
                
                CurrentAction = iCurrentAction
            Else
                MsgBox "Error 12347 getting current type recordset", vbExclamation
            End If
        End If
        Set frmSelectMemberDefinition = Nothing
    End If
End Sub

Private Sub mnuSetMethodToExistentDefinition2_Click()
    mnuSetMethodToExistentDefinition_Click
End Sub

Private Sub mnyDeleteComponentDB_Click()
    On Error Resume Next
    frmSelectComponentDB.Caption = "Select database to delete"
    If Err.Number Then Exit Sub
    On Error GoTo 0
    frmSelectComponentDB.Show vbModal
    If (frmSelectComponentDB.DBPath <> "") And (frmSelectComponentDB.DBPath = mCurrentDBPath) Then
        mSelectedType = 0
        mSelectedID = 0
        mSelectedSecondaryID = 0
        mCurrentAction = ecaDefault
        If Not mDatabase Is Nothing Then
            mDatabase.Close
            Set mDatabase = Nothing
        End If
        ControlsEditZoneVisible = False
        ClearControlsEditZone
        
        mCurrentDBPath = ""
        
        Set mClasses = Nothing
        ShowTree
        trv1.Nodes(1).EnsureVisible
        trv1.SelectedItem = trv1.Nodes(1)
        Me.Caption = App.Title
    End If
    If frmSelectComponentDB.DBPath <> "" Then
        On Error Resume Next
        Kill frmSelectComponentDB.DBPath
        On Error GoTo 0
    End If
    Set frmSelectComponentDB = Nothing
End Sub

Private Sub tmrCheckcboReListVisible_Timer()
    If IsWindowVisible(mcboRefhWndList) = 0 Then
        tmrCheckcboReListVisible.Enabled = False
        txtLongDescription.SetFocus
    End If
End Sub

Private Sub tmrSetFocus_Timer()
    tmrSetFocus.Enabled = False
    On Error Resume Next
    Me.Controls(tmrSetFocus.Tag).SetFocus
End Sub

Private Sub trv1_Click()
    UpdateCurrentSelected
End Sub

Private Sub trv1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Node
    Dim iType As ENodeType
    Dim IID As Long
    Dim iStrs() As String
    
    UpdateCurrentSelected
    If Button = vbRightButton Then
        Set n = trv1.HitTest(X, Y)
        If Not n Is Nothing Then
            Set trv1.SelectedItem = n
            n.Selected = True
            n.Expanded = True
            iStrs = Split(n.Key, "|")
            If UBound(iStrs) >= 1 Then
                iType = Val(iStrs(0))
                IID = Val(iStrs(1))
            End If
            If (iType = entClassesParent) Then
                PopupMenu mnuPopupClasses
            ElseIf (iType = entControlsParent) Then
                PopupMenu mnuPopupControls
            ElseIf (iType = entEnumsParent) Then
                PopupMenu mnuPopupEnums
            ElseIf (iType = entEnum) Then
                mEnums.Index = "PrimaryKey"
                mEnums.Seek "=", mSelectedID
                If mEnums!Ordered_By_Value Then
                    mnuConstantsOrderedByName.Checked = False
                    mnuConstantsOrderedByValue.Checked = True
                Else
                    mnuConstantsOrderedByName.Checked = True
                    mnuConstantsOrderedByValue.Checked = False
                End If
                PopupMenu mnuPopupEnum
            ElseIf (iType = entClass) Or (iType = entControl) Then
                PopupMenu mnuPopupClassOrControl
            ElseIf (iType = entPropertiesParent) Or (iType = entMethodsParent) Or (iType = entEventsParent) Then
                PopupMenu mnuPopupMembersParent
            ElseIf (iType = entProperty) Or (iType = entMethod) Or (iType = entEvent) Then
                PopupMenu mnuPopupMember
            ElseIf (iType = entConstant) Then
                PopupMenu mnuPopupMember
            End If
        End If
    End If
End Sub

Private Sub txtLongDescription_GotFocus()
    cmdBold.Enabled = True
    cmdLink.Enabled = True
    cmdReference.Enabled = True
End Sub

Private Sub txtLongDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) <> 0 Then
        If KeyCode = vbKeyJ Then
            KeyCode = 0
        End If
    End If
End Sub

Private Sub txtLongDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then KeyAscii = 0
End Sub

Private Sub txtLongDescription_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Shift And vbCtrlMask) <> 0 Then
        If KeyCode = vbKeyJ Then
            KeyCode = 0
            cmdReference.Tag = "1"
            cmdReference_Click
        End If
    End If
End Sub

Private Sub txtLongDescription_LostFocus()
    If StrCount(txtLongDescription.Text, "<b>") <> StrCount(txtLongDescription.Text, "</b>") Then
        MsgBox "The count of opening and closing html tags, <b> and </b>, does not match, please correct.", vbExclamation
        tmrSetFocus.Tag = "txtLongDescription"
        tmrSetFocus.Enabled = True
        Exit Sub
    End If
    If StrCount(txtLongDescription.Text, "[") <> StrCount(txtLongDescription.Text, "]") Then
        MsgBox "The count of opening and closing square blackets, [ and ], does not match, please correct.", vbExclamation
        tmrSetFocus.Tag = "txtLongDescription"
        tmrSetFocus.Enabled = True
        Exit Sub
    End If
    
    If Me.ActiveControl Is cmdBold Then
        cmdBold.Tag = "1"
    ElseIf Me.ActiveControl Is cmdLink Then
        cmdLink.Tag = 1
    ElseIf Me.ActiveControl Is cmdReference Then
        cmdReference.Tag = 1
    Else
        cmdBold.Enabled = False
        cmdLink.Enabled = False
        cmdReference.Enabled = False
    End If
End Sub

Private Sub txtLongDescription_Validate(Cancel As Boolean)
    UpdateData
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    UpdateData
End Sub

Private Sub txtParamsInfo_KeyPress(KeyAscii As Integer)
    If mFileImported Then
        If Not mfrmFieldsModifAlertShowed Then
            If Val(GetSetting(App.Title, AppPath4Reg, "HideModifAlert", "0")) = 0 Then
                If KeyAscii <> 3 Then ' Ctrol+C
                    frmFieldsModifAlert.Show vbModal
                    mfrmFieldsModifAlertShowed = True
                End If
            End If
        End If
    End If
End Sub

Private Sub txtParamsInfo_Validate(Cancel As Boolean)
    UpdateData
End Sub

Private Sub txtShortDescription_KeyPress(KeyAscii As Integer)
    If mFileImported Then
        If Not mfrmFieldsModifAlertShowed Then
            If (mCurrentAction = ecaEditProperty) Or (mCurrentAction = ecaEditMethod) Or (mCurrentAction = ecaEditEvent) Then
                If Val(GetSetting(App.Title, AppPath4Reg, "HideModifAlert", "0")) = 0 Then
                    If KeyAscii <> 3 Then ' Ctrol+C
                        frmFieldsModifAlert.Show vbModal
                        mfrmFieldsModifAlertShowed = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtShortDescription_Validate(Cancel As Boolean)
    UpdateData
End Sub


Private Property Get App_Path()
    Static sValue As String
    
    If sValue = "" Then
        sValue = App.Path
        If Right$(sValue, 1) = "\" Then
            sValue = Left$(sValue, Len(sValue) - 1)
        End If
    End If
    App_Path = sValue
End Property

Private Property Get DBPath()
    Static sValue As String
    
    If sValue = "" Then
        sValue = App_Path & "\databases\test.mdb"
    End If
     DBPath = sValue
End Property

Private Sub OpenTheDatabase()
    Dim iRec As Recordset
    Dim iTable As TableDef
    
    Set mDatabase = Nothing
    If mCurrentDBPath = "" Then Exit Sub
    Set mDatabase = DBEngine.OpenDatabase(mCurrentDBPath)
    Dim dbx As Database
    Set mClasses = mDatabase.OpenRecordset("Classes")
    mClasses.Index = "PrimaryKey"
    Set mControls = mDatabase.OpenRecordset("Controls")
    mControls.Index = "PrimaryKey"
    Set mEnums = mDatabase.OpenRecordset("Enums")
    mEnums.Index = "PrimaryKey"
    Set mProperties = mDatabase.OpenRecordset("Properties")
    mProperties.Index = "PrimaryKey"
    Set mMethods = mDatabase.OpenRecordset("Methods")
    mMethods.Index = "PrimaryKey"
    Set mEvents = mDatabase.OpenRecordset("Events")
    mEvents.Index = "PrimaryKey"
    Set mConstants = mDatabase.OpenRecordset("Constants")
    mConstants.Index = "PrimaryKey"
    Set mGeneral_Information = mDatabase.OpenRecordset("General_Information")
    mGeneral_Information.Index = "Name"
    
    mGeneral_Information.Seek "=", "ComponentName"
    If Not mGeneral_Information.NoMatch Then
        mComponentName = mGeneral_Information!Value
        Me.Caption = App.Title & " - " & mComponentName & " (" & GetFileName(mCurrentDBPath) & ")"
    End If
    mComponentVersion = GetGeneralInfo("ComponentVersion")
    mComponentReleaseDate = CDate(Val(GetGeneralInfo("ComponentReleaseDate")))
    
    mNewEnumsOrderedByValue = CBool(Val(GetSettingBase("General", "NewEnumsOrderedByValue", "0")))
    LoadReportingOptions
    mFileImported = CBool(GetSettingBase("General", "FileImported", "0"))
    
    ' Update previous version database
    On Error Resume Next
    Set iTable = mDatabase.TableDefs("Params_Info_Replaced")
    On Error GoTo 0
    If iTable Is Nothing Then
        Dim t As Long
        Dim m As Long
        Dim iField As DAO.Field
        Dim iIndex As DAO.Index
        
        'Table: Params_Info_Replaced
        Set iTable = mDatabase.CreateTableDef("Params_Info_Replaced")
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbAutoIncrField Or dbDescending
        iField.Required = False
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("Params_Info", dbMemo)
        iField.Required = True
        iField.AllowZeroLength = True
        iField.DefaultValue = """"""
        iTable.Fields.Append iField
    
        Set iIndex = iTable.CreateIndex("ID")
        iIndex.Fields.Append iIndex.CreateField("ID_Params_Info_Replaced")
        iIndex.Unique = False
        iIndex.Primary = False
        iIndex.Required = False
        iTable.Indexes.Append iIndex
    
        Set iIndex = iTable.CreateIndex("PrimaryKey")
        iIndex.Fields.Append iIndex.CreateField("ID_Params_Info_Replaced")
        iIndex.Unique = True
        iIndex.Primary = True
        iIndex.Required = True
        iTable.Indexes.Append iIndex
    
        mDatabase.TableDefs.Append iTable
        
        'Table: Short_Description_Replaced
        Set iTable = mDatabase.CreateTableDef("Short_Description_Replaced")
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbAutoIncrField Or dbDescending
        iField.Required = False
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("Short_Description", dbMemo)
        iField.Required = True
        iField.AllowZeroLength = True
        iField.DefaultValue = """"""
        iTable.Fields.Append iField
    
        Set iIndex = iTable.CreateIndex("ID")
        iIndex.Fields.Append iIndex.CreateField("ID_Short_Description_Replaced")
        iIndex.Unique = False
        iIndex.Primary = False
        iIndex.Required = False
        iTable.Indexes.Append iIndex
    
        Set iIndex = iTable.CreateIndex("PrimaryKey")
        iIndex.Fields.Append iIndex.CreateField("ID_Short_Description_Replaced")
        iIndex.Unique = True
        iIndex.Primary = True
        iIndex.Required = True
        iTable.Indexes.Append iIndex
    
        mDatabase.TableDefs.Append iTable
        
        Set iTable = mDatabase.TableDefs("Controls_Properties")
        
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = True
        iField.DefaultValue = 0
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = False
        iField.DefaultValue = 0
        iTable.Fields.Append iField
        
        Set iTable = mDatabase.TableDefs("Controls_Methods")
        
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = True
        iField.DefaultValue = 0
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = False
        iField.DefaultValue = 0
        iTable.Fields.Append iField
        
        Set iTable = mDatabase.TableDefs("Controls_Events")
        
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = True
        iField.DefaultValue = 0
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = False
        iField.DefaultValue = 0
        iTable.Fields.Append iField
        
        
        Set iTable = mDatabase.TableDefs("Classes_Properties")
        
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = True
        iField.DefaultValue = 0
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = False
        iField.DefaultValue = 0
        iTable.Fields.Append iField
        
        Set iTable = mDatabase.TableDefs("Classes_Methods")
        
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = True
        iField.DefaultValue = 0
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = False
        iField.DefaultValue = 0
        iTable.Fields.Append iField
        
        Set iTable = mDatabase.TableDefs("Classes_Events")
        
        Set iField = iTable.CreateField("ID_Params_Info_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = True
        iField.DefaultValue = 0
        iTable.Fields.Append iField
    
        Set iField = iTable.CreateField("ID_Short_Description_Replaced", dbLong)
        iField.Attributes = dbDescending
        iField.Required = False
        iField.DefaultValue = 0
        iTable.Fields.Append iField
        
        For t = 1 To 2
            For m = 1 To 3
                Set iRec = mDatabase.OpenRecordset(mObjectType_p(t) & "_" & mMemberType_p(m))
                If iRec.RecordCount > 0 Then
                    iRec.MoveFirst
                    Do Until iRec.EOF
                        iRec.Edit
                        iRec!ID_Params_Info_Replaced = 0
                        iRec!ID_Short_Description_Replaced = 0
                        iRec.Update
                        iRec.Bookmark = iRec.LastModified
                        iRec.MoveNext
                    Loop
                End If
            Next
        Next
    End If
End Sub

Private Sub UpdateData()
    Dim iStr As String
    Dim c As Long
    Dim iMembers As Recordset
    Dim iCurrentObjectTypePluralStr As String
    Dim iCurrentObjectID As Long
    Dim iUpdateObjects As Boolean
    
    If (Trim(txtName.Text) = "") And (Trim(txtShortDescription.Text) = "") And (Trim(txtLongDescription.Text) = "") Then
        If (mCurrentAction <> ecaEditIntroduction) And (mCurrentAction <> ecaEditEndNotes) Then
            Exit Sub
        End If
    End If
    If mDeletingNode Then Exit Sub
    
    If mCurrentAction = ecaAddClass Then
        iStr = Trim(txtName.Text)
        mClasses.Index = "Name"
        mClasses.Seek "=", iStr
        If mClasses.NoMatch Then
            iStr = GetName
            mClasses.AddNew
            mClasses!Name = iStr
            mClasses!Short_Description = Trim(txtShortDescription.Text)
            mClasses!Long_Description = Trim(txtLongDescription.Text)
            mClasses.Update
            mClasses.Bookmark = mClasses.LastModified
            mSelectedType = entClass
            mSelectedID = mClasses!Class_ID
            ShowTree
            mCurrentAction = ecaEditClass
        Else
            MsgBox "There is already a Class with that name", vbExclamation
        End If
        mClasses.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaAddControl Then
        iStr = Trim(txtName.Text)
        mControls.Index = "Name"
        mControls.Seek "=", iStr
        If mControls.NoMatch Then
            iStr = GetName
            mControls.AddNew
            mControls!Name = iStr
            mControls!Short_Description = Trim(txtShortDescription.Text)
            mControls!Long_Description = Trim(txtLongDescription.Text)
            mControls.Update
            mControls.Bookmark = mControls.LastModified
            mSelectedType = entControl
            mSelectedID = mControls!control_ID
            ShowTree
            mCurrentAction = ecaEditControl
        Else
            MsgBox "There is already a Control with that name", vbExclamation
        End If
        mControls.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaAddEnum Then
        iStr = Trim(txtName.Text)
        mEnums.Index = "Name"
        mEnums.Seek "=", iStr
        If mEnums.NoMatch Then
            iStr = GetName
            mEnums.AddNew
            If mNewEnumsOrderedByValue Then
                mEnums!Ordered_By_Value = True
            End If
            mEnums!Name = iStr
            mEnums!Description = Trim(txtLongDescription.Text)
            mEnums.Update
            mEnums.Bookmark = mEnums.LastModified
            mSelectedType = entEnum
            mSelectedID = mEnums!Enum_ID
            ShowTree
            mCurrentAction = ecaEditEnum
        Else
            MsgBox "There is already a Enum with that name", vbExclamation
        End If
        mEnums.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaEditClass Then
        iStr = GetName
        If mSelectedType <> entClass Then Err.Raise 1234
        mClasses.Index = "PrimaryKey"
        mClasses.Seek "=", mSelectedID
        If mClasses.NoMatch Then Err.Raise 1234
        If mClasses!Name <> Trim$(txtName.Text) Then
            mClasses.Index = "Name"
            mClasses.Seek "=", Trim(txtName.Text)
            If Not mClasses.NoMatch Then
                MsgBox "There is already a Class with that name", vbExclamation
                Exit Sub
            End If
            mClasses.Index = "PrimaryKey"
            mClasses.Seek "=", mSelectedID
        End If
        mClasses.Edit
        If mClasses!Name <> iStr Then
            mClasses!Name = iStr
            iUpdateObjects = True
        End If
        mClasses!Short_Description = Trim(txtShortDescription.Text)
        mClasses!Long_Description = Trim(txtLongDescription.Text)
        mClasses.Update
        If iUpdateObjects Then ShowTree
        mClasses.Bookmark = mClasses.LastModified
    ElseIf mCurrentAction = ecaEditControl Then
        iStr = GetName
        If mSelectedType <> entControl Then Err.Raise 1234
        mControls.Index = "PrimaryKey"
        mControls.Seek "=", mSelectedID
        If mControls.NoMatch Then Err.Raise 1234
        If mControls!Name <> Trim$(txtName.Text) Then
            mControls.Index = "Name"
            mControls.Seek "=", Trim(txtName.Text)
            If Not mControls.NoMatch Then
                MsgBox "There is already a Control with that name", vbExclamation
                Exit Sub
            End If
            mControls.Index = "PrimaryKey"
            mControls.Seek "=", mSelectedID
        End If
        mControls.Edit
        If mControls!Name <> iStr Then
            mControls!Name = iStr
            iUpdateObjects = True
        End If
        mControls!Short_Description = Trim(txtShortDescription.Text)
        mControls!Long_Description = Trim(txtLongDescription.Text)
        mControls.Update
        If iUpdateObjects Then ShowTree
        mControls.Bookmark = mControls.LastModified
    ElseIf mCurrentAction = ecaEditEnum Then
        iStr = GetName
        If mSelectedType <> entEnum Then Err.Raise 1234
        mEnums.Index = "PrimaryKey"
        mEnums.Seek "=", mSelectedID
        If mEnums.NoMatch Then Err.Raise 1234
        If mEnums!Name <> Trim$(txtName.Text) Then
            mEnums.Index = "Name"
            mEnums.Seek "=", Trim(txtName.Text)
            If Not mEnums.NoMatch Then
                MsgBox "There is already a Enum with that name", vbExclamation
                Exit Sub
            End If
            mEnums.Index = "PrimaryKey"
            mEnums.Seek "=", mSelectedID
        End If
        mEnums.Edit
        If mEnums!Name <> iStr Then
            mEnums!Name = iStr
            iUpdateObjects = True
        End If
        mEnums!Description = Trim(txtLongDescription.Text)
        mEnums.Update
        If iUpdateObjects Then ShowTree
        mEnums.Bookmark = mEnums.LastModified
    ElseIf mCurrentAction = ecaAddProperty Then
        iStr = Trim(txtName.Text)
        mProperties.Index = "Name"
        mProperties.Seek "=", iStr
        If mProperties.NoMatch Then
            iStr = GetName
            mProperties.AddNew
            mProperties!Name = iStr
            mProperties!Short_Description = Trim(txtShortDescription.Text)
            mProperties!Long_Description = Trim(txtLongDescription.Text)
            mProperties!Params_Info = Trim(txtParamsInfo.Text)
            mProperties!Auxiliary_Field = 1
            mProperties.Update
            mProperties.Bookmark = mProperties.LastModified
            iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
            iCurrentObjectID = GetCurrentObjectID
            Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_Properties")
            iMembers.AddNew
            iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
            iMembers!Property_ID = mProperties!Property_ID
            iMembers.Update
            mSelectedType = entProperty
            mSelectedID = mProperties!Property_ID
            mSelectedSecondaryID = iCurrentObjectID
            ShowTree
            mCurrentAction = ecaEditProperty
        Else
            If MsgBox("There is already a Property with that name. Do you want to use the existent property? (If select Yes, any data entered now will be lost)", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            mProperties.Edit
            txtShortDescription.Text = mProperties!Short_Description
            txtLongDescription.Text = mProperties!Long_Description
            iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
            iCurrentObjectID = GetCurrentObjectID
            Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_Properties")
            iMembers.AddNew
            iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
            iMembers!Property_ID = mProperties!Property_ID
            iMembers.Update
            mSelectedType = entProperty
            mSelectedID = mProperties!Property_ID
            mSelectedSecondaryID = iCurrentObjectID
            ShowTree
            mCurrentAction = ecaEditProperty
        End If
        mProperties.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaAddMethod Then
        iStr = Trim(txtName.Text)
        mMethods.Index = "Name"
        mMethods.Seek "=", iStr
        If mMethods.NoMatch Then
            iStr = GetName
            mMethods.AddNew
            mMethods!Name = iStr
            mMethods!Short_Description = Trim(txtShortDescription.Text)
            mMethods!Long_Description = Trim(txtLongDescription.Text)
            mMethods!Params_Info = Trim(txtParamsInfo.Text)
            mMethods!Auxiliary_Field = 1
            mMethods.Update
            mMethods.Bookmark = mMethods.LastModified
            iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
            iCurrentObjectID = GetCurrentObjectID
            Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_Methods")
            iMembers.AddNew
            iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
            iMembers!Method_ID = mMethods!Method_ID
            iMembers.Update
            mSelectedType = entMethod
            mSelectedID = mMethods!Method_ID
            mSelectedSecondaryID = iCurrentObjectID
            ShowTree
            mCurrentAction = ecaEditMethod
        Else
            If MsgBox("There is already a Method with that name. Do you want to use the existent Method? (If select Yes, any data entered now will be lost)", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            mMethods.Edit
            txtShortDescription.Text = mMethods!Short_Description
            txtLongDescription.Text = mMethods!Long_Description
            iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
            iCurrentObjectID = GetCurrentObjectID
            Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_Methods")
            iMembers.AddNew
            iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
            iMembers!Method_ID = mMethods!Method_ID
            iMembers.Update
            mSelectedType = entMethod
            mSelectedID = mMethods!Method_ID
            mSelectedSecondaryID = iCurrentObjectID
            ShowTree
            mCurrentAction = ecaEditMethod
        End If
        mMethods.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaAddEvent Then
        iStr = Trim(txtName.Text)
        mEvents.Index = "Name"
        mEvents.Seek "=", iStr
        If mEvents.NoMatch Then
            iStr = GetName
            mEvents.AddNew
            mEvents!Name = iStr
            mEvents!Short_Description = Trim(txtShortDescription.Text)
            mEvents!Long_Description = Trim(txtLongDescription.Text)
            mEvents!Params_Info = Trim(txtParamsInfo.Text)
            mEvents!Auxiliary_Field = 1
            mEvents.Update
            mEvents.Bookmark = mEvents.LastModified
            iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
            iCurrentObjectID = GetCurrentObjectID
            Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_Events")
            iMembers.AddNew
            iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
            iMembers!Event_ID = mEvents!Event_ID
            iMembers.Update
            mSelectedType = entEvent
            mSelectedID = mEvents!Event_ID
            mSelectedSecondaryID = iCurrentObjectID
            ShowTree
            mCurrentAction = ecaEditEvent
        Else
            If MsgBox("There is already a Event with that name. Do you want to use the existent Event? (If select Yes, any data entered now will be lost)", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            mEvents.Edit
            txtShortDescription.Text = mEvents!Short_Description
            txtLongDescription.Text = mEvents!Long_Description
            iCurrentObjectTypePluralStr = GetCurrentObjectTypePluralStr
            iCurrentObjectID = GetCurrentObjectID
            Set iMembers = mDatabase.OpenRecordset(iCurrentObjectTypePluralStr & "_Events")
            iMembers.AddNew
            iMembers.Fields(GetCurrentObjectTypeSingularStr & "_ID").Value = iCurrentObjectID
            iMembers!Event_ID = mEvents!Event_ID
            iMembers.Update
            mSelectedType = entEvent
            mSelectedID = mEvents!Event_ID
            mSelectedSecondaryID = iCurrentObjectID
            ShowTree
            mCurrentAction = ecaEditEvent
        End If
        mEvents.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaEditProperty Then
        iStr = GetName
        If mSelectedType <> entProperty Then Err.Raise 1234
        mProperties.Index = "PrimaryKey"
        mProperties.Seek "=", mSelectedID
        If mProperties.NoMatch Then Err.Raise 1234
        If mProperties!Name <> Trim$(txtName.Text) Then
            mProperties.Index = "Name"
            mProperties.Seek "=", Trim(txtName.Text)
            If Not mProperties.NoMatch Then
                MsgBox "There is already a Property with that name", vbExclamation
                Exit Sub
            End If
            mProperties.Index = "PrimaryKey"
            mProperties.Seek "=", mSelectedID
        End If
        mProperties.Edit
        If mProperties!Name <> iStr Then
            mProperties!Name = iStr
            iUpdateObjects = True
        End If
        mProperties!Short_Description = Trim(txtShortDescription.Text)
        mProperties!Long_Description = Trim(txtLongDescription.Text)
        mProperties!Params_Info = Trim(txtParamsInfo.Text)
        mProperties.Update
        If iUpdateObjects Then ShowTree
        mProperties.Bookmark = mProperties.LastModified
    ElseIf mCurrentAction = ecaEditMethod Then
        iStr = GetName
        If mSelectedType <> entMethod Then Err.Raise 1234
        mMethods.Index = "PrimaryKey"
        mMethods.Seek "=", mSelectedID
        If mMethods.NoMatch Then Err.Raise 1234
        If mMethods!Name <> Trim$(txtName.Text) Then
            mMethods.Index = "Name"
            mMethods.Seek "=", Trim(txtName.Text)
            If Not mMethods.NoMatch Then
                MsgBox "There is already a Method with that name", vbExclamation
                Exit Sub
            End If
            mMethods.Index = "PrimaryKey"
            mMethods.Seek "=", mSelectedID
        End If
        mMethods.Edit
        If mMethods!Name <> iStr Then
            mMethods!Name = iStr
            iUpdateObjects = True
        End If
        mMethods!Short_Description = Trim(txtShortDescription.Text)
        mMethods!Long_Description = Trim(txtLongDescription.Text)
        mMethods!Params_Info = Trim(txtParamsInfo.Text)
        mMethods.Update
        If iUpdateObjects Then ShowTree
        mMethods.Bookmark = mMethods.LastModified
    ElseIf mCurrentAction = ecaEditEvent Then
        iStr = GetName
        If mSelectedType <> entEvent Then Err.Raise 1234
        mEvents.Index = "PrimaryKey"
        mEvents.Seek "=", mSelectedID
        If mEvents.NoMatch Then Err.Raise 1234
        If mEvents!Name <> Trim$(txtName.Text) Then
            mEvents.Index = "Name"
            mEvents.Seek "=", Trim(txtName.Text)
            If Not mEvents.NoMatch Then
                MsgBox "There is already a Event with that name", vbExclamation
                Exit Sub
            End If
            mEvents.Index = "PrimaryKey"
            mEvents.Seek "=", mSelectedID
        End If
        mEvents.Edit
        If mEvents!Name <> iStr Then
            mEvents!Name = iStr
            iUpdateObjects = True
        End If
        mEvents!Name = iStr
        mEvents!Short_Description = Trim(txtShortDescription.Text)
        mEvents!Long_Description = Trim(txtLongDescription.Text)
        mEvents!Params_Info = Trim(txtParamsInfo.Text)
        mEvents.Update
        If iUpdateObjects Then ShowTree
        mEvents.Bookmark = mEvents.LastModified
    ElseIf mCurrentAction = ecaAddConstant Then
        iStr = Trim(txtName.Text)
        If iStr = "" Then Exit Sub
        mConstants.Index = "Name"
        mConstants.Seek "=", iStr
        If Not mConstants.NoMatch Then
            If MsgBox("There is already a Constant with that name. Do you want to use that name anyway?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
        iStr = Trim(txtName.Text)
        mConstants.AddNew
        mConstants!Name = iStr
        mConstants!Enum_ID = GetCurrentObjectID
        mConstants!Description = Trim(txtLongDescription.Text)
        mConstants!Value = Val(txtValue.Text)
        mConstants!Auxiliary_Field = 1
        mConstants.Update
        mConstants.Bookmark = mConstants.LastModified
        mSelectedType = entConstant
        mSelectedID = mConstants!Constant_ID
        ShowTree
        mCurrentAction = ecaEditConstant
        mConstants.Index = "PrimaryKey"
    ElseIf mCurrentAction = ecaEditConstant Then
        iStr = Trim(txtName.Text)
        If mSelectedType <> entConstant Then Err.Raise 1234
        mConstants.Index = "PrimaryKey"
        mConstants.Seek "=", mSelectedID
        If mConstants.NoMatch Then Err.Raise 1234
        If mConstants!Name <> Trim$(txtName.Text) Then
            mConstants.Index = "Name"
            mConstants.Seek "=", Trim(txtName.Text)
            If Not mConstants.NoMatch Then
                MsgBox "There is already a Constant with that name", vbExclamation
                Exit Sub
            End If
            mConstants.Index = "PrimaryKey"
            mConstants.Seek "=", mSelectedID
        End If
        mConstants.Edit
        If mConstants!Name <> iStr Then
            mConstants!Name = iStr
            iUpdateObjects = True
        End If
        mConstants!Name = iStr
        mConstants!Description = Trim(txtLongDescription.Text)
        mConstants!Value = Val(txtValue.Text)
        mConstants.Update
        If iUpdateObjects Then ShowTree
        mConstants.Bookmark = mConstants.LastModified
    ElseIf mCurrentAction = ecaEditIntroduction Then
        If Not mGeneral_Information Is Nothing Then
            mGeneral_Information.Seek "=", "Introduction"
            If mGeneral_Information.NoMatch Then
                Err.Raise 12351
            Else
                mGeneral_Information.Edit
                mGeneral_Information!Value = txtLongDescription.Text
                mGeneral_Information.Update
            End If
        End If
    ElseIf mCurrentAction = ecaEditEndNotes Then
        If Not mGeneral_Information Is Nothing Then
            mGeneral_Information.Seek "=", "EndNotes"
            If mGeneral_Information.NoMatch Then
                Err.Raise 12352
            Else
                mGeneral_Information.Edit
                mGeneral_Information!Value = txtLongDescription.Text
                mGeneral_Information.Update
            End If
        End If
    End If
End Sub

Private Property Let CurrentAction(nValue As ECurrentAction)
    mCurrentAction = nValue
    If mCurrentAction = ecaAddClass Then
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        lblCurrentAction.Caption = "New class:"
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaAddControl Then
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        lblCurrentAction.Caption = "New control:"
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaAddEnum Then
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        lblCurrentAction.Caption = "New Enum:"
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaEditClass Then
        lblCurrentAction.Caption = "Class:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mClasses!Name
        txtShortDescription.Text = mClasses!Short_Description
        txtLongDescription.Text = mClasses!Long_Description
    ElseIf mCurrentAction = ecaEditControl Then
        lblCurrentAction.Caption = "Control:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mControls!Name
        txtShortDescription.Text = mControls!Short_Description
        txtLongDescription.Text = mControls!Long_Description
    ElseIf mCurrentAction = ecaEditEnum Then
        lblCurrentAction.Caption = "Enum:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mEnums!Name
        txtLongDescription.Text = mEnums!Description
    ElseIf mCurrentAction = ecaAddProperty Then
        lblCurrentAction.Caption = "New Property:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaAddMethod Then
        lblCurrentAction.Caption = "New Method:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaAddEvent Then
        lblCurrentAction.Caption = "New Event:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaEditProperty Then
        lblCurrentAction.Caption = "Property:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mProperties!Name
        txtShortDescription.Text = mProperties!Short_Description
        txtLongDescription.Text = mProperties!Long_Description
        txtParamsInfo.Text = mProperties!Params_Info
    ElseIf mCurrentAction = ecaEditMethod Then
        lblCurrentAction.Caption = "Method:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mMethods!Name
        txtShortDescription.Text = mMethods!Short_Description
        txtLongDescription.Text = mMethods!Long_Description
        txtParamsInfo.Text = mMethods!Params_Info
    ElseIf mCurrentAction = ecaEditEvent Then
        lblCurrentAction.Caption = "Event:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mEvents!Name
        txtShortDescription.Text = mEvents!Short_Description
        txtLongDescription.Text = mEvents!Long_Description
        txtParamsInfo.Text = mEvents!Params_Info
    ElseIf mCurrentAction = ecaAddConstant Then
        lblCurrentAction.Caption = "New Constant:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        SetFocusTo txtName
    ElseIf mCurrentAction = ecaEditConstant Then
        lblCurrentAction.Caption = "Constant:"
        ClearControlsEditZone
        ControlsEditZoneVisible = True
        txtName.Text = mConstants!Name
        txtLongDescription.Text = mConstants!Description
        txtValue.Text = mConstants!Value
    ElseIf mCurrentAction = ecaEditIntroduction Then
        If Not mGeneral_Information Is Nothing Then
            lblCurrentAction.Caption = "Intoduction:"
            ClearControlsEditZone
            ControlsEditZoneVisible = True
            mGeneral_Information.Seek "=", "Introduction"
            If Not mGeneral_Information.NoMatch Then
                txtLongDescription.Text = mGeneral_Information!Value
            Else
                Err.Raise 1239
            End If
        End If
    ElseIf mCurrentAction = ecaEditEndNotes Then
        If Not mGeneral_Information Is Nothing Then
            lblCurrentAction.Caption = "EndNotes"
            ClearControlsEditZone
            ControlsEditZoneVisible = True
            mGeneral_Information.Seek "=", "EndNotes"
            If Not mGeneral_Information.NoMatch Then
                txtLongDescription.Text = mGeneral_Information!Value
            Else
                Err.Raise 1239
            End If
        End If
    Else
        lblCurrentAction.Caption = ""
    End If
End Property

Private Property Get CurrentAction() As ECurrentAction
    CurrentAction = mCurrentAction
End Property

Private Sub ShowTree(Optional nSelectCurrentNode As Boolean = True)
    Dim iNodeClasses As Node
    Dim iNodeControls As Node
    Dim iNodeEnums As Node
    Dim iNode As Node
    Dim iNode2 As Node
    Dim iMembers As Recordset
    Dim iFVNodeKey As String
    Dim iFVNode As Node
    
    mShowingTree = True
    Set iFVNode = GetTreeViewFirstVisibleNode(trv1)
    If Not iFVNode Is Nothing Then
        iFVNodeKey = iFVNode.Key
    End If
    trv1.Visible = False
    trv1.Nodes.Clear
    Call trv1.Nodes.Add(, , CStr(enIntroduction) & "|0", "Introduction")
    Set iNodeControls = trv1.Nodes.Add(, , CStr(entControlsParent) & "|0", "[Controls]")
    Set iNodeClasses = trv1.Nodes.Add(, , CStr(entClassesParent) & "|0", "[Classes]")
    Set iNodeEnums = trv1.Nodes.Add(, , CStr(entEnumsParent) & "|0", "[Enums]")
    Call trv1.Nodes.Add(, , CStr(entEndNotes) & "|0", "End Notes")
    
    trv1.Enabled = True
    mnuData.Enabled = True
    mnuReport.Enabled = True
    mnuComponentProperties.Enabled = True
    On Error Resume Next
    mnuListOrphanMembers.Visible = False
    mnuDeleteOrphanMembers.Visible = False
    mnuLoadFromOrphanMember2.Visible = False
    mnuSetMethodToExistentDefinition.Visible = False
    
    On Error GoTo 0
    mnuListOrphanMembers.Enabled = False
    mnuDeleteOrphanMembers.Enabled = False
    If mClasses Is Nothing Then
        trv1.Enabled = False
        mnuData.Enabled = False
        mnuReport.Enabled = False
        mnuComponentProperties.Enabled = False
        GoTo TheExit
    End If
    
    If mClasses.RecordCount > 0 Then
        mClasses.MoveFirst
        Do Until mClasses.EOF
            Set iNode = trv1.Nodes.Add(iNodeClasses, tvwChild, CStr(entClass) & "a|" & CStr(mClasses!Class_ID), mClasses!Name)
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Properties, Properties WHERE (Classes_Properties.Property_ID = Properties.Property_ID) AND (Classes_Properties.Class_ID = " & mClasses!Class_ID & ") ORDER BY Properties.Name")
            If Not iMembers.EOF Then
                Set iNode2 = trv1.Nodes.Add(iNode, tvwChild, CStr(entPropertiesParent) & "a|" & CStr(mClasses!Class_ID), "[Properties]")
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entProperty) & "|" & CStr(iMembers.Fields("Properties.Property_ID") & "a|" & mClasses!Class_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
                iNode2.Text = iNode2.Text & "  (" & iNode2.Children & ")"
            End If
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Methods, Methods WHERE (Classes_Methods.Method_ID = Methods.Method_ID) AND (Classes_Methods.Class_ID = " & mClasses!Class_ID & ") ORDER BY Methods.Name")
            If Not iMembers.EOF Then
                Set iNode2 = trv1.Nodes.Add(iNode, tvwChild, CStr(entMethodsParent) & "a|" & CStr(mClasses!Class_ID), "[Methods]")
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entMethod) & "|" & CStr(iMembers.Fields("Methods.Method_ID") & "a|" & mClasses!Class_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
                iNode2.Text = iNode2.Text & "  (" & iNode2.Children & ")"
            End If
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Classes_Events, Events WHERE (Classes_Events.Event_ID = Events.Event_ID) AND (Classes_Events.Class_ID = " & mClasses!Class_ID & ") ORDER BY Events.Name")
            If Not iMembers.EOF Then
                Set iNode2 = trv1.Nodes.Add(iNode, tvwChild, CStr(entEventsParent) & "a|" & CStr(mClasses!Class_ID), "[Events]")
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entEvent) & "|" & CStr(iMembers.Fields("Events.Event_ID") & "a|" & mClasses!Class_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
                iNode2.Text = iNode2.Text & "  (" & iNode2.Children & ")"
            End If
            mClasses.MoveNext
        Loop
    End If
    
    If mControls.RecordCount > 0 Then
        mControls.MoveFirst
        Do Until mControls.EOF
            Set iNode = trv1.Nodes.Add(iNodeControls, tvwChild, CStr(entControl) & "o|" & CStr(mControls!control_ID), mControls!Name)
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Properties, Properties WHERE (Controls_Properties.Property_ID = Properties.Property_ID) AND (Controls_Properties.Control_ID = " & mControls!control_ID & ") ORDER BY Properties.Name")
            If Not iMembers.EOF Then
                Set iNode2 = trv1.Nodes.Add(iNode, tvwChild, CStr(entPropertiesParent) & "o|" & CStr(mControls!control_ID), "[Properties]")
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entProperty) & "|" & CStr(iMembers.Fields("Properties.Property_ID") & "o|" & mControls!control_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
                iNode2.Text = iNode2.Text & "  (" & iNode2.Children & ")"
            End If
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Methods, Methods WHERE (Controls_Methods.Method_ID = Methods.Method_ID) AND (Controls_Methods.Control_ID = " & mControls!control_ID & ") ORDER BY Methods.Name")
            If Not iMembers.EOF Then
                Set iNode2 = trv1.Nodes.Add(iNode, tvwChild, CStr(entMethodsParent) & "o|" & CStr(mControls!control_ID), "[Methods]")
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entMethod) & "|" & CStr(iMembers.Fields("Methods.Method_ID") & "o|" & mControls!control_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
                iNode2.Text = iNode2.Text & "  (" & iNode2.Children & ")"
            End If
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Controls_Events, Events WHERE (Controls_Events.Event_ID = Events.Event_ID) AND (Controls_Events.Control_ID = " & mControls!control_ID & ") ORDER BY Events.Name")
            If Not iMembers.EOF Then
                Set iNode2 = trv1.Nodes.Add(iNode, tvwChild, CStr(entEventsParent) & "o|" & CStr(mControls!control_ID), "[Events]")
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entEvent) & "|" & CStr(iMembers.Fields("Events.Event_ID") & "o|" & mControls!control_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
                iNode2.Text = iNode2.Text & "  (" & iNode2.Children & ")"
            End If
            mControls.MoveNext
        Loop
    End If
    
    If mEnums.RecordCount > 0 Then
        mEnums.Index = "Name"
        mEnums.MoveFirst
        Do Until mEnums.EOF
            Set iNode2 = trv1.Nodes.Add(iNodeEnums, tvwChild, CStr(entEnum) & "e|" & CStr(mEnums!Enum_ID), mEnums!Name)
            Set iMembers = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE (Constants.Enum_ID = " & mEnums!Enum_ID & ") ORDER BY " & IIf(mEnums!Ordered_By_Value, "Constants.Value", "Constants.Name"))
            If Not iMembers.EOF Then
                iMembers.MoveFirst
                Do Until iMembers.EOF
                    trv1.Nodes.Add iNode2, tvwChild, CStr(entConstant) & "|" & CStr(iMembers!Constant_ID), iMembers!Name
                    iMembers.MoveNext
                Loop
            End If
            mEnums.MoveNext
        Loop
        mEnums.Index = "PrimaryKey"
    End If
    
    mnuListOrphanMembers.Enabled = mFileImported
    mnuDeleteOrphanMembers.Enabled = mnuListOrphanMembers.Enabled
    mnuLoadFromOrphanMember2.Enabled = mnuListOrphanMembers.Enabled
    mnuSetMethodToExistentDefinition.Enabled = mnuListOrphanMembers.Enabled
    
    On Error Resume Next
    mnuListOrphanMembers.Visible = mnuListOrphanMembers.Enabled
    mnuDeleteOrphanMembers.Visible = mnuListOrphanMembers.Enabled
    mnuLoadFromOrphanMember2.Visible = mnuListOrphanMembers.Enabled
    mnuSetMethodToExistentDefinition.Visible = mnuListOrphanMembers.Enabled
    On Error GoTo 0
    
TheExit:
    iNodeClasses.Expanded = True
    iNodeControls.Expanded = True
    iNodeEnums.Expanded = True
    
    trv1.Visible = True
    mShowingTree = False
    
    On Error Resume Next
    SetTreeViewFirstVisibleNode trv1, GetNodeByKey(iFVNodeKey)
    On Error GoTo 0
    If nSelectCurrentNode Then SelectCurrentNode
End Sub

Private Sub SelectCurrentNode()
    Dim iExists As Boolean
    
    If mSelectedID <> 0 Then
        If mSelectedSecondaryID > 0 Then
            If TreeViewNodeExists(trv1, CStr(mSelectedType) & "|" & CStr(mSelectedID) & "|" & CStr(mSelectedSecondaryID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "|" & CStr(mSelectedSecondaryID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "|" & CStr(mSelectedSecondaryID)).Selected = True
            ElseIf TreeViewNodeExists(trv1, CStr(mSelectedType) & "|" & CStr(mSelectedID) & "a|" & CStr(mSelectedSecondaryID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "a|" & CStr(mSelectedSecondaryID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "a|" & CStr(mSelectedSecondaryID)).Selected = True
            ElseIf TreeViewNodeExists(trv1, CStr(mSelectedType) & "|" & CStr(mSelectedID) & "o|" & CStr(mSelectedSecondaryID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "o|" & CStr(mSelectedSecondaryID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "o|" & CStr(mSelectedSecondaryID)).Selected = True
            ElseIf TreeViewNodeExists(trv1, CStr(mSelectedType) & "|" & CStr(mSelectedID) & "e|" & CStr(mSelectedSecondaryID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "e|" & CStr(mSelectedSecondaryID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID) & "e|" & CStr(mSelectedSecondaryID)).Selected = True
            Else
                mSelectedType = entNone
                mSelectedID = 0
            End If
        Else
            If TreeViewNodeExists(trv1, CStr(mSelectedType) & "|" & CStr(mSelectedID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "|" & CStr(mSelectedID)).Selected = True
            ElseIf TreeViewNodeExists(trv1, CStr(mSelectedType) & "a|" & CStr(mSelectedID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "a|" & CStr(mSelectedID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "a|" & CStr(mSelectedID)).Selected = True
            ElseIf TreeViewNodeExists(trv1, CStr(mSelectedType) & "o|" & CStr(mSelectedID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "o|" & CStr(mSelectedID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "o|" & CStr(mSelectedID)).Selected = True
            ElseIf TreeViewNodeExists(trv1, CStr(mSelectedType) & "e|" & CStr(mSelectedID)) Then
                trv1.Nodes.Item(CStr(mSelectedType) & "e|" & CStr(mSelectedID)).EnsureVisible
                trv1.Nodes.Item(CStr(mSelectedType) & "e|" & CStr(mSelectedID)).Selected = True
            Else
                mSelectedType = entNone
                mSelectedID = 0
            End If
        End If
    ElseIf mSelectedType > 0 Then
        If TreeViewNodeExists(trv1, CStr(mSelectedType) & "|0") Then
            trv1.Nodes.Item(CStr(mSelectedType) & "|0").EnsureVisible
            trv1.Nodes.Item(CStr(mSelectedType) & "|0").Selected = True
        End If
    End If
End Sub

Private Function TreeViewNodeExists(trv As TreeView, nKey As String) As Boolean
    Dim n As Node
    
    For Each n In trv.Nodes
        If n.Key = nKey Then
            TreeViewNodeExists = True
            Exit For
        End If
    Next
End Function

Private Sub SetFocusTo(nControl As Control)
    On Error Resume Next
    nControl.SetFocus
End Sub

Private Function GetName() As String
    Dim c As Long
    Dim iCurRec As Recordset
    
    GetName = Trim$(txtName.Text)
    If GetName = "" Then
        GetName = "Unnamed"
        Set iCurRec = CurRec
        iCurRec.Index = "Name"
        iCurRec.Seek "=", GetName
        c = 1
        Do Until iCurRec.NoMatch
            c = c + 1
            GetName = "Unnamed " & CStr(c)
            iCurRec.Seek "=", GetName
        Loop
    End If
End Function

Private Function CurRec() As Recordset
    If (CurrentAction = ecaAddClass) Or (CurrentAction = ecaEditClass) Then
        Set CurRec = mClasses
    ElseIf (CurrentAction = ecaAddControl) Or (CurrentAction = ecaEditControl) Then
        Set CurRec = mControls
    ElseIf (CurrentAction = ecaAddEnum) Or (CurrentAction = ecaEditEnum) Then
        Set CurRec = mEnums
    ElseIf (CurrentAction = ecaAddProperty) Or (CurrentAction = ecaEditProperty) Then
        Set CurRec = mProperties
    ElseIf (CurrentAction = ecaAddMethod) Or (CurrentAction = ecaEditMethod) Then
        Set CurRec = mMethods
    ElseIf (CurrentAction = ecaAddEvent) Or (CurrentAction = ecaEditEvent) Then
        Set CurRec = mEvents
    ElseIf (CurrentAction = ecaAddConstant) Or (CurrentAction = ecaEditConstant) Then
        Set CurRec = mConstants
    End If
End Function

Private Function GetCurrentObjectTypePluralStr() As String
    Dim iSelectedType As ENodeType
    Dim iNode As Node
    
    iSelectedType = mSelectedType
    Set iNode = trv1.SelectedItem
    Do Until GetCurrentObjectTypePluralStr <> ""
        If (iSelectedType = entClass) Then
            GetCurrentObjectTypePluralStr = "Classes"
        ElseIf (iSelectedType = entClassesParent) Then
            GetCurrentObjectTypePluralStr = "Classes"
        ElseIf (iSelectedType = entControl) Then
            GetCurrentObjectTypePluralStr = "Controls"
        ElseIf (iSelectedType = entControlsParent) Then
            GetCurrentObjectTypePluralStr = "Controls"
        ElseIf (iSelectedType = entEnum) Then
            GetCurrentObjectTypePluralStr = "Enums"
        ElseIf (iSelectedType = entEnumsParent) Then
            GetCurrentObjectTypePluralStr = "Enums"
        End If
        Set iNode = iNode.Parent
        If iNode Is Nothing Then Exit Function
        iSelectedType = GetNodeType(iNode)
    Loop
End Function

Private Function GetCurrentObjectTypeSingularStr() As String
    Dim iSelectedType As ENodeType
    Dim iNode As Node
    
    iSelectedType = mSelectedType
    Set iNode = trv1.SelectedItem
    Do Until GetCurrentObjectTypeSingularStr <> ""
        If (iSelectedType = entClass) Then
            GetCurrentObjectTypeSingularStr = "Class"
        ElseIf (iSelectedType = entClassesParent) Then
            GetCurrentObjectTypeSingularStr = "Class"
        ElseIf (iSelectedType = entControl) Then
            GetCurrentObjectTypeSingularStr = "Control"
        ElseIf (iSelectedType = entControlsParent) Then
            GetCurrentObjectTypeSingularStr = "Control"
        ElseIf (iSelectedType = entEnum) Then
            GetCurrentObjectTypeSingularStr = "Enum"
        ElseIf (iSelectedType = entEnumsParent) Then
            GetCurrentObjectTypeSingularStr = "Enum"
        End If
        Set iNode = iNode.Parent
        If iNode Is Nothing Then Exit Function
        iSelectedType = GetNodeType(iNode)
    Loop
End Function

Private Function GetNodeType(nNode As Node) As ENodeType
    Dim iStrs() As String
    
    iStrs = Split(nNode.Key, "|")
    If UBound(iStrs) >= 1 Then
        GetNodeType = Val(iStrs(0))
    End If
End Function

Private Function GetNodeID(nNode As Node) As ENodeType
    Dim iStrs() As String
    
    iStrs = Split(nNode.Key, "|")
    If UBound(iStrs) >= 1 Then
        GetNodeID = Val(iStrs(1))
    End If
End Function

Private Function GetCurrentObjectID() As Long
    Dim iNode As Node
    Dim iNodeType As ENodeType
    
    Set iNode = trv1.SelectedItem
    iNodeType = GetNodeType(iNode)
    Do Until (iNodeType = entClass) Or (iNodeType = entControl) Or (iNodeType = entEnum)
        If (iNodeType = entNone) Or (iNodeType = entClassesParent) Or (iNodeType = entControlsParent) Or (iNodeType = entEnumsParent) Then
            Exit Function
        End If
        Set iNode = iNode.Parent
        iNodeType = GetNodeType(iNode)
    Loop
    GetCurrentObjectID = GetNodeID(iNode)
End Function

Private Sub txtValue_Validate(Cancel As Boolean)
    UpdateData
End Sub

Private Sub PlaceDataControls()
    Dim iLng As Long
    
    lblCurrentAction.Move trv1.Left + trv1.Width + 120, lblCurrentAction.Top
    lblName.Move lblCurrentAction.Left, lblName.Top
    txtName.Move lblName.Left, txtName.Top, Me.ScaleWidth - lblName.Left - 120
    iLng = (trv1.Height - (lblParamsInfo.Top + lblParamsInfo.Height + 60) + 780)
    lblParamsInfo.Move txtName.Left, txtName.Top + txtName.Height + 160
    txtParamsInfo.Move txtName.Left, lblParamsInfo.Top + lblParamsInfo.Height + 60, txtName.Width, iLng * 0.25
    If (mCurrentAction = ecaAddProperty) Or (mCurrentAction = ecaAddMethod) Or (mCurrentAction = ecaAddEvent) Or (mCurrentAction = ecaEditProperty) Or (mCurrentAction = ecaEditMethod) Or (mCurrentAction = ecaEditEvent) Then
        
        lblLongDescription.Move txtName.Left, txtParamsInfo.Top + txtParamsInfo.Height + 160
        txtLongDescription.Move txtName.Left, txtParamsInfo.Top + txtParamsInfo.Height + 470, txtName.Width, iLng * 0.4
        lblShortDescription.Move txtName.Left, txtLongDescription.Top + txtLongDescription.Height + 160
        iLng = txtLongDescription.Top + txtLongDescription.Height + 470
        txtShortDescription.Move txtName.Left, iLng, txtName.Width, IIf(trv1.Top + trv1.Height - iLng < 250, 250, trv1.Top + trv1.Height - iLng)
        txtValue.Move txtShortDescription.Left, txtShortDescription.Top, txtShortDescription.Width
    ElseIf (mCurrentAction = ecaEditIntroduction) Or (mCurrentAction = ecaEditEndNotes) Then
        lblLongDescription.Move txtName.Left, lblName.Top
        txtLongDescription.Move txtName.Left, lblLongDescription.Top + lblLongDescription.Height + 60, txtName.Width, trv1.Height - (lblLongDescription.Top + lblLongDescription.Height - 60)
    Else
        lblLongDescription.Move txtName.Left, lblParamsInfo.Top
        txtLongDescription.Move txtName.Left, txtParamsInfo.Top, txtName.Width, iLng * 0.65
        lblShortDescription.Move txtName.Left, txtLongDescription.Top + txtLongDescription.Height + 160
        iLng = txtLongDescription.Top + txtLongDescription.Height + 470
        txtShortDescription.Move txtName.Left, iLng, txtName.Width, trv1.Top + trv1.Height - iLng
        txtValue.Move txtShortDescription.Left, txtShortDescription.Top, txtShortDescription.Width
    End If
    cmdLongDescriptionMenu.Move (txtLongDescription.Left + txtLongDescription.Width) - cmdLongDescriptionMenu.Width - 250, txtLongDescription.Top - cmdLongDescriptionMenu.Height - 60
    cmdBold.Move txtLongDescription.Left + lblLongDescription.Width + 500, cmdLongDescriptionMenu.Top
    cmdLink.Move cmdBold.Left + 500, cmdLongDescriptionMenu.Top
    cmdReference.Move cmdLink.Left + 500, cmdLongDescriptionMenu.Top
    cmdAppliesTo.Move (txtName.Left + txtName.Width) - cmdAppliesTo.Width - cmdAppliesTo - 250, txtName.Top - cmdAppliesTo.Height - 60
    lblAppliesTo.Move cmdAppliesTo.Left - lblAppliesTo.Width - 60, cmdAppliesTo.Top + (cmdAppliesTo.Height - lblAppliesTo.Height) / 2
End Sub

Private Function GetTypeName(ByVal nVarTypeInfo As VarTypeInfo, Optional nGenericType As Boolean = False) As String
    Dim iStr As String
    Dim iVarType As Long
    Dim iKnownObjectType As Boolean
    
    iVarType = nVarTypeInfo.VarType
    If iVarType <> 0 Then
        Select Case (iVarType And &HFF&)
            Case VT_BOOL
                iStr = "Boolean"
            Case VT_BSTR, VT_LPSTR, VT_LPWSTR
                iStr = "String"
            Case VT_DATE
                iStr = "Date"
            Case VT_INT
                iStr = "Integer"
            Case VT_VARIANT
                iStr = "Variant"
            Case VT_DECIMAL
                iStr = "Decimal"
            Case VT_I4
                iStr = "Long"
            Case VT_I2
                iStr = "Integer"
            Case VT_I8
                iStr = "Unknown"
            Case VT_SAFEARRAY
                iStr = "SafeArray"
            Case VT_CLSID
                iStr = "CLSID"
            Case VT_UINT
                iStr = "UInt"
            Case VT_UI4
'                iStr = "ULong"
                iStr = "Long"
            Case VT_UNKNOWN
                iStr = "Unknown"
            Case VT_VECTOR
                iStr = "Vector"
            Case VT_R4
                iStr = "Single"
            Case VT_R8
                iStr = "Double"
            Case VT_DISPATCH
                iStr = "Object"
            Case VT_UI1
                iStr = "Byte"
            Case VT_CY
                iStr = "Currency"
            Case VT_HRESULT
                iStr = "HRESULT" ' note if this was a function it should be a sub
            Case VT_VOID
                iStr = "Any"
            Case VT_ERROR
                iStr = "Long"
            Case Else
                iStr = "<Unsupported Variant Type"
                Select Case (iVarType And &HFF&)
                    Case VT_UI1
                        iStr = iStr & "(VT_UI1)"
                    Case VT_UI2
                        iStr = iStr & "(VT_UI2)"
                    Case VT_UI4
                        iStr = iStr & "(VT_UI4)"
                    Case VT_UI8
                        iStr = iStr & "(VT_UI8)"
                    Case VT_USERDEFINED
                        iStr = iStr & "(VT_USERDEFINED)"
                End Select
                iStr = iStr & ">"
        End Select
        If (iVarType And VT_ARRAY) = VT_ARRAY Then
            iStr = iStr & "()"
        End If
        
        GetTypeName = iStr
    Else
        On Error Resume Next
        iStr = ""
        iStr = nVarTypeInfo.TypeInfo.Name
        If Left(iStr, 1) = "_" Then
            iStr = Mid$(iStr, 2)
        End If
        iKnownObjectType = False
        Select Case iStr
            Case "Picture", "Font", "Collection", "ContainedControls", "DataObject"
                iKnownObjectType = True
        End Select
        
        If nVarTypeInfo.TypeLibInfoExternal Is Nothing Then
            On Error GoTo 0
            If nGenericType Then
                If Not iKnownObjectType Then
                    GetTypeName = "Object"
                Else
                    GetTypeName = nVarTypeInfo.TypeInfo.Name
                End If
            Else
                GetTypeName = nVarTypeInfo.TypeInfo.Name
            End If
        Else
            If (LCase$(nVarTypeInfo.TypeLibInfoExternal) = "stdole") Then
                On Error GoTo 0
                If nGenericType Then
                    If Not iKnownObjectType Then
                        GetTypeName = "Object"
                    Else
                        GetTypeName = nVarTypeInfo.TypeInfo.Name
                    End If
                Else
                    GetTypeName = nVarTypeInfo.TypeInfo.Name
                End If
            Else
                On Error GoTo 0
                If nGenericType Then
                    If Not iKnownObjectType Then
                        GetTypeName = "Object"
                    Else
                        GetTypeName = nVarTypeInfo.TypeInfo.Name
                    End If
                Else
                    GetTypeName = nVarTypeInfo.TypeInfo.Name
                End If
            End If
        End If
    End If
    If Left(GetTypeName, 1) = "_" Then
        GetTypeName = Mid$(GetTypeName, 2)
    End If
    If Not ((GetTypeName = "Font") Or (GetTypeName = "Picture")) Then
        GetTypeName = "[[" & GetTypeName & "]]"
    End If
    If nGenericType Then
        GetTypeName = Replace$(GetTypeName, "OLE_COLOR", "Long")
    
        If Not nVarTypeInfo.TypeInfo Is Nothing Then
            If (LCase$(nVarTypeInfo.TypeInfo.TypeKindString) = "enum") Then
                GetTypeName = "Long"
            End If
        End If
    End If
End Function

Private Function GetConstantName2(nEnumName As String, nValue As Long) As String
    Dim iConstants As Recordset
    
    mEnums.Index = "Name"
    mEnums.Seek "=", nEnumName
    If Not mEnums.NoMatch Then
        Set iConstants = mDatabase.OpenRecordset("SELECT * FROM Constants WHERE Enum_ID = " & mEnums!Enum_ID)
        If Not iConstants.EOF Then
            iConstants.FindFirst "Value = " & nValue
            If Not iConstants.NoMatch Then
                GetConstantName2 = iConstants!Name & " (" & CStr(nValue) & ")"
            End If
        End If
    End If
    mEnums.Index = "PrimaryKey"
    If GetConstantName2 = "" Then GetConstantName2 = CStr(nValue)
    
End Function

Private Function TxtToHTML(nTxt As String, Optional nSource As String) As String
    Dim iPos1 As Long
    Dim iPos2 As Long
    Dim iStr As String
    Dim iIndex1 As String
    Dim iIndex2 As String
    Dim iRec As Recordset
    Dim iStr2 As String
    
    Dim iClasses_Index As String
    Dim iControls_Index As String
    Dim iEnums_Index As String
    Dim iProperties_Index As String
    Dim iMethods_Index As String
    Dim iEvents_Index As String
    Dim iConstants_Index As String
    
    iClasses_Index = mClasses.Index
    iControls_Index = mControls.Index
    iEnums_Index = mControls.Index
    iProperties_Index = mProperties.Index
    iMethods_Index = mMethods.Index
    iEvents_Index = mEvents.Index
    iConstants_Index = mConstants.Index
    
    mClasses.Index = "Name"
    mControls.Index = "Name"
    mControls.Index = "Name"
    mProperties.Index = "Name"
    mMethods.Index = "Name"
    mEvents.Index = "Name"
    mConstants.Index = "Name"
    
    TxtToHTML = nTxt
    
    ' Enums and objects (Classes)
    iPos1 = InStr(TxtToHTML, "[[")
    If iPos1 > 0 Then
        iIndex1 = mEnums.Index
        mEnums.Index = "Name"
        Do Until iPos1 = 0
            iPos2 = InStr(iPos1 + 2, TxtToHTML, "]]")
            If iPos2 = 0 Then Exit Do
            iStr = Mid$(TxtToHTML, iPos1 + 2, iPos2 - iPos1 - 2)
            Set iRec = Nothing
            mEnums.Seek "=", iStr
            If Not mEnums.NoMatch Then
                Set iRec = mEnums
                iStr2 = "enumeration"
            Else
                If iIndex2 = "" Then
                    iIndex2 = mClasses.Index
                    mClasses.Index = "Name"
                End If
                mClasses.Seek "=", iStr
                If Not mClasses.NoMatch Then
                    Set iRec = mClasses
                    iStr2 = "object"
                End If
            End If
            If iRec Is Nothing Then
                TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & iStr & Mid$(TxtToHTML, iPos2 + 2)
            Else
                TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<a href=""" & LCase(iStr) & "_" & iStr2 & ".html"">" & iStr & "</a>" & Mid$(TxtToHTML, iPos2 + 2)
            End If
            iPos1 = InStr(iPos2 + 2, TxtToHTML, "[[")
        Loop
        mEnums.Index = iIndex1
        If iIndex2 <> "" Then mClasses.Index = iIndex2
    End If
    
    
    ' Controls
    iPos1 = InStr(TxtToHTML, "[c[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToHTML, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToHTML, iPos1 + 3, iPos2 - iPos1 - 3)
        mControls.Seek "=", iStr
        If mControls.NoMatch Then
            AddLinkError iStr & " control in " & nSource
        End If
        TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<a href=""" & LCase(iStr) & "_control.html"">" & iStr & "</a>" & Mid$(TxtToHTML, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToHTML, "[c[")
    Loop
    ' objects
    iPos1 = InStr(TxtToHTML, "[o[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToHTML, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToHTML, iPos1 + 3, iPos2 - iPos1 - 3)
        mClasses.Seek "=", iStr
        If mClasses.NoMatch Then
            AddLinkError iStr & " object in " & nSource
        End If
        TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<a href=""" & LCase(iStr) & "_object.html"">" & iStr & "</a>" & Mid$(TxtToHTML, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToHTML, "[o[")
    Loop
    ' Properties
    iPos1 = InStr(TxtToHTML, "[p[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToHTML, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToHTML, iPos1 + 3, iPos2 - iPos1 - 3)
        If mHTML_Mode = cdHTMLPerMethod Then
            If Not IsMemberNameUnique(1, iStr) Then
                AddToList mNonUniqueMemberNamePages, "1|" & iStr, True
            End If
        End If
        mProperties.Seek "=", iStr
        If mProperties.NoMatch Then
            AddLinkError iStr & " property in " & nSource
        End If
        TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<a href=""" & LCase(iStr) & "_property.html"">" & iStr & "</a>" & Mid$(TxtToHTML, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToHTML, "[p[")
    Loop
    ' Methods
    iPos1 = InStr(TxtToHTML, "[m[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToHTML, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToHTML, iPos1 + 3, iPos2 - iPos1 - 3)
        If mHTML_Mode = cdHTMLPerMethod Then
            If Not IsMemberNameUnique(2, iStr) Then
                AddToList mNonUniqueMemberNamePages, "2|" & iStr, True
            End If
        End If
        mMethods.Seek "=", iStr
        If mMethods.NoMatch Then
            AddLinkError iStr & " method in " & nSource
        End If
        TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<a href=""" & LCase(iStr) & "_method.html"">" & iStr & "</a>" & Mid$(TxtToHTML, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToHTML, "[m[")
    Loop
    ' Events
    iPos1 = InStr(TxtToHTML, "[e[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToHTML, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToHTML, iPos1 + 3, iPos2 - iPos1 - 3)
        If mHTML_Mode = cdHTMLPerMethod Then
            If Not IsMemberNameUnique(3, iStr) Then
                AddToList mNonUniqueMemberNamePages, "3|" & iStr, True
            End If
        End If
        mEvents.Seek "=", iStr
        If mEvents.NoMatch Then
            AddLinkError iStr & " event in " & nSource
        End If
        TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<a href=""" & LCase(iStr) & "_event.html"">" & iStr & "</a>" & Mid$(TxtToHTML, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToHTML, "[e[")
    Loop
    
    ' See also:
    If InStr(LCase$(TxtToHTML), "see also:") > 0 Then
        TxtToHTML = Replace$(TxtToHTML, "See also:", "See Also:")
        TxtToHTML = Replace$(TxtToHTML, "see also:", "See Also:")
        TxtToHTML = Replace$(TxtToHTML, "See Also:", "<h3>See Also:</h3>")
        TxtToHTML = Replace$(TxtToHTML, "<h3>See Also:</h3>" & vbCrLf, "<h3>See Also:</h3>")
    End If
    
    ' [code]
    If False Then
        iStr = LCase$(TxtToHTML)
        Do While InStr(iStr, "code]") > 0
            iPos1 = InStr(iStr, "[code]")
            If iPos1 > 0 Then
                If Mid$(TxtToHTML, iPos1 + 6, 2) = vbCrLf Then
                    TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & Mid$(TxtToHTML, iPos1 + 8)
                Else
                    TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & Mid$(TxtToHTML, iPos1 + 6)
                End If
                iStr = LCase$(TxtToHTML)
            End If
            iPos1 = InStr(iStr, "[/code]")
            If iPos1 > 0 Then
                iStr = TxtToHTML
                TxtToHTML = Left$(TxtToHTML, iPos1 - 1)
                If Mid$(iStr, iPos1 + 7, 2) = vbCrLf Then
                    TxtToHTML = TxtToHTML & Mid$(iStr, iPos1 + 9)
                Else
                    TxtToHTML = TxtToHTML & Mid$(iStr, iPos1 + 7)
                End If
            End If
            iStr = LCase$(TxtToHTML)
        Loop
    Else
        'TxtToHTML = Replace(TxtToHTML, "[code]", "<pre>")
        'TxtToHTML = Replace(TxtToHTML, "[/code]", "</pre>")
        iPos1 = InStr(TxtToHTML, "[code]")
        Do Until iPos1 = 0
            iPos2 = InStr(iPos1 + 5, TxtToHTML, "[/code]")
            If iPos2 > 0 Then
                iStr = Mid$(TxtToHTML, iPos1 + 6, iPos2 - iPos1 - 6)
                iStr = Replace(iStr, vbCr, "")
                TxtToHTML = Left$(TxtToHTML, iPos1 - 1) & "<pre>" & iStr & "</pre>" & Mid$(TxtToHTML, iPos2 + 7)
            End If
            iPos1 = InStr(TxtToHTML, "[code]")
        Loop
    End If
    
    ' done at the end so it doesn't interfere with other replacements that must be done before
    TxtToHTML = Replace(TxtToHTML, vbCrLf, "<br>" & vbCrLf)
    TxtToHTML = Replace(TxtToHTML, vbTab, "&nbsp&nbsp&nbsp&nbsp")

    mClasses.Index = iClasses_Index
    mControls.Index = iControls_Index
    mControls.Index = iEnums_Index
    mProperties.Index = iProperties_Index
    mMethods.Index = iMethods_Index
    mEvents.Index = iEvents_Index
    mConstants.Index = iConstants_Index

End Function

Private Function GetAppliesTo(nType As Long, nMemberID As Long, Optional nPlainText As Boolean) As String
    Dim iRec As Recordset
    Dim iTRec As Recordset
    
    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Controls_" & mMemberType_p(nType) & " WHERE (" & mMemberType_s(nType) & "_ID = " & nMemberID & ")")
    If iRec.RecordCount > 0 Then
        iRec.MoveFirst
        Do Until iRec.EOF
            mControls.Seek "=", iRec!control_ID
            If mControls.NoMatch Then Stop
            If GetAppliesTo <> "" Then GetAppliesTo = GetAppliesTo & ", "
            GetAppliesTo = GetAppliesTo & IIf(nPlainText, "", "[c[") & mControls!Name & IIf(nPlainText, "", "]]")
            iRec.MoveNext
        Loop
    End If
    Set iRec = mDatabase.OpenRecordset("SELECT * FROM Classes_" & mMemberType_p(nType) & " WHERE (" & mMemberType_s(nType) & "_ID = " & nMemberID & ")")
    If iRec.RecordCount > 0 Then
        iRec.MoveFirst
        Do Until iRec.EOF
            mClasses.Seek "=", iRec!Class_ID
            If mClasses.NoMatch Then Stop
            If GetAppliesTo <> "" Then GetAppliesTo = GetAppliesTo & ", "
            GetAppliesTo = GetAppliesTo & IIf(nPlainText, "", "[o[") & mClasses!Name & IIf(nPlainText, "", "]]") & " object"
            iRec.MoveNext
        Loop
    End If
End Function

Private Sub SetFontsTo(nSize As Single)
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        Select Case TypeName(iCtl)
            Case "TextBox", "TreeView"
                iCtl.Font.Size = nSize * mFontPropertion
        End Select
    Next
End Sub

Private Function TxtToRTF(nTxt As String) As String
    Dim iPos1 As Long
    Dim iPos2 As Long
    Dim iPos3 As Long
    Dim iPos4 As Long
    Dim iStr As String
    Dim iIndex1 As String
    Dim iIndex2 As String
    Dim iRec As Recordset
    Dim iStr2 As String
    Dim iStrLeft As String
    Dim iStrRight As String
    
    TxtToRTF = nTxt
    TxtToRTF = Replace$(TxtToRTF, "{", "\{")
    TxtToRTF = Replace$(TxtToRTF, "}", "\}")
    
    iPos1 = InStr(TxtToRTF, "<b>[[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 2, TxtToRTF, "]]</b>")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 5, iPos2 - iPos1 - 5)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 6)
        iPos1 = InStr(iPos2 + 6, TxtToRTF, "[[")
    Loop
    
    ' Enums and objects (Classes)
    iPos1 = InStr(TxtToRTF, "[[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 2, TxtToRTF, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 2, iPos2 - iPos1 - 2)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToRTF, "[[")
    Loop
    
    ' Controls
    iPos1 = InStr(TxtToRTF, "[c[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToRTF, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 3, iPos2 - iPos1 - 3)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToRTF, "[c[")
    Loop
    ' objects
    iPos1 = InStr(TxtToRTF, "[o[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToRTF, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 3, iPos2 - iPos1 - 3)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToRTF, "[o[")
    Loop
    ' Properties
    iPos1 = InStr(TxtToRTF, "[p[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToRTF, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 3, iPos2 - iPos1 - 3)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToRTF, "[p[")
    Loop
    ' Methods
    iPos1 = InStr(TxtToRTF, "[m[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToRTF, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 3, iPos2 - iPos1 - 3)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToRTF, "[m[")
    Loop
    ' Events
    iPos1 = InStr(TxtToRTF, "[e[")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 3, TxtToRTF, "]]")
        If iPos2 = 0 Then Exit Do
        iStr = Mid$(TxtToRTF, iPos1 + 3, iPos2 - iPos1 - 3)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos2 + 2)
        iPos1 = InStr(iPos2 + 2, TxtToRTF, "[e[")
    Loop
    
    ' See also:
    iPos1 = InStr(LCase$(TxtToRTF), "see also:")
    If iPos1 > 0 Then
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1)
    End If
    
    ' [code]
    iPos1 = InStr(TxtToRTF, "[code]")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 5, TxtToRTF, "[/code]")
        If iPos2 > 0 Then
'            If InStr(TxtToRTF, "control and a CommandButton to a form and paste this code") > 0 Then Stop
            iStr = Mid$(TxtToRTF, iPos1 + 6, iPos2 - iPos1 - 8)
            iPos3 = InStr(iStr, "'")
            Do While iPos3 > 0
                iPos4 = InStr(iPos3 + 1, iStr, vbCrLf)
                If (iPos4 > 0) And (iPos4 <= iPos2) Then
                    iStr = Left$(iStr, iPos3 - 1) & "{\cf3 " & Mid$(iStr, iPos3, iPos4 - iPos3) & "} " & Mid$(iStr, iPos4)
                End If
                iPos3 = InStr(iPos4 + 3, iStr, "'")
            Loop
            If Left$(iStr, 2) <> vbCrLf Then
                iStr = vbCrLf & iStr
            End If
            iStr = iStr & vbCrLf
            
            iStrLeft = Left$(TxtToRTF, iPos1 - 1)
            
            Do Until Right$(iStrLeft, 2) <> vbCrLf
                iStrLeft = Left$(iStrLeft, Len(iStrLeft) - 2)
            Loop
            iStrLeft = iStrLeft & vbCrLf
            
            iStrRight = Mid$(TxtToRTF, iPos2 + 7)
            Do Until Left$(iStrRight, 2) <> vbCrLf
                iStrRight = Mid$(iStrRight, 3)
            Loop
            iStrRight = vbCrLf & iStrRight
            
            TxtToRTF = iStrLeft & "{\pard\cf2\f1\li996 " & iStr & "\cf1\f0 }" & iStrRight
        End If
        iPos1 = InStr(TxtToRTF, "[code]")
    Loop
    Do Until Right$(TxtToRTF, 2) <> vbCrLf
        TxtToRTF = Left$(TxtToRTF, Len(TxtToRTF) - 2)
    Loop
    
    ' done at the end so it doesn't interfere with other replacements that must be done before
    TxtToRTF = Replace(TxtToRTF, vbCrLf, "\par " & vbCrLf)
    TxtToRTF = cRTFHeaders & Replace(TxtToRTF, vbTab, "\tab ")
    
    TxtToRTF = TxtToRTF & "}"
    
    TxtToRTF = Replace(TxtToRTF, "</b>", "\b0 ")
    TxtToRTF = Replace(TxtToRTF, "<b>", "\b ")
    TxtToRTF = Replace(TxtToRTF, "</u>", "\ul0 ")
    TxtToRTF = Replace(TxtToRTF, "<u>", "\ul ")
    
    ' <a href="
    iPos1 = InStr(TxtToRTF, "<a href=""")
    Do Until iPos1 = 0
        iPos2 = InStr(iPos1 + 9, TxtToRTF, """>")
        If iPos2 = 0 Then Exit Do
        iPos3 = InStr(iPos2 + 3, TxtToRTF, "</a>")
        If iPos3 = 0 Then iPos3 = iPos2
        iStr = Mid$(TxtToRTF, iPos1 + 9, iPos2 - iPos1 - 9)
        TxtToRTF = Left$(TxtToRTF, iPos1 - 1) & RTFBold(iStr) & Mid$(TxtToRTF, iPos3 + 4)
        iPos1 = InStr(iPos2 + 3, TxtToRTF, "<a href=""")
    Loop
    
End Function

Private Sub AddRTF(nText As String)
    Dim iBold As Boolean
    Dim iUnderline As Boolean
    Dim iItalic As Boolean
    Dim iColor As Long
    Dim iIndent As Single
    Dim iFontSize As Single
    Dim iFontName As String
    
    iFontName = rtbAux.SelFontName
    iBold = rtbAux.SelBold
    iUnderline = rtbAux.SelUnderline
    iItalic = rtbAux.SelItalic
    iColor = rtbAux.SelColor
    iIndent = rtbAux.SelIndent
    iFontSize = rtbAux.SelFontSize
    
    rtbAux.SelRTF = TxtToRTF(nText)
    
    rtbAux.SelFontName = iFontName
    rtbAux.SelBold = iBold
    rtbAux.SelUnderline = iUnderline
    rtbAux.SelItalic = iItalic
    rtbAux.SelColor = iColor
    rtbAux.SelIndent = iIndent
    rtbAux.SelFontSize = iFontSize
    
End Sub

Private Sub AddToTOC(nItem As String, nLevel As Long)
    mTOC_Index = mTOC_Index + 1
    If mTOC_Index > mTOC_Ub Then
        mTOC_Ub = mTOC_Ub + 100
        ReDim Preserve mTOCItems(mTOC_Ub)
        ReDim Preserve mTOCItems_Level(mTOC_Ub)
        ReDim Preserve mTOCItems_Page(mTOC_Ub)
    End If
    mTOCItems(mTOC_Index) = nItem
    mTOCItems_Level(mTOC_Index) = nLevel
    mTOCItems_Page(mTOC_Index) = Printer.Page
End Sub

Private Sub DrawTOCDots(ByVal nLeft As Long, ByVal nRight As Long)
    Dim c As Long
    Dim iTh As Long
    Dim iCY As Long
    
    nLeft = Round(nLeft / 100) * 100
    nRight = Int(nRight / 100) * 100
    
    iCY = Printer.CurrentY
    Printer.DrawWidth = 10
    iTh = Printer.TextHeight("l") * 0.3
    For c = nLeft To nRight Step 100
        Printer.PSet (c, iCY - iTh)
    Next
    Printer.CurrentY = iCY
End Sub

Private Function GetHTMLHeadSection(nTitle As String, nDescription As String, Optional nAdd As String) As String
    GetHTMLHeadSection = mHTML_HeadSection & vbCrLf & vbCrLf
    GetHTMLHeadSection = Replace$(GetHTMLHeadSection, "<!--[PAGE_TITLE]-->", "<title>" & nTitle & "</title>")
    GetHTMLHeadSection = Replace$(GetHTMLHeadSection, "<!--[PAGE_DESCRIPTION]-->", "<meta name=""DESCRIPTION"" content=""" & nTitle & """ />")
    If mExternalCSS Then
        GetHTMLHeadSection = Replace$(GetHTMLHeadSection, "<!--[STYLESHEET_INFO]-->", "<link rel=""stylesheet"" type=""text/css"" href=""styles.css"">")
    Else
        GetHTMLHeadSection = Replace$(GetHTMLHeadSection, "<!--[STYLESHEET_INFO]-->", "<style>" & vbCrLf & mHTML_StyleSheet & vbCrLf & "</style>")
    End If
    If nAdd <> "" Then
        GetHTMLHeadSection = Replace$(GetHTMLHeadSection, "</head>", nAdd & vbCrLf & "</head>", , , vbTextCompare)
    End If
End Function

Private Sub AddLinkError(nError As String)
    ReDim Preserve mLinkErrors(UBound(mLinkErrors) + 1)
    mLinkErrors(UBound(mLinkErrors)) = nError
End Sub

Private Sub UpdateCurrentSelected()
    Dim iStrs() As String
    Static sLastKey As String
    
    If mShowingTree Then Exit Sub
    
    If trv1.SelectedItem Is Nothing Then
        If trv1.Nodes.Count = 0 Then Exit Sub
        trv1.Nodes(1).Selected = True
    End If
    
    If trv1.SelectedItem.Key <> sLastKey Then
        sLastKey = trv1.SelectedItem.Key
    
        mSelectedType = entNone
        mSelectedID = 0
        mSelectedSecondaryID = 0
        
    
        iStrs = Split(trv1.SelectedItem.Key, "|")
        If UBound(iStrs) >= 1 Then
            mSelectedType = Val(iStrs(0))
            mSelectedID = Val(iStrs(1))
            If UBound(iStrs) > 1 Then
                mSelectedSecondaryID = Val(iStrs(2))
            End If
        End If
        
        mnuDataDelete.Enabled = False
        mnuDataAdd.Enabled = True
        If Not mDeletingNode Then ControlsEditZoneVisible = False
        
        If mSelectedType = entClass Then
            mClasses.Index = "PrimaryKey"
            mClasses.Seek "=", mSelectedID
            If Not mClasses.NoMatch Then
                CurrentAction = ecaEditClass
            End If
            mnuDataDelete.Enabled = True
            mnuDataAdd.Enabled = False
        ElseIf mSelectedType = entControl Then
            mControls.Index = "PrimaryKey"
            mControls.Seek "=", mSelectedID
            If Not mControls.NoMatch Then
                CurrentAction = ecaEditControl
            End If
            mnuDataDelete.Enabled = True
            mnuDataAdd.Enabled = False
        ElseIf mSelectedType = entEnum Then
            mEnums.Index = "PrimaryKey"
            mEnums.Seek "=", mSelectedID
            If Not mEnums.NoMatch Then
                CurrentAction = ecaEditEnum
            End If
            mnuDataDelete.Enabled = True
    '        mnuDataAdd.enabled = False
        ElseIf mSelectedType = entProperty Then
            mProperties.Index = "PrimaryKey"
            mProperties.Seek "=", mSelectedID
            If Not mProperties.NoMatch Then
                CurrentAction = ecaEditProperty
            End If
            mnuDataDelete.Enabled = True
            mnuDataAdd.Enabled = False
        ElseIf mSelectedType = entMethod Then
            mMethods.Index = "PrimaryKey"
            mMethods.Seek "=", mSelectedID
            If Not mMethods.NoMatch Then
                CurrentAction = ecaEditMethod
            End If
            mnuDataDelete.Enabled = True
            mnuDataAdd.Enabled = False
        ElseIf mSelectedType = entEvent Then
            mEvents.Index = "PrimaryKey"
            mEvents.Seek "=", mSelectedID
            If Not mEvents.NoMatch Then
                CurrentAction = ecaEditEvent
            End If
            mnuDataDelete.Enabled = True
            mnuDataAdd.Enabled = False
        ElseIf mSelectedType = entConstant Then
            mConstants.Index = "PrimaryKey"
            mConstants.Seek "=", mSelectedID
            If Not mConstants.NoMatch Then
                CurrentAction = ecaEditConstant
            End If
            mnuDataDelete.Enabled = True
            mnuDataAdd.Enabled = False
        ElseIf mSelectedType = enIntroduction Then
            If Not mGeneral_Information Is Nothing Then
                mGeneral_Information.Seek "=", "Introduction"
                If mGeneral_Information.NoMatch Then
                    mGeneral_Information.AddNew
                    mGeneral_Information!Name = "Introduction"
                    mGeneral_Information.Update
                End If
            End If
            CurrentAction = ecaEditIntroduction
        ElseIf mSelectedType = entEndNotes Then
            mGeneral_Information.Seek "=", "EndNotes"
            If mGeneral_Information.NoMatch Then
                mGeneral_Information.AddNew
                mGeneral_Information!Name = "EndNotes"
                mGeneral_Information.Update
            End If
            CurrentAction = ecaEditEndNotes
        Else
            CurrentAction = ecaDefault
        End If
    End If
End Sub

' Return the first visible node of a TreeView
' https://binaryworld.net/Main/CodeDetail.aspx?CodeId=1013
Function GetTreeViewFirstVisibleNode(ByVal TV As TreeView) As Node
  Dim hItem As Long
  Dim selNode As Node
  
  ' remember the node currently selected
  Set selNode = TV.SelectedItem
  ' get the handle of the first visible Node
  hItem = SendMessage(TV.hWnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&)
  ' make it the selected Node
  SendMessage TV.hWnd, TVM_SELECTITEM, TVGN_CARET, ByVal hItem
  ' return the result as a Node object
  Set GetTreeViewFirstVisibleNode = TV.SelectedItem
  ' restore node that was selected
  Set TV.SelectedItem = selNode
End Function

' set the first visible Node of a TreeView control
' https://binaryworld.net/Main/CodeDetail.aspx?CodeId=1025
Sub SetTreeViewFirstVisibleNode(ByVal TV As TreeView, ByVal Node As Node)
  Dim hItem As Long
  Dim selNode As Node
  
  ' remember the node currently selected
  Set selNode = TV.SelectedItem
  ' make the Node the select Node in the control
  
  Set TV.SelectedItem = Node
  ' now we can get its handle
  hItem = SendMessage(TV.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
  ' restore node that was selected
  Set TV.SelectedItem = selNode
  ' make it the fist visible Node
  SendMessage TV.hWnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, ByVal hItem
End Sub

Private Function GetNodeByKey(nKey As String) As Node
    Dim n As Node
    
    For Each n In trv1.Nodes
        If n.Key = nKey Then
            Set GetNodeByKey = n
            Exit For
        End If
    Next
End Function

Private Function GetSettingBase(ByVal Section As String, ByVal Key As String, Optional Default As String) As String
    If mGeneral_Information Is Nothing Then Exit Function
    mGeneral_Information.Seek "=", "Setting_" & Section & "_" & Key
    If mGeneral_Information.NoMatch Then
        GetSettingBase = Default
    Else
        GetSettingBase = mGeneral_Information!Value
    End If
End Function

Private Sub SaveSettingBase(ByVal Section As String, ByVal Key As String, ByVal Setting As String)
    mGeneral_Information.Seek "=", "Setting_" & Section & "_" & Key
    If mGeneral_Information.NoMatch Then
        mGeneral_Information.AddNew
        mGeneral_Information!Name = "Setting_" & Section & "_" & Key
    Else
        mGeneral_Information.Edit
    End If
    mGeneral_Information!Value = Setting
    mGeneral_Information.Update
End Sub

Private Function GetGeneralInfo(nKey As String) As String
    mGeneral_Information.Seek "=", nKey
    If Not mGeneral_Information.NoMatch Then
        GetGeneralInfo = mGeneral_Information!Value
    End If
End Function

Private Sub LoadReportingOptions()
    mHTML_Mode = Val(GetSettingBase("ReportingOptions", "HTML_Mode", 2))
    If (mHTML_Mode < 0) Or (mHTML_Mode > 2) Then mHTML_Mode = 2
    mExternalCSS = CBool(Val(GetSettingBase("ReportingOptions", "HTML_ExternalCSS", 1)))
    mReplaceCSSFile = CBool(Val(GetSettingBase("ReportingOptions", "HTML_ReplaceCSSFile", 1)))
    mPrint_Mode = Val(GetSettingBase("ReportingOptions", "Print_Mode", 0))
    If (mPrint_Mode < 0) Or (mPrint_Mode > 1) Then mPrint_Mode = 0
    
    mHTML_HeadSection = GetSettingBase("ReportingOptions", "HTML_HeadSection", cHTMLDefaultHeadSection)
    If mHTML_HeadSection = "" Then mHTML_HeadSection = cHTMLDefaultHeadSection
    mHTML_StyleSheet = GetSettingBase("ReportingOptions", "HTML_StyleSheet", cHTMLDefaultStyleSheet)
    If mHTML_StyleSheet = "" Then mHTML_StyleSheet = cHTMLDefaultStyleSheet
    mHTML_PageHeaderMP_Template = GetSettingBase("ReportingOptions", "HTML_PageHeaderMP", cHTMLDefaultPageHeaderMP)
    If mHTML_PageHeaderMP_Template = "" Then mHTML_PageHeaderMP_Template = cHTMLDefaultPageHeaderMP
    mHTML_PageHeaderOP_Template = GetSettingBase("ReportingOptions", "HTML_PageHeaderOP", cHTMLDefaultPageHeaderOP)
    If mHTML_PageHeaderOP_Template = "" Then mHTML_PageHeaderOP_Template = cHTMLDefaultPageHeaderOP
    mHTML_PageFooter_Template = GetSettingBase("ReportingOptions", "HTML_PageFooter", cHTMLDefaultPageFooter)
    If mHTML_PageFooter_Template = "" Then mHTML_PageFooter_Template = cHTMLDefaultPageFooter
    
    mHTML_PageHeaderMP = Replace$(mHTML_PageHeaderMP_Template, "[COMPONENT_NAME]", mComponentName)
    mHTML_PageHeaderOP = Replace$(mHTML_PageHeaderOP_Template, "[COMPONENT_NAME]", mComponentName)
    mHTML_PageFooter = Replace$(mHTML_PageFooter_Template, "[COMPONENT_NAME]", mComponentName)
End Sub

Private Sub PrintCentered(Optional nText As String, Optional ByVal nLeft, Optional ByVal nTop, Optional ByVal nWidth, Optional ByVal nHeight)
    Dim iHeight As Single
    Dim iWidth As String
    Dim iText() As String
    Dim c As Long
    
    If IsMissing(nLeft) Then nLeft = mMargin
    If IsMissing(nTop) Then nTop = mMargin
    If IsMissing(nWidth) Then nWidth = Printer.ScaleWidth - mMargin * 2
    If IsMissing(nHeight) Then nHeight = Printer.ScaleHeight - mMargin * 2
    
    Do Until (Printer.TextWidth(nText) <= (Printer.ScaleWidth - mMargin * 2)) And (Printer.TextHeight(nText) <= (Printer.ScaleHeight - mMargin * 2))
        Printer.FontSize = Printer.FontSize - 1
    Loop
    
    iHeight = Printer.TextHeight(nText)
    iText = Split(nText, vbCrLf)
    Printer.CurrentY = (Printer.ScaleHeight - iHeight) / 2
    
    For c = 0 To UBound(iText)
        iWidth = Printer.TextWidth(iText(c))
        Printer.CurrentX = (Printer.ScaleWidth - iWidth) / 2
        Call CheckForNewPage: Printer.Print iText(c)
    Next
End Sub

Private Sub PrintCenteredTop(Optional nText As String, Optional ByVal nLeft, Optional ByVal nTop, Optional ByVal nWidth, Optional ByVal nHeight)
    Dim iWidth As String
    Dim iText() As String
    Dim c As Long
    
    If IsMissing(nLeft) Then nLeft = mMargin
    If IsMissing(nTop) Then nTop = mMargin
    If IsMissing(nWidth) Then nWidth = Printer.ScaleWidth - mMargin * 2
    If IsMissing(nHeight) Then nHeight = Printer.ScaleHeight - mMargin * 2
    
    Do Until (Printer.TextWidth(nText) <= (Printer.ScaleWidth - mMargin * 2)) And (Printer.TextHeight(nText) <= (Printer.ScaleHeight - mMargin * 2))
        Printer.FontSize = Printer.FontSize - 1
    Loop
    
    iText = Split(nText, vbCrLf)
    
    For c = 0 To UBound(iText)
        iWidth = Printer.TextWidth(iText(c))
        Printer.CurrentX = (Printer.ScaleWidth - iWidth) / 2
        Call CheckForNewPage: Printer.Print iText(c)
    Next
End Sub

Private Sub Printer_NewPage()
    Dim f As StdFont
    Dim f2 As StdFont
    
    Set f = Printer.Font
    Printer.NewPage
    
    Printer.CurrentX = Printer.ScaleWidth - Printer.ScaleX(16, vbMillimeters, vbTwips)
    Printer.CurrentY = Printer.ScaleHeight - Printer.ScaleY(16, vbMillimeters, vbTwips)
    
    Set f2 = New StdFont
    Set Printer.Font = f2
    Printer.FontName = "Arial"
    Printer.FontSize = 12
    Printer.Print Printer.Page
    
    Set Printer.Font = f
    Printer.CurrentX = mMargin
    Printer.CurrentY = mMargin
End Sub

Public Sub PrintRTB(nRTB As RichTextBox)
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As RECT
    Dim rcMeasure As RECT
    Dim rcPage As RECT
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long
    Dim iPPScaleMode As Long
    Dim iOldFont As StdFont
    
    Set iOldFont = Printer.Font
    Set Printer.Font = New StdFont
    
    iPPScaleMode = Printer.ScaleMode
    ' Start a Print job to get a valid Printer.hDC
    Printer.Print Space(1);
    Printer.ScaleMode = vbTwips
    Printer.CurrentX = 0
    
    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = mMargin + LeftOffset + Printer.CurrentX
    TopMargin = mMargin + TopOffset '+ Printer.CurrentY
    RightMargin = (Printer.ScaleWidth - mMargin) - LeftOffset '+ 300
    BottomMargin = (Printer.ScaleHeight - mMargin) - TopOffset '+ 300
    
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight
    
    ' Set rect in which to Print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = IIf(TopMargin > Printer.CurrentY, TopMargin, Printer.CurrentY)
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin
    
    rcMeasure.Left = LeftMargin
    rcMeasure.Top = rcDrawTo.Top
    rcMeasure.Right = RightMargin
    rcMeasure.Bottom = BottomMargin
    
    ' Set up the Print instructions
    fr.hDC = Printer.hDC   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hDC  ' Point at Printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text
    
    ' Get length of text in nRTB
    TextLength = Len(nRTB.Text)
    
    ' Loop printing each page until done
    Do
       
       fr.rc = rcMeasure
       Call SendMessage(nRTB.hWnd, EM_FORMATRANGE, 0&, fr)
       rcMeasure = fr.rc
       fr.rc = rcDrawTo
       ' Print the page by sending EM_FORMATRANGE message
       NextCharPosition = SendMessage(nRTB.hWnd, EM_FORMATRANGE, True, fr)
       
       If NextCharPosition >= TextLength Then Exit Do  'If done then exit
       fr.chrg.cpMin = NextCharPosition ' Starting position for next page
       Printer_NewPage                  ' Move on to next page
 '      Printer.Print Space(1);  ' Re-initialize hDC
  '     Printer.CurrentX = 0
   '    Printer.CurrentY = 0
       fr.hDC = Printer.hDC
       fr.hdcTarget = Printer.hDC
       rcDrawTo.Left = LeftMargin
       rcDrawTo.Top = TopMargin
       rcMeasure.Top = rcDrawTo.Top
    Loop

    ' Commit the Print job
 '   Printer.EndDoc

    ' Allow the nRTB to free up memory
    r = SendMessage(nRTB.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
    
    Set Printer.Font = iOldFont
'    Printer.CurrentY = Printer.CurrentY + rcMeasure.Bottom - rcMeasure.Top - Printer.TextHeight("A")
    Printer.CurrentY = rcMeasure.Bottom
End Sub

Private Sub SeparatePrintedItems(Optional ByVal nNumberOfLines As Long = 2)
    If (Printer.CurrentY - mMargin) > 100 Then
        If mPrint_Mode = cdSeparatePages Then Printer_NewPage Else PrintSeparation
    End If
End Sub

Private Sub PrintSeparation(Optional ByVal nNumberOfLines As Long = 2)
    Dim c As Long
    Dim s As String
    Dim iFs As Single
    
    iFs = Printer.FontSize
    Printer.FontSize = 12
    For c = 1 To nNumberOfLines
        s = s & vbCrLf
    Next
    If (Printer.TextHeight(s) + Printer.CurrentY) > (Printer.ScaleHeight - mMargin * 2) Then
        Printer_NewPage
    Else
        If (Printer.CurrentY - mMargin) > 100 Then
            Printer.Print s
        End If
    End If
    Printer.FontSize = iFs
End Sub

Private Sub CheckForNewPage()
    If (Printer.TextHeight("Tq") + Printer.CurrentY) > (Printer.ScaleHeight - mMargin) Then
        Printer_NewPage
    End If
End Sub

Private Function ShowfrmReportSelection(nReportType As String) As Boolean
    Dim iSection As String
    
    iSection = "ReportSelection_" & nReportType
    
    frmReportSelection.SetRec 1, mControls
    frmReportSelection.SetRec 2, mClasses
    frmReportSelection.SetRec 3, mEnums
    frmReportSelection.UnselectedItems(1) = GetSettingBase(iSection, "UnselectedControls", "")
    frmReportSelection.UnselectedItems(2) = GetSettingBase(iSection, "UnselectedClasses", "")
    frmReportSelection.UnselectedItems(3) = GetSettingBase(iSection, "UnselectedEnums", "")
    
    frmReportSelection.InfoVersionAvailable = (mComponentVersion <> "")
    frmReportSelection.InfoReleaseDateAvailable = (mComponentReleaseDate <> 0)
    frmReportSelection.InfoIntroductionAvailable = (GetGeneralInfo("Introduction") <> "")
    frmReportSelection.InfoEndNotesAvailable = (GetGeneralInfo("EndNotes") <> "")
    
    frmReportSelection.chkInfo.Value = Val(GetSettingBase(iSection, "ComponentInfo", 1))
    frmReportSelection.GeneralInfoSettingsStr = GetSettingBase(iSection, "ComponentInfoDetails", "0|0|1|1")
    
    If mControls.RecordCount > 0 Then
        frmReportSelection.cmdSelect(1).Enabled = (mControls.RecordCount > 1)
        If Val(Val(GetSettingBase(iSection, "Controls", 1))) <> 0 Then
            If frmReportSelection.UnselectedItems(1) <> "" Then
                frmReportSelection.chkType(1).Value = 2
            Else
                frmReportSelection.chkType(1).Value = 1
            End If
        Else
            frmReportSelection.chkType(1).Value = 0
        End If
    Else
        frmReportSelection.cmdSelect(1).Enabled = False
        frmReportSelection.chkType(1).Value = 0
        frmReportSelection.chkType(1).Enabled = False
    End If
    If mClasses.RecordCount > 0 Then
        frmReportSelection.cmdSelect(2).Enabled = (mClasses.RecordCount > 1)
        If Val(Val(GetSettingBase(iSection, "Classes", 1))) <> 0 Then
            If frmReportSelection.UnselectedItems(2) <> "" Then
                frmReportSelection.chkType(2).Value = 2
            Else
                frmReportSelection.chkType(2).Value = 1
            End If
        Else
            frmReportSelection.chkType(2).Value = 0
        End If
    Else
        frmReportSelection.cmdSelect(2).Enabled = False
        frmReportSelection.chkType(2).Value = 0
        frmReportSelection.chkType(2).Enabled = False
    End If
    If mEnums.RecordCount > 0 Then
        frmReportSelection.cmdSelect(3).Enabled = (mEnums.RecordCount > 1)
        If Val(Val(GetSettingBase(iSection, "Constants", 1))) <> 0 Then
            If frmReportSelection.UnselectedItems(3) <> "" Then
                frmReportSelection.chkType(3).Value = 2
            Else
                frmReportSelection.chkType(3).Value = 1
            End If
        Else
            frmReportSelection.chkType(3).Value = 0
        End If
    Else
        frmReportSelection.cmdSelect(3).Enabled = False
        frmReportSelection.chkType(3).Value = 0
        frmReportSelection.chkType(3).Enabled = False
    End If
    
    frmReportSelection.Show vbModal
    
    If (frmReportSelection.chkInfo.Value = 0) And (frmReportSelection.chkType(1).Value = 0) And (frmReportSelection.chkType(2).Value = 0) And (frmReportSelection.chkType(3).Value = 0) Then
        Unload frmReportSelection
        MsgBox "Nothing selected.", vbInformation
        Exit Function
    End If
    
    SaveSettingBase iSection, "ComponentInfo", frmReportSelection.chkInfo.Value
    If frmReportSelection.chkInfo.Value Then
        SaveSettingBase iSection, "ComponentInfoDetails", frmReportSelection.GeneralInfoSettingsStr
    End If
    
    If mControls.RecordCount > 0 Then
        SaveSettingBase iSection, "Controls", frmReportSelection.chkType(1).Value
        SaveSettingBase iSection, "UnselectedControls", frmReportSelection.UnselectedItems(1)
    End If
    If mClasses.RecordCount > 0 Then
        SaveSettingBase iSection, "Classes", frmReportSelection.chkType(2).Value
        SaveSettingBase iSection, "UnselectedClasses", frmReportSelection.UnselectedItems(2)
    End If
    If mEnums.RecordCount > 0 Then
        SaveSettingBase iSection, "Constants", frmReportSelection.chkType(3).Value
        SaveSettingBase iSection, "UnselectedEnums", frmReportSelection.UnselectedItems(3)
    End If
    
    ShowfrmReportSelection = True
End Function

Private Sub AddPage(nFileName As String, nPageText As String)
    AddToList mPages, nPageText
    ReDim Preserve mPages_FileNames(UBound(mPages))
    mPages_FileNames(UBound(mPages)) = nFileName
End Sub

Private Function HTMLFormatParameters(nText As String) As String
    Dim iRows() As String
    Dim iCols() As String
    Dim r As Long
    Dim c As Long
    
    If InStr(nText, vbTab) = 0 Then
        HTMLFormatParameters = nText
        Exit Function
    End If
    
    HTMLFormatParameters = "<table><tbody>"
    iRows = Split(nText, vbCrLf)
    For r = 0 To UBound(iRows)
        HTMLFormatParameters = HTMLFormatParameters & "    <tr>"
        iCols = Split(iRows(r), vbTab)
        For c = 0 To UBound(iCols)
            HTMLFormatParameters = HTMLFormatParameters & "    <td>&nbsp&nbsp" & iCols(c) & "&nbsp&nbsp</td>"
        Next
        HTMLFormatParameters = HTMLFormatParameters & "    </tr>"
        If InStr(iRows(r), "Return Type") > 0 Then
            HTMLFormatParameters = HTMLFormatParameters & "<tr><td>&nbsp</td></tr>"
        End If
    Next
    HTMLFormatParameters = HTMLFormatParameters & "  </tbody></table>"
End Function

Private Function TabLines(nText As String, Optional nTabs As Long = 1) As String
    Dim s() As String
    Dim c As Long
    
    s = Split(nText, vbCrLf)
    For c = 0 To UBound(s)
        s(c) = String$(nTabs, vbTab) & s(c)
    Next
    TabLines = Join(s, vbCrLf)
End Function

Private Function GetPageFileName(nType As Long, nRec As Recordset) As String
    If IsMemberNameUnique(nType, nRec!Name) Then
        GetPageFileName = LCase$(nRec!Name) & "_" & LCase$(mMemberType_s(nType)) & ".html"
    Else
        On Error Resume Next
        GetPageFileName = LCase$(nRec!Name) & "_" & LCase$(mMemberType_s(nType)) & CStr(nRec(mMemberType_p(nType) & "." & mMemberType_s(nType) & "_ID").Value) & ".html"
        If Err.Number = 3265 Then
            On Error GoTo 0
            GetPageFileName = LCase$(nRec!Name) & "_" & LCase$(mMemberType_s(nType)) & CStr(nRec(mMemberType_s(nType) & "_ID").Value) & ".html"
        End If
    End If
End Function

Private Function IsMemberNameUnique(nMemberType As Long, nMemberName As String) As Boolean
    Dim iRec As Recordset
    
    Set iRec = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(nMemberType) & " WHERE (Name = '" & nMemberName & "') AND (Auxiliary_Field = 1)")
    iRec.MoveLast
    IsMemberNameUnique = (iRec.RecordCount = 1)
End Function

Private Function ParamsInfoAreTheSame(ByVal nParamsInfo1 As String, ByVal nParamsInfo2 As String)
    nParamsInfo1 = RemoveNAADT(nParamsInfo1)
    nParamsInfo2 = RemoveNAADT(nParamsInfo2)
    nParamsInfo1 = SeparateWordsOneSpace(nParamsInfo1)
    nParamsInfo2 = SeparateWordsOneSpace(nParamsInfo2)
    ParamsInfoAreTheSame = (nParamsInfo1 = nParamsInfo2)
End Function

Private Function SeparateWordsOneSpace(nText As String) As String
    SeparateWordsOneSpace = Replace$(nText, vbTab, " ")
    Do Until InStr(SeparateWordsOneSpace, vbTab) = 0
        SeparateWordsOneSpace = Replace$(SeparateWordsOneSpace, vbTab, " ")
    Loop
    Do Until InStr(SeparateWordsOneSpace, vbCr) = 0
        SeparateWordsOneSpace = Replace$(SeparateWordsOneSpace, vbCr, " ")
    Loop
    Do Until InStr(SeparateWordsOneSpace, vbLf) = 0
        SeparateWordsOneSpace = Replace$(SeparateWordsOneSpace, vbLf, " ")
    Loop
    Do Until InStr(SeparateWordsOneSpace, "  ") = 0
        SeparateWordsOneSpace = Replace$(SeparateWordsOneSpace, "  ", " ")
    Loop
End Function

Private Function ThereAreOrphanMembers() As Boolean
    If (mCurrentAction > 0) And (CurrentType > 0) Then
        ThereAreOrphanMembers = mDatabase.OpenRecordset("SELECT * FROM " & mMemberType_p(CurrentType) & " WHERE (Auxiliary_Field = 0)").RecordCount > 0
    End If
End Function
    
Private Property Get CurrentType() As Long
    If mCurrentAction = ecaEditProperty Then
        CurrentType = 1
    ElseIf mCurrentAction = ecaEditMethod Then
        CurrentType = 2
    ElseIf mCurrentAction = ecaEditEvent Then
        CurrentType = 3
    End If
End Property

Private Sub SetControlsFont()
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "TextBox" Then
            Set ctl.Font = mAppFont
        End If
        Set trv1.Font = mAppFont
    Next
End Sub

Private Function RemoveNAADT(nText As String) As String
    RemoveNAADT = Replace$(nText, " (Not available at design time)", "")
    RemoveNAADT = Replace$(RemoveNAADT, ", not available at design time)", ")")
End Function

Private Function StrCount(nText As String, nFind As String) As Long
    StrCount = UBound(Split(nText, nFind)) + 1
End Function

Private Function GetComboListHwnd(nCombo As Object) As Long
    Dim iCboInf As COMBOBOXINFO
    
    iCboInf.cbSize = Len(iCboInf)
    GetComboBoxInfo nCombo.hWnd, iCboInf
    GetComboListHwnd = iCboInf.hWndList
End Function

