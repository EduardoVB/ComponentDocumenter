VERSION 5.00
Begin VB.Form frmReportSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select items"
   ClientHeight    =   5244
   ClientLeft      =   2796
   ClientTop       =   2160
   ClientWidth     =   4056
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5244
   ScaleWidth      =   4056
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkInfoEndNotes 
      Caption         =   "End Notes"
      Height          =   444
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   2532
   End
   Begin VB.CheckBox chkInfoIntroduction 
      Caption         =   "Introduction"
      Height          =   444
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   2532
   End
   Begin VB.CheckBox chkInfoReleaseDate 
      Caption         =   "Release date"
      Height          =   444
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   2532
   End
   Begin VB.CheckBox chkInfoVersion 
      Caption         =   "Version"
      Height          =   444
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   2532
   End
   Begin VB.CheckBox chkInfoName 
      Caption         =   "Name"
      Height          =   444
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2532
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   350
      Index           =   3
      Left            =   2400
      TabIndex        =   12
      Top             =   3636
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   350
      Index           =   2
      Left            =   2400
      TabIndex        =   10
      Top             =   3156
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   350
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   2676
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CheckBox chkInfo 
      Caption         =   "Component information"
      Height          =   444
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Value           =   2  'Grayed
      Width           =   2532
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Constants"
      Height          =   444
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   3600
      Width           =   2532
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Classes"
      Height          =   444
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   3120
      Width           =   2532
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Controls"
      Height          =   444
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   2640
      Width           =   2532
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   2400
      TabIndex        =   0
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
End
Attribute VB_Name = "frmReportSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InfoVersionAvailable As Boolean
Public InfoReleaseDateAvailable As Boolean
Public InfoIntroductionAvailable As Boolean
Public InfoEndNotesAvailable As Boolean

Private mUnselectedItems(3) As Variant
Private mRec(3) As Recordset

Private Sub chkInfo_Click()
    If chkInfo.Value = 1 Then
        chkInfoName.Value = 1
    End If
    chkInfoName.Enabled = (chkInfo.Value = 1)
    chkInfoVersion.Enabled = (chkInfo.Value = 1) And InfoVersionAvailable
    chkInfoReleaseDate.Enabled = (chkInfo.Value = 1) And InfoReleaseDateAvailable
    chkInfoIntroduction.Enabled = (chkInfo.Value = 1) And InfoIntroductionAvailable
    chkInfoEndNotes.Enabled = (chkInfo.Value = 1) And InfoEndNotesAvailable
End Sub

Private Sub chkInfoName_Click()
    If chkInfoName.Value = 0 Then
        chkInfo.Value = 0
    End If
End Sub

Private Sub cmdOK_Click()
    Dim c As Long
    
    For c = 1 To chkType.UBound
        If chkType(c).Value = 2 Then chkType(c).Value = 1
    Next
    Me.Hide
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim iList() As String
    Dim c As Long
    Dim iStr As String
    Dim iSel As Boolean
    
    iList = mUnselectedItems(Index)
    mRec(Index).MoveFirst
    Do Until mRec(Index).EOF
        frmSelectItems.lstItems.AddItem mRec(Index)!Name
        If Not IsInList(iList, mRec(Index)!Name) Then
            frmSelectItems.lstItems.Selected(frmSelectItems.lstItems.NewIndex) = True
        End If
        mRec(Index).MoveNext
    Loop
    frmSelectItems.lstItems.TopIndex = 0
    frmSelectItems.lstItems.ListIndex = -1
    frmSelectItems.Show vbModal
    If frmSelectItems.OKPressed Then
        iStr = ""
        For c = 0 To frmSelectItems.lstItems.ListCount - 1
            If Not frmSelectItems.lstItems.Selected(c) Then
                If iStr <> "" Then iStr = iStr & "|"
                iStr = iStr & frmSelectItems.lstItems.List(c)
            Else
                iSel = True
            End If
        Next
        If iSel Then
            If iStr = "" Then
                chkType(Index).Value = 1
            Else
                chkType(Index).Value = 2
            End If
            mUnselectedItems(Index) = Split(iStr, "|")
        Else
            chkType(Index).Value = 0
            mUnselectedItems(Index) = Split("")
        End If
        Unload frmSelectItems
    End If
    Set frmSelectItems = Nothing
End Sub

Public Sub SetRec(nIndex As Long, nRec As Recordset)
    Set mRec(nIndex) = nRec
End Sub

Public Property Let UnselectedItems(nIndex As Long, nList As String)
    mUnselectedItems(nIndex) = Split(nList, "|")
End Property

Public Property Get UnselectedItems(nIndex As Long) As String
    UnselectedItems = Join(mUnselectedItems(nIndex), "|")
End Property

Public Function IsItemSelected(nIndex As Long, nItemName As String) As Boolean
    Dim iList() As String
    
    If chkType(nIndex).Value Then
        iList = mUnselectedItems(nIndex)
        IsItemSelected = Not IsInList(iList, nItemName)
    End If
End Function

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

Public Property Get GeneralInfoSettingsStr() As String
    GeneralInfoSettingsStr = chkInfoVersion.Value & "|"
    GeneralInfoSettingsStr = GeneralInfoSettingsStr & chkInfoReleaseDate.Value & "|"
    GeneralInfoSettingsStr = GeneralInfoSettingsStr & chkInfoIntroduction.Value & "|"
    GeneralInfoSettingsStr = GeneralInfoSettingsStr & chkInfoEndNotes.Value
End Property

Public Property Let GeneralInfoSettingsStr(nStr As String)
    Dim iStrs() As String
    
    chkInfoName.Value = 1
    iStrs = Split(nStr, "|")
    If UBound(iStrs) > 2 Then
        chkInfoVersion.Value = Val(iStrs(0))
        chkInfoReleaseDate.Value = Val(iStrs(1))
        chkInfoIntroduction.Value = Val(iStrs(2))
        chkInfoEndNotes.Value = Val(iStrs(3))
    End If
End Property
    
