VERSION 5.00
Begin VB.Form frmSelectMemberDefinition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select member definition"
   ClientHeight    =   6228
   ClientLeft      =   1932
   ClientTop       =   2160
   ClientWidth     =   7128
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
   ScaleHeight     =   6228
   ScaleWidth      =   7128
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDisplay 
      Height          =   336
      ItemData        =   "frmSelectMemberDefinition.frx":0000
      Left            =   960
      List            =   "frmSelectMemberDefinition.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1580
      Width           =   3432
   End
   Begin VB.ListBox lstMembers 
      Height          =   1248
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7112
   End
   Begin VB.TextBox txtDetails 
      Height          =   3132
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2050
      Width           =   7112
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   420
      Left            =   4160
      TabIndex        =   1
      Top             =   5568
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   5648
      TabIndex        =   0
      Top             =   5568
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Display:"
      Height          =   252
      Left            =   180
      TabIndex        =   3
      Top             =   1640
      Width           =   732
   End
   Begin VB.Label lblNote 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   732
      Left            =   120
      TabIndex        =   6
      Top             =   5328
      Width           =   3876
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSelectMemberDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean
Public MemberID As Long
Private mParams() As String
Private mLongDesc() As String
Private mShortDesc() As String
Private mOrigParams As String
Private mOrigShortDesc As String
Public ParamsInfo As String
Public ShortDescription As String

Private Sub cboDisplay_Click()
    If lstMembers.ListIndex > -1 Then
        Select Case cboDisplay.ListIndex
            Case 0 ' Params info
                txtDetails.Text = mParams(lstMembers.ListIndex)
            Case 1 ' Long description
                txtDetails.Text = mLongDesc(lstMembers.ListIndex)
            Case 2 ' Short description
                txtDetails.Text = mShortDesc(lstMembers.ListIndex)
        End Select
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    OKPressed = True
    Unload Me
End Sub

Public Sub LoadList(nDB As Database, nTable As String, nMemberName As String, nMemberType As String, nMemberIDNotToLoad As Long, nOrigParams As String, nOrigShortDesc As String)
    Dim iRec As Recordset
    
    Set iRec = nDB.OpenRecordset("SELECT * FROM " & nTable & " WHERE (Name = '" & nMemberName & "') AND (" & nMemberType & "_ID <> " & CStr(nMemberIDNotToLoad) & ") AND (Auxiliary_Field = 1)")
    If iRec.RecordCount > 0 Then
        iRec.MoveLast
        ReDim mParams(iRec.RecordCount - 1)
        ReDim mLongDesc(iRec.RecordCount - 1)
        ReDim mShortDesc(iRec.RecordCount - 1)
        iRec.MoveFirst
        Do Until iRec.EOF
            lstMembers.AddItem iRec!Name
            lstMembers.ItemData(lstMembers.NewIndex) = iRec.Fields(nMemberType & "_ID").Value
            mParams(lstMembers.NewIndex) = iRec!Params_Info
            mLongDesc(lstMembers.NewIndex) = iRec!Long_Description
            mShortDesc(lstMembers.NewIndex) = iRec!Short_Description
            iRec.MoveNext
        Loop
    End If
    cboDisplay.ListIndex = 0
    mOrigParams = nOrigParams
    mOrigShortDesc = nOrigShortDesc
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

Private Sub lstMembers_Click()
    cboDisplay_Click
    If (mParams(lstMembers.ListIndex) <> mOrigParams) And (mShortDesc(lstMembers.ListIndex) <> mOrigShortDesc) Then
        lblNote.Caption = "Params info and short description are different from originals"
    ElseIf (mParams(lstMembers.ListIndex) <> mOrigParams) Then
        lblNote.Caption = "Params info is different from original"
    ElseIf (mShortDesc(lstMembers.ListIndex) <> mOrigShortDesc) Then
        lblNote.Caption = "Short description is different from original"
    Else
        lblNote.Caption = ""
    End If
    MemberID = lstMembers.ItemData(lstMembers.ListIndex)
    ParamsInfo = mParams(lstMembers.ListIndex)
    ShortDescription = mShortDesc(lstMembers.ListIndex)
End Sub
