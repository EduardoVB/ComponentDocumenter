VERSION 5.00
Begin VB.Form frmSelectOrpahnMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select orphan member"
   ClientHeight    =   6120
   ClientLeft      =   1932
   ClientTop       =   2160
   ClientWidth     =   5124
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
   ScaleHeight     =   6120
   ScaleWidth      =   5124
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   3648
      TabIndex        =   0
      Top             =   5448
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   2160
      TabIndex        =   1
      Top             =   5448
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.TextBox txtDetails 
      Height          =   2352
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2820
      Width           =   5112
   End
   Begin VB.ListBox lstMembers 
      Height          =   2448
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5112
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This will replace the current text"
      Height          =   552
      Left            =   180
      TabIndex        =   5
      Top             =   5280
      Width           =   1692
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2652
   End
End
Attribute VB_Name = "frmSelectOrpahnMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean
Public Text As String
Private mDetails() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lstMembers.ListIndex > -1 Then
        OKPressed = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

Private Sub lstMembers_Click()
    txtDetails.Text = mDetails(lstMembers.ListIndex)
End Sub

Private Sub txtDetails_Change()
    Text = txtDetails.Text
End Sub

Public Sub LoadList(nDB As Database, nTable As String)
    Dim iRec As Recordset
    
    Set iRec = nDB.OpenRecordset("SELECT * FROM " & nTable & " WHERE (Auxiliary_Field = 0)")
    If iRec.RecordCount > 0 Then
        iRec.MoveLast
        ReDim mDetails(iRec.RecordCount - 1)
        iRec.MoveFirst
        Do Until iRec.EOF
            lstMembers.AddItem iRec!Name
            mDetails(lstMembers.NewIndex) = iRec!Long_Description
            iRec.MoveNext
        Loop
    End If
End Sub
