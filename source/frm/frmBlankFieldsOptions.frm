VERSION 5.00
Begin VB.Form frmBlankFieldsOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select fields"
   ClientHeight    =   4692
   ClientLeft      =   1932
   ClientTop       =   2160
   ClientWidth     =   5244
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
   ScaleHeight     =   4692
   ScaleWidth      =   5244
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   3708
      TabIndex        =   1
      Top             =   4020
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   420
      Left            =   2220
      TabIndex        =   0
      Top             =   4020
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CheckBox chkField 
      Caption         =   "Enums descriptions"
      Height          =   372
      Index           =   5
      Left            =   720
      TabIndex        =   7
      Top             =   2670
      Width           =   3132
   End
   Begin VB.CheckBox chkField 
      Caption         =   "Constants descriptions"
      Height          =   372
      Index           =   6
      Left            =   720
      TabIndex        =   8
      Top             =   3120
      Width           =   3132
   End
   Begin VB.CheckBox chkField 
      Caption         =   "Members parameters info"
      Height          =   372
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Value           =   1  'Checked
      Width           =   3132
   End
   Begin VB.CheckBox chkField 
      Caption         =   "Members Long descriptions"
      Height          =   372
      Index           =   3
      Left            =   720
      TabIndex        =   5
      Top             =   1770
      Value           =   1  'Checked
      Width           =   3132
   End
   Begin VB.CheckBox chkField 
      Caption         =   "members Short descriptions"
      Height          =   372
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   2220
      Value           =   1  'Checked
      Width           =   3132
   End
   Begin VB.CheckBox chkField 
      Caption         =   "Controls/Classes Short descriptions"
      Height          =   372
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   870
      Width           =   3132
   End
   Begin VB.CheckBox chkField 
      Caption         =   "Controls/Classes Long descriptions"
      Height          =   372
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   420
      Value           =   1  'Checked
      Width           =   3132
   End
End
Attribute VB_Name = "frmBlankFieldsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean
Private mFields(6) As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim c As Long
    
    For c = 0 To UBound(mFields)
        mFields(c) = (chkField(c).Value = 1)
    Next
    OKPressed = True
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

Public Property Get Field(nIndex As Long) As Boolean
    Field = mFields(nIndex)
End Property

Public Property Get FieldsString() As String
    Dim c As Long
    
    For c = 0 To UBound(mFields)
        FieldsString = FieldsString & CStr(Abs(CLng(mFields(c))))
    Next
End Property

Public Property Let FieldsString(nValue As String)
    Dim c As Long
    
    For c = 0 To UBound(mFields)
        chkField(c).Value = Val(Mid$(nValue, c + 1, 1))
    Next
End Property
