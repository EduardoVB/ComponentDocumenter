VERSION 5.00
Begin VB.Form frmComponentProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Component properties"
   ClientHeight    =   3048
   ClientLeft      =   2820
   ClientTop       =   2160
   ClientWidth     =   4932
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
   ScaleHeight     =   3048
   ScaleWidth      =   4932
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReleaseDate 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1560
      Width           =   1188
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1440
      MaxLength       =   2312
      TabIndex        =   5
      Top             =   1020
      Width           =   1188
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   1884
      TabIndex        =   0
      Top             =   2388
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   3372
      TabIndex        =   1
      Top             =   2388
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   3
      Top             =   456
      Width           =   3228
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Release date:"
      Height          =   348
      Left            =   240
      TabIndex        =   6
      Top             =   1644
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Version:"
      Height          =   348
      Left            =   240
      TabIndex        =   4
      Top             =   1104
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   348
      Left            =   240
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmComponentProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean
Public ComponentName As String
Public ComponentVersion As String
Public ComponentReleaseDate As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtReleaseDate.Text = Trim(txtReleaseDate.Text)
    If txtReleaseDate.Text <> "" Then
        If Not IsDate(txtReleaseDate.Text) Then
            MsgBox "Please enter a valid date or leave blank.", vbExclamation
            txtReleaseDate.SelStart = 0
            txtReleaseDate.SelLength = Len(txtReleaseDate.Text)
            txtReleaseDate.SetFocus
            Exit Sub
        End If
    End If
    ComponentName = txtName.Text
    ComponentVersion = txtVersion.Text
    ComponentReleaseDate = txtReleaseDate.Text
    OKPressed = True
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

