VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "Results"
   ClientHeight    =   5820
   ClientLeft      =   1932
   ClientTop       =   2160
   ClientWidth     =   8544
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
   ScaleHeight     =   5820
   ScaleWidth      =   8544
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   7080
      TabIndex        =   0
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4692
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   7932
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   6612
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

Private Sub Form_Resize()
    If Me.Width < 6000 Then Me.Width = 6000
    If Me.Height < 6000 Then Me.Height = 6000
    
    txtMessage.Move 30, 30, Me.ScaleWidth - 60, Me.ScaleHeight - 750
    cmdOK.Move Me.ScaleWidth - cmdOK.Width - 500, Me.ScaleHeight - cmdOK.Height - 140
    lblNote.Top = txtMessage.Top + txtMessage.Height + 60
    lblNote.Width = cmdOK.Left - 300 - lblNote.Left
    If (lblNote.Height + lblNote.Top) > (Me.ScaleHeight - 90) Then
        txtMessage.Height = Me.ScaleHeight - 180 - lblNote.Height
        lblNote.Top = txtMessage.Top + txtMessage.Height + 60
    End If
End Sub

Public Property Let Message(nTxt As String)
    txtMessage.Text = nTxt
End Property

