VERSION 5.00
Begin VB.Form frmFieldsModifAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alert"
   ClientHeight    =   2160
   ClientLeft      =   1932
   ClientTop       =   2160
   ClientWidth     =   6060
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHide 
      Caption         =   "OK, do not show me it again"
      Height          =   372
      Left            =   300
      TabIndex        =   2
      Top             =   1560
      Width           =   2892
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   420
      Left            =   4480
      TabIndex        =   0
      Top             =   1500
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmFieldsModifAlert.frx":0000
      Height          =   1212
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5632
   End
End
Attribute VB_Name = "frmFieldsModifAlert"
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

Private Sub Form_Unload(Cancel As Integer)
    If chkHide.Value = 1 Then
        SaveSetting App.Title, AppPath4Reg, "HideModifAlert", "1"
    End If
End Sub
