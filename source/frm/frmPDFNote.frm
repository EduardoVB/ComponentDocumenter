VERSION 5.00
Begin VB.Form frmPDFNote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Note about PDF printer drivers"
   ClientHeight    =   2484
   ClientLeft      =   5256
   ClientTop       =   4488
   ClientWidth     =   5700
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
   ScaleHeight     =   2484
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   3960
      TabIndex        =   0
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "OK, do not show me it again"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2892
   End
   Begin VB.Label Label2 
      Caption         =   $"frmPDFNote.frx":0000
      Height          =   732
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5052
   End
   Begin VB.Label Label1 
      Caption         =   "Note about PDF printer drivers:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3732
   End
End
Attribute VB_Name = "frmPDFNote"
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
        SaveSetting App.Title, AppPath4Reg, "HidePDFNote", "1"
    End If
End Sub
