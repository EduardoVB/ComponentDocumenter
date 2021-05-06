VERSION 5.00
Begin VB.Form frmSelectItems 
   Caption         =   "Selection"
   ClientHeight    =   5244
   ClientLeft      =   6348
   ClientTop       =   2676
   ClientWidth     =   5496
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
   ScaleHeight     =   5244
   ScaleWidth      =   5496
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "Deselect all"
      Height          =   420
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select all"
      Height          =   420
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   3888
      TabIndex        =   1
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1092
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
   Begin VB.ListBox lstItems 
      Appearance      =   0  'Flat
      Height          =   4248
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   0
      Width           =   4212
   End
End
Attribute VB_Name = "frmSelectItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    OKPressed = True
    Me.Hide
End Sub

Private Sub cmdSelectAll_Click()
    Dim c As Long
    
    For c = 0 To lstItems.ListCount - 1
        lstItems.Selected(c) = True
    Next
End Sub

Private Sub cmdDeselectAll_Click()
    Dim c As Long
    
    For c = 0 To lstItems.ListCount - 1
        lstItems.Selected(c) = False
    Next
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub
