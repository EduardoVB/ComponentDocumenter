VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   1692
   ClientLeft      =   4020
   ClientTop       =   3384
   ClientWidth     =   5880
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
   ScaleHeight     =   1692
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   2820
      TabIndex        =   1
      Top             =   1008
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4308
      TabIndex        =   2
      Top             =   1008
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdChooseFont 
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
      Height          =   312
      Left            =   5100
      TabIndex        =   0
      Top             =   300
      Width           =   432
   End
   Begin VB.Label lblFont 
      BackColor       =   &H80000005&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1620
      TabIndex        =   4
      Top             =   300
      Width           =   3312
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Editor font:"
      Height          =   324
      Left            =   300
      TabIndex        =   3
      Top             =   288
      Width           =   1164
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AppFont As StdFont
Private mNewFont As StdFont

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChooseFont_Click()
    Dim iDlg As New cDlg
    
    Set iDlg.Font = AppFont
    iDlg.ShowFont
    If Not iDlg.Canceled Then
        Set mNewFont = iDlg.Font
        lblFont.Caption = mNewFont.Name & " " & mNewFont.Size
    End If
End Sub

Private Sub cmdOK_Click()
    AppFont.Name = mNewFont.Name
    AppFont.Size = mNewFont.Size
    AppFont.Bold = mNewFont.Bold
    AppFont.Charset = mNewFont.Charset
    AppFont.Italic = mNewFont.Italic
    AppFont.Weight = mNewFont.Weight
    SaveSetting App.Title, "Preferences", "FontAttr", AppFont.Name & "|" & Int(AppFont.Size * 100) & "|" & CInt(AppFont.Bold) & "|" & AppFont.Charset & "|" & CInt(AppFont.Italic) & "|" & AppFont.Weight
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
    
    Set mNewFont = CloneFont(AppFont)
    lblFont.Caption = mNewFont.Name & " " & mNewFont.Size

End Sub

Private Sub lblFont_Click()
    cmdChooseFont_Click
End Sub

Private Function CloneFont(nOrigFont) As StdFont
    Dim iFont As New StdFont
    
    If nOrigFont Is Nothing Then Exit Function
    If Not TypeOf nOrigFont Is StdFont Then Exit Function
    
    iFont.Name = nOrigFont.Name
    iFont.Size = nOrigFont.Size
    iFont.Bold = nOrigFont.Bold
    iFont.Italic = nOrigFont.Italic
    iFont.Strikethrough = nOrigFont.Strikethrough
    iFont.Underline = nOrigFont.Underline
    iFont.Weight = nOrigFont.Weight
    iFont.Charset = nOrigFont.Charset
    
    Set CloneFont = iFont
End Function

