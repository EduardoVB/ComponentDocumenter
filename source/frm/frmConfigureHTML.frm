VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConfigureHTML 
   Caption         =   "Configure HTML texts"
   ClientHeight    =   7236
   ClientLeft      =   7440
   ClientTop       =   2148
   ClientWidth     =   9384
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
   MinButton       =   0   'False
   ScaleHeight     =   7236
   ScaleWidth      =   9384
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   420
      Left            =   5712
      TabIndex        =   1
      Top             =   6048
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   7200
      TabIndex        =   0
      Top             =   6048
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin TabDlg.SSTab sst1 
      Height          =   5796
      Left            =   96
      TabIndex        =   2
      Top             =   72
      Width           =   8636
      _ExtentX        =   15240
      _ExtentY        =   10224
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   529
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "HEAD"
      TabPicture(0)   =   "frmConfigureHTML.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt(0)"
      Tab(0).Control(1)=   "lblTitle(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Style sheet"
      TabPicture(1)   =   "frmConfigureHTML.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt(1)"
      Tab(1).Control(1)=   "lblTitle(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Page header for multiple pages"
      TabPicture(2)   =   "frmConfigureHTML.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt(2)"
      Tab(2).Control(1)=   "lblTitle(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Page header for one page"
      TabPicture(3)   =   "frmConfigureHTML.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt(3)"
      Tab(3).Control(1)=   "lblTitle(3)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Page footer"
      TabPicture(4)   =   "frmConfigureHTML.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "lblTitle(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txt(4)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5172
         Index           =   4
         Left            =   96
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   700
         Width           =   7788
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5172
         Index           =   3
         Left            =   -74784
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   700
         Width           =   7788
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5412
         Index           =   2
         Left            =   -74712
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   700
         Width           =   7788
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5412
         Index           =   1
         Left            =   -74784
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   700
         Width           =   7788
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5412
         Index           =   0
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   648
         Width           =   7788
      End
      Begin VB.Label lblTitle 
         Caption         =   "HTML code that will be added inmediately before the </body> closing tag"
         ForeColor       =   &H00FA0A22&
         Height          =   300
         Index           =   4
         Left            =   60
         TabIndex        =   11
         Top             =   420
         Width           =   8004
      End
      Begin VB.Label lblTitle 
         Caption         =   "HTML code that will be added inmediately after the <body> opening tag"
         ForeColor       =   &H00FA0A22&
         Height          =   300
         Index           =   2
         Left            =   -74880
         TabIndex        =   7
         Top             =   400
         Width           =   8000
      End
      Begin VB.Label lblTitle 
         Caption         =   "Stylesheet code"
         ForeColor       =   &H00FA0A22&
         Height          =   300
         Index           =   1
         Left            =   -74880
         TabIndex        =   5
         Top             =   400
         Width           =   8000
      End
      Begin VB.Label lblTitle 
         Caption         =   "HTML code correspondig to <head></head> section"
         ForeColor       =   &H00FA0A22&
         Height          =   300
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   400
         Width           =   8000
      End
      Begin VB.Label lblTitle 
         Caption         =   "HTML code that will be added inmediately after the <body> opening tag when using the one page option"
         ForeColor       =   &H00FA0A22&
         Height          =   300
         Index           =   3
         Left            =   -74880
         TabIndex        =   9
         Top             =   400
         Width           =   9000
      End
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "Note: leave blank to restore default. Placeholders are between braces (do not modify). Use valid HTM code, it won't be validated."
      Height          =   720
      Left            =   240
      TabIndex        =   13
      Top             =   5952
      Width           =   4980
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmConfigureHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean

Private mTxt(4) As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    OKPressed = True
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
    sst1.Tab = 0
End Sub

Private Sub Form_Resize()
    Dim t As Long
    Dim tp As Long
    Dim iExtraHeight As Long
    
    If Me.Width < 9400 Then Me.Width = 9400
    If Me.Height < 4000 Then Me.Height = 4000
    
    iExtraHeight = 800
    If (lblNote.Height + 100) > iExtraHeight Then iExtraHeight = lblNote.Height + 100
    
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 500, Me.ScaleHeight - cmdCancel.Height - 140
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 500, cmdCancel.Top
    lblNote.Width = cmdOK.Left - lblNote.Left - 100
    
    sst1.Move 60, 60 + 60, Me.ScaleWidth - 120, Me.ScaleHeight - iExtraHeight
    tp = sst1.Tab
    For t = 0 To sst1.Tabs - 1
        sst1.Tab = t
        lblTitle(t).Move 120, 400
        txt(t).Move Screen.TwipsPerPixelX, Screen.TwipsPerPixelY + lblTitle(t).Top + lblTitle(t).Height, sst1.Width - Screen.TwipsPerPixelX * 2, sst1.Height - Screen.TwipsPerPixelY * 2 - (lblTitle(t).Top + lblTitle(t).Height)
    Next
    sst1.Tab = tp
    
    lblNote.Top = (sst1.Top + sst1.Height + Me.ScaleHeight - lblNote.Height) / 2
End Sub

Private Sub txt_Change(Index As Integer)
    mTxt(Index) = txt(Index).Text
End Sub


Public Property Let HTML_HeadSection(str As String)
    txt(0).Text = str
End Property

Public Property Get HTML_HeadSection() As String
    HTML_HeadSection = mTxt(0)
End Property


Public Property Let HTML_StyleSheet(str As String)
    txt(1).Text = str
End Property

Public Property Get HTML_StyleSheet() As String
    HTML_StyleSheet = mTxt(1)
End Property


Public Property Let HTML_PageHeaderMP(str As String)
    txt(2).Text = str
End Property

Public Property Get HTML_PageHeaderMP() As String
    HTML_PageHeaderMP = mTxt(2)
End Property


Public Property Let HTML_PageHeaderOP(str As String)
    txt(3).Text = str
End Property

Public Property Get HTML_PageHeaderOP() As String
    HTML_PageHeaderOP = mTxt(3)
End Property


Public Property Let HTML_PageFooter(str As String)
    txt(4).Text = str
End Property

Public Property Get HTML_PageFooter() As String
    HTML_PageFooter = mTxt(4)
End Property


