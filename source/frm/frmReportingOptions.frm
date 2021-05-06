VERSION 5.00
Begin VB.Form frmReportingOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporting Options"
   ClientHeight    =   6408
   ClientLeft      =   2820
   ClientTop       =   2160
   ClientWidth     =   4416
   ControlBox      =   0   'False
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
   ScaleHeight     =   6408
   ScaleWidth      =   4416
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Print:"
      Height          =   1352
      Left            =   192
      TabIndex        =   8
      Top             =   4056
      Width           =   3972
      Begin VB.OptionButton optPrint 
         Caption         =   "One item per page"
         Height          =   300
         Index           =   1
         Left            =   264
         TabIndex        =   10
         Top             =   816
         Width           =   3492
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Contiguous"
         Height          =   300
         Index           =   0
         Left            =   264
         TabIndex        =   9
         Top             =   384
         Value           =   -1  'True
         Width           =   3492
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "HTML:"
      Height          =   3684
      Left            =   192
      TabIndex        =   1
      Top             =   192
      Width           =   3972
      Begin VB.CheckBox chkReplaceCSSFile 
         Caption         =   "Replace existent style file"
         Enabled         =   0   'False
         Height          =   396
         Left            =   264
         TabIndex        =   6
         Top             =   2232
         Width           =   3228
      End
      Begin VB.CommandButton cmdConfigHTML 
         Caption         =   "Configure HTML texts"
         Height          =   420
         Left            =   264
         TabIndex        =   7
         Top             =   2904
         UseMaskColor    =   -1  'True
         Width           =   2092
      End
      Begin VB.CheckBox chkExternalCSS 
         Caption         =   "Stylesheet in external file (styles.css)"
         Height          =   396
         Left            =   264
         TabIndex        =   5
         Top             =   1800
         Width           =   3228
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "One page per Property/Method/Event"
         Height          =   300
         Index           =   2
         Left            =   264
         TabIndex        =   4
         Top             =   1296
         Width           =   3492
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "One page per Control/Class"
         Height          =   300
         Index           =   1
         Left            =   264
         TabIndex        =   3
         Top             =   864
         Width           =   3492
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "One page for all"
         Height          =   300
         Index           =   0
         Left            =   264
         TabIndex        =   2
         Top             =   432
         Value           =   -1  'True
         Width           =   3492
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   3048
      TabIndex        =   0
      Top             =   5736
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "These settings are stored per component database"
      ForeColor       =   &H00FF0000&
      Height          =   516
      Left            =   240
      TabIndex        =   11
      Top             =   5566
      Width           =   2464
   End
End
Attribute VB_Name = "frmReportingOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HTML_Mode As Long
Public Print_Mode As Long
Public ExternalCSS As Long
Public ReplaceCSSFile As Long

Public HTML_HeadSection As String
Public HTML_StyleSheet As String
Public HTML_PageHeaderMP As String
Public HTML_PageHeaderOP As String
Public HTML_PageFooter As String

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub chkExternalCSS_Click()
    ExternalCSS = chkExternalCSS.Value
    chkReplaceCSSFile.Enabled = ExternalCSS
End Sub

Private Sub chkReplaceCSSFile_Click()
    ReplaceCSSFile = chkReplaceCSSFile.Value
End Sub

Private Sub cmdConfigHTML_Click()
    frmConfigureHTML.HTML_HeadSection = HTML_HeadSection
    frmConfigureHTML.HTML_StyleSheet = HTML_StyleSheet
    frmConfigureHTML.HTML_PageHeaderMP = HTML_PageHeaderMP
    frmConfigureHTML.HTML_PageHeaderOP = HTML_PageHeaderOP
    frmConfigureHTML.HTML_PageFooter = HTML_PageFooter
    frmConfigureHTML.Show vbModal
    If frmConfigureHTML.OKPressed Then
        HTML_HeadSection = frmConfigureHTML.HTML_HeadSection
        HTML_StyleSheet = frmConfigureHTML.HTML_StyleSheet
        HTML_PageHeaderMP = frmConfigureHTML.HTML_PageHeaderMP
        HTML_PageHeaderOP = frmConfigureHTML.HTML_PageHeaderOP
        HTML_PageFooter = frmConfigureHTML.HTML_PageFooter
    End If
    Set frmConfigureHTML = Nothing
End Sub

Private Sub Form_Load()
    Set Me.Icon = gIcon
End Sub

Private Sub optHTML_Click(Index As Integer)
    HTML_Mode = Index
End Sub

Private Sub optPrint_Click(Index As Integer)
    Print_Mode = Index
End Sub
