VERSION 5.00
Begin VB.Form frmSelectPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Printer"
   ClientHeight    =   1704
   ClientLeft      =   2820
   ClientTop       =   2160
   ClientWidth     =   4884
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
   ScaleHeight     =   1704
   ScaleWidth      =   4884
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   1944
      TabIndex        =   0
      Top             =   1032
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   3432
      TabIndex        =   1
      Top             =   1032
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.ComboBox cboPrinters 
      Height          =   336
      Left            =   1056
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   3660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer:"
      Height          =   348
      Left            =   192
      TabIndex        =   2
      Top             =   360
      Width           =   756
   End
End
Attribute VB_Name = "frmSelectPrinter"
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
    Dim c As Long
    
    For c = 0 To Printers.Count - 1
        If Printers(c).DeviceName = cboPrinters.Text Then
            Set Printer = Printers(c)
            frmMain.PrinterIndex = c
            Exit For
        End If
    Next
    OKPressed = True
    Unload Me
    If InStr(LCase$(Printer.DeviceName), "pdf") Then
        If Val(GetSetting(App.Title, AppPath4Reg, "HidePDFNote", "0")) = 0 Then
            frmPDFNote.Show vbModal
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim p As Printer
    
    Set Me.Icon = gIcon
    
    For Each p In Printers
        cboPrinters.AddItem p.DeviceName
        If p.DeviceName = Printer.DeviceName Then
            cboPrinters.ListIndex = cboPrinters.NewIndex
        End If
    Next
End Sub
