VERSION 5.00
Begin VB.Form frmSelectComponentDB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Component databases"
   ClientHeight    =   5376
   ClientLeft      =   5940
   ClientTop       =   2052
   ClientWidth     =   5520
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
   ScaleHeight     =   5376
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4096
      TabIndex        =   0
      Top             =   4752
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   2608
      TabIndex        =   1
      Top             =   4752
      UseMaskColor    =   -1  'True
      Width           =   1092
   End
   Begin VB.ListBox lstComponentsDB 
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
      Height          =   4440
      Left            =   72
      TabIndex        =   2
      Top             =   96
      Width           =   5380
   End
End
Attribute VB_Name = "frmSelectComponentDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DBPath As String
Private mDBPath As String
Private mPaths() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    DBPath = mDBPath
    Unload Me
End Sub

Private Sub Form_Load()
    Dim iFile As String
    Dim iDb As Database
    Dim iGi As Recordset
    
    Set Me.Icon = gIcon
    If Not FolderExists(App_Path & "\databases") Then
        MsgBox "Database folders does not exist", vbCritical
        Unload Me
        Exit Sub
    End If
    mPaths = Split("")
    iFile = Dir(App_Path & "\databases\*.mdb")
    On Error GoTo FileErr
    Do Until iFile = ""
        Set iDb = DBEngine.OpenDatabase(App_Path & "\databases\" & iFile)
        Set iGi = iDb.OpenRecordset("General_Information")
        iGi.Index = "Name"
        iGi.Seek "=", "ComponentName"
        If Not iGi.NoMatch Then
            ReDim Preserve mPaths(UBound(mPaths) + 1)
            mPaths(UBound(mPaths)) = App_Path & "\databases\" & iFile
            lstComponentsDB.AddItem iGi!Value & " (" & iFile & ")"
        End If
        iDb.Close
NextFile:
        iFile = Dir
    Loop
    Exit Sub
    
FileErr:
    Resume NextFile
End Sub

Private Sub lstComponentsDB_Click()
    If lstComponentsDB.ListIndex > -1 Then
        mDBPath = mPaths(lstComponentsDB.ListIndex)
    End If
End Sub

Private Sub lstComponentsDB_DblClick()
    If lstComponentsDB.ListCount > 0 Then
        If lstComponentsDB.ListIndex = -1 Then
            lstComponentsDB.ListIndex = lstComponentsDB.ListCount - 1
        End If
    End If
    cmdOK.Value = 1
End Sub
