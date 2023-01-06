VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmartConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Separator As String

Private m_MediumGrowSize As Long
Private m_CurrentMediumSize As Long
Private m_CurrentMediumIndex As Long
Private m_MediumStrings() As String

Private m_MaxMediumLength As Long
Private m_SmallTotalLength As Long

Private m_MaxSmallEntries As Long
Private m_SmallIndex As Long
Private m_SmallStrings() As String

Public Function GenerateCurrentString() As String
    If m_SmallIndex Then ClearSmallStrings
    If m_CurrentMediumIndex Then
        'Shrink the size of the string array so that join doesn't get extra stuff
        ReDim Preserve m_MediumStrings(m_CurrentMediumIndex - 1)
        GenerateCurrentString = Join(m_MediumStrings, m_Separator)
        ReDim Preserve m_MediumStrings(m_CurrentMediumSize - 1)
    End If
End Function

Public Sub ClearStrings()
    Dim l As Long
    If m_SmallIndex Then
        For l = 0 To m_SmallIndex - 1
            m_SmallStrings(l) = vbNullString
        Next l
    End If
    m_SmallIndex = 0
    m_SmallTotalLength = 0
    m_CurrentMediumIndex = 0
    m_CurrentMediumSize = 0
    Erase m_MediumStrings
End Sub

Public Sub AddString(NewString As String)
    Dim NewLen As Long
    NewLen = Len(NewString)
    If m_SmallTotalLength + NewLen > m_MaxMediumLength Then
        ClearSmallStrings
    End If
    m_SmallTotalLength = m_SmallTotalLength + NewLen
    m_SmallStrings(m_SmallIndex) = NewString
    m_SmallIndex = m_SmallIndex + 1
    If m_SmallIndex = m_MaxSmallEntries Then
        'Clear out now
        ClearSmallStrings
    End If
End Sub

Public Property Get TempStringCount() As Long
    TempStringCount = m_MaxSmallEntries
End Property

Public Property Let TempStringCount(ByVal RHS As Long)
    If RHS < 1 Then Err.Raise 5
    If m_SmallIndex Then ClearSmallStrings
    m_MaxSmallEntries = RHS
    ReDim Preserve m_SmallStrings(RHS - 1)
End Property

Public Property Get FinalCacheGrowSize() As Long
    FinalCacheGrowSize = m_MediumGrowSize
End Property

Public Property Let FinalCacheGrowSize(ByVal RHS As Long)
    If RHS < 1 Then Err.Raise 5
    If m_CurrentMediumIndex Then Err.Raise 5, , "Call before AddString or after ClearStrings"
    m_MediumGrowSize = RHS
End Property

Public Property Get MaxTempLength() As Long
    MaxTempLength = m_MaxMediumLength
End Property

Public Property Let MaxTempLength(ByVal RHS As Long)
    If RHS < 1 Then Err.Raise 5
    m_MaxMediumLength = RHS
End Property

Public Property Get Separator() As String
    Separator = m_Separator
End Property

Public Property Let Separator(ByVal RHS As String)
    m_Separator = RHS
End Property

Private Function NextMediumIndex() As Long
    If (m_CurrentMediumIndex Mod m_MediumGrowSize) = 0 Then
        m_CurrentMediumSize = m_CurrentMediumSize + m_MediumGrowSize
        ReDim Preserve m_MediumStrings(m_CurrentMediumSize - 1)
    End If
    NextMediumIndex = m_CurrentMediumIndex
    m_CurrentMediumIndex = m_CurrentMediumIndex + 1
End Function

'Use the Join function to generate a medium length string
'and move it to our medium length
Private Sub ClearSmallStrings()
    Dim iNextMediumIndex As Long
    'Debug.Assert m_SmallIndex

    'Temporarily shrink the array to stop Join from adding extra separators
    'This isn't as bad as it seems because we'll be growing back to the same
    'size, so it will likely reoccupy the same memory.  Although you can 't Stop
    'the memory from relocatin, you can see if it actually happened by looking
    'at VarPtr(m_SmallStrings(0)) before and after each ReDim Preserve.
    If m_SmallIndex > 0 Then
        ReDim Preserve m_SmallStrings(m_SmallIndex - 1)

        iNextMediumIndex = NextMediumIndex 'Note: Native/Fast bug, don't do NextMediumIndex inline
        m_MediumStrings(iNextMediumIndex) = Join(m_SmallStrings, m_Separator)

        'We could clear all of the current strings here, but
        'it turns out that it is slightly faster to just leave them
        'alone and let them clear out naturally as the buffer is reused
        'Dim l As Long
        'For l = 0 To m_SmallIndex - 1
        '    m_SmallStrings(l) = vbNullString
        'Next l

        'Put the array size back where it should be
        ReDim Preserve m_SmallStrings(m_MaxSmallEntries - 1)

        'Clear the current length and index
        m_SmallIndex = 0
        m_SmallTotalLength = 0
    End If
End Sub

Private Sub Class_Initialize()
    'Set default and initialize array
    m_MaxMediumLength = 4095
    m_MaxSmallEntries = 128
    m_MediumGrowSize = 64
    ReDim m_SmallStrings(m_MaxSmallEntries - 1)
End Sub
