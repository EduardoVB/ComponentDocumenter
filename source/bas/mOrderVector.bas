Attribute VB_Name = "mOrderVector"
Option Explicit
Option Compare Text

Private mMainVector As Variant
Private mSecondaryVector1 As Variant
Private mHasSecondaryVector1 As Boolean
Private mSecondaryVector1_IsObject As Boolean
Private mSecondaryVector2 As Variant
Private mHasSecondaryVector2 As Boolean
Private mSecondaryVector2_IsObject As Boolean
Private mSecondaryVector3 As Variant
Private mHasSecondaryVector3 As Boolean
Private mSecondaryVector3_IsObject As Boolean
Private mSecondaryVector4 As Variant
Private mHasSecondaryVector4 As Boolean
Private mSecondaryVector4_IsObject As Boolean
Private mSecondaryVector5 As Variant
Private mHasSecondaryVector5 As Boolean
Private mSecondaryVector5_IsObject As Boolean
Private mSecondaryVector6 As Variant
Private mHasSecondaryVector6 As Boolean
Private mSecondaryVector6_IsObject As Boolean

Private mBinaryCompare As Boolean
Private mOrderDescending As Boolean

Public Sub OrderVector(ByRef nMainVector As Variant, Optional ByRef nSecondaryVector1 As Variant, Optional ByRef nSecondaryVector2 As Variant, Optional ByRef nSecondaryVector3 As Variant, Optional ByRef nSecondaryVector4 As Variant, Optional ByRef nSecondaryVector5 As Variant, Optional ByRef nSecondaryVector6 As Variant, Optional nBynaryCompare As Boolean, Optional nOrderDescending As Boolean)
    
    mOrderDescending = nOrderDescending
    mBinaryCompare = nBynaryCompare
    mMainVector = nMainVector
    
    mSecondaryVector1_IsObject = False
    mSecondaryVector2_IsObject = False
    mSecondaryVector3_IsObject = False
    mSecondaryVector4_IsObject = False
    mSecondaryVector5_IsObject = False
    mSecondaryVector6_IsObject = False
    
    If IsMissing(nSecondaryVector1) Then
        mHasSecondaryVector1 = False
    Else
        mHasSecondaryVector1 = True
        mSecondaryVector1 = nSecondaryVector1
        If VarType(nSecondaryVector1(UBound(nSecondaryVector1))) = vbObject Then mSecondaryVector1_IsObject = True
    End If
    
    If IsMissing(nSecondaryVector2) Then
        mHasSecondaryVector2 = False
    Else
        mHasSecondaryVector2 = True
        mSecondaryVector2 = nSecondaryVector2
        If VarType(nSecondaryVector2(UBound(nSecondaryVector2))) = vbObject Then mSecondaryVector2_IsObject = True
    End If
    
    If IsMissing(nSecondaryVector3) Then
        mHasSecondaryVector3 = False
    Else
        mHasSecondaryVector3 = True
        mSecondaryVector3 = nSecondaryVector3
        If VarType(nSecondaryVector3(UBound(nSecondaryVector3))) = vbObject Then mSecondaryVector3_IsObject = True
    End If
    
    If IsMissing(nSecondaryVector4) Then
        mHasSecondaryVector4 = False
    Else
        mHasSecondaryVector4 = True
        mSecondaryVector4 = nSecondaryVector4
        If VarType(nSecondaryVector4(UBound(nSecondaryVector4))) = vbObject Then mSecondaryVector4_IsObject = True
    End If
    
    If IsMissing(nSecondaryVector5) Then
        mHasSecondaryVector5 = False
    Else
        mHasSecondaryVector5 = True
        mSecondaryVector5 = nSecondaryVector5
        If VarType(nSecondaryVector5(UBound(nSecondaryVector5))) = vbObject Then mSecondaryVector5_IsObject = True
    End If
    
    If IsMissing(nSecondaryVector6) Then
        mHasSecondaryVector6 = False
    Else
        mHasSecondaryVector6 = True
        mSecondaryVector6 = nSecondaryVector6
        If VarType(nSecondaryVector6(UBound(nSecondaryVector6))) = vbObject Then mSecondaryVector6_IsObject = True
    End If
    
    OrderElements
    
    nMainVector = mMainVector
    
    If mHasSecondaryVector1 Then
        nSecondaryVector1 = mSecondaryVector1
    End If

    If mHasSecondaryVector2 Then
        nSecondaryVector2 = mSecondaryVector2
    End If

    If mHasSecondaryVector3 Then
        nSecondaryVector3 = mSecondaryVector3
    End If

    If mHasSecondaryVector4 Then
        nSecondaryVector4 = mSecondaryVector4
    End If

    If mHasSecondaryVector5 Then
        nSecondaryVector5 = mSecondaryVector5
    End If

    If mHasSecondaryVector6 Then
        nSecondaryVector6 = mSecondaryVector6
    End If

End Sub

Private Sub OrderElements(Optional nFirstElement As Variant, Optional nLastElement As Variant)
    Dim iFirstElement As Long
    Dim iLastElement As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim iLb As Long
    
    If IsMissing(nFirstElement) Then
        iLb = LBound(mMainVector)
        Select Case VarType(mMainVector(iLb))
            Case vbString
                If mMainVector(iLb) <> "" Then
                    iFirstElement = iLb
                Else
                    If iLb = 0 Then
                        iFirstElement = 1
                    Else
                        iFirstElement = iLb
                    End If
                End If
            Case Else
                If mMainVector(iLb) <> 0 Then
                    iFirstElement = iLb
                Else
                    If iLb = 0 Then
                        iFirstElement = 1
                    Else
                        iFirstElement = iLb
                    End If
                End If
        End Select
    Else
        iFirstElement = nFirstElement
    End If
    If IsMissing(nLastElement) Then
        iLastElement = UBound(mMainVector)
    Else
        iLastElement = nLastElement
    End If
    
    On Error GoTo TheExit:
    
    If iFirstElement < iLastElement Then
        If (iLastElement - iFirstElement) = 1 Then
            If CompareElements(mMainVector(iFirstElement), mMainVector(iLastElement)) > 0 Then
                Call ExchangeElements(iFirstElement, iLastElement)
            End If
        Else
            Call ExchangeElements(iLastElement, Random(iFirstElement, iLastElement))
            iMin = iFirstElement
            iMax = iLastElement
            Do
                Do While (iMin < iMax) And CompareElements(mMainVector(iMin), mMainVector(iLastElement)) <= 0
                    iMin = iMin + 1
                Loop
                Do While (iMin < iMax) And CompareElements(mMainVector(iMax), mMainVector(iLastElement)) >= 0
                    iMax = iMax - 1
                Loop
                If iMin < iMax Then
                    Call ExchangeElements(iMin, iMax)
                End If
            Loop While iMin < iMax
            Call ExchangeElements(iMin, iLastElement)
            If (iMin - iFirstElement) < (iLastElement - iMin) Then
                Call OrderElements(iFirstElement, (iMin - 1))
                Call OrderElements((iMin + 1), iLastElement)
            Else
                Call OrderElements((iMin + 1), iLastElement)
                Call OrderElements(iFirstElement, (iMin - 1))
            End If
        End If
    End If
    Exit Sub
    
TheExit:
End Sub

Private Function CompareElements(nValue1 As Variant, nValue2 As Variant) As Integer
    If mBinaryCompare Then
        CompareElements = StrComp(nValue1, nValue2, vbBinaryCompare)
    Else
        If nValue1 < nValue2 Then
            CompareElements = -1
        Else
            If nValue1 = nValue2 Then
                CompareElements = 0
            Else
                CompareElements = 1
            End If
        End If
    End If
    If mOrderDescending Then
        CompareElements = CompareElements * -1
    End If
End Function


Private Sub ExchangeElements(nIndex1 As Long, nIndex2 As Long)
    Dim Aux As Variant
    Dim iObj As Object
    
    Aux = mMainVector(nIndex1)
    mMainVector(nIndex1) = mMainVector(nIndex2)
    mMainVector(nIndex2) = Aux
    
    If mHasSecondaryVector1 Then
        If mSecondaryVector1_IsObject Then
            Set iObj = mSecondaryVector1(nIndex1)
            Set mSecondaryVector1(nIndex1) = mSecondaryVector1(nIndex2)
            Set mSecondaryVector1(nIndex2) = iObj
        Else
            Aux = mSecondaryVector1(nIndex1)
            mSecondaryVector1(nIndex1) = mSecondaryVector1(nIndex2)
            mSecondaryVector1(nIndex2) = Aux
        End If
    End If

    If mHasSecondaryVector2 Then
        If mSecondaryVector2_IsObject Then
            Set iObj = mSecondaryVector2(nIndex1)
            Set mSecondaryVector2(nIndex1) = mSecondaryVector2(nIndex2)
            Set mSecondaryVector2(nIndex2) = iObj
        Else
            Aux = mSecondaryVector2(nIndex1)
            mSecondaryVector2(nIndex1) = mSecondaryVector2(nIndex2)
            mSecondaryVector2(nIndex2) = Aux
        End If
    End If

    If mHasSecondaryVector3 Then
        If mSecondaryVector3_IsObject Then
            Set iObj = mSecondaryVector3(nIndex1)
            Set mSecondaryVector3(nIndex1) = mSecondaryVector3(nIndex2)
            Set mSecondaryVector3(nIndex2) = iObj
        Else
            Aux = mSecondaryVector3(nIndex1)
            mSecondaryVector3(nIndex1) = mSecondaryVector3(nIndex2)
            mSecondaryVector3(nIndex2) = Aux
        End If
    End If

    If mHasSecondaryVector4 Then
        If mSecondaryVector4_IsObject Then
            Set iObj = mSecondaryVector4(nIndex1)
            Set mSecondaryVector4(nIndex1) = mSecondaryVector4(nIndex2)
            Set mSecondaryVector4(nIndex2) = iObj
        Else
            Aux = mSecondaryVector4(nIndex1)
            mSecondaryVector4(nIndex1) = mSecondaryVector4(nIndex2)
            mSecondaryVector4(nIndex2) = Aux
        End If
    End If

    If mHasSecondaryVector5 Then
        If mSecondaryVector5_IsObject Then
            Set iObj = mSecondaryVector5(nIndex1)
            Set mSecondaryVector5(nIndex1) = mSecondaryVector5(nIndex2)
            Set mSecondaryVector5(nIndex2) = iObj
        Else
            Aux = mSecondaryVector5(nIndex1)
            mSecondaryVector5(nIndex1) = mSecondaryVector5(nIndex2)
            mSecondaryVector5(nIndex2) = Aux
        End If
    End If

    If mHasSecondaryVector6 Then
        If mSecondaryVector6_IsObject Then
            Set iObj = mSecondaryVector6(nIndex1)
            Set mSecondaryVector6(nIndex1) = mSecondaryVector6(nIndex2)
            Set mSecondaryVector6(nIndex2) = iObj
        Else
            Aux = mSecondaryVector6(nIndex1)
            mSecondaryVector6(nIndex1) = mSecondaryVector6(nIndex2)
            mSecondaryVector6(nIndex2) = Aux
        End If
    End If

End Sub


Private Function Random(nFirst As Long, nLast As Long) As Long
    Random = nFirst + Rnd * (nLast - nFirst)
End Function
