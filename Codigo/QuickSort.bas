Attribute VB_Name = "Module1"
Dim vecMenor() As Long, vecMayor() As Long
Dim entre As Boolean
Dim Temp As Long
Dim a, sl, pRight, pLeft, iEnd, Max, star, pos As Long

Public Sub ordenar2(elVector() As Long)

If Not entre Then
   pLeft = star
   pRight = iEnd
   pos = star
   entre = False
End If
   

While elVector(pos) <= elVector(pRight) And (pos <> pRight)
      pRight = pRight - 1
Wend
    
    If pos = pRight Then
       Exit Sub
    ElseIf elVector(pos) > elVector(pRight) Then
            Temp = elVector(pos)
            elVector(pos) = elVector(pRight)
            elVector(pRight) = Temp
            pos = pRight
    End If
    

    While elVector(pLeft) <= elVector(pos) And (pLeft <> pos)
        pLeft = pLeft + 1
    Wend
    
    If pos = pLeft Then
       Exit Sub
    ElseIf elVector(pLeft) > elVector(pos) Then
       Temp = elVector(pos)
       elVector(pos) = elVector(pLeft)
       elVector(pLeft) = Temp
       pos = pLeft
      
    End If
    
    ordenar2 elVector
    
End Sub

Public Function Ordenar(elVector() As Long) As Long()

ReDim vecMenor(UBound(elVector)) As Long
ReDim vecMayor(UBound(elVector)) As Long

 Max = 0
 If UBound(elVector) > 1 Then
    Max = Max + 1
    vecMenor(1) = 1
    vecMayor(1) = UBound(elVector)
 End If
 
 While Max <> 0
   star = vecMenor(Max)
   iEnd = vecMayor(Max)
   Max = Max - 1
 
 ordenar2 elVector
 
 If star < pos - 1 Then
    Max = Max + 1
    vecMenor(Max) = star
    vecMayor(Max) = pos - 1
 End If
 
 If pos + 1 < iEnd Then
    Max = Max + 1
    vecMenor(Max) = pos + 1
    vecMayor(Max) = iEnd
 End If
Wend

Ordenar = elVector

End Function
