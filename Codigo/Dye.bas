Attribute VB_Name = "Dye"
Option Explicit
'MÓDULO PROGRAMADO POR JOSÉ IGNACIO PARODI (DYE)
'PARA TWISTEROS AO 2010
'REPROGRAMADO Y ADAPTADO POR CHOTS
'PARA LAPSUS AO 2.1
'24/11/2010

Public Function DyeDecifro(ByVal Slot As Integer, ByRef data As String) As String
Dim Buffer() As Byte
Dim OutBuffer() As Byte
Dim i As Long

   ReDim Buffer(Len(data) - 1) As Byte
   ReDim OutBuffer(Len(data) - 1) As Byte
    
    Buffer = StrConv(data, vbFromUnicode)
    
    OutBuffer(0) = Buffer(0) Xor UserList(Slot).ClavePublica
    
    For i = 1 To (Len(data) - 1)
        OutBuffer(i) = Buffer(i) Xor Buffer(i - 1)
    Next i
        UserList(Slot).ClavePublica = Buffer(Len(data) - 1)
     
    DyeDecifro = StrConv(OutBuffer, vbUnicode)

End Function


Public Function DyeCifro(ByVal UserIndex As Integer, ByRef inData As String) As String
Dim Buffer() As Byte
Dim OutBuffer() As Byte
Dim i As Long

   Buffer = StrConv(inData, vbFromUnicode)
   ReDim OutBuffer(Len(inData) - 1) As Byte
   OutBuffer(0) = Buffer(0) Xor UserList(UserIndex).ClavePrivada
   For i = 1 To (Len(inData) - 1)
        OutBuffer(i) = Buffer(i) Xor OutBuffer(i - 1)
   Next i

UserList(UserIndex).ClavePrivada = OutBuffer(Len(inData) - 1)
DyeCifro = StrConv(OutBuffer, vbUnicode)
End Function
