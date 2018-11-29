Attribute VB_Name = "Matematicas"
'Lapsus AO 2009
'Copyright (C) 2009 Dalmasso, Juan Andres
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
    Porcentaje = (Total * Porc) / 100
End Function

Public Function SD(ByVal n As Integer) As Integer
'Call LogTarea("Function SD n:" & n)
'Suma digitos

Do
    SD = SD + (n Mod 10)
    n = n \ 10
Loop While (n > 0)

End Function

Public Function SDM(ByVal n As Integer) As Integer
'Call LogTarea("Function SDM n:" & n)
'Suma digitos cada digito menos dos

Do
    SDM = SDM + (n Mod 10) - 1
    n = n \ 10
Loop While (n > 0)

End Function

Public Function Complex(ByVal n As Integer) As Integer
'Call LogTarea("Complex")

If n Mod 2 <> 0 Then
    Complex = n * SD(n)
Else
    Complex = n * SDM(n)
End If

End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
    'Encuentra la distancia entre dos WorldPos
    Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)
End Function

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double

'Encuentra la distancia entre dos puntos

Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function

Public Function RandomNumberOld(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Generates a random number in the range given - recoded to use longs and work properly with ranges
'**************************************************************
    Randomize Timer
    
    RandomNumberOld = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
End Function

'CHOTS | Nuevas funciones de Random Number obtenidas de GS
Public Sub RandomNumberInitialize()
    Randomize GetTickCount
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
