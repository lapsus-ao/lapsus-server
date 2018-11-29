Attribute VB_Name = "Puntos"
'Módulo de puntos de usuario
'Creado por Juan Andrés Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'Para LapsusAO 2.0
'04/09/2010

Public Sub repartirPuntos(ByVal ClanIndex As Integer, ByVal Cantidad As Integer, ByVal Casti As String)
Dim UsersOn As Integer
Dim Total As Integer
Dim i As Integer
UsersOn = 0

'CHOTS | Aca se fija cuantos users hay online del clan y lo guarda en ACUMULADOR
For i = 1 To LastUser
    If UserList(i).GuildIndex = ClanIndex Then
        UsersOn = UsersOn + 1
    End If
Next i

If UsersOn = 0 Then Exit Sub 'CHOTS | Evitamos dividir por cero

'CHOTS | Ahora divide la cantidad de puntos entre los Miembros
Total = Round(Cantidad / UsersOn, 0)

'CHOTS | Entrega el Marrón, digo los puntos...
For i = 1 To LastUser
    If UserList(i).GuildIndex = ClanIndex Then
        UserList(i).Puntos = UserList(i).Puntos + Total
        Call SendData(SendTarget.ToIndex, i, 0, "PRE27" & "," & Casti & "," & Total)
    End If
Next i
    
End Sub

Public Sub CambiarPuntos(ByVal UserIndex As Integer, ByVal item As Byte)
'CHOTS | Sistema de Cambio de Puntos
Dim Puntos As Integer
Dim Premio As Integer

Select Case item
    Case 0 'CHOTS | Espada Mata Dragones
        Puntos = 250
        Premio = 402
    Case 1 ' CHOTS | Baston del Dragon
        Puntos = 250
        Premio = 1006
    Case 2 'CHOTS | Espada Extermina Dragones
        Puntos = 500
        Premio = 1005
    Case 3 'CHOTS | Runa
        Puntos = 100
        Premio = 1121
    Case 4 'CHOTS | Poción Removedora
        Puntos = 200
        Premio = 1122
    Case 5 'CHOTS | Gema de Captura
        Puntos = 500
        Premio = 1138
    Case 6 'CHOTS | Galeon
        Puntos = 100
        Premio = 476
End Select

If UserList(UserIndex).Puntos < Puntos Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z79")
    Exit Sub
End If

UserList(UserIndex).Puntos = UserList(UserIndex).Puntos - Puntos

Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = Premio
If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Call UpdateUserInv(True, UserIndex, 0)

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has cambiado tus puntos. ¡Felicitaciones!." & FONTTYPE_INFON)


End Sub

Public Sub TransferirPuntos(ByVal UserIndex As Integer, ByVal Personaje As String, ByVal Puntos As Integer)
Dim PersonajeIndex As Integer
Dim PuntosTemp As Integer

PersonajeIndex = NameIndex(Personaje)

If (UserList(UserIndex).Puntos < Puntos) Or (Puntos <= 0) Then Exit Sub

If PersonajeIndex <> 0 Then 'Esta Online
    UserList(UserIndex).Puntos = UserList(UserIndex).Puntos - Puntos
    UserList(PersonajeIndex).Puntos = UserList(PersonajeIndex).Puntos + Puntos
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Transferencia exitosa!" & FONTTYPE_ORO)
    Call SendData(SendTarget.ToIndex, PersonajeIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha transferido " & Puntos & " Puntos de Usuario!" & FONTTYPE_ORO)
Else 'No esta Online
    If Not FileExist(CharPath & UCase$(Personaje) & ".chr") Then Exit Sub
    UserList(UserIndex).Puntos = UserList(UserIndex).Puntos - Puntos
    PuntosTemp = val(GetVar(CharPath & UCase$(Personaje) & ".chr", "INIT", "Puntos"))
    PuntosTemp = PuntosTemp + Puntos
    Call WriteVar(CharPath & UCase$(Personaje) & ".chr", "INIT", "Puntos", PuntosTemp)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Transferencia exitosa!" & FONTTYPE_ORO)
End If

End Sub
