Attribute VB_Name = "Duelos2"
'MÓDULO DE DUELOS 2VS2
'CREADO POR JUAN ANDRES DALMASSO (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'EL 07/01/12
'PARA LAPSUS 3.1
Public Const DUELO2_MAPADUELO As Integer = 90
Public Const DUELO2_MINX As Integer = 20
Public Const DUELO2_MAXX As Integer = 40
Public Const DUELO2_MINY As Integer = 20
Public Const DUELO2_MAXY As Integer = 40
Public Type tPareja
    usuario1 As Integer
    usuario2 As Integer
End Type
Public DUELO_PAREJA1 As tPareja
Public DUELO_PAREJA2 As tPareja


Public Function puedePareja(ByVal UserIndex1 As Integer, ByVal UserIndex2 As Integer, ByRef error As String) As Boolean

puedePareja = True

If UserList(UserIndex1).flags.TargetNPC = 0 Then
    error = "Primero debes clickear en el NPC"
    puedePareja = False
    Exit Function
End If

If Npclist(UserList(UserIndex1).flags.TargetNPC).NPCtype <> eNPCType.Duelero Then
    error = "El NPC no organiza duelos!"
    puedePareja = False
    Exit Function
End If

If Distancia(UserList(UserIndex1).Pos, Npclist(UserList(UserIndex1).flags.TargetNPC).Pos) > 2 Then
    error = "Estás demasiado lejos"
    puedePareja = False
    Exit Function
End If

If UserList(UserIndex1).flags.Muerto = 1 Or UserList(UserIndex2).flags.Muerto = 1 Then
    error = "No puedes ingresar a un duelo estando muerto"
    puedePareja = False
    Exit Function
End If

If UserList(UserIndex1).Pos.Map <> UserList(UserIndex2).Pos.Map Then
    error = "Tu pareja no se encuentra en Ullathorpe"
    puedePareja = False
    Exit Function
End If

If Distancia(UserList(UserIndex1).Pos, UserList(UserIndex2).Pos) > 5 Then
    error = "Tu pareja está demasiado lejos"
    puedePareja = False
    Exit Function
End If


If EsNewbie(UserIndex1) Or EsNewbie(UserIndex2) Then
    error = "Los newbies no tienen permitido ingresar a los duelos!"
    puedePareja = False
    Exit Function
End If

If MapInfo(DUELO_MAPADUELO2).NumUsers >= 4 Then
    error = "La Sala de duelos está ocupada"
    puedePareja = False
    Exit Function
End If

End Function

Public Sub ingresarDueloPareja(ByVal UserIndex1 As Integer, ByVal UserIndex2 As Integer)
Dim nPos As WorldPos
Dim Pos As WorldPos

Pos.Map = DUELO2_MAPADUELO
Pos.X = RandomNumber(DUELO2_MINX, DUELO2_MAXX)
Pos.Y = RandomNumber(DUELO2_MINY, DUELO2_MAXY)

UserList(UserIndex1).flags.enDuelo = True
UserList(UserIndex2).flags.enDuelo = True

If MapInfo(DUELO_MAPADUELO).NumUsers = 2 Then 'CHOTS | Ya hay una pareja esperando
    Call SendData(SendTarget.ToMap, 0, DUELO_MAPADUELO, ServerPackages.dialogo & "Duelos2vs2> " & UserList(UserIndex1).Name & " y " & UserList(UserIndex2).Name & " han aceptado el duelo!" & FONTTYPE_DUELO)
    Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos2vs2> " & UserList(UserIndex1).Name & " y " & UserList(UserIndex2).Name & " esperan rival en la sala de duelos..." & FONTTYPE_DUELO)
    DUELO_PAREJA2.usuario1 = UserIndex1
    DUELO_PAREJA2.usuario2 = UserIndex2
ElseIf MapInfo(DUELO_MAPADUELO).NumUsers = 0 Then 'CHOTS | Está vacía la sala
    Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos2vs2> " & UserList(UserIndex1).Name & " y " & UserList(UserIndex2).Name & " esperan rival en la sala de duelos..." & FONTTYPE_DUELO)
    DUELO_PAREJA1.usuario1 = UserIndex1
    DUELO_PAREJA1.usuario2 = UserIndex2
Else
    UserList(UserIndex1).flags.enDuelo = False
    UserList(UserIndex2).flags.enDuelo = False
End If

Call ClosestLegalPos(Pos, nPos)
Call WarpUserChar(UserIndex1, nPos.Map, nPos.X, nPos.Y, True)
Call ClosestLegalPos(Pos, nPos)
Call WarpUserChar(UserIndex2, nPos.Map, nPos.X, nPos.Y, True)


End Sub

Public Sub ganaDueloPareja(ByVal UserIndex As Integer)
Dim ganador1 As Integer
Dim ganador2 As Integer

If DUELO_PAREJA1.usuario1 = UserIndex Or DUELO_PAREJA1.usuario2 = UserIndex Then
    ganador1 = DUELO_PAREJA1.usuario1
    ganador2 = DUELO_PAREJA1.usuario2
Else
    ganador1 = DUELO_PAREJA2.usuario1
    ganador2 = DUELO_PAREJA2.usuario2
End If

Call SendData(SendTarget.ToMap, 0, DUELO_MAPADUELO, ServerPackages.dialogo & "Duelos2vs2> " & UserList(ganador1).Name & " y " & UserList(ganador2).Name & " han ganado el duelo!" & FONTTYPE_DUELO)
Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos2vs2> " & UserList(ganador1).Name & " y " & UserList(ganador2).Name & " han ganado el duelo!" & FONTTYPE_DUELO)


UserList(ganador1).Stats.MinHP = UserList(ganador1).Stats.MaxHP
UserList(ganador1).Stats.MinMAN = UserList(ganador1).Stats.MaxMAN

UserList(ganador2).Stats.MinHP = UserList(ganador2).Stats.MaxHP
UserList(ganador2).Stats.MinMAN = UserList(ganador2).Stats.MaxMAN

If UserList(ganador1).flags.Paralizado = 1 Then
    UserList(ganador1).flags.Paralizado = 0
    Call SendData(SendTarget.ToIndex, ganador1, 0, "DOK")
End If

If UserList(ganador2).flags.Paralizado = 1 Then
    UserList(ganador2).flags.Paralizado = 0
    Call SendData(SendTarget.ToIndex, ganador2, 0, "DOK")
End If


End Sub

Public Sub pierdeDueloPareja(ByVal UserIndex As Integer)
Dim nPos As WorldPos
Dim Pos As WorldPos

Pos.Map = 1
Pos.X = 58
Pos.Y = 45

Call ClosestLegalPos(Pos, nPos)

UserList(UserIndex).flags.enDuelo = False
Call UserDie(UserIndex)
Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)

End Sub

Public Sub salirDueloPareja(ByVal UserIndex As Integer)
Dim jugador1 As Integer
Dim jugador2 As Integer

If DUELO_PAREJA1.usuario1 = UserIndex Or DUELO_PAREJA1.usuario2 = UserIndex Then
    jugador1 = DUELO_PAREJA1.usuario1
    jugador2 = DUELO_PAREJA1.usuario2
Else
    jugador1 = DUELO_PAREJA2.usuario1
    jugador2 = DUELO_PAREJA2.usuario2
End If

Dim nPos As WorldPos
Dim Pos As WorldPos

Pos.Map = 1
Pos.X = 58
Pos.Y = 45

UserList(jugador1).flags.enDuelo = False
UserList(jugador2).flags.enDuelo = False

Call ClosestLegalPos(Pos, nPos)
Call WarpUserChar(jugador1, Pos.Map, Pos.X, Pos.Y, True)

Call ClosestLegalPos(Pos, nPos)
Call WarpUserChar(jugador2, Pos.Map, Pos.X, Pos.Y, True)

Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos2vs2> " & UserList(jugador1).Name & " y " & UserList(jugador2).Name & " ha abandonado la sala de duelos..." & FONTTYPE_DUELO)

End Sub

