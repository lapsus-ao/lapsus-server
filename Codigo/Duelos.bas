Attribute VB_Name = "Duelos"
'MÓDULO DE DUELOS 1VS1
'CREADO POR JUAN ANDRES DALMASSO (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'EL 06/01/12
'PARA LAPSUS 3.1
'REPROGRAMADO POR CHOTS
'EL 06/09/2017
'PARA LAPSUS2017

Public Const DUELO_MAPADUELO As Integer = 61
Public Const DUELO_MINX As Integer = 20
Public Const DUELO_MAXX As Integer = 40
Public Const DUELO_MINY As Integer = 20
Public Const DUELO_MAXY As Integer = 40

Public DUELO_USUARIO1 As Integer
Public DUELO_USUARIO2 As Integer

Public Function puedeDuelo(ByVal UserIndex As Integer, ByRef error As String) As Boolean

puedeDuelo = True

If UserList(UserIndex).flags.Muerto = 1 Then
    error = "No puedes ingresar a un duelo estando muerto"
    puedeDuelo = False
    Exit Function
End If

If UserList(UserIndex).flags.TargetNPC = 0 Then
    error = "Primero debes clickear en el NPC"
    puedeDuelo = False
    Exit Function
End If

If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Duelero Then
    error = "El NPC no organiza duelos!"
    puedeDuelo = False
    Exit Function
End If

If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
    error = "Estás demasiado lejos"
    puedeDuelo = False
    Exit Function
End If

If EsNewbie(UserIndex) Then
    error = "Los newbies no tienen permitido ingresar a los duelos!"
    puedeDuelo = False
    Exit Function
End If

If DUELO_USUARIO1 > 0 And DUELO_USUARIO2 > 0 Then
    error = "La Sala de duelos está ocupada"
    puedeDuelo = False
    Exit Function
End If

End Function

Public Sub ingresarDuelo(ByVal UserIndex As Integer)
Dim nPos As WorldPos
Dim Pos As WorldPos

Pos.Map = DUELO_MAPADUELO
Pos.X = RandomNumber(DUELO_MINX, DUELO_MAXX)
Pos.Y = RandomNumber(DUELO_MINY, DUELO_MAXY)

Call ClosestLegalPos(Pos, nPos)

UserList(UserIndex).flags.enDuelo = True

If DUELO_USUARIO1 > 0 And DUELO_USUARIO2 > 0 Then
    UserList(UserIndex).flags.enDuelo = False
    Exit Sub
ElseIf DUELO_USUARIO1 > 0 Then 'CHOTS | Ya hay uno esperando para duelear
    DUELO_USUARIO2 = UserIndex
    Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " ha aceptado el duelo!" & FONTTYPE_DUELO)
    Call SendData(SendTarget.ToMap, 0, DUELO_MAPADUELO, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " ha aceptado el duelo!" & FONTTYPE_DUELO)
Else 'CHOTS | Está vacía la sala
    DUELO_USUARIO1 = UserIndex
    DUELO_USUARIO2 = 0
    Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " espera rival en la sala de duelos..." & FONTTYPE_DUELO)
End If

Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)

End Sub

Public Sub ganaDuelo(ByVal UserIndex As Integer)

UserList(UserIndex).Stats.DuelosConsec = UserList(UserIndex).Stats.DuelosConsec + 1
Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " ha ganado el duelo!" & FONTTYPE_DUELO)

DUELO_USUARIO1 = UserIndex
DUELO_USUARIO2 = 0

If UserList(UserIndex).Stats.DuelosConsec >= 2 Then
    Call SendData(SendTarget.ToMap, 0, DUELO_MAPADUELO, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " ya lleva " & UserList(UserIndex).Stats.DuelosConsec & " duelos ganados consecutivamente!!!" & FONTTYPE_DUELO)
    Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " ya lleva " & UserList(UserIndex).Stats.DuelosConsec & " duelos ganados consecutivamente!!!" & FONTTYPE_DUELO)

    Call ActualizarRanking(UserIndex, 5)
End If

'CHOTS | Actualizamos sus stats
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
Call EnviarMn(UserIndex)
Call EnviarHP(UserIndex)

If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).flags.Paralizado = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
End If

End Sub

Public Sub pierdeDuelo(ByVal UserIndex As Integer)
Dim nPos As WorldPos
Dim Pos As WorldPos

Pos.Map = 1
Pos.X = 58
Pos.Y = 45

Call ClosestLegalPos(Pos, nPos)

UserList(UserIndex).flags.enDuelo = False
UserList(UserIndex).Stats.DuelosConsec = 0
Call UserDie(UserIndex)
Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)

End Sub

Public Sub salirDuelo(ByVal UserIndex As Integer)
On Local Error Resume Next
Dim nPos As WorldPos
Dim Pos As WorldPos

Pos.Map = 1
Pos.X = 58
Pos.Y = 45

Call ClosestLegalPos(Pos, nPos)

UserList(UserIndex).flags.enDuelo = False
Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)

DUELO_USUARIO1 = 0
DUELO_USUARIO2 = 0

Call SendData(SendTarget.ToMap, 0, 1, ServerPackages.dialogo & "Duelos1vs1> " & UserList(UserIndex).Name & " ha abandonado la sala de duelos..." & FONTTYPE_DUELO)

End Sub

Public Function puedeSalirDuelo(ByVal UserIndex As Integer, ByRef error As String) As Boolean

puedeSalirDuelo = True

If UserList(UserIndex).flags.enDuelo = False Then
    error = "No estás en un duelo."
    puedeSalirDuelo = False
    Exit Function
End If

If DUELO_USUARIO1 > 0 And DUELO_USUARIO2 > 0 Then
    error = "No puedes salir de un duelo si tu contrincante está vivo."
    puedeSalirDuelo = False
    Exit Function
End If

End Function

Public Sub LogDuelo(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\duelos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub