Attribute VB_Name = "Guerras"
'Lapsus2017
'Copyright (C) 2017 Dalmasso, Juan Andres
'
'Modulo de Guerras de clanes
'Programado por CHOTS (Juan Andres Dalmasso)
'Desde Wellington, New Zealand
'10/10/2017

Private Const NPC_ORO_1 As Integer = 593 'Goblin
Private Const NPC_ORO_2 As Integer = 594 'Duende
Private Const NPC_ITEMS_1 As Integer = 595
Private Const NPC_ITEMS_2 As Integer = 596
Private Const NPC_ITEMS_3 As Integer = 597
Private Const NPC_TORRE As Integer = 598
Private Const NPC_CASA As Integer = 599

Private Const MAX_USERS_TEAM = 5

Private Const MAPA_GUERRAS = 89

Private Const COLOR_TEAM_A = "+" & vbRed
Private Const COLOR_TEAM_B = "+" & vbBlue
Private Const COLOR_NEUTRAL = "+" & vbWhite

Public Const GUERRA_ESTADO_NULA = 0
Public Const GUERRA_ESTADO_LOBBY = 1
Public Const GUERRA_ESTADO_INICIADA = 2

Private Const GUERRA_CANT_SALAS = 2

Private Const GUERRA_MINUTOS_TIMEOUT = 5
Private Const GUERRA_MINUTOS_DURACION = 30
Private Const GUERRA_SEGUNDOS_COMIENZO = 10

Private Const GUERRA_TEAM_A = 1
Private Const GUERRA_TEAM_B = 2

Private Const PUNTOS_NPC_ORO_1 = 1
Private Const PUNTOS_NPC_ORO_2 = 2
Private Const PUNTOS_NPC_ITEMS_1 = 3
Private Const PUNTOS_NPC_ITEMS_2 = 4
Private Const PUNTOS_NPC_ITEMS_3 = 5
Private Const PUNTOS_NPC_TORRE = 100
Private Const PUNTOS_FRAG = 15

Private Const ITEMS_ROPA_ALTO_A = 31
Private Const ITEMS_ROPA_BAJO_A = 942
Private Const ITEMS_ROPA_ALTO_B = 32
Private Const ITEMS_ROPA_BAJO_B = 943
Private Const ITEMS_DAGA = 15
Private Const ITEMS_POCION_ROJA = 38
Private Const ITEMS_POCION_AZUL = 37
Private Const ITEMS_MANZANA = 1
Private Const ITEMS_AGUA = 43

Public Type cZonaGuerra
    minX As Byte
    maxX As Byte
    minY As Byte
    maxY As Byte
End Type

Public Type cPosGuerra
    X As Byte
    Y As Byte
End Type

Public Type cSalaGuerra
    Numero As Byte
    mapa As Byte
    nombre As String
    estado As Byte

    mapaLobby As Integer
    zonaLobby As cZonaGuerra

    zonaRespawnTeamA As cZonaGuerra
    zonaNpcsOro1TeamA As cZonaGuerra
    zonaNpcsOro2TeamA As cZonaGuerra
    zonaSeguraTeamA As cZonaGuerra
    posTorre1TeamA As cPosGuerra
    posTorre2TeamA As cPosGuerra
    posTorre3TeamA As cPosGuerra
    posCasaTeamA As cPosGuerra

    zonaRespawnTeamB As cZonaGuerra
    zonaNpcsOro1TeamB As cZonaGuerra
    zonaNpcsOro2TeamB As cZonaGuerra
    zonaSeguraTeamB As cZonaGuerra
    posTorre1TeamB As cPosGuerra
    posTorre2TeamB As cPosGuerra
    posTorre3TeamB As cPosGuerra
    posCasaTeamB As cPosGuerra

    zonaNpcsItems1 As cZonaGuerra
    zonaNpcsItems2 As cZonaGuerra
    zonaNpcsItems3 As cZonaGuerra

    cantidadNpcsOro1 As Byte
    cantidadNpcsOro2 As Byte
    cantidadNpcsItems1 As Byte
    cantidadNpcsItems2 As Byte
    cantidadNpcsItems3 As Byte
End Type

Public Type cGuerra
    Sala As cSalaGuerra
    cantUsers As Byte

    guildIndexA As Integer
    guildIndexB As Integer
    cantUsersGuildA As Byte
    cantUsersGuildB As Byte

    guildA() As Integer
    guildB() As Integer
    puntosPremio As Long
    timeout As Byte 'CHOTS | Minutos guerra
    contador As Byte 'CHOTS | Segundos guerra

    puntosGuildA As Integer
    puntosGuildB As Integer
End Type

Public SalasGuerra(1 To GUERRA_CANT_SALAS) As cSalaGuerra
Public Guerras(1 To GUERRA_CANT_SALAS) As cGuerra

Public Sub inicializarSalasGuerra()
    With SalasGuerra(1)
        .Numero = 1
        .mapa = 91
        .nombre = "Gulfas Morgolock"
        .estado = GUERRA_ESTADO_NULA
        .mapaLobby = 89
        .zonaLobby.minX = 38
        .zonaLobby.maxX = 50
        .zonaLobby.minY = 29
        .zonaLobby.maxY = 34
        
        'CHOTS | Team A
        .zonaRespawnTeamA.minX = 13
        .zonaRespawnTeamA.maxX = 23
        .zonaRespawnTeamA.minY = 38
        .zonaRespawnTeamA.maxY = 43
        
        .posCasaTeamA.X = 19
        .posCasaTeamA.Y = 50
        
        .posTorre1TeamA.X = 34
        .posTorre1TeamA.Y = 30
        
        .posTorre2TeamA.X = 34
        .posTorre2TeamA.Y = 50
        
        .posTorre3TeamA.X = 34
        .posTorre3TeamA.Y = 70
        
        .zonaNpcsOro1TeamA.minX = 13
        .zonaNpcsOro1TeamA.maxX = 23
        .zonaNpcsOro1TeamA.minY = 44
        .zonaNpcsOro1TeamA.maxY = 89
        
        .zonaNpcsOro2TeamA.minX = 25
        .zonaNpcsOro2TeamA.maxX = 33
        .zonaNpcsOro2TeamA.minY = 12
        .zonaNpcsOro2TeamA.maxY = 89

        .zonaSeguraTeamA.minX = 13
        .zonaSeguraTeamA.maxX = 23
        .zonaSeguraTeamA.minY = 12
        .zonaSeguraTeamA.maxY = 21
        
        'CHOTS | Team B
        .zonaRespawnTeamB.minX = 78
        .zonaRespawnTeamB.maxX = 88
        .zonaRespawnTeamB.minY = 38
        .zonaRespawnTeamB.maxY = 43

        .posCasaTeamB.X = 83
        .posCasaTeamB.Y = 50
        
        .posTorre1TeamB.X = 67
        .posTorre1TeamB.Y = 30
        
        .posTorre2TeamB.X = 67
        .posTorre2TeamB.Y = 50
        
        .posTorre3TeamB.X = 67
        .posTorre3TeamB.Y = 70
        
        .zonaNpcsOro1TeamB.minX = 78
        .zonaNpcsOro1TeamB.maxX = 88
        .zonaNpcsOro1TeamB.minY = 44
        .zonaNpcsOro1TeamB.maxY = 89
        
        .zonaNpcsOro2TeamB.minX = 68
        .zonaNpcsOro2TeamB.maxX = 76
        .zonaNpcsOro2TeamB.minY = 12
        .zonaNpcsOro2TeamB.maxY = 89

        .zonaSeguraTeamB.minX = 78
        .zonaSeguraTeamB.maxX = 88
        .zonaSeguraTeamB.minY = 12
        .zonaSeguraTeamB.maxY = 21
        
        'CHOTS | Npcs Items
        .zonaNpcsItems1.minX = 35
        .zonaNpcsItems1.maxX = 66
        .zonaNpcsItems1.minY = 12
        .zonaNpcsItems1.maxY = 89
        
        .zonaNpcsItems2.minX = 35
        .zonaNpcsItems2.maxX = 66
        .zonaNpcsItems2.minY = 12
        .zonaNpcsItems2.maxY = 89
        
        .zonaNpcsItems3.minX = 35
        .zonaNpcsItems3.maxX = 66
        .zonaNpcsItems3.minY = 12
        .zonaNpcsItems3.maxY = 89
        
        'CHOTS | Cantidad NPCs
        .cantidadNpcsItems1 = 3
        .cantidadNpcsItems2 = 2
        .cantidadNpcsItems3 = 1
        .cantidadNpcsOro1 = 5
        .cantidadNpcsOro2 = 3
        
    End With
    
    With SalasGuerra(2)
        .Numero = 2
        .mapa = 90
        .nombre = "Cucsifae"
        .estado = GUERRA_ESTADO_NULA
        .mapaLobby = 89
        .zonaLobby.minX = 43
        .zonaLobby.maxX = 53
        .zonaLobby.minY = 63
        .zonaLobby.maxY = 66
        
        'CHOTS | Team A
        .zonaRespawnTeamA.minX = 13
        .zonaRespawnTeamA.maxX = 23
        .zonaRespawnTeamA.minY = 23
        .zonaRespawnTeamA.maxY = 30
        
        .posCasaTeamA.X = 15
        .posCasaTeamA.Y = 87
        
        .posTorre1TeamA.X = 20
        .posTorre1TeamA.Y = 78
        
        .posTorre2TeamA.X = 23
        .posTorre2TeamA.Y = 72
        
        .posTorre3TeamA.X = 22
        .posTorre3TeamA.Y = 58
        
        .zonaNpcsOro1TeamA.minX = 13
        .zonaNpcsOro1TeamA.maxX = 41
        .zonaNpcsOro1TeamA.minY = 37
        .zonaNpcsOro1TeamA.maxY = 57
        
        .zonaNpcsOro2TeamA.minX = 26
        .zonaNpcsOro2TeamA.maxX = 39
        .zonaNpcsOro2TeamA.minY = 15
        .zonaNpcsOro2TeamA.maxY = 36

        .zonaSeguraTeamA.minX = 12
        .zonaSeguraTeamA.maxX = 23
        .zonaSeguraTeamA.minY = 15
        .zonaSeguraTeamA.maxY = 21
        
        'CHOTS | Team B
        .zonaRespawnTeamB.minX = 67
        .zonaRespawnTeamB.maxX = 78
        .zonaRespawnTeamB.minY = 23
        .zonaRespawnTeamB.maxY = 30

        .posCasaTeamB.X = 75
        .posCasaTeamB.Y = 87
        
        .posTorre1TeamB.X = 71
        .posTorre1TeamB.Y = 77
        
        .posTorre2TeamB.X = 68
        .posTorre2TeamB.Y = 70
        
        .posTorre3TeamB.X = 69
        .posTorre3TeamB.Y = 58
        
        .zonaNpcsOro1TeamB.minX = 44
        .zonaNpcsOro1TeamB.maxX = 78
        .zonaNpcsOro1TeamB.minY = 37
        .zonaNpcsOro1TeamB.maxY = 57
        
        .zonaNpcsOro2TeamB.minX = 51
        .zonaNpcsOro2TeamB.maxX = 64
        .zonaNpcsOro2TeamB.minY = 15
        .zonaNpcsOro2TeamB.maxY = 36

        .zonaSeguraTeamB.minX = 67
        .zonaSeguraTeamB.maxX = 78
        .zonaSeguraTeamB.minY = 15
        .zonaSeguraTeamB.maxY = 22
        
        'CHOTS | Npcs Items
        .zonaNpcsItems1.minX = 30
        .zonaNpcsItems1.maxX = 61
        .zonaNpcsItems1.minY = 59
        .zonaNpcsItems1.maxY = 88
        
        .zonaNpcsItems2.minX = 30
        .zonaNpcsItems2.maxX = 61
        .zonaNpcsItems2.minY = 59
        .zonaNpcsItems2.maxY = 88
        
        .zonaNpcsItems3.minX = 30
        .zonaNpcsItems3.maxX = 61
        .zonaNpcsItems3.minY = 59
        .zonaNpcsItems3.maxY = 88
        
        'CHOTS | Cantidad NPCs
        .cantidadNpcsItems1 = 3
        .cantidadNpcsItems2 = 2
        .cantidadNpcsItems3 = 1
        .cantidadNpcsOro1 = 5
        .cantidadNpcsOro2 = 3
        
    End With
End Sub

Public Sub inicializarGuerras()
    Dim i As Byte
    For i = 1 To GUERRA_CANT_SALAS
        With Guerras(i)
            .Sala = SalasGuerra(i)
            .cantUsers = 0
            .guildIndexA = 0
            .guildIndexB = 0
            .cantUsersGuildA = 0
            .cantUsersGuildB = 0
            .puntosPremio = 0
            .timeout = 0
            .contador = 0
            .puntosGuildA = 0
            .puntosGuildB = 0
        End With
    Next i
End Sub

Public Function PuedeIntentarCrearGuerra(ByVal UserIndex As Integer, ByVal numeroSala As Byte, ByRef error As String) As Boolean
    'CHOTS | Puede ver el formulario para iniciar una guerra?
    PuedeIntentarCrearGuerra = True

    If numeroSala > GUERRA_CANT_SALAS Or numeroSala = 0 Then
        error = "La sala no existe."
        PuedeIntentarCrearGuerra = False
        Exit Function
    End If

    If SalasGuerra(numeroSala).estado <> GUERRA_ESTADO_NULA Then
        error = "La sala elegida esta ocupada."
        PuedeIntentarCrearGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).GuildIndex = 0 Then
        error = "No perteneces a un clan."
        PuedeIntentarCrearGuerra = False
        Exit Function
    End If

    If UCase$(Guilds(UserList(UserIndex).GuildIndex).GetLeader) <> UCase$(UserList(UserIndex).Name) Then
        error = "Solo el lider de un clan puede iniciar una guerra."
        PuedeIntentarCrearGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        error = "No puedes iniciar una guerra estando muerto."
        PuedeIntentarCrearGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).guerra.enGuerra = True Then
        error = "Ya te encuentras en una Guerra."
        PuedeIntentarCrearGuerra = False
        Exit Function
    End If
End Function

Public Function PuedeCrearGuerra(ByVal UserIndex As Integer, ByVal numeroSala As Byte, ByVal cantUsers As Byte, ByVal puntosPremio As Long, ByVal clanEnemigo As String, ByRef error As String) As Boolean
    Dim enemigoGuildIndex As Integer
    enemigoGuildIndex = 0

    'CHOTS | Puede enviar para iniciar una guerra?
    PuedeCrearGuerra = True

    If numeroSala > GUERRA_CANT_SALAS Or numeroSala = 0 Then
        error = "La sala no existe."
        PuedeCrearGuerra = False
        Exit Function
    End If

    If SalasGuerra(numeroSala).estado <> GUERRA_ESTADO_NULA Then
        error = "La sala elegida esta ocupada."
        PuedeCrearGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).guerra.enGuerra = True Then
        error = "Ya te encuentras en una Guerra."
        PuedeCrearGuerra = False
        Exit Function
    End If

    If modGuilds.m_CantidadDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex) < cantUsers Then
        error = "No hay suficientes usuarios conectados de tu clan."
        PuedeCrearGuerra = False
        Exit Function
    End If

    If Guilds(UserList(UserIndex).GuildIndex).GetGuildPoints < puntosPremio Then
        error = "Tu clan no tiene suficientes puntos."
        PuedeCrearGuerra = False
        Exit Function
    End If

    If cantUsers <= 0 Or cantUsers >= MAX_USERS_TEAM Then
        error = "La cantidad de miembros es invalida."
        PuedeCrearGuerra = False
        Exit Function
    End If

    If puntosPremio <= 0 Then
        error = "Los puntos son invalidos."
        PuedeCrearGuerra = False
        Exit Function
    End If

    'CHOTS | Le declara la guerra a alguien
    If clanEnemigo <> "" Then
        enemigoGuildIndex = GuildIndex(clanEnemigo)
        If enemigoGuildIndex = 0 Then
            error = "El clan enemigo no existe."
            PuedeCrearGuerra = False
            Exit Function
        End If

        If enemigoGuildIndex = UserList(UserIndex).GuildIndex Then
            error = "No puedes desafiarte a vos mismo."
            PuedeCrearGuerra = False
            Exit Function
        End If
    End If

End Function

Public Sub CrearGuerra(ByVal UserIndex As Integer, ByVal numeroSala As Byte, ByVal cantUsers As Byte, ByVal puntosPremio As Long, ByVal clanEnemigo As String)
    Dim enemigoGuildIndex As Integer
    enemigoGuildIndex = 0

    'CHOTS | Le declara la guerra a alguien
    If clanEnemigo <> "" Then
        enemigoGuildIndex = GuildIndex(clanEnemigo)
    End If
    
    With SalasGuerra(numeroSala)
        .estado = GUERRA_ESTADO_LOBBY
    End With

    'CHOTS | Creamos la guerra
    With Guerras(numeroSala)
        .Sala = SalasGuerra(numeroSala)
        .guildIndexA = UserList(UserIndex).GuildIndex
        .guildIndexB = enemigoGuildIndex
        .cantUsers = cantUsers
        .cantUsersGuildA = 0
        .cantUsersGuildB = 0
        .puntosPremio = puntosPremio
        .puntosGuildA = 0
        .puntosGuildB = 0

        ReDim .guildA(1 To cantUsers) As Integer
        ReDim .guildB(1 To cantUsers) As Integer

        .timeout = GUERRA_MINUTOS_TIMEOUT
    End With

    If enemigoGuildIndex > 0 Then
        'CHOTS | Avisamos a los miembros del clan que el lider empezo una guerra
        Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "Guerras> " & UserList(UserIndex).Name & ", líder de tu clan, ha iniciado una guerra contra el clan <" & clanEnemigo & ">. Dirigirse a la sala " & UCase$(SalasGuerra(numeroSala).nombre) & " y tipear /IRGUERRA para participar." & FONTTYPE_GUERRA)
        Call SendData(SendTarget.ToGuildMembers, enemigoGuildIndex, 0, ServerPackages.dialogo & "Guerras> " & UserList(UserIndex).Name & ", líder del clan <" & Guilds(UserList(UserIndex).GuildIndex).GuildName & "> ha iniciado una guerra contra tu clan por " & puntosPremio & " puntos. Dirigirse a la sala " & UCase$(SalasGuerra(numeroSala).nombre) & " y tipear /IRGUERRA para participar." & FONTTYPE_GUERRA)
    Else
        'CHOTS | Avisamos a todos que el clan busca rival para una guerra en la sala X
        Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "Guerras> " & UserList(UserIndex).Name & ", líder de tu clan, ha iniciado una guerra. Dirigirse a la sala " & UCase$(SalasGuerra(numeroSala).nombre) & " y tipear /IRGUERRA para participar." & FONTTYPE_GUERRA)
        Call SendData(SendTarget.ToAllButIndex, UserIndex, 0, ServerPackages.dialogo & "Guerras> " & UCase$(UserList(UserIndex).Name) & ", líder del clan <" & Guilds(UserList(UserIndex).GuildIndex).GuildName & "> está buscando un clan rival para una guerra en la sala " & UCase$(SalasGuerra(numeroSala).nombre) & " por " & puntosPremio & " puntos. Dirigete a la sala y tipea /IRGUERRA para participar." & FONTTYPE_GUERRA)
    End If

    'CHOTS | Metemos al userindex en la guerra
    Call IrGuerra(UserIndex, numeroSala)
End Sub

Public Function PuedeIrGuerra(ByVal UserIndex As Integer, ByVal numeroSala As Byte, ByRef error As String) As Boolean
    'CHOTS | Puede entrar a una guerra en estado Lobby
    PuedeIrGuerra = True

    If numeroSala > GUERRA_CANT_SALAS Then
        error = "La sala no existe."
        PuedeIrGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).GuildIndex = 0 Then
        error = "No perteneces a un clan."
        PuedeIrGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        error = "No puedes ir a la guerra estando muerto."
        PuedeIrGuerra = False
        Exit Function
    End If

    If SalasGuerra(numeroSala).estado = GUERRA_ESTADO_NULA Then
        error = "La sala elegida esta vacia."
        PuedeIrGuerra = False
        Exit Function
    End If

    If SalasGuerra(numeroSala).estado = GUERRA_ESTADO_INICIADA Then
        error = "La guerra ya ha comenzado en esta sala."
        PuedeIrGuerra = False
        Exit Function
    End If

    If UserList(UserIndex).guerra.enGuerra = True Then
        error = "Ya te encuentras en una Guerra."
        PuedeIrGuerra = False
        Exit Function
    End If

    If Guerras(numeroSala).guildIndexA = UserList(UserIndex).GuildIndex Then
        'CHOTS | Es miembro del clan que organizo la guerra
        If Guerras(numeroSala).cantUsersGuildA = Guerras(numeroSala).cantUsers Then
            error = "Tu clan ya completo el cupo."
            PuedeIrGuerra = False
            Exit Function
        End If
    Else
        'CHOTS | No es miembro del clan organizador
        If Guerras(numeroSala).guildIndexB <> 0 Then
            'CHOTS | Ya hay dos clanes esperando la guerra
            If Guerras(numeroSala).guildIndexB <> UserList(UserIndex).GuildIndex Then
                error = "Tu clan no pertenece a esta guerra."
                PuedeIrGuerra = False
                Exit Function
            End If

            If Guerras(numeroSala).guildB(1) = 0 Then
                'CHOTS | El lider todavia no acepto el desafio, solo el puede ir a la guerra
                If UCase$(Guilds(UserList(UserIndex).GuildIndex).GetLeader) <> UCase$(UserList(UserIndex).Name) Then
                    error = "Primero el desafio debe ser aceptado por el lider de tu clan."
                    PuedeIrGuerra = False
                    Exit Function
                End If
            Else
                'CHOTS | El lider ya acepto el desafio, cualquier miembro del clan puede entrar
                If Guerras(numeroSala).cantUsersGuildB = Guerras(numeroSala).cantUsers Then
                    error = "Tu clan ya completo el cupo."
                    PuedeIrGuerra = False
                    Exit Function
                End If
            End If
        Else
            'CHOTS | Tiene que ser lider para aceptar la guerra
            If UCase$(Guilds(UserList(UserIndex).GuildIndex).GetLeader) <> UCase$(UserList(UserIndex).Name) Then
                error = "Primero el desafio debe ser aceptado por el lider de tu clan."
                PuedeIrGuerra = False
                Exit Function
            End If
        End If
    End If
    
End Function

Public Sub IrGuerra(ByVal UserIndex As Integer, ByVal numeroSala As Byte)
    Dim i As Byte
    With Guerras(numeroSala)
        If .guildIndexA = UserList(UserIndex).GuildIndex Then
            'CHOTS | Es miembro del clan que inicio el desafio
            For i = 1 To .cantUsers
                If .guildA(i) = 0 Then
                    .guildA(i) = UserIndex
                    .cantUsersGuildA = .cantUsersGuildA + 1
                    Exit For
                End If
            Next i
            UserList(UserIndex).guerra.team = GUERRA_TEAM_A
        Else
            'CHOTS | Es miembro del otro clan
            For i = 1 To .cantUsers
                If .guildB(i) = 0 Then
                    .guildB(i) = UserIndex
                    .cantUsersGuildB = .cantUsersGuildB + 1
                    Exit For
                End If
            Next i
            .guildIndexB = UserList(UserIndex).GuildIndex
            UserList(UserIndex).guerra.team = GUERRA_TEAM_B
        End If
    End With

    Call TelepToLobby(UserIndex, numeroSala)

    Call CheckEquiposCompletos(numeroSala)
End Sub

Public Sub TelepToLobby(ByVal UserIndex As Integer, ByVal numeroSala As Byte)
    Dim respawnPos As WorldPos
    Dim nPos As WorldPos
    UserList(UserIndex).guerra.enGuerra = True
    UserList(UserIndex).guerra.status = GUERRA_ESTADO_LOBBY
    UserList(UserIndex).guerra.Sala = numeroSala

    With SalasGuerra(numeroSala)
        respawnPos.Map = .mapaLobby
        respawnPos.X = RandomNumber(.zonaLobby.minX, .zonaLobby.maxX)
        respawnPos.Y = RandomNumber(.zonaLobby.minY, .zonaLobby.maxY)
        Call ClosestLegalPos(respawnPos, nPos)
        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
    End With
End Sub

Public Sub CheckEquiposCompletos(ByVal numeroSala As Byte)
    With Guerras(numeroSala)
        If .cantUsersGuildA = .cantUsers And .cantUsersGuildB = .cantUsers Then
            'CHOTS | Equipos completos, que empiece la diversion
            .Sala.estado = GUERRA_ESTADO_INICIADA
            .timeout = GUERRA_MINUTOS_DURACION
            .contador = GUERRA_SEGUNDOS_COMIENZO

            'CHOTS | Avisamos a todos que va a empezar una guerra en la sala X
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Guerras> En " & GUERRA_SEGUNDOS_COMIENZO & " segundos comenzará una guerra entre <" & Guilds(.guildIndexA).GuildName & "> y <" & Guilds(.guildIndexB).GuildName & "> en la sala " & UCase$(.Sala.nombre) & "." & FONTTYPE_GUERRA)
        End If
    End With
End Sub

Public Sub TimerMinutosGuerra()
    On Error GoTo chotserror

    'CHOTS | Timer cada un minuto
    Dim i As Byte
    For i = 1 To GUERRA_CANT_SALAS
        With Guerras(i)
            If .timeout > 0 Then
                .timeout = .timeout - 1

                If .timeout = 0 Then
                    If .Sala.estado = GUERRA_ESTADO_INICIADA Then
                        'CHOTS | Se termina la guerra, gana el que tiene mas puntos
                        Call TerminarGuerra(i)
                    ElseIf .Sala.estado = GUERRA_ESTADO_LOBBY Then
                        'CHOTS | Se cancela la guerra por falta de contrincantes
                        Call CancelarGuerra(i)
                    End If
                End If
            End If

            'CHOTS | Avisamos el resultado parcial cada 5 minutos
            If .Sala.estado = GUERRA_ESTADO_INICIADA Then
                If (.timeout Mod 5 = 0) And .timeout > 0 Then
                    Call SendData(SendTarget.ToMap, 0, .Sala.mapa, "!G" & .timeout & " minutos restantes. Resultado Parcial: " & Guilds(.guildIndexA).GuildName & " " & .puntosGuildA & " - " & Guilds(.guildIndexB).GuildName & ": " & .puntosGuildB & COLOR_NEUTRAL)
                End If
            End If

        End With
    Next i

Exit Sub
chotserror:
    Call LogError("Error en TimerMinutosGuerra " & Err.number & " " & Err.Description)
End Sub

Public Sub TimerSegundosGuerra()
    On Error GoTo chotserror

    'CHOTS | Timer cada un segundo
    On Local Error Resume Next
    Dim j As Byte
    Dim i As Byte
    For i = 1 To GUERRA_CANT_SALAS
        With Guerras(i)
            If .contador > 0 Then
                .contador = .contador - 1

                If .contador = 0 Then
                    Call ComenzarGuerra(i)
                End If
            End If

            'CHOTS | Chequeamos que un user no este en la zona segura del otro
            If .Sala.estado = GUERRA_ESTADO_INICIADA Then
                 For j = 1 To .cantUsers
                    If .guildA(j) > 0 Then
                        If UserList(.guildA(j)).Pos.Map = .Sala.mapa And UserList(.guildA(j)).Pos.X >= .Sala.zonaSeguraTeamB.minX And UserList(.guildA(j)).Pos.X <= .Sala.zonaSeguraTeamB.maxX And UserList(.guildA(j)).Pos.Y >= .Sala.zonaSeguraTeamB.minY And UserList(.guildA(j)).Pos.Y <= .Sala.zonaSeguraTeamB.maxY And UserList(.guildA(j)).flags.Muerto = 0 Then
                            Call UserDie(.guildA(j))
                            Call SendData(SendTarget.ToMap, 0, .Sala.mapa, "!G" & UserList(.guildA(j)).Name & " se suicidó." & COLOR_TEAM_B)
                            Call DarPuntosGuera(i, GUERRA_TEAM_A, -1 * PUNTOS_FRAG)
                        End If
                    End If
                    
                    If .guildB(j) > 0 Then
                        If UserList(.guildB(j)).Pos.Map = .Sala.mapa And UserList(.guildB(j)).Pos.X >= .Sala.zonaSeguraTeamA.minX And UserList(.guildB(j)).Pos.X <= .Sala.zonaSeguraTeamA.maxX And UserList(.guildB(j)).Pos.Y >= .Sala.zonaSeguraTeamA.minY And UserList(.guildB(j)).Pos.Y <= .Sala.zonaSeguraTeamA.maxY And UserList(.guildB(j)).flags.Muerto = 0 Then
                            Call UserDie(.guildB(j))
                            Call SendData(SendTarget.ToMap, 0, .Sala.mapa, "!G" & UserList(.guildB(j)).Name & " se suicidó." & COLOR_TEAM_A)
                            Call DarPuntosGuera(i, GUERRA_TEAM_B, -1 * PUNTOS_FRAG)
                        End If
                    End If
                Next j
            End If
        End With
    Next i
Exit Sub
chotserror:
    Call LogError("Error en TimerSegundosGuerra " & Err.number & " " & Err.Description)
End Sub

Public Sub CancelarGuerra(ByVal numeroSala As Byte)
    Dim i As Byte
    
    With Guerras(numeroSala)
        For i = 1 To .cantUsers
            Call RetirarUserGuerra(.guildA(i), False)
            Call RetirarUserGuerra(.guildB(i), False)
        Next i
    End With

    With SalasGuerra(numeroSala)
        .estado = GUERRA_ESTADO_NULA
        
        'CHOTS | Avisamos que la guerra se cancelo
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Guerras> La guerra en la sala " & UCase$(.nombre) & " ha sido cancelada." & FONTTYPE_GUERRA)
    End With
End Sub

Public Sub TerminarGuerra(ByVal numeroSala As Byte)
    Dim i As Byte
    
    With Guerras(numeroSala)
        For i = 1 To .cantUsers
            Call RetirarUserGuerra(.guildA(i), True)
            Call RetirarUserGuerra(.guildB(i), True)
        Next i

        'CHOTS | Avisamos quien gano la guerra por puntos
        If .puntosGuildA > .puntosGuildB Then
            Call Guilds(.guildIndexA).IncrementarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexA).SumarGuerraGanada()
            Call Guilds(.guildIndexB).DescontarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexB).SumarGuerraPerdida()
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Guerras> El clan <" & Guilds(.guildIndexA).GuildName & "> ha ganado la guerra contra el clan <" & Guilds(.guildIndexB).GuildName & "> por puntos (" & .puntosGuildA & " a " & .puntosGuildB & ") en la sala " & UCase$(.Sala.nombre) & "." & FONTTYPE_GUERRA)
            Call SendData(SendTarget.ToGuildMembers, .guildIndexA, 0, ServerPackages.dialogo & "Guerras> Felicitaciones! Tu clan ha ganado " & .puntosPremio & " puntos por haber ganado la guerra!" & FONTTYPE_GUERRA)
        ElseIf .puntosGuildB > .puntosGuildA Then
            Call Guilds(.guildIndexB).IncrementarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexB).SumarGuerraGanada()
            Call Guilds(.guildIndexA).DescontarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexA).SumarGuerraPerdida()
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Guerras> El clan <" & Guilds(.guildIndexB).GuildName & "> ha ganado la guerra contra el clan <" & Guilds(.guildIndexA).GuildName & "> por puntos (" & .puntosGuildB & " a " & .puntosGuildA & ") en la sala " & UCase$(.Sala.nombre) & "." & FONTTYPE_GUERRA)
            Call SendData(SendTarget.ToGuildMembers, .guildIndexB, 0, ServerPackages.dialogo & "Guerras> Felicitaciones! Tu clan ha ganado " & .puntosPremio & " puntos por haber ganado la guerra!" & FONTTYPE_GUERRA)
        Else
            'CHOTS | Empate
        End If

        .timeout = 0
        .contador = 0

    End With

    With SalasGuerra(numeroSala)
        .estado = GUERRA_ESTADO_NULA
    End With
End Sub

Public Sub RetirarUserGuerra(ByVal UserIndex As Integer, ByVal bRestoreInventario As Boolean)
    Dim j As Byte
    Dim i As Byte
    Dim respawnPos As WorldPos
    Dim nPos As WorldPos
    Dim numeroSala As Byte
    
    If UserIndex = 0 Or UserList(UserIndex).guerra.enGuerra = False Then Exit Sub

    numeroSala = UserList(UserIndex).guerra.Sala

    respawnPos.Map = MAPA_GUERRAS
    respawnPos.X = 37
    respawnPos.Y = 47
    Call ClosestLegalPos(respawnPos, nPos)

    UserList(UserIndex).guerra.enGuerra = False
    UserList(UserIndex).guerra.status = 0
    UserList(UserIndex).guerra.team = 0
    UserList(UserIndex).guerra.Sala = 0
    
    With Guerras(numeroSala)
        For i = 1 To .cantUsers
            If .guildA(i) = UserIndex Then
                .guildA(i) = 0
                .cantUsersGuildA = .cantUsersGuildA - 1
            End If
            
            If .guildB(i) = UserIndex Then
                .guildB(i) = 0
                .cantUsersGuildB = .cantUsersGuildB - 1
            End If
        Next i
        If .Sala.estado = GUERRA_ESTADO_LOBBY Then
            If .cantUsersGuildB = 0 Then
                .guildIndexB = 0
            End If
            
            If .cantUsersGuildA = 0 Then
                Call CancelarGuerra(numeroSala)
            End If
            
        End If
    End With

    If bRestoreInventario = True Then
        'CHOTS | Restoreamos su inventario
        Call RestoreInventario(UserIndex)

        'CHOTS | Le damos el oro
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.Banco
        UserList(UserIndex).Stats.Banco = 0

        'CHOTS | Le limpiamos el old inventario
        UserList(UserIndex).guerra.OldInvent.NroItems = 0
        For j = 1 To MAX_INVENTORY_SLOTS
            If UserList(UserIndex).guerra.OldInvent.Object(j).ObjIndex > 0 Then
                UserList(UserIndex).guerra.OldInvent.Object(j).ObjIndex = 0
                UserList(UserIndex).guerra.OldInvent.Object(j).Amount = 0
                UserList(UserIndex).guerra.OldInvent.Object(j).Equipped = 0
            End If
        Next j
        Call UpdateUserInv(True, UserIndex, 0)

        Call EnviarOro(UserIndex)
    End If
    
    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
End Sub

Public Sub ComenzarGuerra(ByVal numeroSala As Byte)
    On Error Resume Next
    'CHOTS | Limpiamos los NPCs del mapa
    Call QuitarNpcsSala(numeroSala)

    'CHOTS | Respawneamos las casas
    Call RespawnCasasSala(numeroSala)

    'CHOTS | Respawneamos las torres
    Call RespawnTorresSala(numeroSala)

    'CHOTS | Respawneamos los NPCs items 1
    Call RespawnNpcsItemsSala(numeroSala)

    'CHOTS | Respawneamos los NPCs oro 1
    Call RespawnNpcsOroSala(numeroSala)

    'CHOTS | Respawneamos los Users
    Call TelepUsersSala(numeroSala)
    
    'CHOTS | Avisamos por consola de clan
    Call SendData(SendTarget.ToMap, 0, SalasGuerra(numeroSala).mapa, "!GLa Guerra ha comenzado! Tienes " & GUERRA_MINUTOS_DURACION & " minutos para destruir la base de tu enemigo. Buena suerte!" & COLOR_NEUTRAL)
End Sub

Public Sub QuitarNpcsSala(ByVal numeroSala As Byte)
    Dim Y As Integer
    Dim X As Integer

    With SalasGuerra(numeroSala)
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
                If MapData(.mapa, X, Y).NpcIndex > 0 Then
                    If Npclist(MapData(.mapa, X, Y).NpcIndex).Numero > 500 Then Call QuitarNPC(MapData(.mapa, X, Y).NpcIndex)
                End If
            Next X
        Next Y
    End With
End Sub

Public Sub RespawnCasasSala(ByVal numeroSala As Byte)
    Dim nPos As WorldPos
    Dim nIndex As Integer

    With SalasGuerra(numeroSala)
        'CHOTS | Team A
        nPos.Map = .mapa
        nPos.X = .posCasaTeamA.X
        nPos.Y = .posCasaTeamA.Y
        nIndex = SpawnNpc(NPC_CASA, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_A
        Npclist(nIndex).guerra.enGuerra = True

        'CHOTS | Team B
        nPos.Map = .mapa
        nPos.X = .posCasaTeamB.X
        nPos.Y = .posCasaTeamB.Y
        nIndex = SpawnNpc(NPC_CASA, nPos, False, False)
        Npclist(nIndex).guerra.enGuerra = True
    End With
End Sub

Public Sub RespawnTorresSala(ByVal numeroSala As Byte)
    Dim nPos As WorldPos
    Dim nIndex As Integer

    With SalasGuerra(numeroSala)
        'CHOTS | Torre 1 Team A
        nPos.Map = .mapa
        nPos.X = .posTorre1TeamA.X
        nPos.Y = .posTorre1TeamA.Y
        nIndex = SpawnNpc(NPC_TORRE, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_A
        Npclist(nIndex).guerra.enGuerra = True

        'CHOTS | Torre 2 Team A
        nPos.Map = .mapa
        nPos.X = .posTorre2TeamA.X
        nPos.Y = .posTorre2TeamA.Y
        nIndex = SpawnNpc(NPC_TORRE, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_A
        Npclist(nIndex).guerra.enGuerra = True

        'CHOTS | Torre 3 Team A
        nPos.Map = .mapa
        nPos.X = .posTorre3TeamA.X
        nPos.Y = .posTorre3TeamA.Y
        nIndex = SpawnNpc(NPC_TORRE, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_A
        Npclist(nIndex).guerra.enGuerra = True

        'CHOTS | Torre 1 Team B
        nPos.Map = .mapa
        nPos.X = .posTorre1TeamB.X
        nPos.Y = .posTorre1TeamB.Y
        nIndex = SpawnNpc(NPC_TORRE, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_B
        Npclist(nIndex).guerra.enGuerra = True

        'CHOTS | Torre 2 Team B
        nPos.Map = .mapa
        nPos.X = .posTorre2TeamB.X
        nPos.Y = .posTorre2TeamB.Y
        nIndex = SpawnNpc(NPC_TORRE, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_B
        Npclist(nIndex).guerra.enGuerra = True

        'CHOTS | Torre 3 Team B
        nPos.Map = .mapa
        nPos.X = .posTorre3TeamB.X
        nPos.Y = .posTorre3TeamB.Y
        nIndex = SpawnNpc(NPC_TORRE, nPos, False, False)
        Npclist(nIndex).guerra.team = GUERRA_TEAM_B
        Npclist(nIndex).guerra.enGuerra = True
    End With
End Sub

Public Sub RespawnNpcsItemsSala(ByVal numeroSala As Byte)
    Dim nPos As WorldPos
    Dim nIndex As Integer
    Dim i As Byte

    With SalasGuerra(numeroSala)
        'CHOTS | Tipo 1
        For i = 1 To .cantidadNpcsItems1
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsItems1, NPC_ITEMS_1)
            Npclist(nIndex).guerra.enGuerra = True
        Next i

        'CHOTS | Tipo 2
        For i = 1 To .cantidadNpcsItems2
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsItems2, NPC_ITEMS_2)
            Npclist(nIndex).guerra.enGuerra = True
        Next i

        'CHOTS | Tipo 3
        For i = 1 To .cantidadNpcsItems3
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsItems3, NPC_ITEMS_3)
            Npclist(nIndex).guerra.enGuerra = True
        Next i
    End With
End Sub

Public Sub RespawnNpcsOroSala(ByVal numeroSala As Byte)
    Dim nIndex As Integer
    Dim i As Byte

    With SalasGuerra(numeroSala)
        'CHOTS | Tipo 1
        For i = 1 To .cantidadNpcsOro1
            'CHOTS | Team 1
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsOro1TeamA, NPC_ORO_1)
            Npclist(nIndex).guerra.team = GUERRA_TEAM_A
            Npclist(nIndex).guerra.enGuerra = True

            'CHOTS | Team 2
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsOro1TeamB, NPC_ORO_1)
            Npclist(nIndex).guerra.team = GUERRA_TEAM_B
            Npclist(nIndex).guerra.enGuerra = True
        Next i

        'CHOTS | Tipo 2
        For i = 1 To .cantidadNpcsOro2
            'CHOTS | Team 1
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsOro2TeamA, NPC_ORO_2)
            Npclist(nIndex).guerra.team = GUERRA_TEAM_A
            Npclist(nIndex).guerra.enGuerra = True

            'CHOTS | Team 2
            nIndex = SpawnNpcZonaSala(.mapa, .zonaNpcsOro2TeamB, NPC_ORO_2)
            Npclist(nIndex).guerra.team = GUERRA_TEAM_B
            Npclist(nIndex).guerra.enGuerra = True
        Next i

    End With
End Sub

Public Function SpawnNpcZonaSala(ByVal mapa As Integer, ByRef zona As cZonaGuerra, ByVal numNpc As Integer) As Integer
    Dim nPos As WorldPos
    nPos.Map = mapa
    nPos.X = RandomNumber(zona.minX, zona.maxX)
    nPos.Y = RandomNumber(zona.minY, zona.maxY)
    SpawnNpcZonaSala = SpawnNpc(numNpc, nPos, False, False)
End Function

Public Sub TelepUsersSala(ByVal numeroSala As Byte)
    Dim i As Byte
    Dim respawnPos As WorldPos
    Dim nPos As WorldPos

    With Guerras(numeroSala)
        For i = 1 To .cantUsers
            'CHOTS | Les cambiamos el inventario
            Call DarItemsGuerra(.guildA(i), GUERRA_TEAM_A)
            Call DarItemsGuerra(.guildB(i), GUERRA_TEAM_B)

            'CHOTS | Team 1
            respawnPos.Map = .Sala.mapa
            respawnPos.X = RandomNumber(.Sala.zonaRespawnTeamA.minX, .Sala.zonaRespawnTeamA.maxX)
            respawnPos.Y = RandomNumber(.Sala.zonaRespawnTeamA.minY, .Sala.zonaRespawnTeamA.maxY)
            Call ClosestLegalPos(respawnPos, nPos)
            Call WarpUserChar(.guildA(i), nPos.Map, nPos.X, nPos.Y, True)
            UserList(.guildA(i)).guerra.status = GUERRA_ESTADO_INICIADA
            
            'CHOTS | Team 2
            respawnPos.Map = .Sala.mapa
            respawnPos.X = RandomNumber(.Sala.zonaRespawnTeamB.minX, .Sala.zonaRespawnTeamB.maxX)
            respawnPos.Y = RandomNumber(.Sala.zonaRespawnTeamB.minY, .Sala.zonaRespawnTeamB.maxY)
            Call ClosestLegalPos(respawnPos, nPos)
            Call WarpUserChar(.guildB(i), nPos.Map, nPos.X, nPos.Y, True)
            UserList(.guildB(i)).guerra.status = GUERRA_ESTADO_INICIADA
        Next i
        
    End With
End Sub

Public Sub DarItemsGuerra(ByVal UserIndex As Integer, ByVal team As Byte)
    Dim j As Byte

    If UserIndex = 0 Or UserList(UserIndex).guerra.enGuerra = False Then Exit Sub

    'CHOTS | Backupeamos su inventario
    Call BackupInventario(UserIndex)
    
    'CHOTS | Le damos el oro
    UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + UserList(UserIndex).Stats.GLD
    UserList(UserIndex).Stats.GLD = 0

    'CHOTS | Le limpiamos el inventario
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
        End If
    Next j

    'CHOTS | Le damos los items de guerra
    ' Vestimenta
    ' Daga
    ' 100 rojas
    ' 100 azules si corresponde

    UserList(UserIndex).Invent.NroItems = 6

    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        If team = GUERRA_TEAM_A Then
            UserList(UserIndex).Invent.Object(1).ObjIndex = ITEMS_ROPA_BAJO_A
        Else
            UserList(UserIndex).Invent.Object(1).ObjIndex = ITEMS_ROPA_BAJO_B
        End If
    Else
        If team = GUERRA_TEAM_A Then
            UserList(UserIndex).Invent.Object(1).ObjIndex = ITEMS_ROPA_ALTO_A
        Else
            UserList(UserIndex).Invent.Object(1).ObjIndex = ITEMS_ROPA_ALTO_B
        End If
    End If
    
    UserList(UserIndex).Invent.Object(1).Amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1
    UserList(UserIndex).Invent.ArmourEqpSlot = 1
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(1).ObjIndex

    UserList(UserIndex).Invent.Object(2).ObjIndex = ITEMS_DAGA
    UserList(UserIndex).Invent.Object(2).Amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1
    UserList(UserIndex).Invent.WeaponEqpSlot = 2
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(2).ObjIndex
    
    UserList(UserIndex).char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
    UserList(UserIndex).char.WeaponAnim = 9
    
    UserList(UserIndex).flags.Desnudo = 0

    UserList(UserIndex).Invent.Object(3).ObjIndex = ITEMS_MANZANA
    UserList(UserIndex).Invent.Object(3).Amount = 100
    UserList(UserIndex).Invent.Object(4).ObjIndex = ITEMS_AGUA
    UserList(UserIndex).Invent.Object(4).Amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = ITEMS_POCION_ROJA

    If UserList(UserIndex).Stats.MaxMAN = 0 Then
        UserList(UserIndex).Invent.Object(5).Amount = 100
        UserList(UserIndex).Invent.NroItems = 5
    Else
        UserList(UserIndex).Invent.Object(5).Amount = 50

        UserList(UserIndex).Invent.Object(6).ObjIndex = ITEMS_POCION_AZUL
        UserList(UserIndex).Invent.Object(6).Amount = 50
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call EnviarOro(UserIndex)
End Sub

Public Sub MuereNpcGuerra(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim npcTeam As Byte
    Dim numeroSala As Byte
    Dim nIndex As Integer
    Dim Puntos As Integer
    Dim puntosRecibe As Byte
    npcTeam = Npclist(NpcIndex).guerra.team
    numeroSala = GetSalaMapa(Npclist(NpcIndex).Pos.Map)

    If numeroSala = 0 Then
        Call QuitarNPC(NpcIndex)
        Exit Sub
    End If

    Puntos = 0

    With Guerras(numeroSala)
        Select Case Npclist(NpcIndex).Numero
            Case NPC_CASA:
                'CHOTS | Mataron la casa, gana la guerra el otro team
                Call GanaGuerra(numeroSala, UserList(UserIndex).guerra.team)
                Exit Sub

            Case NPC_TORRE:
                'CHOTS | Mataron una torre de defensa, damos puntos al otro team
                Puntos = PUNTOS_NPC_TORRE
                puntosRecibe = IIf(Npclist(NpcIndex).guerra.team = GUERRA_TEAM_B, GUERRA_TEAM_A, GUERRA_TEAM_B)
                Call SendData(SendTarget.ToMap, 0, SalasGuerra(numeroSala).mapa, "!G" & UserList(UserIndex).Name & " ha destruido una torre de defensa!" & IIf(UserList(UserIndex).guerra.team = GUERRA_TEAM_A, COLOR_TEAM_A, COLOR_TEAM_B))
                
            Case NPC_ITEMS_1:
                'CHOTS | Muere un NPC de Items, damos puntos y respawn
                nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsItems1, NPC_ITEMS_1)
                Puntos = PUNTOS_NPC_ITEMS_1
                puntosRecibe = UserList(UserIndex).guerra.team
                Call SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsItems1, NPC_ITEMS_1)

            Case NPC_ITEMS_2:
                'CHOTS | Muere un NPC de Items, damos puntos y respawn
                nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsItems2, NPC_ITEMS_2)
                Puntos = PUNTOS_NPC_ITEMS_2
                puntosRecibe = UserList(UserIndex).guerra.team
                Call SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsItems2, NPC_ITEMS_2)

            Case NPC_ITEMS_3:
                'CHOTS | Muere un NPC de Items, damos puntos y respawn
                nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsItems3, NPC_ITEMS_3)
                Puntos = PUNTOS_NPC_ITEMS_3
                puntosRecibe = UserList(UserIndex).guerra.team
                Call SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsItems3, NPC_ITEMS_3)

            Case NPC_ORO_1:
                'CHOTS | Muere un NPC de oro, damos puntos y respawn en la zona del team
                If npcTeam = GUERRA_TEAM_A Then
                    nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsOro1TeamA, NPC_ORO_1)
                    Npclist(nIndex).guerra.team = GUERRA_TEAM_A
                Else
                    nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsOro1TeamB, NPC_ORO_1)
                    Npclist(nIndex).guerra.team = GUERRA_TEAM_B
                End If
                Puntos = PUNTOS_NPC_ORO_1
                puntosRecibe = UserList(UserIndex).guerra.team

            Case NPC_ORO_2:
                'CHOTS | Muere un NPC de oro, damos puntos y respawn en la zona del team
                If npcTeam = GUERRA_TEAM_A Then
                    nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsOro2TeamA, NPC_ORO_2)
                    Npclist(nIndex).guerra.team = GUERRA_TEAM_A
                Else
                    nIndex = SpawnNpcZonaSala(.Sala.mapa, .Sala.zonaNpcsOro2TeamB, NPC_ORO_2)
                    Npclist(nIndex).guerra.team = GUERRA_TEAM_B
                End If
                Puntos = PUNTOS_NPC_ORO_2
                puntosRecibe = UserList(UserIndex).guerra.team
        End Select
    End With

    If Puntos > 0 Then
        Call DarPuntosGuera(numeroSala, puntosRecibe, Puntos)
    End If

    Dim MiNPC As Npc
    MiNPC = Npclist(NpcIndex)

    Call QuitarNPC(NpcIndex)

    Call NPC_TIRAR_ITEMS(MiNPC)
    
    Call NPCTirarOro(MiNPC, UserIndex)
    
    Call EnviarOro(UserIndex)

End Sub

Public Sub MuereUserGuerra(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    Dim numeroSala As Byte
    numeroSala = GetSalaMapa(UserList(AttackerIndex).Pos.Map)
    
    If numeroSala = 0 Then Exit Sub
    
    'CHOTS | Avisamos por consola de clan
    Call SendData(SendTarget.ToMap, 0, SalasGuerra(numeroSala).mapa, "!G" & UserList(AttackerIndex).Name & " ha matado a " & UserList(VictimIndex).Name & IIf(UserList(AttackerIndex).guerra.team = GUERRA_TEAM_A, COLOR_TEAM_A, COLOR_TEAM_B))

    Call DarPuntosGuera(numeroSala, UserList(AttackerIndex).guerra.team, PUNTOS_FRAG)
End Sub

Public Sub GanaGuerra(ByVal numeroSala As Byte, ByVal team As Byte)
    Dim i As Byte
    
    With Guerras(numeroSala)
        For i = 1 To .cantUsers
            Call RetirarUserGuerra(.guildA(i), True)
            Call RetirarUserGuerra(.guildB(i), True)
        Next i

        'CHOTS | Sumamos puntos
        If team = GUERRA_TEAM_A Then
            Call Guilds(.guildIndexA).IncrementarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexA).SumarGuerraGanada()
            Call Guilds(.guildIndexB).DescontarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexB).SumarGuerraPerdida()
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Guerras> El clan <" & Guilds(.guildIndexA).GuildName & "> ha ganado la guerra contra el clan <" & Guilds(.guildIndexB).GuildName & "> en la sala " & UCase$(.Sala.nombre) & "." & FONTTYPE_GUERRA)
            Call SendData(SendTarget.ToGuildMembers, .guildIndexA, 0, ServerPackages.dialogo & "Guerras> Felicitaciones! Tu clan ha ganado " & .puntosPremio & " puntos por haber ganado la guerra!" & FONTTYPE_GUERRA)
        ElseIf team = GUERRA_TEAM_B Then
            Call Guilds(.guildIndexB).IncrementarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexB).SumarGuerraGanada()
            Call Guilds(.guildIndexA).DescontarGuildPoints(.puntosPremio)
            Call Guilds(.guildIndexA).SumarGuerraPerdida()
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Guerras> El clan <" & Guilds(.guildIndexB).GuildName & "> ha ganado la guerra contra el clan <" & Guilds(.guildIndexA).GuildName & "> en la sala " & UCase$(.Sala.nombre) & "." & FONTTYPE_GUERRA)
            Call SendData(SendTarget.ToGuildMembers, .guildIndexB, 0, ServerPackages.dialogo & "Guerras> Felicitaciones! Tu clan ha ganado " & .puntosPremio & " puntos por haber ganado la guerra!" & FONTTYPE_GUERRA)
        Else
            'CHOTS | Error
        End If

        .timeout = 0
        .contador = 0
        .Sala.estado = GUERRA_ESTADO_NULA

    End With

    With SalasGuerra(numeroSala)
        .estado = GUERRA_ESTADO_NULA
    End With

End Sub

Public Function GetSalaMapa(ByVal mapa As Integer) As Byte
    Dim i As Byte

    GetSalaMapa = 0

    For i = 1 To GUERRA_CANT_SALAS
        With SalasGuerra(i)
            If .mapa = mapa Then
                GetSalaMapa = .Numero
            End If
        End With
    Next i
End Function

Public Sub DarPuntosGuera(ByVal numeroSala As Byte, ByVal team As Byte, ByVal Puntos As Integer)
    If team = GUERRA_TEAM_A Then
        If Guerras(numeroSala).puntosGuildA > 32000 Then Exit Sub
        Guerras(numeroSala).puntosGuildA = Guerras(numeroSala).puntosGuildA + Puntos
    Else
        If Guerras(numeroSala).puntosGuildB > 32000 Then Exit Sub
        Guerras(numeroSala).puntosGuildB = Guerras(numeroSala).puntosGuildB + Puntos
    End If
End Sub

Public Sub BackupInventario(ByVal UserIndex As Integer)
    Dim i As Byte

    With UserList(UserIndex).guerra.OldInvent
        .WeaponEqpObjIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
        .WeaponEqpSlot = UserList(UserIndex).Invent.WeaponEqpSlot
        .ArmourEqpObjIndex = UserList(UserIndex).Invent.ArmourEqpObjIndex
        .ArmourEqpSlot = UserList(UserIndex).Invent.ArmourEqpSlot
        .EscudoEqpObjIndex = UserList(UserIndex).Invent.EscudoEqpObjIndex
        .EscudoEqpSlot = UserList(UserIndex).Invent.EscudoEqpSlot
        .CascoEqpObjIndex = UserList(UserIndex).Invent.CascoEqpObjIndex
        .CascoEqpSlot = UserList(UserIndex).Invent.CascoEqpSlot
        .MunicionEqpObjIndex = UserList(UserIndex).Invent.MunicionEqpObjIndex
        .MunicionEqpSlot = UserList(UserIndex).Invent.MunicionEqpSlot
        .HerramientaEqpObjIndex = UserList(UserIndex).Invent.HerramientaEqpObjIndex
        .HerramientaEqpSlot = UserList(UserIndex).Invent.HerramientaEqpSlot
        .MonturaObjIndex = UserList(UserIndex).Invent.MonturaObjIndex
        .MonturaSlot = UserList(UserIndex).Invent.MonturaSlot
        .BarcoObjIndex = UserList(UserIndex).Invent.BarcoObjIndex
        .BarcoSlot = UserList(UserIndex).Invent.BarcoSlot
        .NroItems = UserList(UserIndex).Invent.NroItems

        For i = 1 To MAX_INVENTORY_SLOTS
            .Object(i).ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
            .Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount
            .Object(i).Equipped = UserList(UserIndex).Invent.Object(i).Equipped
            .Object(i).ProbTirar = UserList(UserIndex).Invent.Object(i).ProbTirar
        Next i
    End With

End Sub

Public Sub RestoreInventario(ByVal UserIndex As Integer)
    Dim i As Byte

    With UserList(UserIndex).Invent
        .WeaponEqpObjIndex = UserList(UserIndex).guerra.OldInvent.WeaponEqpObjIndex
        .WeaponEqpSlot = UserList(UserIndex).guerra.OldInvent.WeaponEqpSlot
        .ArmourEqpObjIndex = UserList(UserIndex).guerra.OldInvent.ArmourEqpObjIndex
        .ArmourEqpSlot = UserList(UserIndex).guerra.OldInvent.ArmourEqpSlot
        .EscudoEqpObjIndex = UserList(UserIndex).guerra.OldInvent.EscudoEqpObjIndex
        .EscudoEqpSlot = UserList(UserIndex).guerra.OldInvent.EscudoEqpSlot
        .CascoEqpObjIndex = UserList(UserIndex).guerra.OldInvent.CascoEqpObjIndex
        .CascoEqpSlot = UserList(UserIndex).guerra.OldInvent.CascoEqpSlot
        .MunicionEqpObjIndex = UserList(UserIndex).guerra.OldInvent.MunicionEqpObjIndex
        .MunicionEqpSlot = UserList(UserIndex).guerra.OldInvent.MunicionEqpSlot
        .HerramientaEqpObjIndex = UserList(UserIndex).guerra.OldInvent.HerramientaEqpObjIndex
        .HerramientaEqpSlot = UserList(UserIndex).guerra.OldInvent.HerramientaEqpSlot
        .MonturaObjIndex = UserList(UserIndex).guerra.OldInvent.MonturaObjIndex
        .MonturaSlot = UserList(UserIndex).guerra.OldInvent.MonturaSlot
        .BarcoObjIndex = UserList(UserIndex).guerra.OldInvent.BarcoObjIndex
        .BarcoSlot = UserList(UserIndex).guerra.OldInvent.BarcoSlot
        .NroItems = UserList(UserIndex).guerra.OldInvent.NroItems

        For i = 1 To MAX_INVENTORY_SLOTS
            .Object(i).ObjIndex = UserList(UserIndex).guerra.OldInvent.Object(i).ObjIndex
            .Object(i).Amount = UserList(UserIndex).guerra.OldInvent.Object(i).Amount
            .Object(i).Equipped = UserList(UserIndex).guerra.OldInvent.Object(i).Equipped
            .Object(i).ProbTirar = UserList(UserIndex).guerra.OldInvent.Object(i).ProbTirar
        Next i

        If .ArmourEqpObjIndex > 0 Then _
            UserList(UserIndex).char.Body = ObjData(.ArmourEqpObjIndex).Ropaje
            
        If .EscudoEqpObjIndex > 0 Then
            UserList(UserIndex).char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim
        Else
            UserList(UserIndex).char.ShieldAnim = NingunEscudo
        End If
        
        If .WeaponEqpObjIndex > 0 Then
            UserList(UserIndex).char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim
        Else
            UserList(UserIndex).char.WeaponAnim = NingunArma
        End If

        If .CascoEqpObjIndex > 0 Then
            UserList(UserIndex).char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim
        Else
            UserList(UserIndex).char.CascoAnim = NingunCasco
        End If
        
    End With
End Sub

Public Sub SendStatusGuerras(ByVal UserIndex As Integer)
    Dim i As Byte
    For i = 1 To GUERRA_CANT_SALAS
        Dim currentStatus As String
        With Guerras(i)
            currentStatus = "Sala " & .Sala.nombre & ": "
            Select Case .Sala.estado
                Case GUERRA_ESTADO_NULA:
                    currentStatus = currentStatus & " la sala está vacía."

                Case GUERRA_ESTADO_INICIADA:
                    currentStatus = currentStatus & "<" & Guilds(.guildIndexA).GuildName & "> (" & .puntosGuildA & ") vs <" & Guilds(.guildIndexB).GuildName & "> (" & .puntosGuildB & "). " & .timeout & " minutos restantes."

                Case GUERRA_ESTADO_LOBBY:
                    If .guildIndexB > 0 Then
                        currentStatus = currentStatus & "Los clanes <" & Guilds(.guildIndexA).GuildName & "> y <" & Guilds(.guildIndexB).GuildName & "> se están preparando para la guerra."
                    Else
                        currentStatus = currentStatus & "El clan <" & Guilds(.guildIndexA).GuildName & "> está esperando rival para la guerra."
                    End If
            End Select

            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & currentStatus & FONTTYPE_GUERRA)
        End With
    Next i
End Sub
