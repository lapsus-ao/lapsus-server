Attribute VB_Name = "Ranking"
'Modulo de Ranking
'Creado por Juan Andrés Dalmasso(CHOTS)
'Para Lapsus AO
'CHOTS_AO@HOTMAIL.COM
'Modificado para Lapsus 3 por CHOTS
'17/10/2011

Public Ranking_Trofeos_Nick As String
Public Ranking_Trofeos_Cant As Integer
Public Ranking_Frags_Nick As String
Public Ranking_Frags_Cant As Integer
Public Ranking_Torneos_Nick As String
Public Ranking_Torneos_Cant As Integer

Public Type tRanking_Usuario
    nombre As String
    Level As Byte
    Exp As Double
End Type
Public Ranking_Level() As tRanking_Usuario
Sub GuardarRanking()
'CHOTS | Sistema de Ranking
'CHOTS | Guarda los datos
Dim n As Integer
Dim i As Byte
n = FreeFile
Open (IniPath & "RANKING.INI") For Output As #n
Print #n, "[TROFEOS]"
Print #n, "Nombre=" & Ranking_Trofeos_Nick
Print #n, "Cantidad=" & Ranking_Trofeos_Cant
Print #n, ""

Print #n, "[FRAGS]"
Print #n, "Nombre=" & Ranking_Frags_Nick
Print #n, "Cantidad=" & Ranking_Frags_Cant
Print #n, ""

Print #n, "[TORNEOS]"
Print #n, "Nombre=" & Ranking_Torneos_Nick
Print #n, "Cantidad=" & Ranking_Torneos_Cant
Print #n, ""

Print #n, "[LEVEL]"
For i = 1 To 10
    Print #n, "Nombre" & i & "=" & Ranking_Level(i).nombre
    Print #n, "Nivel" & i & "=" & Ranking_Level(i).Level
    Print #n, "Exp" & i & "=" & Ranking_Level(i).Exp
Next i

Close #n

'CHOTS | Sistema de Ranking
End Sub
Sub CargarRanking()
'CHOTS | Sistema de Ranking
'CHOTS | Carga los datos
Ranking_Trofeos_Nick = GetVar(IniPath & "RANKING.INI", "TROFEOS", "Nombre")
Ranking_Trofeos_Cant = val(GetVar(IniPath & "RANKING.INI", "TROFEOS", "Cantidad"))
Ranking_Frags_Nick = GetVar(IniPath & "RANKING.INI", "FRAGS", "Nombre")
Ranking_Frags_Cant = val(GetVar(IniPath & "RANKING.INI", "FRAGS", "Cantidad"))
Ranking_Torneos_Nick = GetVar(IniPath & "RANKING.INI", "TORNEOS", "Nombre")
Ranking_Torneos_Cant = val(GetVar(IniPath & "RANKING.INI", "TORNEOS", "Cantidad"))
'CHOTS | Sistema de Ranking

ReDim Ranking_Level(1 To 10) As tRanking_Usuario
Dim i As Byte
Dim usuario As tRanking_Usuario

For i = 1 To 10
    usuario.nombre = getNombre(i)
    usuario.Level = getLevel(i)
    usuario.Exp = getExp(i)
    Ranking_Level(i) = usuario
Next i

End Sub
Sub ActualizarRanking(ByVal UserIndex As Integer, ByVal Ranking As Byte)
'CHOTS | Tipos de Ranking
'CHOTS | 1: Trofeos
'CHOTS | 2: Frags
'CHOTS | 3: Level
'CHOTS | 4: Torneos Auto

If UserList(UserIndex).flags.Privilegios <> PlayerType.User Then Exit Sub

Select Case Ranking
    Case 1
        If UserList(UserIndex).Stats.TrofOro > Ranking_Trofeos_Cant Then
            Ranking_Trofeos_Nick = UserList(UserIndex).Name
            Ranking_Trofeos_Cant = UserList(UserIndex).Stats.TrofOro
        End If
        Exit Sub
        
    Case 2
        If (UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados) > Ranking_Frags_Cant Then
            Ranking_Frags_Nick = UserList(UserIndex).Name
            Ranking_Frags_Cant = UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados
        End If
        Exit Sub
        
    Case 3
        Call ActualizarRankingLevel(UserIndex)
        Exit Sub
        
    Case 4
        If UserList(UserIndex).Stats.TorneosAuto > Ranking_Torneos_Cant Then
            Ranking_Torneos_Nick = UserList(UserIndex).Name
            Ranking_Torneos_Cant = UserList(UserIndex).Stats.TorneosAuto
        End If
        Exit Sub
        
End Select

End Sub

Private Function getNombre(ByVal indice As Byte) As String
    getNombre = GetVar(IniPath & "RANKING.INI", "LEVEL", "Nombre" & indice)
End Function
Private Function getLevel(ByVal indice As Byte) As Byte
    getLevel = GetVar(IniPath & "RANKING.INI", "LEVEL", "Nivel" & indice)
End Function
Private Function getExp(ByVal indice As Byte) As Double
    getExp = GetVar(IniPath & "RANKING.INI", "LEVEL", "Exp" & indice)
End Function

Public Function ActualizarRankingLevel(ByVal UserIndex As Integer)
    Dim usuario As tRanking_Usuario
    Dim posicion As Byte
    
    If UserList(UserIndex).Stats.ELV = STAT_MAXELV Then
        UserList(UserIndex).Stats.Exp = 0
        UserList(UserIndex).Stats.ELU = 1
    End If

    If Not estaEnRanking(UserList(UserIndex).Name) Then
    
        If UserList(UserIndex).Stats.ELV > Ranking_Level(10).Level Then
            usuario.nombre = UserList(UserIndex).Name
            usuario.Level = UserList(UserIndex).Stats.ELV
            usuario.Exp = Round(((UserList(UserIndex).Stats.Exp * 100) / UserList(UserIndex).Stats.ELU), 2)
            Ranking_Level(10) = usuario
            Call ordenarRankingLevel
        ElseIf UserList(UserIndex).Stats.ELV = Ranking_Level(10).Level Then
            If Round(((UserList(UserIndex).Stats.Exp * 100) / UserList(UserIndex).Stats.ELU), 2) > Ranking_Level(10).Exp Then
                usuario.nombre = UserList(UserIndex).Name
                usuario.Level = UserList(UserIndex).Stats.ELV
                usuario.Exp = Round(((UserList(UserIndex).Stats.Exp * 100) / UserList(UserIndex).Stats.ELU), 2)
                Ranking_Level(10) = usuario
                Call ordenarRankingLevel
            End If
        End If
        
    Else
        posicion = getPosRanking(UserList(UserIndex).Name)
        Ranking_Level(posicion).Level = UserList(UserIndex).Stats.ELV
        Ranking_Level(posicion).Exp = Round(((UserList(UserIndex).Stats.Exp * 100) / UserList(UserIndex).Stats.ELU), 2)
        Call ordenarRankingLevel
    End If
    
    
End Function

Private Function estaEnRanking(ByVal nick As String) As Boolean
Dim i As Byte
estaEnRanking = False

For i = 1 To 10
    If UCase$(Ranking_Level(i).nombre) = UCase$(nick) Then
        estaEnRanking = True
        Exit Function
    End If
Next i
End Function
Private Function getPosRanking(ByVal nick As String) As Byte
Dim i As Byte
getPosRanking = 0
For i = 1 To 10
    If UCase$(Ranking_Level(i).nombre) = UCase$(nick) Then
        getPosRanking = i
        Exit Function
    End If
Next i
End Function

Private Sub ordenarRankingLevel()
Dim i As Byte
Dim j As Byte
Dim aux As tRanking_Usuario

For i = 1 To 10
    For j = (i + 1) To 10
        If Ranking_Level(i).Level < Ranking_Level(j).Level Then
            aux = Ranking_Level(i)
            Ranking_Level(i) = Ranking_Level(j)
            Ranking_Level(j) = aux
        ElseIf (Ranking_Level(i).Level = Ranking_Level(j).Level) And (Ranking_Level(i).Exp < Ranking_Level(j).Exp) Then
            aux = Ranking_Level(i)
            Ranking_Level(i) = Ranking_Level(j)
            Ranking_Level(j) = aux
        End If
    Next j
Next i
            
End Sub

'CHOTS | Sistema de Ranking
Sub ActualizarWebUsuarios(Optional ByVal number As Integer = -1)
On Local Error Resume Next
Dim baseUrl As String
baseUrl = "http://www.lapsus2017.com/update.php?param="
Dim enviar As String

enviar = "1@" & IIf(number >= 0, number, NumUsers) & "@" & recordusuarios

frmMain.InetUsers.Execute baseUrl & enviar, "GET"

End Sub

Sub ActualizarWeb()
On Local Error Resume Next
'CHOTS | Envía la web

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Actualizando Web..." & FONTTYPE_SERVER)

Dim i, j, aux As Integer
Dim clan1, clan2, clan3, clan4, clan5, clan6, clan7, clan8, clan9, clan10 As String
Dim punto1, punto2, punto3, punto4, punto5, punto6, punto7, punto8, punto9, punto10 As Long
Dim guerra1, guerra2, guerra3, guerra4, guerra5, guerra6, guerra7, guerra8, guerra9, guerra10 As String
Dim guerrasGanadas1, guerrasGanadas2, guerrasGanadas3, guerrasGanadas4, guerrasGanadas5, guerrasGanadas6, guerrasGanadas7, guerrasGanadas8, guerrasGanadas9, guerrasGanadas10 As Integer
Dim guerrasPerdidas1, guerrasPerdidas2, guerrasPerdidas3, guerrasPerdidas4, guerrasPerdidas5, guerrasPerdidas6, guerrasPerdidas7, guerrasPerdidas8, guerrasPerdidas9, guerrasPerdidas10 As Integer
Dim enviar As String
Dim baseUrl As String
baseUrl = "http://www.lapsus2017.com/update.php?param="

If CANTIDADDECLANES < 10 Then
    clan1 = "Game Masters"
    clan2 = "Game Masters"
    clan3 = "Game Masters"
    clan4 = "Game Masters"
    clan5 = "Game Masters"
    clan6 = "Game Masters"
    clan7 = "Game Masters"
    clan8 = "Game Masters"
    clan9 = "Game Masters"
    clan10 = "Game Masters"

    punto1 = 10
    punto2 = 9
    punto3 = 8
    punto4 = 7
    punto5 = 6
    punto6 = 5
    punto7 = 4
    punto8 = 3
    punto9 = 2
    punto10 = 1

    guerra1 = "Game Masters1"
    guerra2 = "Game Masters2"
    guerra3 = "Game Masters3"
    guerra4 = "Game Masters4"
    guerra5 = "Game Masters5"
    guerra6 = "Game Masters6"
    guerra7 = "Game Masters7"
    guerra8 = "Game Masters8"
    guerra9 = "Game Masters9"
    guerra10 = "Game Masters10"

    guerrasGanadas1 = 1
    guerrasGanadas2 = 1
    guerrasGanadas3 = 1
    guerrasGanadas4 = 1
    guerrasGanadas5 = 1
    guerrasGanadas6 = 1
    guerrasGanadas7 = 1
    guerrasGanadas8 = 1
    guerrasGanadas9 = 1
    guerrasGanadas10 = 1

    guerrasPerdidas1 = 0
    guerrasPerdidas2 = 0
    guerrasPerdidas3 = 0
    guerrasPerdidas4 = 0
    guerrasPerdidas5 = 0
    guerrasPerdidas6 = 0
    guerrasPerdidas7 = 0
    guerrasPerdidas8 = 0
    guerrasPerdidas9 = 0
    guerrasPerdidas10 = 0

Else

    'CHOTS | Ranking de clanes por puntos
    Dim clanesRankeados As Integer
    clanesRankeados = 0

    ReDim vecChuts(1 To CANTIDADDECLANES) As Long

    'CHOTS | Almacena los clanes en un Array y los ordena
    For i = 1 To CANTIDADDECLANES
        If Guilds(i).GetGuildPoints >= minPuntosClan Then
            clanesRankeados = clanesRankeados + 1
            vecChuts(clanesRankeados) = i
        End If
    Next i
    
    ReDim vecClan(1 To clanesRankeados) As Long

    For i = 1 To clanesRankeados
        vecClan(i) = vecChuts(i)
    Next i

    For i = 1 To (clanesRankeados - 1)
        For j = (i + 1) To clanesRankeados
            If Guilds(vecClan(i)).GetGuildPoints < Guilds(vecClan(j)).GetGuildPoints Then
                aux = vecClan(i)
                vecClan(i) = vecClan(j)
                vecClan(j) = aux
            End If
        Next j
    Next i
    'CHOTS | Almacena los clanes en un Array y los ordena
    
    'CHOTS | Carga Puntos y Nombres
    clan1 = Guilds(vecClan(1)).GuildName
    clan2 = Guilds(vecClan(2)).GuildName
    clan3 = Guilds(vecClan(3)).GuildName
    clan4 = Guilds(vecClan(4)).GuildName
    clan5 = Guilds(vecClan(5)).GuildName
    clan6 = Guilds(vecClan(6)).GuildName
    clan7 = Guilds(vecClan(7)).GuildName
    clan8 = Guilds(vecClan(8)).GuildName
    clan9 = Guilds(vecClan(9)).GuildName
    clan10 = Guilds(vecClan(10)).GuildName

    punto1 = Guilds(vecClan(1)).GetGuildPoints
    punto2 = Guilds(vecClan(2)).GetGuildPoints
    punto3 = Guilds(vecClan(3)).GetGuildPoints
    punto4 = Guilds(vecClan(4)).GetGuildPoints
    punto5 = Guilds(vecClan(5)).GetGuildPoints
    punto6 = Guilds(vecClan(6)).GetGuildPoints
    punto7 = Guilds(vecClan(7)).GetGuildPoints
    punto8 = Guilds(vecClan(8)).GetGuildPoints
    punto9 = Guilds(vecClan(9)).GetGuildPoints
    punto10 = Guilds(vecClan(10)).GetGuildPoints
    'CHOTS | Carga Puntos y Nombres
    
    minPuntosClan = punto10 - 50
    Call WriteVar(App.Path & "\Server.ini", "OTROS", "minPuntosClan", minPuntosClan)


    'CHOTS | Ranking de clanes por guerras

    ReDim vecClan(1 To CANTIDADDECLANES) As Long
    For i = 1 To CANTIDADDECLANES
        vecClan(i) = i
    Next i

    For i = 1 To (CANTIDADDECLANES - 1)
        For j = (i + 1) To CANTIDADDECLANES
            If (Guilds(vecClan(i)).GetGuerrasGanadas() - Guilds(vecClan(i)).GetGuerrasPerdidas()) < (Guilds(vecClan(j)).GetGuerrasGanadas() - Guilds(vecClan(j)).GetGuerrasPerdidas()) Then
                aux = vecClan(i)
                vecClan(i) = vecClan(j)
                vecClan(j) = aux
            End If
        Next j
    Next i

    guerra1 = Guilds(vecClan(1)).GuildName
    guerra2 = Guilds(vecClan(2)).GuildName
    guerra3 = Guilds(vecClan(3)).GuildName
    guerra4 = Guilds(vecClan(4)).GuildName
    guerra5 = Guilds(vecClan(5)).GuildName
    guerra6 = Guilds(vecClan(6)).GuildName
    guerra7 = Guilds(vecClan(7)).GuildName
    guerra8 = Guilds(vecClan(8)).GuildName
    guerra9 = Guilds(vecClan(9)).GuildName
    guerra10 = Guilds(vecClan(10)).GuildName

    guerrasGanadas1 = Guilds(vecClan(1)).GetGuerrasGanadas()
    guerrasGanadas2 = Guilds(vecClan(2)).GetGuerrasGanadas()
    guerrasGanadas3 = Guilds(vecClan(3)).GetGuerrasGanadas()
    guerrasGanadas4 = Guilds(vecClan(4)).GetGuerrasGanadas()
    guerrasGanadas5 = Guilds(vecClan(5)).GetGuerrasGanadas()
    guerrasGanadas6 = Guilds(vecClan(6)).GetGuerrasGanadas()
    guerrasGanadas7 = Guilds(vecClan(7)).GetGuerrasGanadas()
    guerrasGanadas8 = Guilds(vecClan(8)).GetGuerrasGanadas()
    guerrasGanadas9 = Guilds(vecClan(9)).GetGuerrasGanadas()
    guerrasGanadas10 = Guilds(vecClan(10)).GetGuerrasGanadas()

    guerrasPerdidas1 = Guilds(vecClan(1)).GetGuerrasPerdidas()
    guerrasPerdidas2 = Guilds(vecClan(2)).GetGuerrasPerdidas()
    guerrasPerdidas3 = Guilds(vecClan(3)).GetGuerrasPerdidas()
    guerrasPerdidas4 = Guilds(vecClan(4)).GetGuerrasPerdidas()
    guerrasPerdidas5 = Guilds(vecClan(5)).GetGuerrasPerdidas()
    guerrasPerdidas6 = Guilds(vecClan(6)).GetGuerrasPerdidas()
    guerrasPerdidas7 = Guilds(vecClan(7)).GetGuerrasPerdidas()
    guerrasPerdidas8 = Guilds(vecClan(8)).GetGuerrasPerdidas()
    guerrasPerdidas9 = Guilds(vecClan(9)).GetGuerrasPerdidas()
    guerrasPerdidas10 = Guilds(vecClan(10)).GetGuerrasPerdidas()
    
End If

'CHOTS | Castillos
enviar = "2@" & CASTILLO_NORTE_DUEÑO & "@" & CASTILLO_SUR_DUEÑO & "@" & CASTILLO_ESTE_DUEÑO & "@" & CASTILLO_OESTE_DUEÑO
frmMain.InetCastillos.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking general
enviar = "3@" & Ranking_Frags_Nick & "@" & Ranking_Frags_Cant & "@" & Ranking_Trofeos_Nick & "@" & Ranking_Trofeos_Cant & "@" & Ranking_Torneos_Nick & "@" & Ranking_Torneos_Cant
frmMain.InetRanking.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking Users
enviar = "4@" & Ranking_Level(1).nombre & "@" & Ranking_Level(1).Level & "@" & Ranking_Level(1).Exp & "@" & Ranking_Level(2).nombre & "@" & Ranking_Level(2).Level & "@" & Ranking_Level(2).Exp & "@" & Ranking_Level(3).nombre & "@" & Ranking_Level(3).Level & "@" & Ranking_Level(3).Exp & "@" & Ranking_Level(4).nombre & "@" & Ranking_Level(4).Level & "@" & Ranking_Level(4).Exp & "@" & Ranking_Level(5).nombre & "@" & Ranking_Level(5).Level & "@" & Ranking_Level(5).Exp & "@" & Ranking_Level(6).nombre & "@" & Ranking_Level(6).Level & "@" & Ranking_Level(6).Exp & "@" & Ranking_Level(7).nombre & "@" & Ranking_Level(7).Level & "@" & Ranking_Level(7).Exp & "@" & Ranking_Level(8).nombre & "@" & Ranking_Level(8).Level & "@" & Ranking_Level(8).Exp & "@" & Ranking_Level(9).nombre & "@" & Ranking_Level(9).Level & "@" & Ranking_Level(9).Exp & "@" & Ranking_Level(10).nombre & "@" & Ranking_Level(10).Level & "@" & Ranking_Level(10).Exp
frmMain.InetRankingUsers.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking Clanes
enviar = "5@" & clan1 & "@" & punto1 & "@" & clan2 & "@" & punto2 & "@" & clan3 & "@" & punto3 & "@" & clan4 & "@" & punto4 & "@" & clan5 & "@" & punto5 & "@" & clan6 & "@" & punto6 & "@" & clan7 & "@" & punto7 & "@" & clan8 & "@" & punto8 & "@" & clan9 & "@" & punto9 & "@" & clan10 & "@" & punto10
frmMain.InetClanes.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking Guerras
enviar = "6@" & guerra1 & "@" & guerrasGanadas1 & "@" & guerrasPerdidas1 & "@" & guerra2 & "@" & guerrasGanadas2 & "@" & guerrasPerdidas2 & "@" & guerra3 & "@" & guerrasGanadas3 & "@" & guerrasPerdidas3 & "@" & guerra4 & "@" & guerrasGanadas4 & "@" & guerrasPerdidas4 & "@" & guerra5 & "@" & guerrasGanadas5 & "@" & guerrasPerdidas5 & "@" & guerra6 & "@" & guerrasGanadas6 & "@" & guerrasPerdidas6 & "@" & guerra7 & "@" & guerrasGanadas7 & "@" & guerrasPerdidas7 & "@" & guerra8 & "@" & guerrasGanadas8 & "@" & guerrasPerdidas8 & "@" & guerra9 & "@" & guerrasGanadas9 & "@" & guerrasPerdidas9 & "@" & guerra10 & "@" & guerrasGanadas10 & "@" & guerrasPerdidas10 & "@"
frmMain.InetGuerras.Execute baseUrl & enviar, "GET"

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Web Actualizada..." & FONTTYPE_SERVER)

End Sub
