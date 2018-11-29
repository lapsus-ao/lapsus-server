Attribute VB_Name = "modCastillos"
'Modulo Castillos obtenido de Empires AO
'Adaptado y Reprogramado por CHOTS, Para LAPSUS AO 2009
'Martes, 27 de Octubre de 2009
'Adaptado y Reprogramado por CHOTS, para LAPSUS AO 2010
'Lunes, 1 de Marzo de 2010
'Agregado Castillos Faccionarios
'LapsusAO 2.0
'Reprogramado por CHOTS
'10/09/2010
'Reprogramado por CHOTS para Lapsus 3
'29/10/11
'Agregada la Fortaleza
'Reprogramado por CHOTS para Lapsus2017
'24/10/17
'Agregado dungeon Fortaleza
Option Explicit

Public CASTILLO_NORTE_DUEÑO As String
Public CASTILLO_SUR_DUEÑO As String
Public CASTILLO_ESTE_DUEÑO As String
Public CASTILLO_OESTE_DUEÑO As String

Public CASTILLO_NORTE_TIEMPO As Integer
Public CASTILLO_SUR_TIEMPO As Integer
Public CASTILLO_ESTE_TIEMPO As Integer
Public CASTILLO_OESTE_TIEMPO As Integer

Public Const CASTILLO_PREMIO_TOMAR As Integer = 50
Public Const CASTILLO_PREMIO_PERDER As Integer = 50
Public Const CASTILLO_PREMIO_MANTENER As Integer = 10
Public Const CASTILLO_PREMIO_MANTENER_PUNTOS_USUARIO As Integer = 10
 
Public Const ReyNpcN As Integer = 579

Public Const MAXDEFENSORES As Byte = 3

Public Const CastilloOeste As Integer = 20
Public Const CastilloEste As Integer = 40
Public Const CastilloSur As Integer = 22
Public Const CastilloNorte As Integer = 49
Public Const CastilloFortaleza As Integer = 50
Public Const CastilloFortalezaDungeon1 As Integer = 92
Public Const CastilloFortalezaDungeon2 As Integer = 93

Public Sub AtacandoCastillo(ByVal UserIndex As Integer, NpcIndex As Integer)
 
On Error GoTo errorh

Static avisado As Byte
Dim Castillo As Integer
Dim ClanAtaca As String
 
Castillo = 0
 
If UserList(UserIndex).Pos.Map = CastilloOeste Then Castillo = 1
If UserList(UserIndex).Pos.Map = CastilloEste Then Castillo = 2
If UserList(UserIndex).Pos.Map = CastilloSur Then Castillo = 3
If UserList(UserIndex).Pos.Map = CastilloNorte Then Castillo = 4
 
If Castillo = 0 Then Exit Sub

ClanAtaca = Guilds(UserList(UserIndex).GuildIndex).GuildName
 
avisado = avisado + 1

If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP And Npclist(NpcIndex).Stats.MinHP > 2000 Then
    If avisado >= 10 Then
        Call SendData(SendTarget.ToAll, 0, 0, "PRE20," & ClanAtaca & "," & Castillo)
        avisado = 0
    End If
 
ElseIf Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP And Npclist(NpcIndex).Stats.MinHP < 2000 And Npclist(NpcIndex).Numero = ReyNpcN Then
    If avisado >= 5 Then
        Call SendData(SendTarget.ToAll, 0, 0, "PRE22," & ClanAtaca & "," & Castillo)
        avisado = 0
    End If
 
End If

 
Exit Sub
 
errorh: Call LogError("AtacandoCastillo:" & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
 
End Sub
 

Public Sub MuereRey(ByVal UserIndex As Integer, NpcIndex As Integer)
 
Dim reNpcPos As WorldPos
Dim itemPos As WorldPos

Dim itemIndex As Obj
Dim reNpcIndex As Integer
Dim ClanTomo As String
Dim Castillo As Integer
Dim nombreExClan As String
Dim exClan As Integer
Dim num As Byte

Castillo = 0
 
If UserList(UserIndex).Pos.Map = CastilloOeste Then
    Castillo = 1
    nombreExClan = CASTILLO_OESTE_DUEÑO
End If

If UserList(UserIndex).Pos.Map = CastilloEste Then
    Castillo = 2
    nombreExClan = CASTILLO_ESTE_DUEÑO
End If

If UserList(UserIndex).Pos.Map = CastilloSur Then
    Castillo = 3
    nombreExClan = CASTILLO_SUR_DUEÑO
End If

If UserList(UserIndex).Pos.Map = CastilloNorte Then
    Castillo = 4
    nombreExClan = CASTILLO_NORTE_DUEÑO
End If

exClan = cstGuildIndex(nombreExClan)
 
If Castillo = 0 Then Exit Sub
 
reNpcPos.Map = UserList(UserIndex).Pos.Map
reNpcPos.X = 50
reNpcPos.Y = 44

itemPos.Map = UserList(UserIndex).Pos.Map
itemPos.X = 50
itemPos.Y = 45

itemIndex.Amount = 1

num = RandomNumber(1, 2)

If num = 1 Then
    itemIndex.ObjIndex = 371
Else
    itemIndex.ObjIndex = 1066
End If
 
reNpcIndex = NpcIndex
 
ClanTomo = Guilds(UserList(UserIndex).GuildIndex).GuildName
 
If Castillo = 4 Then
    CASTILLO_NORTE_DUEÑO = ClanTomo
    CASTILLO_NORTE_TIEMPO = 0
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte", ClanTomo)
    Call TirarItemAlPiso(itemPos, itemIndex)
ElseIf Castillo = 3 Then
    CASTILLO_SUR_DUEÑO = ClanTomo
    CASTILLO_SUR_TIEMPO = 0
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur", ClanTomo)
ElseIf Castillo = 2 Then
    CASTILLO_ESTE_DUEÑO = ClanTomo
    CASTILLO_ESTE_TIEMPO = 0
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este", ClanTomo)
ElseIf Castillo = 1 Then
    CASTILLO_OESTE_DUEÑO = ClanTomo
    CASTILLO_OESTE_TIEMPO = 0
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste", ClanTomo)
End If
 
Call Guilds(UserList(UserIndex).GuildIndex).IncrementarGuildPoints(CASTILLO_PREMIO_TOMAR)  'CHOTS | Incrementa GuildPoints del clan del usuario que mato al Rey
If exClan <> 0 Then Call Guilds(exClan).DescontarGuildPoints(CASTILLO_PREMIO_PERDER)  'CHOTS | Descuenta GuildPoints del clan del usuario que mato al Rey

Call QuitarNPC(NpcIndex)
 
Call SendData(SendTarget.ToAll, 0, 0, "PRE24," & ClanTomo & "," & Castillo)

Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "PRE87")
 
Call SpawnNpc(ReyNpcN, reNpcPos, True, False)
 
Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE25")

'CHOTS | Saca a los users de la fortaleza
Call SacarUsersFortaleza
 
End Sub
 

 
Public Sub SendCastellOwner(ByVal UserIndex As Integer)

Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE13," & CASTILLO_NORTE_DUEÑO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE14," & CASTILLO_OESTE_DUEÑO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE15," & CASTILLO_ESTE_DUEÑO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE16," & CASTILLO_SUR_DUEÑO)

 
End Sub
Public Sub CargarCastillos()
'CHOTS | 20/11/10
CASTILLO_OESTE_DUEÑO = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste")
CASTILLO_ESTE_DUEÑO = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este")
CASTILLO_NORTE_DUEÑO = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte")
CASTILLO_SUR_DUEÑO = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur")

CASTILLO_OESTE_TIEMPO = 0
CASTILLO_ESTE_TIEMPO = 0
CASTILLO_NORTE_TIEMPO = 0
CASTILLO_SUR_TIEMPO = 0

End Sub
 
Public Sub TelepToCasti(ByVal UserIndex As Integer, castiStrng As String)
On Error GoTo errh
Dim IsKingPos As WorldPos
 
 If UserList(UserIndex).flags.Paralizado = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE75")
    Exit Sub
 End If
 
 If UserList(UserIndex).Counters.Pena > 0 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE76")
    Exit Sub
 End If
 
If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE89")
    Exit Sub
End If

If UserList(UserIndex).flags.enTorneoAuto = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al castillo si estas en Torneo!" & FONTTYPE_TORNEOAUTO)
    Exit Sub
End If
 
 'CHOTS | Guerras
If UserList(UserIndex).guerra.enGuerra = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al castillo si estas en una guerra." & FONTTYPE_GUERRA)
    Exit Sub
End If

 
IsKingPos.Map = 0

If UCase$(castiStrng) = "NORTE" Then IsKingPos.Map = CastilloNorte
If UCase$(castiStrng) = "OESTE" Then IsKingPos.Map = CastilloOeste
If UCase$(castiStrng) = "ESTE" Then IsKingPos.Map = CastilloEste
If UCase$(castiStrng) = "SUR" Then IsKingPos.Map = CastilloSur

If IsKingPos.Map = 0 Then Exit Sub
If IsKingPos.Map = CastilloNorte And UCase$(CASTILLO_NORTE_DUEÑO) = UCase$(Guilds(UserList(UserIndex).GuildIndex).GuildName) Or _
IsKingPos.Map = CastilloEste And UCase$(CASTILLO_ESTE_DUEÑO) = UCase$(Guilds(UserList(UserIndex).GuildIndex).GuildName) Or _
IsKingPos.Map = CastilloOeste And UCase$(CASTILLO_OESTE_DUEÑO) = UCase$(Guilds(UserList(UserIndex).GuildIndex).GuildName) Or _
IsKingPos.Map = CastilloSur And UCase$(CASTILLO_SUR_DUEÑO) = UCase$(Guilds(UserList(UserIndex).GuildIndex).GuildName) Then


 If UserList(UserIndex).Pos.Map = IsKingPos.Map Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE77")
Exit Sub
 End If

IsKingPos.X = RandomNumber(46, 57)
IsKingPos.Y = RandomNumber(66, 70)
 
Do While LegalPos(IsKingPos.Map, IsKingPos.X, IsKingPos.Y, False) = False
IsKingPos.X = RandomNumber(46, 57)
IsKingPos.Y = RandomNumber(66, 70)
Loop

Call WarpUserChar(UserIndex, IsKingPos.Map, IsKingPos.X, IsKingPos.Y, True)
 
 End If
 
errh:
Debug.Print "ERROR!" & Err.number & "; " & Err.Source & "; " & Err.Description
 
End Sub
 
Public Sub DarPremioCastillos()
'CHOTS | Sistema de Castillos

Dim ClanIndex As Integer

CASTILLO_NORTE_TIEMPO = CASTILLO_NORTE_TIEMPO + 1
CASTILLO_SUR_TIEMPO = CASTILLO_SUR_TIEMPO + 1
CASTILLO_ESTE_TIEMPO = CASTILLO_ESTE_TIEMPO + 1
CASTILLO_OESTE_TIEMPO = CASTILLO_OESTE_TIEMPO + 1

If CASTILLO_NORTE_TIEMPO >= 30 Then
    ClanIndex = CHOTSGuildIndex(CASTILLO_NORTE_DUEÑO)
    If ClanIndex > 0 Then
        Call Guilds(ClanIndex).IncrementarGuildPoints(CASTILLO_PREMIO_MANTENER)
        Call SendData(SendTarget.ToGuildMembers, ClanIndex, 0, "PRE26,NORTE")
        Call repartirPuntos(ClanIndex, CASTILLO_PREMIO_MANTENER_PUNTOS_USUARIO, "Norte")
    End If
    CASTILLO_NORTE_TIEMPO = 0
End If

If CASTILLO_SUR_TIEMPO >= 30 Then
    ClanIndex = CHOTSGuildIndex(CASTILLO_SUR_DUEÑO)
    If ClanIndex > 0 Then
        Call Guilds(ClanIndex).IncrementarGuildPoints(CASTILLO_PREMIO_MANTENER)
        Call SendData(SendTarget.ToGuildMembers, ClanIndex, 0, "PRE26,SUR")
        Call repartirPuntos(ClanIndex, CASTILLO_PREMIO_MANTENER_PUNTOS_USUARIO, "Sur")
    End If
    CASTILLO_SUR_TIEMPO = 0
End If

If CASTILLO_ESTE_TIEMPO >= 30 Then
    ClanIndex = CHOTSGuildIndex(CASTILLO_ESTE_DUEÑO)
    If ClanIndex > 0 Then
        Call Guilds(ClanIndex).IncrementarGuildPoints(CASTILLO_PREMIO_MANTENER)
        Call SendData(SendTarget.ToGuildMembers, ClanIndex, 0, "PRE26,ESTE")
        Call repartirPuntos(ClanIndex, CASTILLO_PREMIO_MANTENER_PUNTOS_USUARIO, "Este")
    End If
    CASTILLO_ESTE_TIEMPO = 0
End If

If CASTILLO_OESTE_TIEMPO >= 30 Then
    ClanIndex = CHOTSGuildIndex(CASTILLO_OESTE_DUEÑO)
    If ClanIndex > 0 Then
        Call Guilds(ClanIndex).IncrementarGuildPoints(CASTILLO_PREMIO_MANTENER)
        Call SendData(SendTarget.ToGuildMembers, ClanIndex, 0, "PRE26,OESTE")
        Call repartirPuntos(ClanIndex, CASTILLO_PREMIO_MANTENER_PUNTOS_USUARIO, "Oeste")
    End If
    CASTILLO_OESTE_TIEMPO = 0
End If
 
End Sub
 
Public Function cstGuildIndex(ByVal Clan As String) As Integer

Dim LoopC As Integer
 
For LoopC = 1 To LastUser
 If UserList(LoopC).GuildIndex <> 0 Then
 If Guilds(UserList(LoopC).GuildIndex).GuildName = Clan Then
cstGuildIndex = UserList(LoopC).GuildIndex
Exit Function
 End If
 End If
Next LoopC
 
cstGuildIndex = 0
 
End Function

Public Function EstaEnCastillo(ByVal UserIndex As Integer) As Boolean

'CHOTS | Sistema de Castillos

If UserList(UserIndex).Pos.Map = CastilloNorte Or UserList(UserIndex).Pos.Map = CastilloSur Or _
UserList(UserIndex).Pos.Map = CastilloEste Or UserList(UserIndex).Pos.Map = CastilloOeste Or UserList(UserIndex).Pos.Map = CastilloFortaleza Then
    EstaEnCastillo = True
    Exit Function
End If

EstaEnCastillo = False
 
End Function

Public Function EstaEnFortaleza(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Pos.Map = CastilloFortaleza Or UserList(UserIndex).Pos.Map = CastilloFortalezaDungeon1 Or UserList(UserIndex).Pos.Map = CastilloFortalezaDungeon2 Then
    EstaEnFortaleza = True
    Exit Function
End If

EstaEnFortaleza = False
 
End Function

Public Sub SacarUsersFortaleza()
    'CHOTS | Si hay algun user en fortaleza lo saca afuera del casti norte
    Dim i As Integer

    For i = 1 To LastUser
        If EstaEnFortaleza(i) = True Then
            Call TelepUserAfueraCastiNorte(i)
        End If
    Next i
End Sub

Public Sub TelepUserAfueraCastiNorte(ByVal UserIndex As Integer)
    Dim respawnPos As WorldPos
    Dim nPos As WorldPos
    respawnPos.Map = 42
    respawnPos.X = 53
    respawnPos.Y = 16
    Call ClosestLegalPos(respawnPos, nPos)

    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
End Sub
