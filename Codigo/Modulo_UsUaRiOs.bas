Attribute VB_Name = "UsUaRiOs"
'Lapsus AO 2009
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer

DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp
If UserList(AttackerIndex).Stats.Exp > MAXEXP Then _
    UserList(AttackerIndex).Stats.Exp = MAXEXP

'CHOTS | Optimizado
Call SendData(SendTarget.ToIndex, AttackerIndex, 0, "MAÑ" & UserList(VictimIndex).Name & "," & DaExp)

Call SendData(SendTarget.ToIndex, VictimIndex, 0, ServerPackages.dialogo & UserList(AttackerIndex).Name & " te ha matado!" & FONTTYPE_FIGHT)

If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
    If (Not Criminal(VictimIndex)) Then
        If UserList(AttackerIndex).flags.enTorneoAuto = False And Not EstaEnCastillo(AttackerIndex) Then
            UserList(AttackerIndex).Reputacion.AsesinoRep = UserList(AttackerIndex).Reputacion.AsesinoRep + vlASESINO * 2
            If UserList(AttackerIndex).Reputacion.AsesinoRep > MAXREP Then _
               UserList(AttackerIndex).Reputacion.AsesinoRep = MAXREP
            UserList(AttackerIndex).Reputacion.BurguesRep = 0
            UserList(AttackerIndex).Reputacion.NobleRep = 0
            UserList(AttackerIndex).Reputacion.PlebeRep = 0
        End If
    Else
         UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble
         If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.NobleRep = MAXREP
    End If
End If

'CHOTS | Torneos Automáticos
If UserList(AttackerIndex).flags.enTorneoAuto = True Then
    Call ganaUsuario(AttackerIndex)
Else
    Call UserDie(VictimIndex)
End If
'CHOTS | Torneos Automáticos

'CHOTS | Guerras
If UserList(AttackerIndex).guerra.enGuerra = True And UserList(VictimIndex).guerra.enGuerra = True Then
    Call MuereUserGuerra(AttackerIndex, VictimIndex)
End If
'CHOTS | Guerras


If UserList(AttackerIndex).Stats.UsuariosMatados < 32000 Then _
    UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1


Call ActualizarRanking(AttackerIndex, 2) 'CHOTS | Sistema de Ranking

End Sub

Sub QuitarCiuda(ByVal UserIndex As Integer)
'CHOTS | El secuas te quita el ciuda
Dim Soborno As Long
Soborno = 100000

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & "Vete de aquí!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados = 0 Then
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & "Tu no tienes ciudadanos matados." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.GLD < Soborno Then
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & "Tu no tienes " & Soborno & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.CiudadanosMatados = UserList(UserIndex).Faccion.CiudadanosMatados - 1
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Soborno

Call EnviarOro(UserIndex)
Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & FONTTYPE_TALK & "°" & "Ahora tienes un ciudadano matado menos!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    
End Sub
Sub DioResu(ByVal UserIndex As Integer)
'CHOTS | Tiro el hechizo Resu
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 6 Then
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP / 2
    UserList(UserIndex).Stats.MinHam = 0
    UserList(UserIndex).Stats.MinAGU = 0
    UserList(UserIndex).flags.Hambre = 1
    UserList(UserIndex).flags.Sed = 1
    Call EnviarHP(UserIndex)
    Call EnviarhambreYsed(UserIndex)
End If
End Sub

Sub Resucitar(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
UserList(UserIndex).Stats.MinSta = 0

'No puede estar empollando
UserList(UserIndex).flags.EstaEmpo = 0
UserList(UserIndex).EmpoCont = 0

Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
Call SendUserStatsBox(UserIndex)

End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 35
UserList(UserIndex).Stats.MinMAN = 0
UserList(UserIndex).Stats.MinSta = 0
UserList(UserIndex).Stats.MinHam = 0
UserList(UserIndex).Stats.MinAGU = 0
UserList(UserIndex).flags.Hambre = 1
UserList(UserIndex).flags.Sed = 1

'No puede estar empollando
UserList(UserIndex).flags.EstaEmpo = 0
UserList(UserIndex).EmpoCont = 0

If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
Call SendUserStatsBox(UserIndex)
Call EnviarhambreYsed(UserIndex)

End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
                    ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    UserList(UserIndex).char.Body = Body
    UserList(UserIndex).char.Head = Head
    UserList(UserIndex).char.Heading = Heading
    UserList(UserIndex).char.WeaponAnim = Arma
    UserList(UserIndex).char.ShieldAnim = Escudo
    UserList(UserIndex).char.CascoAnim = Casco
    
    'CHOTS | Monturas
    If UserList(UserIndex).flags.montandoMascota = 1 Then
        Arma = NingunArma
        Escudo = NingunEscudo
    End If
    'CHOTS | Monturas
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(UserIndex, "CP" & UserList(UserIndex).char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).char.FX & "," & UserList(UserIndex).char.loops & "," & Casco)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).char.FX & "," & UserList(UserIndex).char.loops & "," & Casco)
    End If
End Sub

'CHOTS | Full estadisticas
Public Sub EnviarFullEstadisticas(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim cad As String
    ' CHOTS | Atrib, Fama, Skills, Stats

    With UserList(UserIndex)
        'ATR
        For i = 1 To NUMATRIBUTOS
            cad = cad & .Stats.UserAtributos(i) & ","
        Next i

        'FAMA
        cad = cad & .Reputacion.AsesinoRep & ","
        cad = cad & .Reputacion.BandidoRep & ","
        cad = cad & .Reputacion.BurguesRep & ","
        cad = cad & .Reputacion.LadronesRep & ","
        cad = cad & .Reputacion.NobleRep & ","
        cad = cad & .Reputacion.PlebeRep & ","

        Dim L As Long

        L = (-.Reputacion.AsesinoRep) + _
            (-.Reputacion.BandidoRep) + _
            .Reputacion.BurguesRep + _
            (-.Reputacion.LadronesRep) + _
            .Reputacion.NobleRep + _
            .Reputacion.PlebeRep
        L = L / 6

        .Reputacion.Promedio = L

        cad = cad & .Reputacion.Promedio & ","

        'ESKILS
        For i = 1 To NUMSKILLS
           cad = cad & .Stats.UserSkills(i) & ","
        Next i

        'MEST
        cad = cad & .Faccion.CiudadanosMatados & "," & _
            .Faccion.CriminalesMatados & "," & .Stats.UsuariosMatados & "," & _
            .Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena

    End With

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "XEST" & cad)
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).char.CharIndex) = 0
    
    If UserList(UserIndex).char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
Dim code As String
    code = str(UserList(UserIndex).char.CharIndex)
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén en el mismo mapa
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(UserIndex, ServerPackages.borrarChar & code)
        Call QuitarUser(UserIndex, UserList(UserIndex).Pos.Map)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, ServerPackages.borrarChar & code)
    End If
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.number & ": " & Err.Description)

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(Map, X, Y) Then
        'If needed make a new character in list
        If UserList(UserIndex).char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
        End If
        
        'Place character on map
        MapData(Map, X, Y).UserIndex = UserIndex
        
        'Send make character command to clients
        Dim klan As String
        If UserList(UserIndex).GuildIndex > 0 Then
            klan = Guilds(UserList(UserIndex).GuildIndex).GuildName
        End If
        
        Dim bCr As Byte
        Dim SendPrivilegios As Byte
       
        bCr = Criminal(UserIndex)

        If klan <> "" Then
            If sndRoute = SendTarget.ToIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearChar & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan & ">" & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                        Else
                            'Hide the name and clan
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearChar & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & ",," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                        End If
                    Else
                        Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearChar & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 4, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
                Else
#End If
Dim code As String

                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            code = UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan & ">" & "," & bCr & "," & UserList(UserIndex).flags.Privilegios
                            Call SendData(sndRoute, sndIndex, sndMap, ServerPackages.crearChar & code) 'mandamos el CC encriptado
                        Else
                            'Hide the name and clan
                            code = UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & ",," & bCr & "," & UserList(UserIndex).flags.Privilegios
                            Call SendData(sndRoute, sndIndex, sndMap, ServerPackages.crearChar & code)
                        End If
                    Else
                        code = UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan & ">" & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 5, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0))
                        Call SendData(sndRoute, sndIndex, sndMap, ServerPackages.crearChar & code)
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map)
                Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
            End If
        Else 'if tiene clan
            If sndRoute = SendTarget.ToIndex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearNpc & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                        Else
                            'Hide the name
                            Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearNpc & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & ",," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                        End If
                    Else
                        Call SendCryptedData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearNpc & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 5, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
                Else
#End If
                    If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                        If UserList(UserIndex).showName Then
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearNpc & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                        Else
                            Call SendData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearNpc & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & ",," & bCr & "," & UserList(UserIndex).flags.Privilegios)
                        End If
                    Else
                        Call SendData(SendTarget.ToIndex, sndIndex, sndMap, ServerPackages.crearNpc & UserList(UserIndex).char.Body & "," & UserList(UserIndex).char.Head & "," & UserList(UserIndex).char.Heading & "," & UserList(UserIndex).char.CharIndex & "," & X & "," & Y & "," & UserList(UserIndex).char.WeaponAnim & "," & UserList(UserIndex).char.ShieldAnim & "," & UserList(UserIndex).char.FX & "," & 999 & "," & UserList(UserIndex).char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & IIf(UserList(UserIndex).flags.PertAlCons = 1, 5, IIf(UserList(UserIndex).flags.PertAlConsCaos = 1, 6, 0)))
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map)
                Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
            End If
       End If   'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
'CHOTS | Optimizado para Lapsus AO 2.1
'11 de Noviembre de 2010
'(en dos días me voy a ver al Indio :D)
'LEVskills,sta,mana,hp,hit
Dim Aenviar As String


On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim WasNewbie As Boolean
Dim envioExp As Boolean
envioExp = False

AumentoHIT = 0
AumentoMANA = 0
AumentoSTA = 0

'¿Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(UserIndex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU
    Aenviar = "LEV"
    If UserList(UserIndex).Stats.ELV = 1 Then
        Pts = 10
    Else
        Pts = 5
    End If
    
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts

    Aenviar = Aenviar & Pts & "@"
    
    If UserList(UserIndex).Stats.ELV <= STAT_MAXELV Then UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp - UserList(UserIndex).Stats.ELU
    
    If Not EsNewbie(UserIndex) And WasNewbie Then
        Call QuitarNewbieObj(UserIndex)
        If UCase$(MapInfo(UserList(UserIndex).Pos.Map).Restringir) = "SI" Then
            Call WarpUserChar(UserIndex, 1, 58, 45, True)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z90")
        End If
    End If
    

    'LAPSUS 2.0
    'If UserList(UserIndex).Stats.ELV < 11 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    'ElseIf UserList(UserIndex).Stats.ELV < 25 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 36 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    'ElseIf UserList(UserIndex).Stats.ELV < 46 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 50 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.4
    'Else
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    'End If
    
    
    'LAPSUS 2.1
    'If UserList(UserIndex).Stats.ELV < 11 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    'ElseIf UserList(UserIndex).Stats.ELV < 25 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 36 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    'ElseIf UserList(UserIndex).Stats.ELV < 50 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'Else
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    'End If
    
    'LAPSUS 3
    'If UserList(UserIndex).Stats.ELV < 10 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    'ElseIf UserList(UserIndex).Stats.ELV < 25 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 36 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    'ElseIf UserList(UserIndex).Stats.ELV < 50 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 54 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.6
    'ElseIf UserList(UserIndex).Stats.ELV <= 56 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 2.2
    'End If

    'LAPSUS 2017
    'If UserList(UserIndex).Stats.ELV < 10 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    'ElseIf UserList(UserIndex).Stats.ELV < 25 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 36 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    'ElseIf UserList(UserIndex).Stats.ELV < 46 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    'ElseIf UserList(UserIndex).Stats.ELV < 50 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.6
    'ElseIf UserList(UserIndex).Stats.ELV <= 54 Then
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.9
    'End If

    'LAPSUS 2017 OFI
    If UserList(UserIndex).Stats.ELV < 10 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.5
    ElseIf UserList(UserIndex).Stats.ELV < 25 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    ElseIf UserList(UserIndex).Stats.ELV < 36 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    ElseIf UserList(UserIndex).Stats.ELV < 41 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    ElseIf UserList(UserIndex).Stats.ELV < 46 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.4
    ElseIf UserList(UserIndex).Stats.ELV < 50 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.6
    ElseIf UserList(UserIndex).Stats.ELV <= 54 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.9
    End If

    Dim AumentoHP As Integer
    
    'CHOTS | Indices de vidas (con dados 18)
    '22 = Orco
    '21 = Enano
    '20 = Humano
    '19 = Elfo / Drow
    '18 = Gnomo
    'CHOTS | Indices de vidas (con dados 18)
    
    Select Case UCase$(UserList(UserIndex).Clase)
        Case "GUERRERO"
            Select Case UserList(UserIndex).Stats.UserAtributos(Constitucion)
                Case 22
                    AumentoHP = RandomNumber(11, 13)
                Case 21
                    AumentoHP = RandomNumber(10, 13)
                Case 20
                    AumentoHP = RandomNumber(9, 12)
                Case 19
                    AumentoHP = RandomNumber(9, 11)
                Case 18
                    AumentoHP = RandomNumber(8, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "CAZADOR"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(10, 13)
                Case 21
                    AumentoHP = RandomNumber(9, 13)
                Case 20
                    AumentoHP = RandomNumber(8, 12)
                Case 19
                    AumentoHP = RandomNumber(8, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select

            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "PALADIN"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(10, 12)
                Case 21
                    AumentoHP = RandomNumber(9, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 12)
                Case 19
                    AumentoHP = RandomNumber(8, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 11)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
            
        Case "DRUIDA"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(8, 11)
                Case 21
                    AumentoHP = RandomNumber(8, 10)
                Case 20
                    AumentoHP = RandomNumber(7, 9)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(6, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 2)
            AumentoMANA = Round((2.3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)), 0)
            AumentoSTA = AumentoSTMago
            
        Case "MAGO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(8, 10)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 9)
                Case 19
                    AumentoHP = RandomNumber(6, 8)
                Case 18
                    AumentoHP = RandomNumber(5, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case "CLERIGO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(9, 11)
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 10)
                Case 19
                    AumentoHP = RandomNumber(7, 9)
                Case 18
                    AumentoHP = RandomNumber(6, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "ASESINO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(9, 11)
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 10)
                Case 19
                    AumentoHP = RandomNumber(7, 9)
                Case 18
                    AumentoHP = RandomNumber(6, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            End Select
            
            AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "BARDO"
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(8, 11)
                Case 21
                    AumentoHP = RandomNumber(8, 10)
                Case 20
                    AumentoHP = RandomNumber(7, 9)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(6, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case Else
            Select Case UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = 9
                Case 21
                    AumentoHP = 8
                Case 20
                    AumentoHP = 7
                Case 19
                    AumentoHP = 6
                Case 18
                    AumentoHP = 5
                Case Else
                    AumentoHP = RandomNumber(5, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select

            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + AumentoHP
    If UserList(UserIndex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(UserIndex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(UserIndex).Stats.MaxSta = UserList(UserIndex).Stats.MaxSta + AumentoSTA
    If UserList(UserIndex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(UserIndex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + AumentoMANA

    If UserList(UserIndex).Stats.MaxMAN > STAT_MAXMAN Then _
        UserList(UserIndex).Stats.MaxMAN = STAT_MAXMAN
    
    'Actualizamos Golpe Máximo
    UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT + AumentoHIT
    If UserList(UserIndex).Stats.ELV < 36 Then
        If UserList(UserIndex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(UserIndex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(UserIndex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(UserIndex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT + AumentoHIT
    If UserList(UserIndex).Stats.ELV < 36 Then
        If UserList(UserIndex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(UserIndex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(UserIndex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(UserIndex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    Aenviar = Aenviar & AumentoSTA & "@" & AumentoMANA & "@" & AumentoHP & "@" & AumentoHIT & "@" & UserList(UserIndex).Stats.Exp & "@" & UserList(UserIndex).Stats.ELU
    
    Call LogDesarrollo(Date & " " & UserList(UserIndex).Name & " paso a nivel " & UserList(UserIndex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, Aenviar)
    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_NIVEL)
    Call EnviarExp(UserIndex)
    envioExp = True
    
Loop

If Not envioExp Then Call EnviarExp(UserIndex)

Exit Sub


errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

Dim nPos As WorldPos
    
    nPos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(UserIndex)) Then
        If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...

            Call SendToUserAreaButindex(UserIndex, ServerPackages.moverChar & UserList(UserIndex).char.CharIndex & "," & nPos.X & "," & nPos.Y)
        End If
        
        'Update map and user pos
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).char.Heading = nHeading
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex

        Call DoTileEvents(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.updatePos & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)

    UserList(UserIndex).Invent.Object(Slot) = Object
    
    If Object.ObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef & "," _
        & ObjData(Object.ObjIndex).Valor \ 3)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & ",0")
    End If

End Sub
Sub ChangeUserInvConecta(UserIndex As Integer, Slot As Byte, Object As UserOBJ)
'CHOTS | Optimizado, solo envía los items que tiene, si no tiene nada no envía nada
    UserList(UserIndex).Invent.Object(Slot) = Object
    
    If Object.ObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef & "," _
        & ObjData(Object.ObjIndex).Valor \ 3)
    Else
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "CSI" & Slot & ",0")
    End If

End Sub


Function NextOpenCharIndex() As Integer
'Modificada por el oso para codificar los MP1234,2,1 en 2 bytes
'para lograrlo, el charindex no puede tener su bit numero 6 (desde 0) en 1
'y tampoco puede ser un charindex que tenga el bit 0 en 1.

On Local Error GoTo hayerror

Dim LoopC As Integer
    
    LoopC = 1
    
    While LoopC < MAXCHARS
        If CharList(LoopC) = 0 And Not ((LoopC And &HFFC0&) = 64) Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        Else
            LoopC = LoopC + 1
        End If
    Wend

Exit Function
hayerror:
LogError ("NextOpenCharIndex: num: " & Err.number & " desc: " & Err.Description)

End Function
Function NextOpenUser() As Integer
    Dim LoopC As Long
       
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
       
    NextOpenUser = LoopC
End Function
Sub SendUserArma(ByVal UserIndex As Integer)
Dim CHOTSminArma As Integer
Dim CHOTSmaxArma As Integer

If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    CHOTSminArma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT
    CHOTSmaxArma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT
Else
    CHOTSminArma = "0"
    CHOTSmaxArma = "0"
End If

Call SendData(ToIndex, UserIndex, 0, "CHA" & CHOTSminArma & "," & CHOTSmaxArma)

End Sub
Sub SendUserEscu(ByVal UserIndex As Integer)
Dim CHOTSminEscu As Integer
Dim CHOTSmaxEscu As Integer

If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    CHOTSminEscu = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef
    CHOTSmaxEscu = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef
Else
    CHOTSminEscu = "0"
    CHOTSmaxEscu = "0"
End If

Call SendData(ToIndex, UserIndex, 0, "CHE" & CHOTSminEscu & "," & CHOTSmaxEscu)

End Sub
Sub SendUserCasco(ByVal UserIndex As Integer)

Dim CHOTSminCasco As Integer
Dim CHOTSmaxCasco As Integer

If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    CHOTSminCasco = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef
    CHOTSmaxCasco = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef
Else
    CHOTSminCasco = "0"
    CHOTSmaxCasco = "0"
End If

Call SendData(ToIndex, UserIndex, 0, "CHC" & CHOTSminCasco & "," & CHOTSmaxCasco)

End Sub
Sub SendUserDefMag(ByVal UserIndex As Integer)
Dim CHOTSminMag As Integer
Dim CHOTSmaxMag As Integer

If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    CHOTSminMag = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin
    CHOTSmaxMag = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax
Else
    CHOTSminMag = "0"
    CHOTSmaxMag = "0"
End If

If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    CHOTSminMag = CHOTSminMag + ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin
    CHOTSmaxMag = CHOTSmaxMag + ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax
End If

If UserList(UserIndex).flags.montandoMascota = 1 Then
    CHOTSminMag = CHOTSminMag + ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinMag
    CHOTSmaxMag = CHOTSmaxMag + ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxMag
End If

Call SendData(ToIndex, UserIndex, 0, "CHD" & CHOTSminMag & "," & CHOTSmaxMag)

End Sub
Sub SendUserRopa(ByVal UserIndex As Integer)
Dim CHOTSminArmor As Integer
Dim CHOTSmaxArmor As Integer

If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    CHOTSminArmor = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef
    CHOTSmaxArmor = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef
Else
    CHOTSminArmor = "0"
    CHOTSmaxArmor = "0"
End If

If UserList(UserIndex).flags.montandoMascota = 1 Then
    CHOTSminArmor = CHOTSminArmor + ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinDefensa
    CHOTSmaxArmor = CHOTSmaxArmor + ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxDefensa
End If

Call SendData(ToIndex, UserIndex, 0, "CHV" & CHOTSminArmor & "," & CHOTSmaxArmor)

End Sub
Sub SendUserHitBoxMuerto(ByVal UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "ARX")
End Sub
Sub EnviarDopa(ByVal UserIndex As Integer)
Dim Amarilla As Byte
Dim Verde As Byte
Verde = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
Amarilla = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
Call SendData(SendTarget.ToIndex, UserIndex, 0, "DRR" & Amarilla & "," & Verde)
End Sub
Sub SendUserStatsBox(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "EXT" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp)
End Sub
Sub SendUserConecta(ByVal UserIndex As Integer)

Dim Amarilla As Byte
Dim Verde As Byte
Verde = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
Amarilla = val(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

Dim lagaminarma As Integer
Dim lagamaxarma As Integer
Dim lagaminarmor As Integer
Dim lagamaxarmor As Integer
Dim lagaminescu As Integer
Dim lagamaxescu As Integer
Dim lagamincasc As Integer
Dim lagamaxcasc As Integer
Dim CHOTSminMag As Integer
Dim CHOTSmaxMag As Integer
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
lagaminarma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT
lagamaxarma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT
Else
lagaminarma = "0"
lagamaxarma = "0"
End If
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
lagaminarmor = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef
lagamaxarmor = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef
CHOTSminMag = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin
CHOTSmaxMag = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax
Else
lagaminarmor = "0"
lagamaxarmor = "0"
CHOTSminMag = "0"
CHOTSmaxMag = "0"
End If
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
lagaminescu = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef
lagamaxescu = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef
Else
lagaminescu = "0"
lagamaxescu = "0"
End If
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
lagamincasc = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef
lagamaxcasc = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef
CHOTSminMag = CHOTSminMag + ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin
CHOTSmaxMag = CHOTSmaxMag + ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax
Else
lagamincasc = "0"
lagamaxcasc = "0"
CHOTSminMag = CHOTSminMag
CHOTSmaxMag = CHOTSmaxMag
End If

Call SendData(SendTarget.ToIndex, UserIndex, 0, "CNC" & UserList(UserIndex).Stats.MaxHP & "@" & UserList(UserIndex).Stats.MinHP & "@" & UserList(UserIndex).Stats.MaxMAN & "@" & UserList(UserIndex).Stats.MinMAN & "@" & UserList(UserIndex).Stats.MaxSta & "@" & UserList(UserIndex).Stats.MinSta & "@" & UserList(UserIndex).Stats.GLD & "@" & UserList(UserIndex).Stats.ELV & "@" & UserList(UserIndex).Stats.ELU & "@" & UserList(UserIndex).Stats.Exp & "@" _
                & lagaminarma & "@" & lagamaxarma & "@" & lagaminarmor & "@" & lagamaxarmor & "@" & lagaminescu & "@" & lagamaxescu & "@" & lagamincasc & "@" & lagamaxcasc & "@" & CHOTSminMag & "@" & CHOTSmaxMag & "@" _
                & UserList(UserIndex).Stats.MinAGU & "@" & UserList(UserIndex).Stats.MinHam & "@" _
                & Amarilla & "@" & Verde & "@" _
                & CASTILLO_NORTE_DUEÑO & "@" & CASTILLO_SUR_DUEÑO & "@" & CASTILLO_ESTE_DUEÑO & "@" & CASTILLO_OESTE_DUEÑO & "@" & UserList(UserIndex).Stats.SkillPts)
End Sub
Sub EnviarMuereVive(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "MUE" & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MinSta)
End Sub
Sub EnviarHP(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "VHP" & UserList(UserIndex).Stats.MinHP)
'CHOTS | Espiando al user
If Espia_Espiador <> 0 And UserIndex = Espia_Espiado Then Call SendData(SendTarget.ToIndex, Espia_Espiador, 0, "VHÑ" & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxHP)
End Sub
Sub EnviarPocionRoja(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "VHR")
'CHOTS | Espiando al user
If Espia_Espiador <> 0 And UserIndex = Espia_Espiado Then Call SendData(SendTarget.ToIndex, Espia_Espiador, 0, "VHÑ" & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxHP)
End Sub
Sub EnviarMn(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "MN" & UserList(UserIndex).Stats.MinMAN)
'CHOTS | Espiando al user
If Espia_Espiador <> 0 And UserIndex = Espia_Espiado Then Call SendData(SendTarget.ToIndex, Espia_Espiador, 0, "MÑ" & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxMAN)
End Sub
Sub EnviarPocionAzul(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "MB")
'CHOTS | Espiando al user
If Espia_Espiador <> 0 And UserIndex = Espia_Espiado Then Call SendData(SendTarget.ToIndex, Espia_Espiador, 0, "MÑ" & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxMAN)
End Sub
Sub EnviarSta(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "STT" & UserList(UserIndex).Stats.MinSta)
End Sub
Sub EnviarOro(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "OLD" & UserList(UserIndex).Stats.GLD)
End Sub
Sub EnviarExp(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "ESP" & UserList(UserIndex).Stats.Exp)
End Sub
Sub EnviarhambreYsed(ByVal UserIndex As Integer)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
Dim GuildI As Integer


    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Clase: " & UserList(UserIndex).Clase & FONTTYPE_INFO)
    
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    GuildI = UserList(UserIndex).GuildIndex
    If GuildI > 0 Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Clan: " & Guilds(GuildI).GuildName & FONTTYPE_INFO)
        If UCase$(Guilds(GuildI).GetLeader) = UCase$(UserList(sendIndex).Name) Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Status: Lider" & FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
        'Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Trofeos Ganados: " & UserList(UserIndex).Stats.TrofOro & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Torneos Ganados: " & UserList(UserIndex).Stats.TorneosAuto & FONTTYPE_INFO)


End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
With UserList(UserIndex)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Pj: " & .Name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Clase: " & .Clase & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Pj: " & CharName & FONTTYPE_INFO)
        ' 3 en uno :p
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Ban: " & Ban & FONTTYPE_INFO)
        If Ban = "1" Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "El pj no existe: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next

    Dim j As Long
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
            End If
        Next j
    Else
        Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Usuario inexistente: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " SkillLibres:" & UserList(UserIndex).Stats.SkillPts & FONTTYPE_INFO)
End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
nombre = UCase$(nombre)

Do Until UCase$(UserList(LoopC).Name) = nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(SendTarget.ToIndex, Npclist(NpcIndex).MaestroUser, 0, ServerPackages.dialogo & "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)


'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

If Npclist(NpcIndex).Numero = ReyNpcN Then Exit Sub  'CHOTS | Sistema de Castillos


If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            UserList(UserIndex).Reputacion.NobleRep = 0
            UserList(UserIndex).Reputacion.PlebeRep = 0
            UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 200
            If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
       Else
            If Not Npclist(NpcIndex).MaestroUser > 0 Then   'mascotas nooo!
                UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + vlASALTO
                If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.BandidoRep = MAXREP
            End If
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2
       If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP
    End If
    
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
    
End If

End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UCase$(UserList(UserIndex).Clase) = "ASESINO") And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
'CHOTS | Optimizado
Dim Aumenta As Integer
Aumenta = RandomNumber(1, 100)

Dim lvl As Integer
lvl = UserList(UserIndex).Stats.ELV
    
If lvl > UBound(LevelSkill) Or Aumenta < 40 Or UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
If UserList(UserIndex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
    UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
    'Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
    Call SendData(ToIndex, UserIndex, 0, "SKI" & Skill & "," & UserList(UserIndex).Stats.UserSkills(Skill))
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + 100
    Call CheckUserLevel(UserIndex)
End If

End Sub

Sub UserDie(ByVal UserIndex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UCase$(UserList(UserIndex).Genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    If UserList(UserIndex).flags.montandoMascota = 1 Then
        Call DesMascotar(UserIndex)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).char.CharIndex)
    
    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1
    
    
    Dim aN As Integer
    
    aN = UserList(UserIndex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""
    End If
    
    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
        
    End If
    
    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")
    End If
    
    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
    End If
    
    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
    End If
    
    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.Invisible = 0
        'no hace falta encriptar este NOVER
        Dim ChotsNover As String
        ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
        'ChotsNover = Encriptar(ChotsNover)
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
    End If
    
    If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
            Call TirarTodo(UserIndex)
        End If
    End If
    
    Call TirarBandera(UserIndex)
    
    'CHOTS | Seguro Resu
    If UserList(UserIndex).flags.SeguroResu = False And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 6 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGRON")
        UserList(UserIndex).flags.SeguroResu = True
    End If
    'CHOTS | Seguro Resu
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
    End If
    'desequipar municiones
    If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(UserIndex).char.loops = LoopAdEternum Then
        UserList(UserIndex).char.FX = 0
        UserList(UserIndex).char.loops = 0
    End If
    
    'BYSNACK | Retos
    If UserList(UserIndex).Reto.enReto = True Then
        Select Case UserList(UserIndex).Reto.TipoReto
            Case 1
                Call EndReto(UserIndex)
        
            Case 2
                If UserList(NameIndex(UserList(UserIndex).Reto.Pareja)).flags.Muerto = 1 Then Call EndReto(UserIndex)

        End Select
    End If
    'BYSNACK | Retos
    
    ' << Restauramos el mimetismo
    'If UserList(UserIndex).flags.Mimetizado = 1 Then
    '    UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
    '    UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
    '    UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    '    UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    '    UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    '    UserList(UserIndex).Counters.Mimetismo = 0
    '    UserList(UserIndex).flags.Mimetizado = 0
    'End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).char.Body = iCuerpoMuerto
        UserList(UserIndex).char.Head = iCabezaMuerto
        UserList(UserIndex).char.ShieldAnim = NingunEscudo
        UserList(UserIndex).char.WeaponAnim = NingunArma
        UserList(UserIndex).char.CascoAnim = NingunCasco
    Else
        UserList(UserIndex).char.Body = iFragataFantasmal ';)
    End If
    
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
                    UserList(UserIndex).MascotasIndex(i) = 0
                    UserList(UserIndex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(UserIndex).NroMacotas = 0
    
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, NingunArma, NingunEscudo, NingunCasco)
    Call EnviarMuereVive(UserIndex)
    Call SendUserHitBoxMuerto(UserIndex)
    
    '<<Castigos por party>>
    If UserList(UserIndex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(UserIndex, UserList(UserIndex).Stats.ELV * -100 * mdParty.CantMiembros(UserIndex), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    End If
    

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.number & " Descripción: " & Err.Description)
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    If UserList(Muerto).flags.enTorneoAuto Then Exit Sub 'CHOTS | Torneos automáticos
    If UserList(Muerto).guerra.enGuerra Then Exit Sub 'CHOTS | Guerras
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
            Call LogAsesinato(UserList(Atacante).Name & " asesinó a " & UserList(Muerto).Name)
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).Name
            If UserList(Atacante).Faccion.CriminalesMatados < 65000 Then _
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = 0
            UserList(Atacante).Faccion.Amatar = 0
        End If
        
    Else
        'CHOTS | Armadas no suman frag
        If UserList(Atacante).Faccion.ArmadaReal = 0 Then
            If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
                UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).Name
                Call LogAsesinato(UserList(Atacante).Name & " asesinó a " & UserList(Muerto).Name)
                If UserList(Atacante).Faccion.CiudadanosMatados < 65000 Then _
                    UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
            End If
        End If
        'CHOTS | Armadas no suman frag
        
        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CiudadanosMatados = 0
            UserList(Atacante).Faccion.Amatar = 0
        End If
    End If
    
    'CHOTS | Sistema de GuildPoints
    If UserList(Muerto).GuildIndex > 0 And UserList(Atacante).GuildIndex > 0 Then
        If Guilds(UserList(Muerto).GuildIndex).GetGuildPoints > 0 Then
            Call Guilds(UserList(Atacante).GuildIndex).IncrementarGuildPoints(1)
            Call Guilds(UserList(Muerto).GuildIndex).DescontarGuildPoints(1)
        End If
    End If
    'CHOTS | Sistema de GuildPoints
    
    
    'CHOTS | Sistema de Puntos
    If UserList(Muerto).Puntos > 0 And UserList(Atacante).Puntos > 0 Then
        UserList(Atacante).Puntos = UserList(Atacante).Puntos + 1
        UserList(Muerto).Puntos = UserList(Muerto).Puntos - 1
    End If
    'CHOTS | Sistema de Puntos

End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                If LegalPos(nPos.Map, tX, tY) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = Pos.X + LoopC
                        tY = Pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    
If UserIndex = 0 Then
    Call LogError("ERROR EN WARPUSERCHAR! Destino: " & Map & "," & X & "," & Y)
    Exit Sub
End If

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

    'Quitar el dialogo
    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).char.CharIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "QTDL")
    
    OldMap = UserList(UserIndex).Pos.Map
    OldX = UserList(UserIndex).Pos.X
    OldY = UserList(UserIndex).Pos.Y
    
    Call EraseUserChar(SendTarget.ToMap, 0, OldMap, UserIndex)
        
    If OldMap <> Map Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.cargarMapa & Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(UserIndex).Pos.X = X
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.Map = Map
    
    If MapData(Map, X, Y).UserIndex <> 0 Then
       UserList(UserIndex).Pos = DamePos(UserList(UserIndex).Pos)
    End If
    
    Call MakeUserChar(SendTarget.ToMap, 0, Map, UserIndex, Map, X, Y)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).char.CharIndex)
    Call DoTileEvents(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
    'Seguis invisible al pasar de mapa
    If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Dim ChotsNover As String
        ChotsNover = UserList(UserIndex).char.CharIndex & ",1"
        'ChotsNover = Encriptar(ChotsNover)
        Call SendToUserArea(UserIndex, Nover(5) & ChotsNover, EncriptarProtocolosCriticos)
    End If
    
    If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXWARP & ",0")
    End If
    
    'CHOTS | Monturas
    If (UserList(UserIndex).flags.montandoMascota = 1) And (MapInfo(Map).Pk = False Or MapInfo(Map).zona = "DUNGEON" Or Map > 59) Then
        Call DesMascotar(UserIndex)
        Exit Sub
    End If
    'CHOTS | Monturas
    
    Call WarpMascotas(UserIndex)
End Sub

Sub UpdateUserMap(ByVal UserIndex As Integer)

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

'EnviarNoche UserIndex

On Error GoTo 0

Map = UserList(UserIndex).Pos.Map

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).UserIndex > 0 And UserIndex <> MapData(Map, X, Y).UserIndex Then
            Call MakeUserChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).UserIndex, Map, X, Y)
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).UserIndex).flags.Oculto = 1 Then Call SendCryptedData(SendTarget.ToIndex, UserIndex, 0, Nover(5) & UserList(MapData(Map, X, Y).UserIndex).char.CharIndex & ",1")
            Else
#End If
                Dim ChotsNover As String
                ChotsNover = UserList(MapData(Map, X, Y).UserIndex).char.CharIndex & ",1"
                'ChotsNover = Encriptar(ChotsNover)
                If UserList(MapData(Map, X, Y).UserIndex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).UserIndex).flags.Oculto = 1 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, Nover(5) & ChotsNover)
#If SeguridadAlkon Then
            End If
#End If
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType <> eOBJType.otArboles Then
                Call MakeObj(SendTarget.ToIndex, UserIndex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                          Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                          Call Bloquear(SendTarget.ToIndex, UserIndex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                End If
            End If
        End If
        
    Next X
Next Y

End Sub


Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(UserIndex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    
 For i = 1 To MAXMASCOTAS
If UserList(UserIndex).MascotasIndex(i) > 0 Then
             ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Or EstaEnCastillo(UserIndex) = True Then
                 Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
                 UserList(UserIndex).MascotasIndex(i) = 0
                 InvocadosMatados = InvocadosMatados + 1
                 NroPets = NroPets - 1
        End If
    End If
     Next i
    
    If InvocadosMatados > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z51")
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
            UserList(UserIndex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(UserIndex).MascotasIndex(i) = 0 Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
        End If
    Next i
    
    UserList(UserIndex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal Tiempo As Integer = -1)
    'BYSNACK | Retos
    If UserList(UserIndex).Reto.enReto = True Then
        Call UserDie(UserIndex)
        If UserList(UserList(UserIndex).Reto.Pareja).flags.Muerto = 0 Then
            Call UserDie(UserList(UserIndex).Reto.Pareja)
        End If
    End If
    'BYSNACK | Retos
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = IIf(UserList(UserIndex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList(UserIndex).Pos.Map).Pk, 0, Tiempo)
        
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Cerrando...Se cerrará el juego en " & UserList(UserIndex).Counters.Salir & " segundos..." & FONTTYPE_INFO)
    End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).Name

If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If

End Sub

Public Sub Empollando(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    UserList(UserIndex).flags.EstaEmpo = 1
Else
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0
End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)

If FileExist(CharPath & nombre & ".chr", vbArchive) = False Then
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Pj Inexistente" & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Estadisticas de: " & nombre & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Nivel: " & GetVar(CharPath & nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "elu") & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Vitalidad: " & GetVar(CharPath & nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "maxSta") & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Salud: " & GetVar(CharPath & nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxMAN") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & nombre & ".chr", "stats", "MaxHIT") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "GLD") & FONTTYPE_INFO)
End If
Exit Sub

End Sub
Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & CharName & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco." & FONTTYPE_INFO)
    Else
    Call SendData(SendTarget.ToIndex, sendIndex, 0, ServerPackages.dialogo & "Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub


