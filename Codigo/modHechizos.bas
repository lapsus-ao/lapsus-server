Attribute VB_Name = "modHechizos"
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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

'CHOTS | Sistema de Castillos
If Npclist(NpcIndex).NPCtype = eNPCType.reydelcastillo Then
    If UserList(UserIndex).GuildIndex = 0 Then Exit Sub

    If UserList(UserIndex).Pos.Map = CastilloOeste And CASTILLO_OESTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
        UserList(UserIndex).Pos.Map = CastilloEste And CASTILLO_ESTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
        UserList(UserIndex).Pos.Map = CastilloNorte And CASTILLO_NORTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
        UserList(UserIndex).Pos.Map = CastilloSur And CASTILLO_SUR_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
        UserList(UserIndex).Pos.Map = CastilloFortaleza Then Exit Sub
End If
'CHOTS | Sistema de Castillos

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    Call EnviarHP(UserIndex)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        
        'CHOTS | Monturas
        If UserList(UserIndex).flags.montandoMascota = 1 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinMag, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxMag)
        End If
        'CHOTS | Monturas
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)

        Call EnviarHP(UserIndex)
        
        'Muere
        If UserList(UserIndex).Stats.MinHP < 1 Then
            UserList(UserIndex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                RestarCriminalidad (UserIndex)
            End If
            Call UserDie(UserIndex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(UserIndex).flags.Paralizado = 1
          UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.updatePos & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).Pos.Map, "CXF" & Npclist(TargetNPC).char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(UserIndex).char.CharIndex
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbCyan & "°" & s & "°" & ind)
    Exit Sub
End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.X = UserList(UserIndex).flags.TargetX
    wp2.Y = UserList(UserIndex).flags.TargetY
    
    'CHOTS | Torneos Automáticos
    If UserList(UserIndex).flags.enTorneoAuto = True And UserList(UserIndex).flags.enDueloTorneoAuto = False Then
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < Hechizos(HechizoIndex).MinLevel Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Para tirar este hechizo debes ser nivel " & Hechizos(HechizoIndex).MinLevel & FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z43")
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z44")
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
    
    
    If Hechizos(HechizoIndex).NeedStaff >= 2 Then
        If UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> FLAUTAMAGICA And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> FLAUTAUNICA Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Para Lanzar este conjuro necesitas una Flauta Mágica!" & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
                
                
                
                
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z1")
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z2")
            PuedeLanzar = False
        End If
    Else 'CHOTS | Se fija las gemas
        If EstaLiberado(UserIndex) Then
            
            'CHOTS | Gemas (Druida Elfo)
            If UCase$(UserList(UserIndex).Raza) = "ELFO" And UCase$(UserList(UserIndex).Clase) = "DRUIDA" And UCase$(Hechizos(HechizoIndex).nombre) = "IMPLORAR AYUDA" Then
                If UserList(UserIndex).Stats.MinMAN < 1500 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
                    PuedeLanzar = False
                    Exit Function
                End If
            End If
            
            'CHOTS | Gemas (Mago Enano)
            If UCase$(UserList(UserIndex).Raza) = "ENANO" And UCase$(UserList(UserIndex).Clase) = "MAGO" And UCase$(Hechizos(HechizoIndex).nombre) = "APOCALIPSIS" Then
                If UserList(UserIndex).Stats.MinMAN < 1000 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
                    PuedeLanzar = False
                    Exit Function
                End If
            End If
            
            'CHOTS | Gemas (Druida Humano)
            If UCase$(UserList(UserIndex).Raza) = "HUMANO" And UCase$(UserList(UserIndex).Clase) = "DRUIDA" And UCase$(Hechizos(HechizoIndex).nombre) = "EXPLOSIÓN FAUSTICA" Then
                If UserList(UserIndex).Stats.MinMAN < 1250 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
                    PuedeLanzar = False
                    Exit Function
                End If
            End If
            
            'CHOTS | Gemas (Druida Orco)
            If UCase$(UserList(UserIndex).Raza) = "ORCO" And UCase$(UserList(UserIndex).Clase) = "DRUIDA" And UCase$(Hechizos(HechizoIndex).nombre) = "LLAMARADA" Then
                If UserList(UserIndex).Stats.MinMAN < 500 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
                    PuedeLanzar = False
                    Exit Function
                End If
            End If
            
            'CHOTS | Gemas (Bardo Enano)
            If UCase$(UserList(UserIndex).Raza) = "ENANO" And UCase$(UserList(UserIndex).Clase) = "BARDO" And UCase$(Hechizos(HechizoIndex).nombre) = "RÁFAGA ARDIENTE" Then
                If UserList(UserIndex).Stats.MinMAN < 900 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
                    PuedeLanzar = False
                    Exit Function
                End If
            End If
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
            PuedeLanzar = False
            
        Else
        
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z3")
            PuedeLanzar = False
        End If
    End If
Else
   Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z4")
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer

    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        PosCasteadaX = UserList(UserIndex).flags.TargetX
        PosCasteadaY = UserList(UserIndex).flags.TargetY
        PosCasteadaM = UserList(UserIndex).flags.TargetMap
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)
    End If

End Sub

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Or UserList(UserIndex).Pos.Map = 81 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z5")
    Exit Sub
End If


'CHOTS | Sistema de Castillos
If EstaEnCastillo(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE80")
    Exit Sub
End If
'CHOTS | Sistema de Castillos

'CHOTS | Guerras
If UserList(UserIndex).guerra.enGuerra = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes invocar mascotas en una Guerra!" & FONTTYPE_GUERRA)
    Exit Sub
End If

'CHOTS | Invocaciones
If UserList(UserIndex).Pos.Map = INVOCACION_MAPA Or UserList(UserIndex).Pos.Map = Torneo_MAPATORNEO Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z80")
    Exit Sub
End If
'CHOTS | Invocaciones

'CHOTS | Duelos
If UserList(UserIndex).Pos.Map = DUELO_MAPADUELO Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No se permite invocar en un Duelo" & FONTTYPE_DUELO)
    Exit Sub
End If
'CHOTS | Duelos

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.X = UserList(UserIndex).flags.TargetX
TargetPos.Y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).numNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(UserIndex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(UserIndex, b)
    
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call EnviarMn(UserIndex)
    Call EnviarSta(UserIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    
    If EstaLiberado(UserIndex) Then
    
        If UCase$(UserList(UserIndex).Raza) = "ELFO" And UCase$(UserList(UserIndex).Clase) = "DRUIDA" And UCase$(Hechizos(uh).nombre) = "IMPLORAR AYUDA" Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - val(1500) 'CHOTS | Gemas (Druida Elfo)
            If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        ElseIf UCase$(UserList(UserIndex).Raza) = "ENANO" And UCase$(UserList(UserIndex).Clase) = "MAGO" And UCase$(Hechizos(uh).nombre) = "APOCALIPSIS" Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - val(1000) 'CHOTS | Gemas (Mago Enano)
            If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        ElseIf UCase$(UserList(UserIndex).Raza) = "HUMANO" And UCase$(UserList(UserIndex).Clase) = "DRUIDA" And UCase$(Hechizos(uh).nombre) = "EXPLOSIÓN FAUSTICA" Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - val(1250) 'CHOTS | Gemas (Druida Humano)
            If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        ElseIf UCase$(UserList(UserIndex).Raza) = "ORCO" And UCase$(UserList(UserIndex).Clase) = "DRUIDA" And UCase$(Hechizos(uh).nombre) = "LLAMARADA" Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - val(500) 'CHOTS | Gemas (Druida Orco)
            If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        ElseIf UCase$(UserList(UserIndex).Raza) = "ENANO" And UCase$(UserList(UserIndex).Clase) = "BARDO" And UCase$(Hechizos(uh).nombre) = "RÁFAGA ARDIENTE" Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - val(900) 'CHOTS | Gemas (Bardo Enano)
            If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        Else
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
            If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        End If
    
   
    Else 'CHOTS | No está liberado
    
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    
    End If
    
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call EnviarSta(UserIndex)
    Call EnviarMn(UserIndex)
    Call EnviarHP(UserList(UserIndex).flags.TargetUser)
    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call EnviarMn(UserIndex)
    Call EnviarSta(UserIndex)
End If

End Sub


Sub lanzarHechizo(Index As Integer, UserIndex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case TargetType.uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z60")
            End If
        Case TargetType.uNPC
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z61")
            End If
        Case TargetType.uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(UserIndex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                End If
            ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(UserIndex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z26")
            End If
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If

If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)



Dim H As Integer, TU As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
TU = UserList(UserIndex).flags.TargetUser


If Hechizos(H).Invisibilidad = 1 Then
   
    If UserList(TU).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Está muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'CHOTS | Sistema de Castillos
    If EstaEnCastillo(TU) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE79")
        b = False
        Exit Sub
    End If
    'CHOTS | Sistema de Castillos
    
    'CHOTS | Guerras
    If UserList(UserIndex).guerra.enGuerra = True Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes hacerte invisible en una Guerra!" & FONTTYPE_GUERRA)
        Exit Sub
    End If
    
    'CHOTS | Duelos
    If UserList(UserIndex).Reto.enReto = True Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No se permite invisibilidad en los retos!" & FONTTYPE_DUELO)
        b = False
        Exit Sub
    End If
    
    'CHOTS | No tirar invi en torneos y en invocaciones
    If UserList(TU).Pos.Map = 81 Or UserList(TU).Pos.Map = 82 Or UserList(TU).Pos.Map = INVOCACION_MAPA Or UserList(TU).Pos.Map = INVOCACIONTELEP_MAPA Or UserList(TU).Pos.Map = Torneo_MAPATORNEO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z72")
        b = False
        Exit Sub
    End If
    'CHOTS | No tirar invi en torneos y en invocaciones
    
    
    If Criminal(TU) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
            Exit Sub
        Else
            Call VolverCriminal(UserIndex)
        End If
    End If
    
    UserList(TU).flags.Invisible = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(TU).Pos.Map, Nover(5) & UserList(TU).char.CharIndex & ",1")
    Else
#End If
        Dim ChotsNover As String
        ChotsNover = UserList(TU).char.CharIndex & ",1"
        'ChotsNover = Encriptar(ChotsNover)
        Call SendData(SendTarget.ToMap, 0, UserList(TU).Pos.Map, Nover(5) & ChotsNover)
#If SeguridadAlkon Then
    End If
#End If
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(UserIndex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Privilegios >= PlayerType.Consejero Then
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'copio el char original al mimetizado
    
    With UserList(UserIndex)
        .CharMimetizado.Body = .char.Body
        .CharMimetizado.Head = .char.Head
        .CharMimetizado.CascoAnim = .char.CascoAnim
        .CharMimetizado.ShieldAnim = .char.ShieldAnim
        .CharMimetizado.WeaponAnim = .char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .char.Body = UserList(TU).char.Body
        .char.Head = UserList(TU).char.Head
        .char.CascoAnim = UserList(TU).char.CascoAnim
        .char.ShieldAnim = UserList(TU).char.ShieldAnim
        .char.WeaponAnim = UserList(TU).char.WeaponAnim
    
        Call ChangeUserChar(SendTarget.ToMap, 0, .Pos.Map, UserIndex, .char.Body, .char.Head, .char.Heading, .char.WeaponAnim, .char.ShieldAnim, .char.CascoAnim)
    End With
   
   Call InfoHechizo(UserIndex)
   b = True
End If


If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
            
            If UserIndex = TU Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z62")
                Exit Sub
            End If
            
            If UserIndex <> TU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TU)
            End If
            
            Call InfoHechizo(UserIndex)
            b = True
            If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.ToIndex, TU, 0, ServerPackages.dialogo & " Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " ¡El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
                Call SendData(SendTarget.ToIndex, TU, 0, "DOK")
                Call SendData(SendTarget.ToIndex, TU, 0, ServerPackages.updatePos & UserList(TU).Pos.X & "," & UserList(TU).Pos.Y)

    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
        If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
                Exit Sub
            Else
                Call VolverCriminal(UserIndex)
            End If
        End If
        
        UserList(TU).flags.Paralizado = 0
        'no need to crypt this
        Call SendData(SendTarget.ToIndex, TU, 0, "DOK")
        Call InfoHechizo(UserIndex)
        b = True
    End If
End If

If Hechizos(H).RemoverEstupidez = 1 Then
    If Not UserList(TU).flags.Estupidez = 0 Then
                UserList(TU).flags.Estupidez = 0
                'no need to crypt this
                Call SendData(SendTarget.ToIndex, TU, 0, "NESTUP")
                Call InfoHechizo(UserIndex)
                b = True
    End If
End If


If Hechizos(H).Revivir = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        If Criminal(TU) And Not Criminal(UserIndex) Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
                Exit Sub
            Else
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call VolverCriminal(UserIndex)
                Else
                    Exit Sub 'CHOTS | Si es armada no tira el Resu
                End If
                
            End If
        End If
        
        If UserList(TU).flags.SeguroResu = True Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z69")
            Exit Sub
        End If

        'revisamos si necesita vara
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Necesitas un mejor báculo para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
            
        ElseIf UCase$(UserList(UserIndex).Clase) = "BARDO" Then
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> LAUDMAGICO And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> LAUDUNICO Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z71")
                b = False
                Exit Sub
            End If
            
        ElseIf UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then 'CHOTS | El Druida necesita flauta
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> FLAUTAMAGICA And UserList(UserIndex).Invent.HerramientaEqpObjIndex <> FLAUTAUNICA Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z71")
                b = False
                Exit Sub
            End If
        End If
        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> UserIndex Then
                UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + 500
                If UserList(UserIndex).Reputacion.NobleRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.NobleRep = MAXREP
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)
            End If
        End If
        UserList(TU).Stats.MinMAN = 0
        '/Pablo Toxic Waste
        
        b = True
        Call InfoHechizo(UserIndex)
        Call RevivirUsuario(TU)
        Call DioResu(UserIndex)
        If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " resucitó a " & UserList(TU).Name & " en el mapa " & UserList(TU).Pos.Map & "." & FONTTYPE_SERVER)

    Else
        b = False
    End If

End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado / 3
#If SeguridadAlkon Then
        Call SendCryptedData(SendTarget.ToIndex, TU, 0, "CEGU")
#Else
        Call SendData(SendTarget.ToIndex, TU, 0, "CEGU")
#End If
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.ToIndex, TU, 0, "DUMB")
        Else
#End If
            Call SendData(SendTarget.ToIndex, TU, 0, "DUMB")
#If SeguridadAlkon Then
        End If
#End If
        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z7")
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
            Exit Sub
        Else
            UserList(UserIndex).Reputacion.NobleRep = 0
            UserList(UserIndex).Reputacion.PlebeRep = 0
            UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 200
            If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
        
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z7")
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
            Exit Sub
        Else
            UserList(UserIndex).Reputacion.NobleRep = 0
            UserList(UserIndex).Reputacion.PlebeRep = 0
            UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 200
            If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
                Exit Sub
            Else
                UserList(UserIndex).Reputacion.NobleRep = 0
                UserList(UserIndex).Reputacion.PlebeRep = 0
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 500
                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizadoNpc
        b = True
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z9")
    End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = UserIndex Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z10")
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
                Exit Sub
            Else
                UserList(UserIndex).Reputacion.NobleRep = 0
                UserList(UserIndex).Reputacion.PlebeRep = 0
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 500
                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(UserIndex)
        b = True
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z9")
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)

Dim daño As Long


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has curado " & daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
    
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z7")
        b = False
        Exit Sub
    End If
    
   

'CHOTS | Sistema de Castillos
If Npclist(NpcIndex).NPCtype = eNPCType.reydelcastillo Then

    If UserList(UserIndex).GuildIndex = 0 Then
         Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE17")
         Exit Sub
    Else
        If UserList(UserIndex).Pos.Map = CastilloOeste And CASTILLO_OESTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
             UserList(UserIndex).Pos.Map = CastilloEste And CASTILLO_ESTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
             UserList(UserIndex).Pos.Map = CastilloNorte And CASTILLO_NORTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
             UserList(UserIndex).Pos.Map = CastilloSur And CASTILLO_SUR_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
             UserList(UserIndex).Pos.Map = CastilloFortaleza Then
             Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE19")
             Exit Sub

        Else
             Call AtacandoCastillo(UserIndex, NpcIndex)
        End If
    End If
End If

'CHOTS | Sistema de Castillos
    
    If Npclist(NpcIndex).NPCtype = 2 And UserList(UserIndex).flags.Seguro Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z8")
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    'CHOTS | Bastón del Dragón
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).VaraDragon = 1 And Npclist(NpcIndex).NPCtype = DRAGON Then
            daño = daño * 30
        End If
    End If
    'CHOTS | Bastón del Dragón
    
    'CHOTS | Monturas
    If UserList(UserIndex).flags.montandoMascota = 1 Then
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinAtaMag, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxAtaMag)
    End If
    'CHOTS | Monturas

    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If

    End If
    
    'CHOTS | El Bardo baja el Daño
    If UCase$(UserList(UserIndex).Clase) = "BARDO" Then
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
            daño = daño * 1.02
        ElseIf UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDUNICO Then
            daño = daño * 1.01
        Else
            daño = daño * 0.98
        End If
    End If
    'CHOTS | El Bardo baja el Daño
    
    'CHOTS | El Druida baja el Daño
    If UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = FLAUTAMAGICA Then
            daño = daño * 0.98
        ElseIf UserList(UserIndex).Invent.HerramientaEqpObjIndex = FLAUTAUNICA Then
            daño = daño * 0.96
        Else
            daño = daño * 0.95
        End If
    End If
    'CHOTS | El Druida baja el Daño

    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    SendData SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has causado " & daño & " puntos de daño a la criatura!" & FONTTYPE_FIGHT
    Call CalcularDarExp(UserIndex, NpcIndex, daño)

If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
        'CHOTS | Saca la Vara Draco
        If Npclist(NpcIndex).NPCtype = DRAGON Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = BacuDragonIndex Then
                Call QuitarObjetos(UserList(UserIndex).Invent.WeaponEqpObjIndex, 1, UserIndex)
            End If
        End If
        'CHOTS | Saca la Vara Draco
Else
    'Mascotas atacan a la criatura.
    Call CheckPets(NpcIndex, UserIndex, True)
End If
End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserList(UserIndex).flags.TargetUser).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Map, "CXF" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToIndex, UserList(UserIndex).flags.TargetUser, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim daño As Long
Dim tempChr As Integer
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim TempX As Integer
Dim TempY As Integer

    
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.TargetUser

'CHOTS | Detectar Invi
If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
    PosCasteadaX = UserList(tempChr).Pos.X
    PosCasteadaY = UserList(tempChr).Pos.Y
    PosCasteadaM = UserList(tempChr).Pos.Map
    
    For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
        For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
            If InMapBounds(PosCasteadaM, TempX, TempY) Then
                If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                    'hay un user
                    If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
                    End If
                End If
            End If
        Next TempY
    Next TempX

    Call InfoHechizo(UserIndex)

    b = True
End If
'CHOTS | Detectar Invi
      
'hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has restaurado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarhambreYsed(tempChr)
    UserList(tempChr).flags.Hambre = 0
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(UserIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has quitado " & daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarhambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If UserIndex <> tempChr Then
      Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
      Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has restaurado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).flags.Sed = 0
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has quitado " & daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = IntervaloDroga
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
    
    If UCase$(UserList(tempChr).Raza) = "HUMANO" Then
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) >= STAT_MAXATRIBUTOS_HUMANO Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = STAT_MAXATRIBUTOS_HUMANO
    Else
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) >= STAT_MAXATRIBUTOS Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = STAT_MAXATRIBUTOS
    End If

    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) >= (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    Call EnviarDopa(tempChr)
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = IntervaloDroga
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    Call EnviarDopa(tempChr)
    ' BysNacK - Clero dropa 40 solo a clero
ElseIf Hechizos(H).SubeAgilidad = 3 Then
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = IntervaloDroga
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño

    If UCase$(UserList(tempChr).Raza) = "HUMANO" Then
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) >= STAT_MAXATRIBUTOS_HUMANO Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = STAT_MAXATRIBUTOS_HUMANO
    Else
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) >= STAT_MAXATRIBUTOS Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = STAT_MAXATRIBUTOS
    End If

    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) >= (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)
    
    If UCase$(UserList(tempChr).Clase) = "CLERIGO" Then
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + 2
    End If
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    Call EnviarDopa(tempChr)
    ' BysNacK - Clero dropa 40 solo a clero
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = IntervaloDroga

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño
    
    If UCase$(UserList(tempChr).Raza) = "HUMANO" Then
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) >= STAT_MAXATRIBUTOS_HUMANO Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = STAT_MAXATRIBUTOS_HUMANO
    Else
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) >= STAT_MAXATRIBUTOS Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = STAT_MAXATRIBUTOS
    End If

    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) >= (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    Call EnviarDopa(tempChr)
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = IntervaloDroga
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    Call EnviarDopa(tempChr)
    ' BysNacK - Clero dropa 40 solo a clero
ElseIf Hechizos(H).SubeFuerza = 3 Then
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(UserIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = IntervaloDroga
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

    If UCase$(UserList(tempChr).Raza) = "HUMANO" Then
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) >= STAT_MAXATRIBUTOS_HUMANO Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = STAT_MAXATRIBUTOS_HUMANO
    Else
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) >= STAT_MAXATRIBUTOS Then _
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = STAT_MAXATRIBUTOS
    End If

    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) >= (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = (UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)
    
    If UCase$(UserList(tempChr).Clase) = "CLERIGO" Then
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + 2
    End If
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    Call EnviarDopa(tempChr)
    ' BysNacK - Clero dropa 40 solo a clero
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    
    If Criminal(tempChr) And Not Criminal(UserIndex) Then
        If UserList(UserIndex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z6")
            Exit Sub
        Else
            Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has restaurado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
        'If UserList(UserIndex).flags.SeguroClan Then
        'If Guilds(UserList(tempChr).GuildIndex).GuildName = Guilds(UserList(UserIndex).GuildIndex).GuildName And Guilds(UserList(UserIndex).GuildIndex).GuildName <> "" Then
        '    Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_FIGHT)
        '    Exit Sub
        'End If
    'End If
    
    If UserIndex = tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z62")
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " danio, minhp, maxhp " & daño & " " & Hechizos(H).MinHP & " " & Hechizos(H).MaxHP & FONTTYPE_VENENO)
'End If
    
    
    daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
    
    If Hechizos(H).StaffAffected Then
        If UCase$(UserList(UserIndex).Clase) = "MAGO" Then
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    
    If UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.02  'laud magico de los bardos
    ElseIf UserList(UserIndex).Invent.HerramientaEqpObjIndex = LAUDUNICO Then
        daño = daño * 1.01
    End If
    
    If EstaLiberado(UserIndex) And UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" And UCase$(UserList(UserIndex).Clase) = "MAGO" Then
        daño = daño + 10  'CHOTS | Gemas (Mago EO)
    End If
    
    
    'CHOTS | El Druida baja el Daño
    If UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = FLAUTAMAGICA Then
            daño = daño * 0.98
        ElseIf UserList(UserIndex).Invent.HerramientaEqpObjIndex = FLAUTAUNICA Then
            daño = daño * 0.96
        Else
            daño = daño * 0.95
        End If
    End If
    'CHOTS | El Druida baja el Daño
    
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
    End If
    
    'CHOTS | Monturas
    If UserList(UserIndex).flags.montandoMascota = 1 Then
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinAtaMag, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxAtaMag)
    End If

    If UserList(tempChr).flags.montandoMascota = 1 Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.MonturaObjIndex).MinMag, ObjData(UserList(tempChr).Invent.MonturaObjIndex).MaxMag)
    End If
    'CHOTS | Monturas
    
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    If daño < 0 Then daño = 0
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
    Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        'Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte, Optional Conecta As Boolean = False)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

If Conecta Then
    For LoopC = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        End If
    Next LoopC
    Exit Sub
End If

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).nombre)

Else

    Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHS" & Slot & "," & "N") 'CHOTS | Optimizado

End If


End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, UserIndex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, UserIndex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, UserIndex, CualHechizo)

End Sub


Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

    'Si estamos en la arena no hacemos nada
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
    
    'pierdo nobleza...
    UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep - NoblePts
    If UserList(UserIndex).Reputacion.NobleRep < 0 Then
        UserList(UserIndex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + BandidoPts
    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
        UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PN")
    If Criminal(UserIndex) Then If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End Sub
