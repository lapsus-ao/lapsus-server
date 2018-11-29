Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, X, Y) Then
    
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        FxFlag = (ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTELEPORT)
    End If
    
    If MapData(Map, X, Y).TileExit.Map > 0 Then
    
        'CHOTS | Fortaleza
        If MapData(Map, X, Y).TileExit.Map = CastilloFortaleza Or MapData(Map, X, Y).TileExit.Map = CastilloFortalezaDungeon1 Or MapData(Map, X, Y).TileExit.Map = CastilloFortalezaDungeon2 Then
        
            If UserList(UserIndex).GuildIndex > 0 Then
            
                If Not CASTILLO_OESTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or Not _
                CASTILLO_ESTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or Not _
                CASTILLO_NORTE_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Or Not _
                CASTILLO_SUR_DUEÑO = Guilds(UserList(UserIndex).GuildIndex).GuildName Then
                    
                    If UserList(UserIndex).flags.UltimoMensaje <> 218 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE91")
                        UserList(UserIndex).flags.UltimoMensaje = 218
                    End If
                    
                    Exit Sub
                    
                End If
            Else
                If UserList(UserIndex).flags.UltimoMensaje <> 218 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE91")
                    UserList(UserIndex).flags.UltimoMensaje = 218
                End If
                
                Exit Sub
                
            End If
        
        End If
        'CHOTS | Fortaleza

    
        If UserList(UserIndex).GuildIndex = 0 Then
            If MapData(Map, X, Y).TileExit.Map = CastilloNorte Or MapData(Map, X, Y).TileExit.Map = CastilloOeste Or MapData(Map, X, Y).TileExit.Map = CastilloEste Or MapData(Map, X, Y).TileExit.Map = CastilloSur Then

                If UserList(UserIndex).flags.UltimoMensaje <> 15 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE78")
                    UserList(UserIndex).flags.UltimoMensaje = 15
                End If

                Exit Sub

            End If
        End If
    
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                    End If
                End If
            Else 'No es newbie
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z45")
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        ElseIf MapInfo(MapData(Map, X, Y).TileExit.Map).MinLevel > 0 Then
            'CHOTS | El mapa requiere un nivel minimo
            If UserList(UserIndex).Stats.ELV < MapInfo(MapData(Map, X, Y).TileExit.Map).MinLevel Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Sientes la sangre helarse cuando te acercas a la puerta, y el miedo te impide seguir. Volveré cuando sea nivel " & MapInfo(MapData(Map, X, Y).TileExit.Map).MinLevel & "." & FONTTYPE_WARNING)
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            Else
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                    End If
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub

Function InRangoVision(ByVal UserIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
    If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
    If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
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
Sub ClosestLegalPos2(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, True)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY, True) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
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
Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
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

Function NameIndex(ByRef Name As String) As Integer

Dim UserIndex As Integer
'¿Nombre valido?
If Name = "" Then
    NameIndex = 0
    Exit Function
End If

Name = UCase$(Replace(Name, "+", " "))

UserIndex = 1
Do Until UCase$(UserList(UserIndex).Name) = Name
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        NameIndex = 0
        Exit Function
    End If
    
Loop
 
NameIndex = UserIndex
 
End Function



Function IP_Index(ByVal inIP As String) As Integer
 
Dim UserIndex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = UserIndex

Exit Function

End Function


Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Long
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        
        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If Head = eHeading.NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = eHeading.SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = eHeading.EAST Then
    nX = X + 1
    nY = Y
End If

If Head = eHeading.WEST Then
    nX = X - 1
    nY = Y
End If

'Devuelve valores
Pos.X = nX
Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False) As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
  
If Not PuedeAgua Then
    LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
    (MapData(Map, X, Y).UserIndex = 0) And _
    (MapData(Map, X, Y).NpcIndex = 0) And _
    (Not HayAgua(Map, X, Y))
Else
    LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
    (MapData(Map, X, Y).UserIndex = 0) And _
    (MapData(Map, X, Y).NpcIndex = 0) And _
    (HayAgua(Map, X, Y))
End If
   
End If

End Function
Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

    If (Map <= 0 Or Map > NumMaps) Or _
        (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If
    

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA) _
     And Not HayAgua(Map, X, Y)
 Else
   LegalPosNPC = (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
      HayAgua(Map, X, Y)
 End If
 
'End If


End Function
Function LegalPosNPC2(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

    If (Map <= 0 Or Map > NumMaps) Or _
        (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC2 = False
        Exit Function
    End If
    
    LegalPosNPC2 = True
End Function


Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(SendTarget.ToIndex, Index, 0, ServerPackages.dialogo & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).char.CharIndex & FONTTYPE_INFO)
    End If
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error GoTo errhandler
'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim stat As String
Dim OBJType As Integer

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = X
    UserList(UserIndex).flags.TargetY = Y
    '¿Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = X
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.ObjIndex
        If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & ObjData(UserList(UserIndex).flags.TargetObj).Name & " - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).OBJInfo.Amount & "" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & ObjData(UserList(UserIndex).flags.TargetObj).Name & FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
            
            If UserList(TempCharIndex).DescRM = "" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "VES" & Vesa(TempCharIndex))
            Else
                stat = UserList(TempCharIndex).DescRM & " " & FONTTYPE_INFON
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & stat)
            End If
            

            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
             
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).char.CharIndex & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, VesNpc(TempCharIndex, UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia), UserList(UserIndex).flags.Privilegios >= PlayerType.Ot))
            End If
            
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If

Exit Sub

errhandler: Debug.Print "Error en LookAtTILE" & Err.number & Err.Description
End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otFOROS And _
            ObjData(Index).OBJType <> eOBJType.otCARTELES And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTELEPORT
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otFOROS And _
            ObjData(Index).OBJType <> eOBJType.otCARTELES And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTELEPORT
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otFOROS Or _
               OBJType = eOBJType.otCARTELES Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function
Public Function VesNpc(ByVal NpcIndex As Integer, ByVal Supervivencia As Byte, ByVal EsGm As Byte) As String
'Sistema de optimizaciones al clickear un Npc
'Programado por Juan Andrés Dalmasso (CHOTS)
'Lapsus Corp | Todos los derechos reservados
'Programado el 20/11/2010
'Referencias:
'0=Intacto
'1=Sano
'2=Levemente Herido
'3=Herido
'4=Malherido
'5=Muy malherido
'6=Casi Muerto
'7=Agonizando
'8=Dudoso
'9=Error
'
'Procedimiento:
'Tradicionalmente envía ||(Dudoso) Orco Brujo.
'Aca enviaría NPCOrco Brujo,8
'
'Cuando ve la vida mucho no se puede optimizar, pero igualmente rinde
Dim estatus As String
Dim VeVida As Boolean
VeVida = False
estatus = Npclist(NpcIndex).Name & ","
            
If EsGm Then
    estatus = estatus & Npclist(NpcIndex).Stats.MinHP & "," & Npclist(NpcIndex).Stats.MaxHP
    VeVida = True
Else
    If Supervivencia >= 0 And Supervivencia <= 10 Then
        estatus = estatus & 8
    ElseIf Supervivencia > 10 And Supervivencia <= 20 Then
        If Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP / 2) Then
            estatus = estatus & 3
        Else
            estatus = estatus & 1
        End If
    ElseIf Supervivencia > 20 And Supervivencia <= 30 Then
        If Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.5) Then
            estatus = estatus & 4
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.75) Then
            estatus = estatus & 3
        Else
            estatus = estatus & 1
        End If
    ElseIf Supervivencia > 30 And Supervivencia <= 40 Then
        If Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.25) Then
            estatus = estatus & 5
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.5) Then
            estatus = estatus & 3
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.75) Then
            estatus = estatus & 2
        Else
            estatus = estatus = estatus & 1
        End If
        
    ElseIf Supervivencia > 40 And Supervivencia < 60 Then
        If Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.05) Then
            estatus = estatus & 7
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.1) Then
            estatus = estatus & 6
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.25) Then
            estatus = estatus & 5
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.5) Then
            estatus = estatus & 3
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP * 0.75) Then
            estatus = estatus & 2
        ElseIf Npclist(NpcIndex).Stats.MinHP < (Npclist(NpcIndex).Stats.MaxHP) Then
            estatus = estatus & 1
        Else
            estatus = estatus & 0
        End If
    ElseIf Supervivencia >= 60 Then
        estatus = estatus & Npclist(NpcIndex).Stats.MinHP & "," & Npclist(NpcIndex).Stats.MaxHP
        VeVida = True
    Else
        estatus = estatus & 9
    End If
End If

If Npclist(NpcIndex).MaestroUser > 0 Then
    estatus = estatus & "," & UserList(Npclist(NpcIndex).MaestroUser).Name
End If

If VeVida Then
    VesNpc = "NPZ" & estatus
Else
    VesNpc = "NPÑ" & estatus
End If

End Function
Public Function Vesa(ByVal UserIndex As Integer) As String
'Sistema de optimización de click al usuario
'Desarrollado por Juan Andrés Dalmasso (CHOTS)
'Programado el 11 de Febrero de 2010 para LapsusAO
'Lapsus Corp | Todos los derechos reservados
'
'Como funciona
'Nick,Nw,Caos o Armada,Titulo,Clan,Casado,Pareja,Desc,Consejo,Ciuda Crimi Gm
'
'Ejemplo:
'Ves a KillerProCjs <Legion Oscura> <Esbirro> <DoN ErnesTo> <Casado con Aramir> CHOTS <Criminal>
'Peso: 105 bytes
'
'Optimizado quedaria
'KillerProCjs,2,1,Don ErnesTo,Aramir,CHOTS,0,1
'Peso: 45 bytes
'Optimizado más del 50%

Dim stat As String 'CHOTS | Acá se almacena todo


stat = UserList(UserIndex).Name & ","


If EsNewbie(UserIndex) Then
    stat = stat & "1,"
Else
    stat = stat & "0,"
End If


If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    stat = stat & "1," & UserList(UserIndex).Faccion.Jerarquia & ","
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    stat = stat & "2," & UserList(UserIndex).Faccion.Jerarquia & ","
Else
    stat = stat & "0,0,"
End If


If UserList(UserIndex).GuildIndex > 0 Then
    stat = stat & Guilds(UserList(UserIndex).GuildIndex).GuildName & ","
Else
    stat = stat & "0,"
End If


If UserList(UserIndex).flags.Casado = 1 And UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
    stat = stat & "1," & UserList(UserIndex).Pareja & ","
ElseIf UserList(UserIndex).flags.Casado = 1 And UCase$(UserList(UserIndex).Genero) = "MUJER" Then
    stat = stat & "2," & UserList(UserIndex).Pareja & ","
Else
    stat = stat & "0,0,"
End If


If Len(UserList(UserIndex).Desc) > 1 Then
    stat = stat & UserList(UserIndex).Desc & ","
Else
    stat = stat & "0,"
End If

If UserList(UserIndex).flags.PertAlCons > 0 Then
    stat = stat & "1,"
ElseIf UserList(UserIndex).flags.PertAlConsCaos > 0 Then
    stat = stat & "2,"
Else
    stat = stat & "0,"
End If

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
    stat = stat & "2"
ElseIf Criminal(UserIndex) Then
    stat = stat & "1"
Else
    stat = stat & "0"
End If

Vesa = stat


End Function

Public Sub CambiarItem(ByVal UserIndex As Integer, ByVal item As Byte)

'CHOTS | Sistema de Cambio de Trofeos
'Programado por Lucho para Land of Dragons 2008
'Reprogramado y adaptado por CHOTS para Silv AO 2008
'Reprogramado y adaptado por CHOTS para LapsusAO 2010

'CHOTS | Constantes
'Trofeo de Oro = 931
'Trofeo de Plata = 932
'Gema = 413 (GEMAVIOLETA)
'CHOTS | Constantes


'CHOTS | 28 de Septiembre de 2010
'Este es el Ermitaño, osea, cambia trofeos por items grosos ;)


Dim trofeosoro As Byte
Dim trofeosplata As Byte
Dim Premio As Integer

trofeosoro = 0
trofeosplata = 0

Select Case item
    Case 0 'CHOTS | Gema
        trofeosoro = 30
        Premio = GEMAVIOLETA
    Case 1 'CHOTS | Armadura Pesada Alto
        trofeosoro = 20
        Premio = 1106
    Case 2 'CHOTS | Armadura Pesada Bajo
        trofeosoro = 20
        Premio = 1107
    Case 3 'CHOTS | Armadura Liviana Alto
        trofeosoro = 20
        Premio = 1108
    Case 4 'CHOTS | Armadura Liviana Bajo
        trofeosoro = 20
        Premio = 1109
    Case 5 'CHOTS | Tunica Alto
        trofeosoro = 20
        Premio = 1110
    Case 6 'CHOTS | Tunica Bajo
        trofeosoro = 20
        Premio = 1111
    Case 7 'CHOTS | Arma
        trofeosoro = 20
        Premio = 1114
    Case 8 'CHOTS | Daga
        trofeosoro = 20
        Premio = 1115
    Case 9 'CHOTS | Baculo
        trofeosoro = 20
        Premio = 1116
    Case 10 'CHOTS | Arco
        trofeosoro = 20
        Premio = 1120
    Case 11 'CHOTS | Escudo Pesado
        trofeosoro = 10
        Premio = 1112
    Case 12 'CHOTS | Escudo liviano
        trofeosoro = 10
        Premio = 1113
    Case 13 'CHOTS | Casco
        trofeosoro = 10
        Premio = 1118
    Case 14 'CHOTS | Corona
        trofeosoro = 10
        Premio = 1117
    Case 15 'CHOTS | Laúd
        trofeosoro = 20
        Premio = 1133
    Case 16 'CHOTS | Flauta
        trofeosoro = 20
        Premio = 1132
End Select

If Not TieneObjetos(931, trofeosoro, UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes trofeos de oro." & FONTTYPE_INFO)
    Exit Sub
End If

If Not TieneObjetos(932, trofeosplata, UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes trofeos de plata." & FONTTYPE_INFO)
    Exit Sub
End If

Call QuitarObjetos(931, trofeosoro, UserIndex)
Call QuitarObjetos(932, trofeosplata, UserIndex)

Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = Premio
    
If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Call UpdateUserInv(True, UserIndex, 0)

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has cambiado tus trofeos. ¡Felicitaciones!." & FONTTYPE_INFON)
End Sub

Public Sub CambiarTrofeosRunas(ByVal UserIndex As Integer, ByVal item As Byte)

'CHOTS | Sistema de Cambio de Trofeos y Runas



Dim runas As Byte
Dim trofeosplata As Byte
Dim Premio As Integer

trofeosplata = 0
runas = 0

Select Case item
    Case 0 'CHOTS | Trofeo de Oro
        trofeosplata = 3
        Premio = 931
    Case 1 'CHOTS | Trofeo de Plata
        runas = 3
        Premio = 932
End Select

If Not TieneObjetos(932, trofeosplata, UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes Trofeos de plata." & FONTTYPE_INFO)
    Exit Sub
End If

If Not TieneObjetos(1121, runas, UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes Runas." & FONTTYPE_INFO)
    Exit Sub
End If

Call QuitarObjetos(1121, runas, UserIndex)
Call QuitarObjetos(932, trofeosplata, UserIndex)

Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = Premio
    
If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Call UpdateUserInv(True, UserIndex, 0)

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Aquí están tus items ¡Felicitaciones!." & FONTTYPE_INFON)
End Sub
