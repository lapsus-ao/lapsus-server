Attribute VB_Name = "TCP_HandleData1"
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

Public Sub HandleData_1(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer
Dim Cantidad As Single

Procesado = True 'ver al final del sub

    Select Case UCase$(Left$(rData, 1))
        Case ClientPackages.hablar
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
        
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(UserIndex).Name, "Dijo: " & rData, True)
            End If
            
            ind = UserList(UserIndex).char.CharIndex
            
            'piedra libre para todos los compas!
            If UserList(UserIndex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub
            End If
            
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                If UserList(UserIndex).flags.Invisible = 0 Then
                    Dim ChotsNover As String
                    ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
                    'ChotsNover = Encriptar(ChotsNover)
                    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
                End If
            End If
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & "12632256°" & rData & "°" & CStr(ind))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & rData & "°" & CStr(ind))
            End If
            Exit Sub
        Case ClientPackages.gritar
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                    Exit Sub
            End If
                        If UserList(UserIndex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(UserIndex).Name, "Grito: " & rData, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                If UserList(UserIndex).flags.Invisible = 0 Then
                    'Dim ChotsNover As String
                    ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
                    'ChotsNover = Encriptar(ChotsNover)
                    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
                End If
            End If
    
    
            ind = UserList(UserIndex).char.CharIndex
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbRed & "°" & rData & "°" & str(ind))
            Exit Sub
        Case "\" 'Susurrar al oido
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            If UserList(UserIndex).flags.Silenciado = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estas Silenciado!" & FONTTYPE_WARNING)
            Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            tName = ReadField(1, rData, 32)
            
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If (EsDios(tName) Or EsAdmin(tName)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes susurrarle a los Dioses y Admins." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
            If UserList(UserIndex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName) Or EsConsejero(tName)) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes susurrarle a los GMs" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If Len(rData) <> Len(tName) Then
                    tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                If Not EstaPCarea(UserIndex, tIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z14")
                    Exit Sub
                End If
                ind = UserList(UserIndex).char.CharIndex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                
                '[Consejeros]
                If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(UserList(UserIndex).Name, "Le dijo a '" & UserList(tIndex).Name & "' " & tMessage, True)
                End If
    
                Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbBlue & "°" & tMessage & "°" & str(ind))
                Call SendData(SendTarget.ToIndex, tIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbBlue & "°" & tMessage & "°" & str(ind))
                '[CDT 17-02-2004]
                If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
                    Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbYellow & "°" & "a " & UserList(tIndex).Name & "> " & tMessage & "°" & str(ind))
                End If
                '[/CDT]
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z13")
            Exit Sub
        Case ClientPackages.moverse
            rData = Right$(rData, Len(rData) - 1)
            
            'salida parche
            If UserList(UserIndex).Counters.Saliendo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z15")
                UserList(UserIndex).Counters.Saliendo = False
                UserList(UserIndex).Counters.Salir = 0
            End If
            
            If UserList(UserIndex).Counters.Hogar > 0 Then
                If Not UserList(UserIndex).flags.UltimoMensaje = 88 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes moverte mientras estás yendo a tu hogar!" & FONTTYPE_HOGAR)
                    UserList(UserIndex).flags.UltimoMensaje = 88
                End If
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Paralizado = 0 Then
                If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando Then
                    Call MoveUserChar(UserIndex, val(rData))
                ElseIf UserList(UserIndex).flags.Descansar Then
                    UserList(UserIndex).flags.Descansar = False
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has dejado de descansar." & FONTTYPE_INFO)
                    Call MoveUserChar(UserIndex, val(rData))
                ElseIf UserList(UserIndex).flags.Meditando Then
                    UserList(UserIndex).flags.Meditando = False
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z16")
                    UserList(UserIndex).char.FX = 0
                    UserList(UserIndex).char.loops = 0
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & 0 & "," & 0)
                End If
            Else    'paralizado
                '[CDT 17-02-2004] (<- emmmmm ?????)
                If Not UserList(UserIndex).flags.UltimoMensaje = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z17")
                    UserList(UserIndex).flags.UltimoMensaje = 1
                End If
                '[/CDT]
            End If
            
            'CHOTS | Con vestimenta de la oscuridad no se va el Ocu
            If UserList(UserIndex).flags.Oculto = 1 Then
                If UserList(UserIndex).Invent.ArmourEqpObjIndex <> 906 Then
                    UserList(UserIndex).flags.Oculto = 0
                    If UserList(UserIndex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
                        'Dim ChotsNover As String
                        ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
                        'ChotsNover = Encriptar(ChotsNover)
                        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
                    End If
                End If
            End If
            'CHOTS | Con vestimenta de la oscuridad no se va el Ocu
            
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Empollando(UserIndex)
            Else
                UserList(UserIndex).flags.EstaEmpo = 0
                UserList(UserIndex).EmpoCont = 0
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rData)
        'Implementaciones del anti cheat By NicoNZ
        Case "TENGOSH"
            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Sistema Anti Cheat 2> " & UserList(UserIndex).Name & " ha sido expulsado por el Anti Cheat. Por favor, que algun gm lo siga ya que es muy probable que tenga un programa externo corriendo." & FONTTYPE_SERVER)
            Call CloseSocket(UserIndex)
            Exit Sub
        
        Case "LAG" 'Pedido de actualizacion de la posicion
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.updatePos & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
        Exit Sub
        
        Case ClientPackages.atacar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z19")
                    Exit Sub
                End If
            End If
            Call UsuarioAtaca(UserIndex)
            Exit Sub
        Case ClientPackages.agarrarObjeto
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And Not UserList(UserIndex).flags.EsRolesMaster Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes tomar ningun objeto. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call GetObj(UserIndex)
            Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z21")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
                UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            End If
            Exit Sub
        Case "ACTUALIZAR"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.updatePos & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Sub
        Case "PTINFO"
            tStr = "PT" & UserList(UserIndex).Puntos
            
            If UserList(UserIndex).GuildIndex <> 0 Then
                tStr = tStr & "," & Guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros & "," & Guilds(UserList(UserIndex).GuildIndex).GetMemberList("@")
            Else
                tStr = tStr & ",0"
            End If
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
            Exit Sub
        Case "GLINFO"
            tStr = SendGuildLeaderInfo(UserIndex)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "GL" & SendGuildsList(UserIndex))
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "LEADERI" & tStr)
            End If
            Exit Sub
        Case "XEST" 'CHOTS | Full estadisticas
            Call EnviarFullEstadisticas(UserIndex)
            Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(UserIndex).ComUsu.DestUsu > 0 And _
                UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINBANOK")
            Exit Sub
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(UserIndex)
            Exit Sub
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[/Alejo]
    
    
    End Select
    
    
    
    Select Case UCase$(Left$(rData, 2))
        Case ClientPackages.tirarItem
                If UserList(UserIndex).flags.Navegando = 1 Or _
                   UserList(UserIndex).flags.Muerto = 1 Or _
                   (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And Not UserList(UserIndex).flags.EsRolesMaster) Then Exit Sub
                   '[Consejeros]
                
                rData = Right$(rData, Len(rData) - 2)
                Arg1 = ReadField(1, rData, 44)
                Arg2 = ReadField(2, rData, 44)
                If val(Arg1) = FLAGORO Then
                    
                    Call TirarOro(val(Arg2), UserIndex)
                    
                    Call EnviarOro(UserIndex)
                    Exit Sub
                Else
                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                    Else
                        Exit Sub
                    End If
                End If
                Exit Sub
        Case ClientPackages.lanzarHechizo
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 2)
            UserList(UserIndex).flags.Hechizo = val(rData)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Magia)
            Exit Sub
        Case ClientPackages.leftClick
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Sub
        Case ClientPackages.rightClick
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Sub
        Case ClientPackages.usarSkill
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
    
            rData = Right$(rData, Len(rData) - 2)
            Select Case val(rData)
                Case Robar
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Robar)
                'CHOTS | Movido al lanzarHechizo
                'Case Magia
                '    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(UserIndex).flags.Navegando = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                            UserList(UserIndex).flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    If UserList(UserIndex).flags.Oculto = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(UserIndex).flags.UltimoMensaje = 2 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z28")
                            UserList(UserIndex).flags.UltimoMensaje = 2
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    Call DoOcultarse(UserIndex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 3))
         Case "UMH" ' Usa macro de hechizos
            Call SendData(SendTarget.ToAdmins, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " fue expulsado por Anti-macro de hechizos " & FONTTYPE_VENENO)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros" & FONTTYPE_INFO)
            Call CloseSocket(UserIndex)
            Exit Sub
        
        
    Case "FPS"
               rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Los FPS del Usuario " & UserSolicitadoFPS & " Son: " & Arg1 & FONTTYPE_SERVER)
               Exit Sub

    Case "FPI"
               rData = Right$(rData, Len(rData) - 3)
               Arg1 = ReadField(1, rData, 44)
               Arg2 = ReadField(2, rData, 44)
               Arg3 = ReadField(3, rData, 44)
               Arg4 = ReadField(4, rData, 44)
               Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Intervalos de " & UserSolicitadoFPS & ": " & FONTTYPE_SERVER)
               Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Ataque: " & Arg1 & FONTTYPE_SERVER)
               Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Pots: " & Arg2 & FONTTYPE_SERVER)
               Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Combo: " & Arg3 & FONTTYPE_SERVER)
               Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Click: " & Arg4 & FONTTYPE_SERVER)
               Exit Sub
        
        Case ClientPackages.usarItem 'CHOTS | Encriptado (17/11/10)
            On Local Error Resume Next
            Dim Numero As Byte
            Dim item As Byte
            rData = Right$(rData, Len(rData) - 3)
            Call IncrementarUseNum(UserIndex)
            
            rData = DecryptStr(rData, UserList(UserIndex).UseAcum)

            item = CByte(ReadField(1, rData, 44))
            Numero = CByte(ReadField(2, rData, 44))
            
            If val(item) <= MAX_INVENTORY_SLOTS And val(item) > 0 Then
                If UserList(UserIndex).Invent.Object(val(item)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            
            'CHOTS | Anti editores de paquetes
            If UserList(UserIndex).UseNum <> Numero Then Exit Sub
            'CHOTS | Anti editores de paquetes
            
            Call UseInvItem(UserIndex, val(item))
            Exit Sub
      Case "CNS" ' Construye herreria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(ReadField(1, rData, 44))
            Cantidad = CSng(ReadField(2, rData, 44))
            If X < 1 Then Exit Sub
            If ObjData(X).SkHerreria = 0 Then Exit Sub
            Call HerreroConstruirItem(UserIndex, X, Cantidad)
            Exit Sub
        Case "AND" ' Construye Pociones
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(ReadField(1, rData, 44))
            Cantidad = CSng(ReadField(2, rData, 44))
            If X < 1 Or ObjData(X).SkAlquimia = 0 Then Exit Sub
            Call DruidaConstruirItem(UserIndex, X, Cantidad)
            Exit Sub
        Case "CNC" ' Construye carpinteria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(ReadField(1, rData, 44))
            Cantidad = CSng(ReadField(2, rData, 44))
            If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
            Call CarpinteroConstruirItem(UserIndex, X, Cantidad)
            Exit Sub
        Case "CND" ' Construye Sastreria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(ReadField(1, rData, 44))
            Cantidad = CSng(ReadField(2, rData, 44))
            If X < 1 Or ObjData(X).SkSastreria = 0 Then Exit Sub
            Call SastreConstruirItem(UserIndex, X, Cantidad)
            Exit Sub
        Case "TPT" ' CHOTS | Transfiere Puntos de Usuario
            rData = Right$(rData, Len(rData) - 3)
            Call TransferirPuntos(UserIndex, ReadField(1, rData, 44), val(ReadField(2, rData, 44)))
            Exit Sub
        Case "CPT" ' CHOTS | Cambia Puntos
            rData = Right$(rData, Len(rData) - 3)
            Call CambiarPuntos(UserIndex, rData)
            Exit Sub
        Case "CTR" ' CHOTS | Cambia trofeos
            rData = Right$(rData, Len(rData) - 3)
            Call CambiarItem(UserIndex, rData)
            Exit Sub
        Case "CRN" ' CHOTS | Cambia trofeos y runas
            rData = Right$(rData, Len(rData) - 3)
            Call CambiarTrofeosRunas(UserIndex, rData)
            Exit Sub
        Case ClientPackages.trabajoClick
            rData = Right$(rData, Len(rData) - 3)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            Arg3 = ReadField(3, rData, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            X = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(UserIndex).flags.Muerto = 1 Or _
               UserList(UserIndex).flags.Descansar Or _
               UserList(UserIndex).flags.Meditando Or _
               Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Sub
            
            If Not InRangoVision(UserIndex, X, Y) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.updatePos & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                If Not IntervaloPermiteAtacar(UserIndex, False) Or Not IntervaloPermiteUsarArcos(UserIndex) Then
                    Exit Sub
                End If

                DummyInt = 0

                If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.WeaponEqpSlot < 1 Or UserList(UserIndex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.MunicionEqpSlot < 1 Or UserList(UserIndex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes municiones." & FONTTYPE_INFO)
                    End If
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                
                DummyInt = 0
                'Quitamos stamina
                If UserList(UserIndex).Stats.MinSta >= 10 Then
                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estas muy cansado para luchar." & FONTTYPE_INFO)
                     Exit Sub
                End If
                 
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
                
                TU = UserList(UserIndex).flags.TargetUser
                tN = UserList(UserIndex).flags.TargetNPC
                
                'Sólo permitimos atacar si el otro nos puede atacar también
                If TU > 0 Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub
                    End If
                ElseIf tN > 0 Then
                    If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub
                    End If
                End If
                
                
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z22")
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                        UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, UserIndex, DummyInt)
                        UserList(UserIndex).Invent.MunicionEqpSlot = 0
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If

                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                ElseIf TU > 0 Then
                    If UserList(UserIndex).flags.Seguro Then
                        If Not Criminal(TU) Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Para atacar ciudadanos desactiva el seguro!" & FONTTYPE_FIGHT)
                            Exit Sub
                        End If
                    End If
                    Call UsuarioAtacaUsuario(UserIndex, TU)
                End If
                
            Case Magia
                If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Una fuerza oscura te impide canalizar tu energía" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(UserIndex).Pos.Map
                wp2.X = X
                wp2.Y = Y
                                
                If UserList(UserIndex).flags.Hechizo > 0 Then
                    If IntervaloPermiteLanzarSpell(UserIndex) Then
                        Call lanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                        'UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(UserIndex).flags.Hechizo = 0
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z67")
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(UserIndex).Pos.X - wp2.X) > 9 Or Abs(UserList(UserIndex).Pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(UserIndex).Name & "(" & UserList(UserIndex).Pos.Map & "/" & UserList(UserIndex).Pos.X & "/" & UserList(UserIndex).Pos.Y & ") ip: " & UserList(UserIndex).ip & " a la posicion (" & wp2.Map & "/" & wp2.X & "/" & wp2.Y & ") "
                    If UserList(UserIndex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).nombre
                    End If
                    If MapData(wp2.Map, wp2.X, wp2.Y).UserIndex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.X, wp2.Y).UserIndex).Name
                    ElseIf MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex).Name
                    End If
                    
                    Call LogCheating(txt)
                End If
                
            
            
            
            Case Pesca
                        
                AuxInd = UserList(UserIndex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                    'Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes pescar desde donde te encuentras." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PESCAR)
                    
                    Select Case AuxInd
                    Case CAÑA_PESCA
                        Call DoPescar(UserIndex)
                    Case RED_PESCA
                        With UserList(UserIndex)
                            wpaux.Map = .Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(UserIndex).Pos, wpaux) > 2 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(UserIndex)
                    End Select
    
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                End If
                
            Case Robar
               If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                       If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = val(ReadField(1, rData, 44))
                            wpaux.Y = val(ReadField(2, rData, 44))
                            If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            
                            Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No a quien robarle!." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                End If
            Case Talar
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Deberías equiparte el hacha." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(UserIndex).Pos) = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes talar desde allí." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SND_TALAR)
                        Call DoTalar(UserIndex)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
                
                Case Botanica
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Deberías equiparte la tijera." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> TIJERA_DRUIDA Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(UserIndex).Pos) = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes Sacar Raices desde allí." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SND_TALAR)
                        Call DoSacarChala(UserIndex)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
                
                
            Case Mineria
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otYacimiento Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_MINERO)
                        Call DoMineria(UserIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
              CI = UserList(UserIndex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
                                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            Call DoDomar(UserIndex, CI)
                        Else
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes domar a esa criatura." & FONTTYPE_INFO)
                        End If
              Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ninguna criatura alli!." & FONTTYPE_INFO)
              End If
              
            Case FundirMetal
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otFragua Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex <> UserList(UserIndex).flags.TargetObjInvIndex Then
                            If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes mas minerales" & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & ">>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        Call FundirMineral(UserIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                End If
                
            'CHOTS | Monturas
            Case CapturarNpc
            
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If Not TieneObjetos(GEMACAPTURA, 1, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes la gema purificada." & FONTTYPE_MONTURA)
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex
                
                
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 6 Then
                        Call SendData(ToIndex, UserIndex, 0, "Z27")
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(UserIndex).Pos) = 0 Then
                        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'CHOTS | Acá se deben indicar el rango de animales capturables
                    If Npclist(AuxInd).Numero >= 586 And Npclist(AuxInd).Numero <= 591 Then

                        If Npclist(AuxInd).MaestroUser = UserIndex Then
                            
                            'CHOTS | Vemos si captura o rompe la gema
                            If RandomNumber(1, 100) <= Npclist(AuxInd).probCaptura Then
                                Call SendData(SendTarget.ToNPCArea, AuxInd, Npclist(AuxInd).Pos.Map, "CXF" & Npclist(AuxInd).char.CharIndex & "," & FXIDs.FXCAPTURA & ",0")
                                Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SND_CAPTURA)
                                Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Has capturado el animal en tu gema!." & FONTTYPE_MONTURA)
                                Call LogGM("CAPTURAS", UserList(UserIndex).Name & " capturó un " & Npclist(AuxInd).Name, False)
                                Call CrearMontura(UserIndex, Npclist(AuxInd).Numero)
                            Else
                                Call SendData(SendTarget.ToNPCArea, AuxInd, Npclist(AuxInd).Pos.Map, "CXF" & Npclist(AuxInd).char.CharIndex & "," & FXIDs.FXROTURA & "," & 1)
                                Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SND_ROTURA)
                                Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡El animal ha roto tu gema!" & FONTTYPE_MONTURA)
                                Call LogGM("CAPTURAS", UserList(UserIndex).Name & " rompió gema tratando de capturar un " & Npclist(AuxInd).Name, False)
                            End If
                                
                            'CHOTS | QUITAMOS LA GEMA
                            Call QuitarObjetos(GEMACAPTURA, 1, UserIndex)
    
                            'CHOTS | MATAMOS EL NPC CAPTURADO
                            Call MuereNpc(AuxInd, 0)
                        Else
                            Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡El animal debe ser domado antes de utilizar la gema!" & FONTTYPE_MONTURA)
                        End If
                    Else
                        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "Debes clickear un animal capturable." & FONTTYPE_MONTURA)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ningun animal ahí." & FONTTYPE_MONTURA)
                End If
            Exit Sub
            'CHOTS | Monturas
                
            Case Herreria
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SFH")
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ahi no hay ningun yunque." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ahi no hay ningun yunque." & FONTTYPE_INFO)
                End If
                
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
            
        Case "CIN" 'CHOTS | Denuncias
            UserList(UserIndex).flags.YaDenuncio = 0
        Exit Sub
        
        
        Case "CIG"
            rData = Right$(rData, Len(rData) - 3)
            
            If modGuilds.CrearNuevoClan(rData, UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " fundó el clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " !!!" & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToAll, 0, 0, "TW" & SONIDOS_GUILD.SND_CREACIONCLAN)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 4))
    
        Case "PCGF" 'CHOTS | Ver Procesos
            On Local Error Resume Next
            Dim proceso As String
            rData = Right$(rData, Len(rData) - 4)
            proceso = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.ToIndex, tIndex, 0, "PCGN" & proceso & "," & UserList(UserIndex).Name)
            Exit Sub
            
        Case "PCWC" 'CHOTS | Ver Rutas
            On Local Error Resume Next
            Dim proseso As String
            rData = Right$(rData, Len(rData) - 4)
            proseso = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.ToIndex, tIndex, 0, "PCSS" & proseso & "," & UserList(UserIndex).Name)
            Exit Sub
            
        Case "PCCC" 'CHOTS | Ver Captions
            On Local Error Resume Next
            Dim caption As String
            rData = Right$(rData, Len(rData) - 4)
            caption = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.ToIndex, tIndex, 0, "PCCC" & caption & "," & UserList(UserIndex).Name)
            Exit Sub

        Case "PFTF" 'CHOTS | Ver Foto
            On Local Error Resume Next
            Dim ftUserName As String
            rData = Right$(rData, Len(rData) - 4)
            ftUserName = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "La foto de " & UCase$(ftUserName) & " se ha subido con exito!" & FONTTYPE_SERVER)
            Exit Sub

        Case "PFTE" 'CHOTS | Ver Foto
            On Local Error Resume Next
            Dim fteUserName As String
            rData = Right$(rData, Len(rData) - 4)
            fteUserName = ReadField(1, rData, 44)
            tIndex = ReadField(3, rData, 44)
            Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "Error al subir la foto de " & UCase$(fteUserName) & ". Error: " & ReadField(2, rData, 44) & FONTTYPE_SERVER)
            Exit Sub
               
        Case "REAQ" 'CHOTS | Realiza Quest
            rData = Right$(rData, Len(rData) - 4)
            Call AceptarQuest(UserIndex, rData)
            Exit Sub
            
        Case "DADI" 'BysNacK | Drag And Drop Items
            rData = Right$(rData, Len(rData) - 4)
            Dim ObjAMover1 As Byte
            Dim ObjAMover2 As Byte
            ObjAMover1 = ReadField(1, rData, 44)
            ObjAMover2 = ReadField(2, rData, 44)
            If UserList(UserIndex).flags.Comerciando Then Exit Sub 'Si esta comerciando no puede moverlo (evitamos que intenten dupear con banco, comercio, comercio usu)
            Call IntercambiarObjetos(UserIndex, ObjAMover1, ObjAMover2)
            Exit Sub
    
        Case "INFS" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) > 0 And val(rData) < MAXUSERHECHIZOS + 1 Then
                    Dim H As Integer
                    H = UserList(UserIndex).Stats.UserHechizos(val(rData))
                    If H > 0 And H < NumeroHechizos + 1 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Nombre:" & Hechizos(H).nombre & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Stamina necesaria: " & Hechizos(H).StaRequerido & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
                End If
                Exit Sub
        Case ClientPackages.equiparItem
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                    Exit Sub
                End If
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                     If UserList(UserIndex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(UserIndex, val(rData))
                Exit Sub
        Case "CHEA" 'Cambiar Heading ;-)
            rData = Right$(rData, Len(rData) - 4)
            If val(rData) > 0 And val(rData) < 5 Then
                UserList(UserIndex).char.Heading = rData
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
            End If
            Exit Sub
        Case "SKSE" 'Modificar skills
            Dim Sumatoria As Integer
            Dim incremento As Integer
            rData = Right$(rData, Len(rData) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                
                If incremento < 0 Then
                    'Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                    UserList(UserIndex).Stats.SkillPts = 0
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                Sumatoria = Sumatoria + incremento
            Next i
            
            If Sumatoria > UserList(UserIndex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
                UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
                If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            
            rData = Right$(rData, Len(rData) - 4)
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rData) > 0 And val(rData) < Npclist(UserList(UserIndex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNPC).Criaturas(val(rData)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Pos, True, False)
                        If SpawnedNpc > 0 Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNPC
                            Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas = Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            End If
            
            Exit Sub
        Case "COMP"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            'User compra el item del slot rdata
            If UserList(UserIndex).flags.Comerciando = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No estas comerciando " & FONTTYPE_INFO)
                Exit Sub
            End If
            'listindex+1, cantidad
            Call NPCVentaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rData = Right(rData, Len(rData) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            '¿El target es un NPC valido?
            tInt = val(ReadField(1, rData, 44))
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
'           rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            Call NPCCompraItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right(rData, Len(rData) - 5)
            'User deposita el item del slot rdata
            Call UserDepositaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select

    Select Case UCase$(Left$(rData, 5))
        Case "DEMSG"
            If UserList(UserIndex).flags.TargetObj > 0 Then
            rData = Right$(rData, Len(rData) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rData, 176)
            msg = ReadField(2, rData, 176)
            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 6))
        Case "DESPHE" 'Mover Hechizo de lugar
            rData = Right(rData, Len(rData) - 6)
            Call DesplazarHechizo(UserIndex, CInt(ReadField(1, rData, 44)), CInt(ReadField(2, rData, 44)))
            Exit Sub
        Case "DESCOD" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 6)
                Call modGuilds.ActualizarCodexYDesc(rData, UserList(UserIndex).GuildIndex)
                Exit Sub
                
        Case "GEMHAB" 'CHOTS | Liberar Gemas Habilidad
                rData = Right$(rData, Len(rData) - 6)
                
                If Not TieneObjetos(GEMAVIOLETA, 1, UserIndex) Then 'CHOTS | No tiene la gema
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes una Gema Sagrada!" & FONTTYPE_GEMA)
                Exit Sub
                End If
                
                If EstaLiberado(UserIndex) Then 'CHOTS | Ya la chupo (a la gema)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya has liberado tus Habilidades!" & FONTTYPE_GEMA)
                Exit Sub
                End If
                
                If UserList(UserIndex).Stats.ELV < 46 Then 'CHOTS | No es lvl alto
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Necesitas ser Nivel 46 para liberar tus habilidades!" & FONTTYPE_GEMA)
                Exit Sub
                End If
                
                Call LiberarHabilidad(UserIndex)
                
                Call QuitarObjetos(GEMAVIOLETA, 1, UserIndex) 'CHOTS | Rompe la Gema
                
                Exit Sub
                
                
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(rData, 7))
        Case "BANEAME"
            rData = Right(rData, Len(rData) - 7)
            H = FreeFile
            Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As H
            
            Print #H, "########################################################################"
            Print #H, "Usuario: " & UserList(UserIndex).Name
            Print #H, "Fecha: " & Date
            Print #H, "Hora: " & Time
            Print #H, "CHEAT: " & rData
            Print #H, "########################################################################"
            Print #H, " "
            Close #H
            
            'UserList(UserIndex).flags.Ban = 1
        
            'Avisamos a los admins
            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Sistema Anticheat> " & UserList(UserIndex).Name & " ha sido Echado por uso de " & rData & FONTTYPE_SERVER)
            'Call CloseSocket(UserIndex)
            Exit Sub
    Case "OFRECER"
            rData = Right$(rData, Len(rData) - 7)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))

            If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
                Exit Sub
            End If
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else
                'esta vivo ?
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                End If
                '//Tiene la cantidad que ofrece ??//'
                If val(Arg1) = FLAGORO Then
                    'oro
                    If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                Else
                    'inventario
                    If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                If UserList(UserIndex).ComUsu.Objeto > 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes cambiar tu oferta." & FONTTYPE_TALK)
                    Exit Sub
                End If
                'No permitimos vender barcos mientras están equipados (no podés desequiparlos y causa errores)
                If UserList(UserIndex).flags.Navegando = 1 Then
                    If UserList(UserIndex).Invent.BarcoSlot = val(Arg1) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podés vender tu barco mientras lo estés usando." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                
                UserList(UserIndex).ComUsu.Objeto = val(Arg1)
                UserList(UserIndex).ComUsu.Cant = val(Arg2)
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Exit Sub
    End Select
    '[/Alejo]
    
    Select Case UCase$(Left$(rData, 8))
    
    
        'clanesnuevo
        Case "ACEPPEAT" 'aceptar paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDePaz(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "Tu clan ha firmado la paz con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & "Tu clan ha firmado la paz con " & UserList(UserIndex).Name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPALIA" 'rechazar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "Tu clan rechazado la propuesta de alianza de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha rechazado nuestra propuesta de alianza con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPPEAT" 'rechazar propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDePaz(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "Tu clan rechazado la propuesta de paz de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha rechazado nuestra propuesta de paz con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACEPALIA" 'aceptar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "Tu clan ha firmado la alianza con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & "Tu clan ha firmado la paz con " & UserList(UserIndex).Name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "PEACEOFF"
            'un clan solicita propuesta de paz a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(UserIndex, Arg1, PAZ, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Propuesta de paz enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEOFF" 'un clan solicita propuesta de alianza a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(UserIndex, Arg1, ALIADOS, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Propuesta de alianza enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEDET"
            'un clan pide los detalles de una propuesta de ALIANZA
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(UserIndex, rData, ALIADOS, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ALLIEDE" & tStr)
            End If
            Exit Sub
        Case "PEACEDET" '-"ALLIEDET"
            'un clan pide los detalles de una propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(UserIndex, rData, PAZ, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PEACEDE" & tStr)
            End If
            Exit Sub
        Case "ENVCOMEN"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            If rData = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesAspirante(UserIndex, rData)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " El personaje no ha mandado solicitud, o no estás habilitado para verla." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PETICIO" & tStr)
            End If
            Exit Sub
        Case "ENVALPRO" 'enviame la lista de propuestas de alianza
            tIndex = modGuilds.r_CantidadDePropuestas(UserIndex, ALIADOS)
            tStr = "ALLIEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(UserIndex, ALIADOS)
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
            Exit Sub
        Case "ENVPROPP" 'enviame la lista de propuestas de paz
            tIndex = modGuilds.r_CantidadDePropuestas(UserIndex, PAZ)
            tStr = "PEACEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(UserIndex, PAZ)
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
            Exit Sub
        Case "DECGUERR" 'declaro la guerra
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_DeclararGuerra(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                'WAR shall be!
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & " TU CLAN HA ENTRADO EN GUERRA CON " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & " " & UserList(UserIndex).Name & " LE DECLARA LA GUERRA A TU CLAN" & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "NEWWEBSI"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarWebSite(UserIndex, rData)
            Exit Sub
        Case "ACEPTARI"
            rData = Right$(rData, Len(rData) - 8)
            If Not modGuilds.a_AceptarAspirante(UserIndex, rData, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(rData)
                If tInt > 0 Then
                    Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(UserIndex).GuildIndex)
                End If
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & rData & " ha sido aceptado como miembro del clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "TW" & SONIDOS_GUILD.SND_ACEPTADOCLAN)
            End If
            Exit Sub
        Case "RECHAZAR"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If Not modGuilds.a_RechazarAspirante(UserIndex, Arg1, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & Arg3 & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(Arg1)
                tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo
                If tInt > 0 Then
                    Call SendData(SendTarget.ToIndex, tInt, 0, ServerPackages.dialogo & " " & tStr & FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
                    Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(UserIndex).GuildIndex, Arg2)
                End If
            End If
            Exit Sub
        Case "ECHARCLA"
            'el lider echa de clan a alguien
            rData = Trim$(Right$(rData, Len(rData) - 8))
            tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData)
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & rData & " fue expulsado del clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No puedes expulsar ese personaje del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACTGNEWS"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarNoticias(UserIndex, rData)
            Exit Sub
        Case "1HRINFO<"
            rData = Right$(rData, Len(rData) - 8)
            If Trim$(rData) = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesPersonaje(UserIndex, rData, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "CHRINFO" & tStr)
            End If
            Exit Sub
        Case "ABREELEC"
            If Not modGuilds.v_AbrirElecciones(UserIndex, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogo & "¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & UserList(UserIndex).Name & FONTTYPE_GUILD)
            End If
            Exit Sub
    End Select
    

    Select Case UCase$(Left$(rData, 9))
        Case "SOLICITUD"
             rData = Right$(rData, Len(rData) - 9)
             Arg1 = ReadField(1, rData, Asc(","))
             Arg2 = ReadField(2, rData, Asc(","))
             If Not modGuilds.a_NuevoAspirante(UserIndex, Arg1, Arg2, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
             Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & Arg1 & "." & FONTTYPE_GUILD)
             End If
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "CLANDETAILS"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then Exit Sub
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "CLANDET" & modGuilds.SendGuildDetails(rData))
            Exit Sub
    End Select
    
Procesado = False
    
End Sub
