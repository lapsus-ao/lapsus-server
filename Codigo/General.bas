Attribute VB_Name = "General"
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

'Global ANpc As Long
'Global Anpc_host As Long

Option Explicit

Global LeerNPCs As New clsIniReader
Global LeerNPCsHostiles As New clsIniReader
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Double
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function
Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)

Select Case UCase$(UserList(UserIndex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 21
                    Else
                        UserList(UserIndex).char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 39
                    Else
                        UserList(UserIndex).char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 32
                    Else
                        UserList(UserIndex).char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 40
                    Else
                        UserList(UserIndex).char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 53
                    Else
                        UserList(UserIndex).char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 60
                    Else
                        UserList(UserIndex).char.Body = 60
                    End If
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 53
                    Else
                        UserList(UserIndex).char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 60
                    Else
                        UserList(UserIndex).char.Body = 60
                    End If
      End Select

    Case "ELFO"
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 21
                    Else
                        UserList(UserIndex).char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 39
                    Else
                        UserList(UserIndex).char.Body = 39
                    End If
      End Select

    Case Else 'CHOTS |Raza Orco
      Select Case UCase$(UserList(UserIndex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 208
                    Else
                        UserList(UserIndex).char.Body = 208
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(UserIndex).CharMimetizado.Body = 209
                    Else
                        UserList(UserIndex).char.Body = 209
                    End If
      End Select
    
End Select

UserList(UserIndex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Sub LimpiarObjs()

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Limpiando mundo..." & FONTTYPE_SERVER)
Dim i As Integer
Dim u As Integer
Dim Y As Integer
Dim X As Integer
Dim tInt As String



For i = 1 To NumMaps
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            If (i < 62 Or (i > 62 And i < 81)) And X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(i, X, Y).OBJInfo.ObjIndex > 0 Then
                    tInt = ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).OBJType
                    If tInt <> otArboles And tInt <> otPuertas And tInt <> otCONTENEDORES And _
                        tInt <> otCARTELES And tInt <> otFOROS And tInt <> otYacimiento And _
                        tInt <> otTELEPORT And tInt <> otYunque And tInt <> otFragua And _
                        tInt <> otMANCHAS And tInt <> otMuebles Then
                        Call EraseObj(ToMap, 0, i, MapData(i, X, Y).OBJInfo.Amount, i, X, Y)
                    End If
                End If
            End If
            
        Next X
    Next Y
Next i

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Limpieza de mundo terminada." & FONTTYPE_SERVER)

End Sub

Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer


For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(SendTarget.ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i

Call SecurityIp.IpSecurityMantenimientoLista



End Sub
Public Sub hambreYSed(UserIndex As Integer)
Dim fenviarAyS As Boolean
fenviarAyS = False
'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
          UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
          UserList(UserIndex).Counters.AGUACounter = 0
          UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
                            
          If UserList(UserIndex).Stats.MinAGU <= 0 Then
               UserList(UserIndex).Stats.MinAGU = 0
               UserList(UserIndex).flags.Sed = 1
          End If
          fenviarAyS = True
    End If
End If

'hambre
If UserList(UserIndex).Stats.MinHam > 0 Then
   If UserList(UserIndex).Counters.COMCounter < Intervalohambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
        If UserList(UserIndex).Stats.MinHam <= 0 Then
               UserList(UserIndex).Stats.MinHam = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

If fenviarAyS Then Call EnviarhambreYsed(UserIndex)

End Sub
Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(SendTarget.ToIndex, UserIndex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call BanIpCargar

Prision.Map = 14
Libertad.Map = 14

Prision.X = 34
Prision.Y = 33
Libertad.X = 31
Libertad.Y = 41

'CHOTS | Escuchar Clan
Clan_ClanIndex = 0
Clan_EscuchadorIndex = 0
'CHOTS | Escuchar Clan

'CHOTS | Invocaciones
INVOCACION_INVOCADO = False
'CHOTS | Invocaciones

'CHOTS | Espiar Users
Espia_Espiador = 0
Espia_Espiado = 0
'CHOTS | Espiar Users

LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

'CHOTS - BysNacK | AntiBots
NumIps = 0


ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Parties(1 To MAX_PARTIES) As clsParty
ReDim Guilds(1 To MAX_GUILDS) As clsClan

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 6
LevelSkill(3).LevelValue = 9
LevelSkill(4).LevelValue = 12
LevelSkill(5).LevelValue = 15
LevelSkill(6).LevelValue = 18
LevelSkill(7).LevelValue = 21
LevelSkill(8).LevelValue = 24
LevelSkill(9).LevelValue = 27
LevelSkill(10).LevelValue = 30
LevelSkill(11).LevelValue = 33
LevelSkill(12).LevelValue = 36
LevelSkill(13).LevelValue = 39
LevelSkill(14).LevelValue = 42
LevelSkill(15).LevelValue = 45
LevelSkill(16).LevelValue = 48
LevelSkill(17).LevelValue = 51
LevelSkill(18).LevelValue = 54
LevelSkill(19).LevelValue = 57
LevelSkill(20).LevelValue = 60
LevelSkill(21).LevelValue = 63
LevelSkill(22).LevelValue = 66
LevelSkill(23).LevelValue = 69
LevelSkill(24).LevelValue = 72
LevelSkill(25).LevelValue = 75
LevelSkill(26).LevelValue = 78
LevelSkill(27).LevelValue = 81
LevelSkill(28).LevelValue = 84
LevelSkill(29).LevelValue = 87
LevelSkill(30).LevelValue = 90
LevelSkill(31).LevelValue = 93
LevelSkill(32).LevelValue = 96
LevelSkill(33).LevelValue = 100
LevelSkill(34).LevelValue = 100
LevelSkill(35).LevelValue = 100
LevelSkill(36).LevelValue = 100
LevelSkill(37).LevelValue = 100
LevelSkill(38).LevelValue = 100
LevelSkill(39).LevelValue = 100
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100
LevelSkill(51).LevelValue = 100
LevelSkill(52).LevelValue = 100
LevelSkill(53).LevelValue = 100
LevelSkill(54).LevelValue = 100

ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
ListaRazas(6) = "Orco" 'CHOTS |Raza Orco

Torneo_Clases_Validas(1) = "Guerrero"
Torneo_Clases_Validas(2) = "Mago"
Torneo_Clases_Validas(3) = "Paladin"
Torneo_Clases_Validas(4) = "Clerigo"
Torneo_Clases_Validas(5) = "Bardo"
Torneo_Clases_Validas(6) = "Asesino"
Torneo_Clases_Validas(7) = "Druida"
Torneo_Clases_Validas(8) = "Cazador"

Torneo_Alineacion_Validas(1) = "Criminal"
Torneo_Alineacion_Validas(2) = "Ciudadano"
Torneo_Alineacion_Validas(3) = "Armada CAOS"
Torneo_Alineacion_Validas(4) = "Armada REAL"

ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "Alquimia"
SkillsNames(23) = "Sastreria"
SkillsNames(24) = "Botanica"

frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.caption = frmMain.caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).caption = "Iniciando Arrays..."

Call LoadGuildsDB

'CHOTS | Nueva randomnumber obtenida de GS
Call RandomNumberInitialize

'CHOTS | Inicializa las salas de reto
Call InicializarMapasReto

Call CargarSpawnList
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargarRanking 'CHOTS | Sistema de Ranking
Call CargaApuestas

'CHOTS | Guerras
Call inicializarSalasGuerra
Call inicializarGuerras

'*************************************************
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).caption = "Cargando Objetos"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).caption = "Cargando Hechizos"
Call CargarHechizos

frmCargando.Label1(2).caption = "Cargando Castillos"
Call CargarCastillos
    
frmCargando.Label1(2).caption = "Cargando Objetos Extras"
Call CargarHechizos
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero
Call LoadObjSastre
Call LoadObjDruida
Call LoadArmadaCaos

If BootDelBackUp Then
    
    frmCargando.Label1(2).caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo

'CHOTS | Monturas
cargarMapasMonturas
RespawnearMonturas
'CHOTS | Monturas

'CHOTS | Seguridad
Call inicializarSeguridad


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .AutoSave.Enabled = True
    .tPiqueteC.Enabled = True
    .Timer1.Enabled = True
    .GameTimer.Enabled = True
    .NuevoGameTimer.Enabled = True
    .tmrInvocaciones.Enabled = True
    .tmrSegundosGuerra.Enabled = True
    .tmrMinutosGuerra.Enabled = True
    .tmrCastillos.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .KillLog.Enabled = False
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Unload frmCargando

'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #n

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
'Call InicializaEstadisticas

End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(File, FileType) <> ""
End Function

Function ReadField(ByVal Pos As Integer, ByVal text As String, ByVal SepASCII As Integer) As String

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(text)
    CurChar = mid$(text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid$(text, LastPos + 1, (InStr(LastPos + 1, text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = Pos Then
    ReadField = mid$(text, LastPos + 1)
End If

End Function

Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub



Public Sub LogGM(nombre As String, texto As String, Consejero As Boolean)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If Consejero Then
    Open App.Path & "\logs\consejeros\" & nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
''On Error GoTo errhandler
''
''Dim nfile As Integer
''nfile = FreeFile ' obtenemos un canal
''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
''
''Print #nfile, "<stats>"
''Print #nfile, "<ao>"
''Print #nfile, "<dia>" & Date & "</dia>"
''Print #nfile, "<hora>" & Time & "</hora>"
''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
''Print #nfile, "</ao>"
''Print #nfile, "</stats>"
''
''
''Close #nfile
Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.caption = "Reiniciando."

Dim LoopC As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

ReDim UserList(1 To MaxUsers)

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As Npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."

'Log it
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " servidor reiniciado."
Close #n

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    If MapInfo(UserList(UserIndex).Pos.Map).zona <> "DUNGEON" Then
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)

Dim modifi As Integer
Dim tieneCero As Boolean
tieneCero = False 'CHOTS | Evitamos paquetes innecesarios

If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
  UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
Else
  If MapInfo(UserList(UserIndex).Pos.Map).Terreno = Dungeon Then
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z82")
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
    If UserList(UserIndex).Stats.MinHP < 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z83")
            UserList(UserIndex).Stats.MinHP = 0
            Call UserDie(UserIndex)
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)
  Else
    modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
    If UserList(UserIndex).Stats.MinSta = 0 Then tieneCero = True
    Call QuitarSta(UserIndex, modifi)
    If Not tieneCero Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASS" & UserList(UserIndex).Stats.MinSta)
  End If
  
  UserList(UserIndex).Counters.Frio = 0
  
  
End If

End Sub

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(UserIndex).Counters.Mimetismo = UserList(UserIndex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        
    
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
    Call ChangeUserChar(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
End If

End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

If EstaEnCastillo(UserIndex) Or UserList(UserIndex).Reto.enReto = True Then
    Call QuitarInvisibilidad(UserIndex)
    Exit Sub
End If

If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
Else
    Call QuitarInvisibilidad(UserIndex)
End If

End Sub

Public Sub QuitarInvisibilidad(ByVal UserIndex As Integer)
    UserList(UserIndex).Counters.Invisibilidad = 0
    UserList(UserIndex).flags.Invisible = 0
    If UserList(UserIndex).flags.Oculto = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
        Dim ChotsNover As String
        ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
    End If
End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NSEGUE")
    End If
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "NESTUP")
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Paralisis > 0 Then
    UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
Else
    UserList(UserIndex).flags.Paralizado = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya no Estas Paralizado" & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
End If

End Sub

Public Sub RecStamina(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

If UserList(UserIndex).flags.Desnudo = 1 Then Exit Sub

Dim massta As Integer
If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
   If UserList(UserIndex).Counters.STACounter < Intervalo Then
       UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
   Else
       EnviarStats = True
       UserList(UserIndex).Counters.STACounter = 0
       massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 6))
       UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta
       If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(UserIndex As Integer)
Dim n As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z35")
  UserList(UserIndex).Counters.Veneno = 0
  n = RandomNumber(1, 5)
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - n
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  Call SendData(SendTarget.ToIndex, UserIndex, 0, "ASH" & UserList(UserIndex).Stats.MinHP)
End If

End Sub

Public Sub DuracionPociones(UserIndex As Integer)

'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto > 0 Then
   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
   If UserList(UserIndex).flags.DuracionEfecto = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z46")
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
        UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
        Call EnviarDopa(UserIndex)
   End If
End If

End Sub

Public Sub Sanar(UserIndex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
   If UserList(UserIndex).Counters.HPCounter < Intervalo Then
      UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
   Else
      mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
                           
      UserList(UserIndex).Counters.HPCounter = 0
      UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
      If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z36")
      EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
Call LeerNPCs.Initialize(npcfile)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Call LeerNPCsHostiles.Initialize(npcfile)

End Sub

Sub PasarSegundo()
    Dim i As Integer
    For i = 1 To LastUser

    'CHOTS | /HOGAR por tiempo
    If UserList(i).Counters.Hogar > 0 Then
        UserList(i).Counters.Hogar = UserList(i).Counters.Hogar - 1
        
        If UserList(i).Counters.Hogar = 0 Then
            Dim HogarTemp As WorldPos
        
            If UCase$(UserList(i).Hogar) = "NIX" Then
                HogarTemp = Nix
            ElseIf UCase$(UserList(i).Hogar) = "ULLATHORPE" Then
                HogarTemp = Ullathorpe
            ElseIf UCase$(UserList(i).Hogar) = "BANDERBILL" Then
                HogarTemp = Banderbill
            ElseIf UCase$(UserList(i).Hogar) = "LINDOS" Then
                HogarTemp = Lindos
            End If
            
            If LegalPos(HogarTemp.Map, HogarTemp.X, HogarTemp.Y, False) Then
                Call WarpUserChar(i, HogarTemp.Map, HogarTemp.X, HogarTemp.Y, True)
            Else
                Dim NuevaPos As WorldPos
                Call ClosestLegalPos(HogarTemp, NuevaPos)
                Call WarpUserChar(i, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
            End If
            
            Call SendData(SendTarget.ToIndex, i, 0, ServerPackages.dialogo & "Has vuelto a tu hogar!" & FONTTYPE_HOGAR)
            Call SendData(SendTarget.ToIndex, i, 0, "DOK")
        
        End If
        
    End If
    'CHOTS | /HOGAR por tiempo

    'Cerrar usuario
    If UserList(i).Counters.Saliendo Then
        UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
        If UserList(i).Counters.Salir <= 0 Then
            Call SendData(SendTarget.ToIndex, i, 0, ServerPackages.logout)
            Call CloseSocket(i)
            Exit Sub
        End If
    
    'ANTIEMPOLLOS
    ElseIf UserList(i).flags.EstaEmpo = 1 Then
         UserList(i).EmpoCont = UserList(i).EmpoCont + 1
         If UserList(i).EmpoCont = 20 Then
             
             Call SendData(SendTarget.ToIndex, i, 0, "!! Fuiste expulsado por permanecer muerto sobre un item")
             
             UserList(i).EmpoCont = 0
             Call CloseSocket(i)
             Exit Sub
         ElseIf UserList(i).EmpoCont = 10 Then
             Call SendData(SendTarget.ToIndex, i, 0, ServerPackages.dialogo & "LLevas 10 segundos bloqueando el item, muévete o serás desconectado." & FONTTYPE_WARNING)
         End If
    End If
    
    'Bysnack | Retos
    If UserList(i).Reto.GanoReto = True Or UserList(i).Reto.PerdioReto = True Then
        If UserList(i).Reto.TimeReto > 0 Then
            UserList(i).Reto.TimeReto = UserList(i).Reto.TimeReto - 1
        Else
            Dim Pos As WorldPos
            Dim nPos As WorldPos
            Pos.Map = 1
            Pos.X = 58
            Pos.Y = 45
    
            Call ClosestLegalPos(Pos, nPos)
            Call WarpUserChar(i, nPos.Map, nPos.X, nPos.Y, True)
    
            UserList(i).Reto.GanoReto = False
            UserList(i).Reto.PerdioReto = False
        End If
    End If
    
Next i
    
    If CuentaRegresiva > 0 Then
        If CuentaRegresiva > 1 Then
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & CuentaRegresiva - 1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "YA!!!" & FONTTYPE_FIGHT)
        End If
        CuentaRegresiva = CuentaRegresiva - 1
    End If
    
    'CHOTS | Torneos Automáticos
    If Torneo_CR.segundos > 0 Then
        Torneo_CR.segundos = Torneo_CR.segundos - 1
        If Torneo_CR.segundos = 0 Then finalizarCuenta
    End If
    
    
    If Torneo_CuentaPelea > 0 Then
    
        Torneo_CuentaPelea = Torneo_CuentaPelea - 1
        
        If Torneo_CuentaPelea > 0 Then
            Call SendData(SendTarget.ToMap, 0, Torneo_MAPATORNEO, ServerPackages.dialogo & Torneo_CuentaPelea & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToMap, 0, Torneo_MAPATORNEO, ServerPackages.dialogo & "YA!!!" & FONTTYPE_FIGHT)
            MapInfo(Torneo_MAPATORNEO).Pk = True
        End If
        
        
    End If
    'CHOTS | Torneos Automáticos

End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Grabando Personajes" & FONTTYPE_SERVER)
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Personajes Grabados" & FONTTYPE_SERVER)
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")

    haciendoBK = False
End Sub

Public Function Tilde(data As String) As String
 
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function
Public Function DamePos(ByRef original_Pos As WorldPos) As WorldPos
 
'
' @ Devuelve un tile libre.
 
Dim iRange      As Long
Dim iX          As Long
Dim iY          As Long
Dim now_Index  As Integer
Dim no_User    As Boolean
Dim not_Pos    As WorldPos
 
not_Pos = original_Pos
DamePos.Map = original_Pos.Map
 
With original_Pos
    For iRange = 1 To 3
        For iX = (.X - iRange) To (.X + iRange)
            For iY = (.Y - iRange) To (.Y + iRange)
               
                now_Index = MapData(.Map, iX, iY).UserIndex
               
                'No hay n usuario
                If (now_Index = 0) Then
                    DamePos.X = iX
                    DamePos.Y = iY
                    no_User = True
                End If
               
                'No hay usuario, revisa npc
                If (no_User = True) Then
                    now_Index = MapData(.Map, iX, iY).NpcIndex
                 
                    'No hay un npc.
                    If (now_Index = 0) Then
                      DamePos.X = iX
                      DamePos.Y = iY
                      Exit Function
                    Else
                      no_User = False
                    End If
                End If
 
            Next iY
        Next iX
    Next iRange
End With
 
'Llega acá, devuelve la posición original.
DamePos = not_Pos
 
End Function
Public Sub IntercambiarObjetos(ByVal UserIndex As Integer, ByVal ObjAMover1 As Byte, ByVal ObjAMover2 As Byte)
Dim tmpUserObj As UserOBJ
 
    With UserList(UserIndex)
               
        'Cambiamos si alguno es una herramienta
        If .Invent.HerramientaEqpSlot = ObjAMover1 Then
            .Invent.HerramientaEqpSlot = ObjAMover2
        ElseIf .Invent.HerramientaEqpSlot = ObjAMover2 Then
            .Invent.HerramientaEqpSlot = ObjAMover1
        End If
       
        'Cambiamos si alguno es un armor
        If .Invent.ArmourEqpSlot = ObjAMover1 Then
            .Invent.ArmourEqpSlot = ObjAMover2
        ElseIf .Invent.ArmourEqpSlot = ObjAMover2 Then
            .Invent.ArmourEqpSlot = ObjAMover1
        End If
       
        'Cambiamos si alguno es un barco
        If .Invent.BarcoSlot = ObjAMover1 Then
            .Invent.BarcoSlot = ObjAMover2
        ElseIf .Invent.BarcoSlot = ObjAMover2 Then
            .Invent.BarcoSlot = ObjAMover1
        End If
       
        'Cambiamos si alguno es un casco
        If .Invent.CascoEqpSlot = ObjAMover1 Then
            .Invent.CascoEqpSlot = ObjAMover2
        ElseIf .Invent.CascoEqpSlot = ObjAMover2 Then
            .Invent.CascoEqpSlot = ObjAMover1
        End If
       
        'Cambiamos si alguno es un escudo
        If .Invent.EscudoEqpSlot = ObjAMover1 Then
            .Invent.EscudoEqpSlot = ObjAMover2
        ElseIf .Invent.EscudoEqpSlot = ObjAMover2 Then
            .Invent.EscudoEqpSlot = ObjAMover1
        End If
       
        'Cambiamos si alguno es munición
        If .Invent.MunicionEqpSlot = ObjAMover1 Then
            .Invent.MunicionEqpSlot = ObjAMover2
        ElseIf .Invent.MunicionEqpSlot = ObjAMover2 Then
            .Invent.MunicionEqpSlot = ObjAMover1
        End If
       
        'Cambiamos si alguno es un arma
        If .Invent.WeaponEqpSlot = ObjAMover1 Then
            .Invent.WeaponEqpSlot = ObjAMover2
        ElseIf .Invent.WeaponEqpSlot = ObjAMover2 Then
            .Invent.WeaponEqpSlot = ObjAMover1
        End If
       
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjAMover1)
        .Invent.Object(ObjAMover1) = .Invent.Object(ObjAMover2)
        .Invent.Object(ObjAMover2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, UserIndex, ObjAMover1)
        Call UpdateUserInv(False, UserIndex, ObjAMover2)
    End With
End Sub
