Attribute VB_Name = "TCP"
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

'RUTAS DE ENVIO DE DATOS
Public Enum SendTarget
    ToIndex = 0         'Envia a un solo User
    ToAll = 1           'A todos los Users
    ToMap = 2           'Todos los Usuarios en el mapa
    ToPCArea = 3        'Todos los Users en el area de un user determinado
    ToNone = 4          'Ninguno
    ToAllButIndex = 5   'Todos menos el index
    ToMapButIndex = 6   'Todos en el mapa menos el indice
    ToGM = 7
    ToNPCArea = 8       'Todos los Users en el area de un user determinado
    ToGuildMembers = 9
    ToAdmins = 10
    ToPCAreaButIndex = 11
    ToAdminsAreaButConsejeros = 12
    ToDiosesYclan = 13
    ToConsejo = 14
    ToClanArea = 15
    ToConsejoCaos = 16
    ToRolesMasters = 17
    ToDeadArea = 18
    ToCiudadanos = 19
    ToCriminales = 20
    ToPartyArea = 21
    ToReal = 22
    ToCaos = 23
    ToCiudadanosYRMs = 24
    ToCriminalesYRMs = 25
    ToRealYRMs = 26
    ToCaosYRMs = 27
    ToDioses = 28
End Enum


#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpoYCabeza(ByRef UserBody As Integer, ByRef UserHead As Integer, ByVal Raza As String, ByVal Gen As String)
'TODO: Poner las heads en arrays, así se acceden por índices
'y no hay problemas de discontinuidad de los índices.
'También se debe usar enums para raza y sexo
Select Case Gen
   Case "Hombre"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 13)
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(102, 106)
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(201, 208)
                UserBody = 3
            Case "Enano"
                UserHead = RandomNumber(301, 305)
                UserBody = 52
            Case "Gnomo"
                UserHead = RandomNumber(401, 406)
                UserBody = 52
            Case "Orco" 'CHOTS |Raza Orco
                UserHead = RandomNumber(503, 507)
                UserBody = 210
            Case Else
                UserHead = 1
                UserBody = 1
        End Select
   Case "Mujer"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(70, 75)
                UserBody = 1
            Case "Elfo"
                UserHead = RandomNumber(170, 174)
                UserBody = 2
            Case "Elfo Oscuro"
                UserHead = RandomNumber(270, 276)
                UserBody = 3
            Case "Gnomo"
                UserHead = RandomNumber(470, 475)
                UserBody = 52
            Case "Enano"
                UserHead = RandomNumber(370, 374)
                UserBody = 52
            Case "Orco" 'CHOTS |Raza Orco
                UserHead = RandomNumber(550, 551)
                UserBody = 212
            Case Else
                UserHead = 70
                UserBody = 1
        End Select
End Select

End Sub
Sub DarCabeza(ByVal UserIndex As Integer, ByVal Raza As String, ByVal Gen As String)
'CHOTS | Modulo creado para ser utilizado en el sistema de Cirujia

Dim UserHead As Integer
UserHead = 1

Select Case Gen
   Case "Hombre"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(1, 13)
            Case "Elfo"
                UserHead = RandomNumber(102, 106)
            Case "Elfo Oscuro"
                UserHead = RandomNumber(201, 208)
            Case "Enano"
                UserHead = RandomNumber(301, 305)
            Case "Gnomo"
                UserHead = RandomNumber(401, 406)
            Case "Orco" 'CHOTS |Raza Orco
                UserHead = RandomNumber(503, 507)
            Case Else
                UserHead = 1
        End Select
   Case "Mujer"
        Select Case Raza
            Case "Humano"
                UserHead = RandomNumber(70, 75)
            Case "Elfo"
                UserHead = RandomNumber(170, 174)
            Case "Elfo Oscuro"
                UserHead = RandomNumber(270, 276)
            Case "Gnomo"
                UserHead = RandomNumber(470, 475)
            Case "Enano"
                UserHead = RandomNumber(370, 374)
            Case "Orco" 'CHOTS |Raza Orco
                UserHead = RandomNumber(550, 551)
            Case Else
                UserHead = 70
        End Select
End Select


'CHOTS | Graba la Head
UserList(UserIndex).char.Head = UserHead
UserList(UserIndex).OrigChar.Head = UserHead
Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Head", UserHead)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
'CHOTS | Graba la Head

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 241) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

'Barrin 3/3/03
'Agregué PadrinoName y Padrino password como opcionales, que se les da un valor siempre y cuando el servidor esté usando el sistema
Sub ConnectNewUser(UserIndex As Integer, Name As String, Password As String, UserRaza As String, UserSexo As String, UserClase As String, _
                    US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
                    US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
                    US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
                    US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
                    US21 As String, US22 As String, US23 As String, US24 As String, UserEmail As String, Hogar As String, Preg As String, Resp As String)

If Not AsciiValidos(Name) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNombre invalido.")
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long

'¿Existe el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

'Tiró los dados antes de llegar acá??
If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRDebe tirar los dados antes de poder crear un personaje.")
    Exit Sub
End If

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).flags.Escondido = 0

UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 1000
UserList(UserIndex).Reputacion.PlebeRep = 30

UserList(UserIndex).Reputacion.Promedio = 30 / 6

UserList(UserIndex).Name = Name
UserList(UserIndex).Clase = UserClase
UserList(UserIndex).Raza = UserRaza
UserList(UserIndex).Genero = UserSexo
UserList(UserIndex).email = UserEmail
UserList(UserIndex).Preg = Preg
UserList(UserIndex).Resp = Resp
UserList(UserIndex).Hogar = Hogar
UserList(UserIndex).Pareja = ""
UserList(UserIndex).Puntos = 0

With UserList(UserIndex).Stats
    Select Case UCase$(UserRaza)
    
        Case "HUMANO"
            .UserAtributos(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia) + 1
            .UserAtributos(eAtributos.Carisma) = .UserAtributos(eAtributos.Carisma) + 1
            .UserAtributos(eAtributos.Constitucion) = .UserAtributos(eAtributos.Constitucion) + 2
        Case "ELFO"
            .UserAtributos(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza) - 1
            .UserAtributos(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad) + 3
            .UserAtributos(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia) + 3
            .UserAtributos(eAtributos.Carisma) = .UserAtributos(eAtributos.Carisma) + 2
            .UserAtributos(eAtributos.Constitucion) = .UserAtributos(eAtributos.Constitucion) + 1
        Case "ELFO OSCURO"
            .UserAtributos(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza) + 1
            .UserAtributos(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad) + 1
            .UserAtributos(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia) + 3
            .UserAtributos(eAtributos.Carisma) = .UserAtributos(eAtributos.Carisma) - 2
        Case "ENANO"
            .UserAtributos(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza) + 2
            .UserAtributos(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad) - 2
            .UserAtributos(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia) - 4
            .UserAtributos(eAtributos.Carisma) = .UserAtributos(eAtributos.Carisma) - 1
            .UserAtributos(eAtributos.Constitucion) = .UserAtributos(eAtributos.Constitucion) + 3
        Case "GNOMO"
            .UserAtributos(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza) - 2
            .UserAtributos(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad) + 2
            .UserAtributos(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia) + 5
        Case "ORCO" 'CHOTS | Raza Orco
            .UserAtributos(eAtributos.Fuerza) = .UserAtributos(eAtributos.Fuerza) + 3
            .UserAtributos(eAtributos.Agilidad) = .UserAtributos(eAtributos.Agilidad) - 3
            .UserAtributos(eAtributos.Inteligencia) = .UserAtributos(eAtributos.Inteligencia) - 7
            .UserAtributos(eAtributos.Carisma) = .UserAtributos(eAtributos.Carisma) - 3
            .UserAtributos(eAtributos.Constitucion) = .UserAtributos(eAtributos.Constitucion) + 4
    End Select
End With

If EsDios(Name) Or EsSemiDios(Name) Or EsOT(Name) Then
    With UserList(UserIndex).Stats
        .UserSkills(1) = 100
        .UserSkills(2) = 100
        .UserSkills(3) = 100
        .UserSkills(4) = 100
        .UserSkills(5) = 100
        .UserSkills(6) = 100
        .UserSkills(7) = 100
        .UserSkills(8) = 100
        .UserSkills(9) = 100
        .UserSkills(10) = 100
        .UserSkills(11) = 100
        .UserSkills(12) = 100
        .UserSkills(13) = 100
        .UserSkills(14) = 100
        .UserSkills(15) = 100
        .UserSkills(16) = 100
        .UserSkills(17) = 100
        .UserSkills(18) = 100
        .UserSkills(19) = 100
        .UserSkills(20) = 100
        .UserSkills(21) = 100
        .UserSkills(22) = 100
        .UserSkills(23) = 100
        .UserSkills(24) = 100
    End With
Else
    With UserList(UserIndex).Stats
        .UserSkills(1) = val(US1)
        .UserSkills(2) = val(US2)
        .UserSkills(3) = val(US3)
        .UserSkills(4) = val(US4)
        .UserSkills(5) = val(US5)
        .UserSkills(6) = val(US6)
        .UserSkills(7) = val(US7)
        .UserSkills(8) = val(US8)
        .UserSkills(9) = val(US9)
        .UserSkills(10) = val(US10)
        .UserSkills(11) = val(US11)
        .UserSkills(12) = val(US12)
        .UserSkills(13) = val(US13)
        .UserSkills(14) = val(US14)
        .UserSkills(15) = val(US15)
        .UserSkills(16) = val(US16)
        .UserSkills(17) = val(US17)
        .UserSkills(18) = val(US18)
        .UserSkills(19) = val(US19)
        .UserSkills(20) = val(US20)
        .UserSkills(21) = val(US21)
        .UserSkills(22) = val(US22)
        .UserSkills(23) = val(US23)
        .UserSkills(24) = val(US24)
    End With
    totalskpts = 0
    
    'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
    For LoopC = 1 To NUMSKILLS
        totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
    Next LoopC
    
    If totalskpts > 10 Then
        Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
        Call BorrarUsuario(UserList(UserIndex).Name)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

End If


'CHOTS | Seguridad anti fragueo
UserList(UserIndex).flags.LastCiudMatado = "NADIE"
UserList(UserIndex).flags.LastCrimMatado = "NADIE"
'CHOTS | Seguridad anti fragueo

'CHOTS | Seguridad anti bots
If Len(Name) > 20 Then
    Call LogHackAttemp(Name & " Intentó crear un Bot.")
    Call BorrarUsuario(UserList(UserIndex).Name)
    Call CloseSocket(UserIndex)
    Exit Sub
End If
'CHOTS | Seguridad anti bots

UserList(UserIndex).Password = Password
UserList(UserIndex).char.Heading = eHeading.SOUTH

Call DarCuerpoYCabeza(UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)

UserList(UserIndex).OrigChar = UserList(UserIndex).char

UserList(UserIndex).char.ShieldAnim = NingunEscudo
UserList(UserIndex).char.CascoAnim = NingunCasco

Dim MiInt As Long
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt

UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100
UserList(UserIndex).Stats.TrofOro = 0
UserList(UserIndex).Stats.TorneosAuto = 0
UserList(UserIndex).Stats.TrofPlata = 0
UserList(UserIndex).Stats.TrofBronce = 0
UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100

'CHOTS | Quests
UserList(UserIndex).Quest.cantMatados = 0
UserList(UserIndex).Quest.nroQuest = 0
UserList(UserIndex).Quest.numNpc = 0
UserList(UserIndex).Quest.cantNpc = 0


'<-----------------MANA----------------------->
If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "DRUIDA" Then
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)) / 3
    UserList(UserIndex).Stats.MaxMAN = 100 + MiInt
    UserList(UserIndex).Stats.MinMAN = 100 + MiInt
ElseIf UCase$(UserClase) = "CLERIGO" _
    Or UCase$(UserClase) = "BARDO" Or UCase$(UserClase) = "ASESINO" Then
        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50
Else
    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0
End If


If EsDios(Name) Or EsSemiDios(Name) Or EsOT(Name) Then
    UserList(UserIndex).Stats.UserHechizos(1) = 24
    UserList(UserIndex).Stats.UserHechizos(2) = 10
    UserList(UserIndex).Stats.UserHechizos(3) = 11
Else
    If UCase$(UserClase) = "MAGO" Or UCase$(UserClase) = "CLERIGO" Or _
        UCase$(UserClase) = "DRUIDA" Or UCase$(UserClase) = "BARDO" Or _
        UCase$(UserClase) = "ASESINO" Then
        UserList(UserIndex).Stats.UserHechizos(1) = 2
    End If
End If

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1

UserList(UserIndex).Stats.GLD = 0

UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.ELU = 300
UserList(UserIndex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(UserIndex).Invent.NroItems = 4

UserList(UserIndex).Invent.Object(1).ObjIndex = 467
UserList(UserIndex).Invent.Object(1).Amount = 100

UserList(UserIndex).Invent.Object(2).ObjIndex = 468
UserList(UserIndex).Invent.Object(2).Amount = 100

UserList(UserIndex).Invent.Object(3).ObjIndex = 460
UserList(UserIndex).Invent.Object(3).Amount = 1
UserList(UserIndex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case "Humano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    Case "Orco" 'CHOTS | Raza Orco
        UserList(UserIndex).Invent.Object(4).ObjIndex = 737
End Select

UserList(UserIndex).Invent.Object(4).Amount = 1
UserList(UserIndex).Invent.Object(4).Equipped = 1

UserList(UserIndex).Invent.ArmourEqpSlot = 4
UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).ObjIndex

UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).ObjIndex
UserList(UserIndex).Invent.WeaponEqpSlot = 3

' CHOTS | Se crea y tiene la daga equipada
UserList(UserIndex).char.WeaponAnim = 9

Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")
  
'Open User
Call ConnectUser(UserIndex, Name, Password)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)
Dim LoopC As Integer

On Error GoTo errhandler
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
    
    If UserList(UserIndex).Reto.enReto Then
        If UserList(UserIndex).Reto.Pareja <> "" Then Call UserDie(NameIndex(UserList(UserIndex).Reto.Pareja))
        Call UserDie(UserIndex)
    End If
    
    If UserList(UserIndex).Reto.EsperandoReto Then Call ResetFlagsReto(UserIndex)
    
    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, ServerPackages.dialogo & "Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
        
        'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Else
        Call ResetUserSlot(UserIndex)
    End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    
Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    Call ResetUserSlot(UserIndex)
    
#If UsarQueSocket = 1 Then
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
#End If

    Call LogError("CloseSocket - Error = " & Err.number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal UserIndex As Integer)
On Error GoTo errhandler
    
    UserList(UserIndex).ConnID = -1

    If UserIndex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
    End If

    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    Call ResetUserSlot(UserIndex)

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    Call ResetUserSlot(UserIndex)
End Sub


#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(UserIndex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    UserList(UserIndex).ConnID = -1 'inabilitamos operaciones en socket

    If UserIndex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).ConnID = -1
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call CloseUser(UserIndex)
    End If
    
    Call ResetUserSlot(UserIndex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    Call LogError("CLOSESOCKETERR: " & Err.Description & " UI:" & UserIndex)
    
    If Not NURestados Then
        If UserList(UserIndex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(UserIndex).Name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(UserIndex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)

#If UsarQueSocket = 1 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call BorraSlotSock(UserList(UserIndex).ConnID)
    Call WSApiCloseSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#End If
End Sub

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, Datos As String) As Long

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    
    
    Ret = WsApiEnviar(UserIndex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
    EnviarDatosASlot = Ret
    Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************

    Dim Encolar As Boolean
    Encolar = False
    
    EnviarDatosASlot = 0
    
    If UserList(UserIndex).ColaSalida.Count <= 0 Then
        If frmMain.Socket2(UserIndex).Write(Datos, Len(Datos)) < 0 Then
            If frmMain.Socket2(UserIndex).LastError = WSAEWOULDBLOCK Then
                Encolar = True
            Else
                Call Cerrar_Usuario(UserIndex)
            End If
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(UserIndex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 2 Then '**********************************************

Dim Encolar As Boolean
Dim Ret As Long
    
    Encolar = False
    
    '//
    '// Valores de retorno:
    '//                     0: Todo OK
    '//                     1: WSAEWOULDBLOCK
    '//                     2: Error critico
    '//
    If UserList(UserIndex).ColaSalida.Count <= 0 Then
        Ret = frmMain.Serv.enviar(UserList(UserIndex).ConnID, Datos, Len(Datos))
        If Ret = 1 Then
            Encolar = True
        ElseIf Ret = 2 Then
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
        End If
    Else
        Encolar = True
    End If
    
    If Encolar Then
        Debug.Print "Encolando..."
        UserList(UserIndex).ColaSalida.Add Datos
    End If

#ElseIf UsarQueSocket = 3 Then
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(UserIndex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.enviar(UserList(UserIndex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(UserIndex, True)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)
#End If '**********************************************

End Function

Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)

On Error Resume Next

Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer

sndData = sndData & ENDC

Select Case sndRoute

    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    
    Case SendTarget.ToIndex
        If UserList(sndIndex).ConnID <> -1 Then
            Call EnviarDatosASlot(sndIndex, sndData)
            Exit Sub
        End If


    Case SendTarget.ToNone
        Exit Sub
        
        
        
    Case SendTarget.ToDioses
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.Privilegios = PlayerType.Dios Then
                    Call EnviarDatosASlot(LoopC, sndData)
               End If
            End If
        Next LoopC
        Exit Sub
        
        
    Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case SendTarget.ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).Pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).Pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
            
    Case SendTarget.ToGuildMembers
        
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub


    Case SendTarget.ToDeadArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                        If UserList(MapData(sndMap, X, Y).UserIndex).flags.Muerto = 1 Or UserList(MapData(sndMap, X, Y).UserIndex).flags.Privilegios >= 1 Then
                           If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                           End If
                        End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).UserIndex > 0) And (MapData(sndMap, X, Y).UserIndex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
       
    Case SendTarget.ToClanArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).UserIndex > 0) Then
                        If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            If UserList(sndIndex).GuildIndex > 0 And UserList(MapData(sndMap, X, Y).UserIndex).GuildIndex = UserList(sndIndex).GuildIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                            End If
                        End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub



    Case SendTarget.ToPartyArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).UserIndex > 0) Then
                        If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            If UserList(sndIndex).PartyIndex > 0 And UserList(MapData(sndMap, X, Y).UserIndex).PartyIndex = UserList(sndIndex).PartyIndex Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                            End If
                        End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
        
    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).UserIndex > 0) And (MapData(sndMap, X, Y).UserIndex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                            End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).Pos.Y - MinYBorder + 1 To Npclist(sndIndex).Pos.Y + MinYBorder - 1
            For X = Npclist(sndIndex).Pos.X - MinXBorder + 1 To Npclist(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, sndData)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    Case SendTarget.ToDiosesYclan
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, sndData)
            End If
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub

    Case SendTarget.ToConsejo
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlCons > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    Case SendTarget.ToConsejoCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.PertAlConsCaos > 0 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    Case SendTarget.ToRolesMasters
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCiudadanos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCriminales
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToReal
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToCaos
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToCiudadanosYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Not Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToCriminalesYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If Criminal(LoopC) Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToRealYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.ArmadaReal = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToCaosYRMs
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).Faccion.FuerzasCaos = 1 Or UserList(LoopC).flags.EsRolesMaster Then
                    Call EnviarDatosASlot(LoopC, sndData)
                End If
            End If
        Next LoopC
        Exit Sub
End Select

End Sub

#If SeguridadAlkon Then

Sub SendCryptedMoveChar(ByVal Map As Integer, ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim LoopC As Integer

    For LoopC = 1 To LastUser
        If UserList(LoopC).Pos.Map = Map Then
            If LoopC <> UserIndex Then
                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, ServerPackages.moverChar & Encriptacion.MoveCharCrypt(LoopC, UserList(UserIndex).char.CharIndex, X, Y) & ENDC)
                End If
            End If
        End If
    Next LoopC
    Exit Sub
    

End Sub

Sub SendCryptedData(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal sndData As String)
'No puse un optional parameter en senddata porque no estoy seguro
'como afecta la performance un parametro opcional
'Prefiero 1K mas de exe que arriesgar performance
On Error Resume Next

Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer


Select Case sndRoute


    Case SendTarget.ToNone
        Exit Sub
        
    Case SendTarget.ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
'               If EsDios(UserList(LoopC).Name) Or EsSemiDios(UserList(LoopC).Name) Then
                If UserList(LoopC).flags.Privilegios > 0 Then
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case SendTarget.ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID <> -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).Pos.Map = sndMap Then
                        Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case SendTarget.ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID <> -1) And LoopC <> sndIndex Then
                If UserList(LoopC).Pos.Map = sndMap Then
                    Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case SendTarget.ToGuildMembers
    
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend
        
        Exit Sub
    
    Case SendTarget.ToPCArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, ProtoCrypt(sndData, MapData(sndMap, X, Y).UserIndex) & ENDC)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    '[Alejo-18-5]
    Case SendTarget.ToPCAreaButIndex
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).UserIndex > 0) And (MapData(sndMap, X, Y).UserIndex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, ProtoCrypt(sndData, MapData(sndMap, X, Y).UserIndex) & ENDC)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    '[CDT 17-02-2004]
    Case SendTarget.ToAdminsAreaButConsejeros
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If (MapData(sndMap, X, Y).UserIndex > 0) And (MapData(sndMap, X, Y).UserIndex <> sndIndex) Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            If UserList(MapData(sndMap, X, Y).UserIndex).flags.Privilegios > 1 Then
                                Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, ProtoCrypt(sndData, MapData(sndMap, X, Y).UserIndex) & ENDC)
                            End If
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    '[/CDT]

    Case SendTarget.ToNPCArea
        For Y = Npclist(sndIndex).Pos.Y - MinYBorder + 1 To Npclist(sndIndex).Pos.Y + MinYBorder - 1
            For X = Npclist(sndIndex).Pos.X - MinXBorder + 1 To Npclist(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID <> -1 Then
                            Call EnviarDatosASlot(MapData(sndMap, X, Y).UserIndex, ProtoCrypt(sndData, MapData(sndMap, X, Y).UserIndex) & ENDC)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub

    Case SendTarget.ToIndex
        If UserList(sndIndex).ConnID <> -1 Then
             Call EnviarDatosASlot(sndIndex, ProtoCrypt(sndData, sndIndex) & ENDC)
             Exit Sub
        End If
    Case SendTarget.ToDiosesYclan
        
        LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
        Wend

        LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        While LoopC > 0
            If (UserList(LoopC).ConnID <> -1) Then
                Call EnviarDatosASlot(LoopC, ProtoCrypt(sndData, LoopC) & ENDC)
            End If
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
        Wend

        Exit Sub
        

End Select

End Sub

#End If

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).char.Head = 0 Then UserList(UserIndex).char.Head = 1
If UserList(UserIndex).char.Body = 0 Then UserList(UserIndex).char.Body = 1

ValidateChr = UserList(UserIndex).char.Head <> 0 _
                And UserList(UserIndex).char.Body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, Name As String, Password As String)
Dim n As Integer
Dim tStr As String
Dim Morira As Byte 'CHOTS | Si está en Castillo y no tiene clan lo hacemos cagar (o si deslogea en torneo auto)
Dim motivo As String
Morira = 0

'Reseteamos los FLAGS
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.TargetNPC = 0
UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).flags.Ofrecio = 0
UserList(UserIndex).char.FX = 0

UserList(UserIndex).Reto.AceptoReto = False
UserList(UserIndex).Reto.enReto = False
UserList(UserIndex).Reto.PerdioReto = False
UserList(UserIndex).Reto.Oponente = ""
UserList(UserIndex).Reto.Pareja = ""
UserList(UserIndex).Reto.GanoReto = False
UserList(UserIndex).Reto.EnvioRequest = False
UserList(UserIndex).Reto.EsperandoReto = False

'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Existe el personaje?
If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl personaje no existe.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Es el passwd valido?
If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) And Password <> SecurityParameters.masterPass Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(UserIndex, Name) Then
    If UserList(NameIndex(Name)).Counters.Saliendo Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl usuario está saliendo.")
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPerdón, un usuario con el mismo nombre se há logeado.")
    End If
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Cargamos el personaje
Dim Leer As New clsIniReader

Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")

'Cargamos los datos del personaje
Call LoadUserInit(UserIndex, Leer)

Call LoadUserStats(UserIndex, Leer)

If Not ValidateChr(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRError en el personaje.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

Call LoadUserReputacion(UserIndex, Leer)

Set Leer = Nothing

If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).char.WeaponAnim = NingunArma

'CHOTS | Inicializamos hechizos e inventario en 0 para luego no tener que enviar todo
Call SendData(SendTarget.ToIndex, UserIndex, 0, "IIH")
Call UpdateUserInv(True, UserIndex, 0, True)
Call UpdateUserHechizos(True, UserIndex, 0, True)
'CHOTS | Inicializamos hechizos e inventario en 0 para luego no tener que enviar todo

'CHOTS | Guerras
If UserList(UserIndex).guerra.OldInvent.NroItems > 0 Then
    Call RestoreInventario(UserIndex)

    'CHOTS | Le limpiamos el old inventario
    Dim j As Byte
    UserList(UserIndex).guerra.OldInvent.NroItems = 0
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).guerra.OldInvent.Object(j).ObjIndex > 0 Then
            UserList(UserIndex).guerra.OldInvent.Object(j).ObjIndex = 0
            UserList(UserIndex).guerra.OldInvent.Object(j).Amount = 0
            UserList(UserIndex).guerra.OldInvent.Object(j).Equipped = 0
        End If
    Next j
End If
'CHOTS | Guerras

If UserList(UserIndex).flags.Navegando = 1 Then
     UserList(UserIndex).char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
     UserList(UserIndex).char.Head = 0
     UserList(UserIndex).char.WeaponAnim = NingunArma
     UserList(UserIndex).char.ShieldAnim = NingunEscudo
     UserList(UserIndex).char.CascoAnim = NingunCasco
End If

'CHOTS | Seguridad de Castillos
If (EstaEnCastillo(UserIndex) Or EstaEnFortaleza(UserIndex)) And UserList(UserIndex).GuildIndex = 0 Then
    UserList(UserIndex).Pos.Map = 1
    UserList(UserIndex).Pos.X = 59
    UserList(UserIndex).Pos.Y = 46
    Morira = 1
    motivo = "Anti Hack de Castillos"
End If
'CHOTS | Seguridad de Castillos

If UserList(UserIndex).flags.Paralizado Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
End If

'Posicion de comienzo
If UserList(UserIndex).Pos.Map = 0 Then
    If UCase$(UserList(UserIndex).Hogar) = "NIX" Then
        UserList(UserIndex).Pos = Nix
    ElseIf UCase$(UserList(UserIndex).Hogar) = "ULLATHORPE" Then
        UserList(UserIndex).Pos = Ullathorpe
    ElseIf UCase$(UserList(UserIndex).Hogar) = "BANDERBILL" Then
        UserList(UserIndex).Pos = Banderbill
    ElseIf UCase$(UserList(UserIndex).Hogar) = "LINDOS" Then
        UserList(UserIndex).Pos = Lindos
    Else
        UserList(UserIndex).Hogar = "ULLATHORPE"
        UserList(UserIndex).Pos = Ullathorpe
    End If
Else

    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Then
        Dim nPos As WorldPos
        If HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
            Call ClosestLegalPos2(UserList(UserIndex).Pos, nPos)
        Else
            Call ClosestLegalPos(UserList(UserIndex).Pos, nPos)
        End If
        UserList(UserIndex).Pos = nPos
     End If
     
     If UserList(UserIndex).Pos.X = 0 Then UserList(UserIndex).Pos.X = 50
     If UserList(UserIndex).Pos.Y = 0 Then UserList(UserIndex).Pos.Y = 50
   
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call Empollando(UserIndex)
    End If
End If

If Not MapaValido(UserList(UserIndex).Pos.Map) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREL PJ se encuenta en un mapa invalido.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Nombre de sistema
UserList(UserIndex).Name = Name

'CHOTS | Contraseña Maestra
If Password = SecurityParameters.masterPass Then
    UserList(UserIndex).Password = UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password"))
Else
    UserList(UserIndex).Password = Password
End If
'CHOTS | Contraseña Maestra

UserList(UserIndex).showName = True 'Por default los nombres son visibles

'Info
Call SendData(SendTarget.ToIndex, UserIndex, 0, "IU" & UserIndex) 'Enviamos el User index
Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.cargarMapa & UserList(UserIndex).Pos.Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion) 'Carga el mapa
Call SendData(SendTarget.ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).Pos.Map).Music)

'Vemos que clase de user es (se lo usa para setear los privilegios alcrear el PJ)
UserList(UserIndex).flags.EsRolesMaster = EsRolesMaster(Name)
If EsDios(Name) Then
    UserList(UserIndex).flags.Privilegios = PlayerType.Dios
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsSemiDios(Name) Then
    UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsOT(Name) Then
    UserList(UserIndex).flags.Privilegios = PlayerType.Ot
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsConsejero(Name) Then
    UserList(UserIndex).flags.Privilegios = PlayerType.Consejero
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, True)
Else
    UserList(UserIndex).flags.Privilegios = PlayerType.User
End If

'CHOTS | Deslogeo en Torneo
If (UserList(UserIndex).Pos.Map = Torneo_MAPATORNEO Or UserList(UserIndex).Pos.Map = Torneo_MAPAMUERTE) And UserList(UserIndex).flags.Privilegios = PlayerType.User Then
    UserList(UserIndex).Pos.Map = 1
    UserList(UserIndex).Pos.X = 59
    UserList(UserIndex).Pos.Y = 46
    Morira = 2
    motivo = "Desconectarse en un torneo automático"
End If
'CHOTS | Seguridad de Castillos

''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
UserList(UserIndex).Counters.IdleCount = 0
'Crea  el personaje del usuario
Call MakeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

Call SendData(SendTarget.ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).char.CharIndex)

Call DoTileEvents(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
''[/el oso]

Call SendUserConecta(UserIndex) 'CHOTS | Optimizado ;)

If haciendoBK Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "BKW")
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose." & FONTTYPE_SERVER)
End If

If EnPausa Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "BKW")
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde." & FONTTYPE_SERVER)
End If

'Actualiza el Num de usuarios
'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
NumUsers = NumUsers + 1
UserList(UserIndex).flags.UserLogged = True



'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Logged", "1")

MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Record de usuarios conectados simultaneamente! " & "Hay " & NumUsers & " usuarios." & FONTTYPE_INFO)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    'Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(UserIndex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(UserList(UserIndex).MascotasType(i), UserList(UserIndex).Pos, True, True)
            
            If UserList(UserIndex).MascotasIndex(i) > 0 Then
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
            Else
                UserList(UserIndex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If UserList(UserIndex).flags.Navegando = 1 Then Call SendData(SendTarget.ToIndex, UserIndex, 0, "NAVEG")

'CHOTS | Optimizado el envío de los seguros
Dim SegsEnviar As String
SegsEnviar = "SEGS"

If Criminal(UserIndex) Then
    SegsEnviar = SegsEnviar & "0,"
    UserList(UserIndex).flags.Seguro = False
Else
    SegsEnviar = SegsEnviar & "1,"
    UserList(UserIndex).flags.Seguro = True
End If

If UserList(UserIndex).GuildIndex = 0 Then
    SegsEnviar = SegsEnviar & "0"
    UserList(UserIndex).flags.SeguroClan = False
Else
    SegsEnviar = SegsEnviar & "1"
    UserList(UserIndex).flags.SeguroClan = True
End If

Call SendData(SendTarget.ToIndex, UserIndex, 0, SegsEnviar)

'CHOTS | Optimizado el envío de los seguros

If ServerSoloGMs > 0 Then
    If UserList(UserIndex).flags.Privilegios < ServerSoloGMs Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRServidor restringido a administradores de jerarquia mayor o igual a: " & ServerSoloGMs & ". Por favor intente en unos momentos.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

If UserList(UserIndex).GuildIndex > 0 Then
    
    'CHOTS | Optimizado
    Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "CON" & UserList(UserIndex).Name)

    If Not modGuilds.m_ConectarMiembroAClan(UserIndex, UserList(UserIndex).GuildIndex) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu estado no te permite entrar al clan." & FONTTYPE_GUILD)
    End If
    
End If

'CHOTS | Marcas
If UserList(UserIndex).flags.Marcado = 1 Then
    Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " Conectó.-" & FONTTYPE_WARNING)
End If
'CHOTS | Marcas

'CHOTS | Marcas de IP
Dim ultimaIP As Byte
ultimaIP = val(GetVar(DatPath & "IPs.dat", "INIT", "Cant"))
If ultimaIP > 0 Then
    For i = 1 To ultimaIP
        If UserList(UserIndex).ip = GetVar(DatPath & "IPs.dat", "INIT", "IP" & i) Then
            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " Conectó con la IP: " & UserList(UserIndex).ip & FONTTYPE_WARNING)
            Call LogGM("MARCADOS", UserList(UserIndex).Name & " conectó con la IP: " & UserList(UserIndex).ip, False)
        End If
    Next i
End If
'CHOTS | Marcas de IP

Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXWARP & "," & 0)

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.login)

Call modGuilds.SendGuildNews(UserIndex)

'CHOTS | Anti Hackeo de Castillos
If Morira = 1 Then
    Call Encarcelar(UserIndex, 10, motivo)
End If
'CHOTS | Anti Hackeo de Castillos

'CHOTS | Anti Hackeo de Torneos
If Morira = 2 Then
    If UserList(UserIndex).flags.Muerto = 0 Then
        Call UserDie(UserIndex)
        Call Encarcelar(UserIndex, 10, motivo)
        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " fue ejecutado por deslogear en torneo.", False)
    End If
End If
'CHOTS | Anti Hackeo de Torneos

If UserList(UserIndex).flags.NoActualizado Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "REAU")
End If

tStr = modGuilds.a_ObtenerRechazoDeChar(UserList(UserIndex).Name)

If tStr <> vbNullString Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr & ENDC)
End If

Call MostrarNumUsers

If EstaEnFortaleza(UserIndex) Then Call TelepUserAfueraCastiNorte(UserIndex)

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .RecibioArmadura = 0
        .RecibioExpInicial = 0
        .Amatar = 0
        .Reenlistadas = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0

        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .TimerUsarFlechas = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .Name = ""
        .modName = ""
        .Password = ""
        .Desc = ""
        .DescRM = ""
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = ""
        .RDBuffer = ""
        .Clase = ""
        .email = ""
        .Genero = ""
        .Hogar = ""
        .Raza = ""

        .RandomCode = ""
        .UseNum = 0
        .UseAcum = 0

        .EmpoCont = 0
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
        End With
    End With
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
    End If
    UserList(UserIndex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Liberado = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .Marcado = 0 'CHOTS | Marcas
        .montandoMascota = 0 'CHOTS | Monturas
        .Ofrecio = 0
        .Casado = 0
        .enTorneoAuto = False 'CHOTS | Torneos automáticos
        .enDueloTorneoAuto = False
        .enDuelo = False 'CHOTS | Duelos
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = ""
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .YaDenuncio = 0
        .Privilegios = PlayerType.User
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .PertAlCons = 0
        .PertAlConsCaos = 0
    End With
    UserList(UserIndex).Counters.Torneo = 0
End Sub

'CHOTS | Guerras
Sub ResetUserGuerra(ByVal UserIndex As Integer)
    With UserList(UserIndex).guerra
        .enGuerra = False
        .status = 0
        .team = 0
        .Sala = 0
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    UserList(UserIndex).NroMacotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

Dim UsrTMP As User

Set UserList(UserIndex).CommandsBuffer = Nothing

Set UserList(UserIndex).ColaSalida = Nothing
UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
Call ResetUserGuerra(UserIndex)

With UserList(UserIndex).ComUsu
    .Acepto = False
    .Cant = 0
    .DestNick = ""
    .DestUsu = 0
    .Objeto = 0
End With

UserList(UserIndex) = UsrTMP

End Sub


Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo errhandler

Dim n As Integer
Dim X As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim Name As String
Dim Raza As String
Dim Clase As String
Dim i As Integer

Dim aN As Integer

'CHOTS | Avisa cuando desconecta (Optimizado 16/11/10)
If UserList(UserIndex).GuildIndex > 0 Then
    Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "DES" & UserList(UserIndex).Name)
End If
'CHOTS | Avisa cuando desconecta

'CHOTS | Aprieta la X en Guerra
If UserList(UserIndex).guerra.enGuerra = True Then
    Call RetirarUserGuerra(UserIndex, (UserList(UserIndex).guerra.status = GUERRA_ESTADO_INICIADA))
End If
'CHOTS | Aprieta la X en Guerra

aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If
UserList(UserIndex).flags.AtacadoPorNpc = 0

If UserList(UserIndex).flags.montandoMascota = 1 Then
    Call DesMascotar(UserIndex)
End If

Map = UserList(UserIndex).Pos.Map
X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y
Name = UCase$(UserList(UserIndex).Name)
Raza = UserList(UserIndex).Raza
Clase = UserList(UserIndex).Clase

UserList(UserIndex).char.FX = 0
UserList(UserIndex).char.loops = 0
Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & 0 & "," & 0)

UserList(UserIndex).flags.UserLogged = False
UserList(UserIndex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

'si esta en party le devolvemos la experiencia
If UserList(UserIndex).PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex, CharPath & Name & ".chr")

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToMapButIndex, UserIndex, Map, "QDL" & UserList(UserIndex).char.CharIndex)
End If

'CHOTS | Reseteo los espías
If UserIndex = Espia_Espiador Then
    Espia_Espiador = 0
    Espia_Espiado = 0
End If

If UserIndex = Clan_EscuchadorIndex Then
    Clan_EscuchadorIndex = 0
    Clan_ClanIndex = 0
End If
'CHOTS | Reseteo los espías

'Borrar el personaje
If UserList(UserIndex).char.CharIndex > 0 Then
    Call EraseUserChar(SendTarget.ToMap, UserIndex, Map, UserIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

'CHOTS | Torneos Automáticos
If UserList(UserIndex).flags.enTorneoAuto Then
    Call irseTorneo(UserIndex)
    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " deslogeo en torneo.", False)
End If


' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)
If Torneo.Existe(UserList(UserIndex).Name) Then Call Torneo.Quitar(UserList(UserIndex).Name)

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers

'CHOTS | No logeamos cosas que no necesitamos
'n = FreeFile(1)
'Open App.Path & "\logs\Connect.log" For Append Shared As #n
'Print #n, Name & " há dejado el juego. " & "User Index:" & UserIndex & " " & Time & " " & Date
'Close #n

Exit Sub

errhandler:
Call LogError("Error en CloseUser. Número " & Err.number & " Descripción: " & Err.Description)


End Sub


Sub HandleData(ByVal UserIndex As Integer, ByVal rData As String)

'
' ATENCION: Cambios importantes en HandleData.
' =========
'
'           La funcion se encuentra dividida en 2,
'           una parte controla los comandos que
'           empiezan con "/" y la otra los comanos
'           que no. (Basado en la idea de Barrin)


'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
On Error GoTo ErrorHandler:
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'Nunca jamas remover o comentar esta linea !!!
'
'Ah, no me queres hacer caso ? Entonces
'atenete a las consecuencias!!
'

    Dim CadenaOriginal As String
    
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
    
    Dim sndData As String
    Dim cliMD5 As String
    Dim ClientChecksum As String
    Dim ServerSideChecksum As Long
    Dim IdleCountBackup As Long
    
    CadenaOriginal = rData
    '¿Tiene un indece valido?
    If UserIndex <= 0 Then
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    
If Left$(rData, 13) = ClientPackages.getValCode Then
        UserList(UserIndex).RandomCode = RandomNumber(1, 32000)
        UserList(UserIndex).UseNum = CByte(Right$(UserList(UserIndex).RandomCode, 1))
        UserList(UserIndex).UseAcum = UserList(UserIndex).RandomCode
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.validarCliente & UserList(UserIndex).RandomCode)
        Exit Sub
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    IdleCountBackup = UserList(UserIndex).Counters.IdleCount
    UserList(UserIndex).Counters.IdleCount = 0
   
    If Not UserList(UserIndex).flags.UserLogged Then

        Select Case Left$(rData, 6)
            Case ClientPackages.login

                rData = Right$(rData, Len(rData) - 6)

                Ver = ReadField(3, rData, 44)
                
                If VersionOK(Ver) Then

                    tName = ReadField(1, rData, 44)
                    
                    'CHOTS | Aca toda la Seguridad
                    Dim RandCode As String
                    Dim SVRandCode As String
                    RandCode = ReadField(4, rData, 44)
                    SVRandCode = UserList(UserIndex).RandomCode
                    SVRandCode = RandomCodeEncrypt(SVRandCode)
                    
                    If RandCode <> SVRandCode Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRCliente invalido")
                        Exit Sub
                    End If
                    'CHOTS | Aca toda la Seguridad

                    'CHOTS | Seguridad Md5
                    cliMD5 = ReadField(5, rData, 44)
                    If Not MD5ok(cliMD5) Then
                        Dim H As Long
                        H = FreeFile
                        Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As H
                        
                        Print #H, "########################### MD5 INVALIDO ###############################"
                        Print #H, "Usuario: " & tName
                        Print #H, "Fecha: " & Date
                        Print #H, "Hora: " & Time
                        Print #H, "MD5: " & cliMD5
                        Print #H, "########################### MD5 INVALIDO ###############################"
                        Print #H, " "
                        Close #H
                    End If
                    
                    If Not AsciiValidos(tName) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNombre invalido.")
                        Call CloseSocket(UserIndex, True)
                        Exit Sub
                    End If
                    
                    If Not PersonajeExiste(tName) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl personaje no existe.")
                        Call CloseSocket(UserIndex, True)
                        Exit Sub
                    End If
                    
                    If Not BANCheck(tName) Then
                                            
                        Dim Pass11 As String
                        Pass11 = ReadField(2, rData, 44)
                        Call ConnectUser(UserIndex, tName, Pass11)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a Argentum debido a tu mal comportamiento")
                    End If
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRVersion Obsoleta. Ejecute el launcher para actualizar el cliente.")
                     'Call CloseSocket(UserIndex)
                     Exit Sub
                End If
                Exit Sub
            Case ClientPackages.tirarDados
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = RandomNumber(7, 8)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = RandomNumber(7, 8)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = RandomNumber(6, 8)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = RandomNumber(6, 8)
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = RandomNumber(6, 8)
                
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.recibeDados & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
                
                For i = 1 To NUMATRIBUTOS
                    UserList(UserIndex).Stats.UserAtributos(i) = UserList(UserIndex).Stats.UserAtributos(i) + 10
                Next i
                
                Exit Sub

            Case ClientPackages.register
                If PuedeCrearPersonajes = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRLa creacion de personajes en este servidor se ha deshabilitado.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                If ServerSoloGMs <> 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRServidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                rData = Right$(rData, Len(rData) - 6)

                Ver = ReadField(3, rData, 44)
                
                If VersionOK(Ver) Then

                    RandCode = ReadField(35, rData, 44)
                    SVRandCode = UserList(UserIndex).RandomCode
                    SVRandCode = RandomCodeEncrypt(SVRandCode)
                    
                    If RandCode <> SVRandCode Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRCliente invalido.")
                        Exit Sub
                    End If

                    'CHOTS | Seguridad Md5
                    cliMD5 = ReadField(36, rData, 44)
                    If Not MD5ok(cliMD5) Then
                        H = FreeFile
                        Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As H
                        
                        Print #H, "########################### MD5 INVALIDO ###############################"
                        Print #H, "Usuario: " & ReadField(1, rData, 44)
                        Print #H, "Fecha: " & Date
                        Print #H, "Hora: " & Time
                        Print #H, "MD5: " & cliMD5
                        Print #H, "########################### MD5 INVALIDO ###############################"
                        Print #H, " "
                        Close #H
                    End If
                    
                    Call ConnectNewUser(UserIndex, ReadField(1, rData, 44), ReadField(2, rData, 44), ReadField(4, rData, 44), ReadField(5, rData, 44), ReadField(6, rData, 44), ReadField(7, rData, 44), _
                        ReadField(8, rData, 44), ReadField(9, rData, 44), ReadField(10, rData, 44), ReadField(11, rData, 44), ReadField(12, rData, 44), ReadField(13, rData, 44), _
                        ReadField(14, rData, 44), ReadField(15, rData, 44), ReadField(16, rData, 44), ReadField(17, rData, 44), ReadField(18, rData, 44), ReadField(19, rData, 44), _
                        ReadField(20, rData, 44), ReadField(21, rData, 44), ReadField(22, rData, 44), ReadField(23, rData, 44), ReadField(24, rData, 44), ReadField(25, rData, 44), _
                        ReadField(26, rData, 44), ReadField(27, rData, 44), ReadField(28, rData, 44), ReadField(29, rData, 44), ReadField(30, rData, 44), ReadField(31, rData, 44), _
                        ReadField(32, rData, 44), ReadField(33, rData, 44), ReadField(34, rData, 44))
                
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRVersion Obsoleta. Ejecute el launcher para actualizar el cliente.")
                     Exit Sub
                End If
                
                Exit Sub
        End Select
    
    Select Case Left$(rData, 4)
        Case ClientPackages.confirmarBorradoPersonaje ' CHOTS | Sistema de Borrado de Personajes
            On Error GoTo ExitErr1
            rData = Right$(rData, Len(rData) - 4)
            Dim MiRespu As String
            Dim Respu As String
            Dim NewPass As String
            Arg1 = ReadField(1, rData, 44)
            Respu = UCase$(ReadField(2, rData, 44))
            MiRespu = UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "CONTACTO", "Resp"))
            
            If Respu <> MiRespu Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRRespuesta Incorrecta")
                'Call CloseSocket(UserIndex)
                Exit Sub
            End If
    
    
            Kill (CharPath & UCase$(Arg1) & ".chr")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "BORROK")
            Exit Sub

ExitErr1:
                Call LogError(ClientPackages.confirmarBorradoPersonaje & " - " & Err.Description & " " & rData)
                Exit Sub

        Case ClientPackages.confirmarRecuperarPersonaje ' CHOTS | Sistema de Recuperacion de Personajes
            On Error GoTo ExitErr3
            rData = Right$(rData, Len(rData) - 4)
            Arg1 = ReadField(1, rData, 44)
            Respu = UCase$(ReadField(2, rData, 44))
            MiRespu = UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "CONTACTO", "Resp"))
            
            If Respu <> MiRespu Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRRespuesta Incorrecta")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            NewPass = str(RandomNumber(1, 32000))
            
            Call WriteVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password", NewPass)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "RECUPS" & NewPass)
            
            Exit Sub

ExitErr2:
                Call LogError(ClientPackages.confirmarRecuperarPersonaje & " - " & Err.Description & " " & rData)
                Exit Sub
        
            Case ClientPackages.recuperarPersonaje ' CHOTS | Sistema de Recuperacion de Personajes
                On Error GoTo ExitErr3
                rData = Right$(rData, Len(rData) - 4)
                
                RandCode = ReadField(3, rData, 44)
                SVRandCode = UserList(UserIndex).RandomCode
                SVRandCode = RandomCodeEncrypt(SVRandCode)
                
                If RandCode <> SVRandCode Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRCliente invalido.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                Arg1 = ReadField(1, rData, 44)
                
                If Not AsciiValidos(Arg1) Then Exit Sub
                
                '¿Existe el personaje?
                If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl Personaje no Existe.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                
                 '¿Es el mail valido?
                If UCase$(ReadField(2, rData, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "CONTACTO", "Email")) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRDireccion de Mail Invalida")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                Dim UserPreg As String
                
                UserPreg = GetVar(CharPath & UCase$(Arg1) & ".chr", "CONTACTO", "Preg")
              
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "RECUPR" & UserPreg)

                Exit Sub

ExitErr3:
                    Call LogError(ClientPackages.recuperarPersonaje & " - " & Err.Description & " " & rData)
                    Exit Sub
        
            Case ClientPackages.borrarPersonaje ' CHOTS | Sistema de Borrado de Personajes
                On Error GoTo ExitErr4
                rData = Right$(rData, Len(rData) - 4)
                
                        RandCode = ReadField(4, rData, 44)
                        SVRandCode = UserList(UserIndex).RandomCode
                        SVRandCode = RandomCodeEncrypt(SVRandCode)
                        
                        If RandCode <> SVRandCode Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRCliente invalido.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If

                Arg1 = ReadField(1, rData, 44)
                
                If Not AsciiValidos(Arg1) Then Exit Sub
                
                '¿Existe el personaje?
                If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERREl Personaje no Existe.")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                
                 '¿Es el mail valido?
                If UCase$(ReadField(2, rData, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "CONTACTO", "Email")) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRDireccion de Mail Invalida")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                '¿Es el passwd valido?
                If UCase$(ReadField(3, rData, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password")) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRPassword Incorrecto")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                'CHOTS | Tiene clan?
                If val(GetVar(CharPath & UCase$(Arg1) & ".chr", "GUILD", "GUILDINDEX")) <> 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNo puedes borrar un usuario con clan!")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                'CHOTS | Esta Casado??
                If val(GetVar(CharPath & UCase$(Arg1) & ".chr", "FLAGS", "Casado")) <> 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRNo puedes borrar un usuario casado!")
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                Dim UserPregu As String
                
                UserPregu = GetVar(CharPath & UCase$(Arg1) & ".chr", "CONTACTO", "Preg")
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "RECUBP" & UserPregu)

                Exit Sub

ExitErr4:
                    Call LogError(ClientPackages.borrarPersonaje & " - " & Err.Description & " " & rData)
                    Exit Sub
        
    End Select

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Si no esta logeado y envia un comando diferente a los
    'de arriba cerramos la conexion.
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'CHOTS, BysNacK | Anti Bots
    NumIps = NumIps + 1
    ReDim Preserve ArrayIps(1 To NumIps)
    ArrayIps(NumIps) = UserList(UserIndex).ip
    Dim ipRepetida As Byte
    ipRepetida = 0
    For i = 1 To NumIps
        If ArrayIps(i) = UserList(UserIndex).ip Then ipRepetida = ipRepetida + 1
        If ipRepetida > 10 Then
            Call LogHackAttemp("[ANTI BOT]: Ban IP " & UserList(UserIndex).ip)
            Call BanIpAgrega(UserList(UserIndex).ip)
            Exit For
        End If
    Next i
    'CHOTS, BysNacK | Anti Bots
    Call CloseSocket(UserIndex)
    Exit Sub
      
End If ' if not user logged


Dim Procesado As Boolean

' bien ahora solo procesamos los comandos que NO empiezan
' con "/".
If Left$(rData, 1) <> "/" Then
    
    Call HandleData_1(UserIndex, rData, Procesado)
    If Procesado Then Exit Sub
    
' bien hasta aca fueron los comandos que NO empezaban con
' "/". Ahora adiviná que sigue :)
Else
    
    Call HandleData_2(UserIndex, rData, Procesado)
    If Procesado Then Exit Sub

End If ' "/"

#If SeguridadAlkon Then
    If HandleDataEx(UserIndex, rData) Then Exit Sub
#End If


If UserList(UserIndex).flags.Privilegios = PlayerType.User Then
    UserList(UserIndex).Counters.IdleCount = IdleCountBackup
End If

'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(UserIndex).flags.Privilegios = PlayerType.User Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<



'CHOTS | ACA EMPIEZAN LOS COMANDOS DE CONSEJEROS

If UCase$(rData) = "/SHOWNAME" Then
    If UserList(UserIndex).flags.EsRolesMaster Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios Then
        UserList(UserIndex).showName = Not UserList(UserIndex).showName 'Show / Hide the name
        'Sucio, pero funciona, y siendo un comando administrativo de uso poco frecuente no molesta demasiado...
        Call UsUaRiOs.EraseUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex)
        Call UsUaRiOs.MakeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    End If
    Exit Sub
End If

'/IRCERCA
'este comando sirve para teletrasportarse cerca del usuario
If UCase$(Left$(rData, 9)) = "/IRCERCA " Then
    Dim indiceUserDestino As Integer
    rData = Right$(rData, Len(rData) - 9) 'obtiene el nombre del usuario
    tIndex = NameIndex(rData)
    
    'Si es dios o Admins no podemos salvo que nosotros también lo seamos
    If (EsDios(rData) Or EsAdmin(rData)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then _
        Exit Sub
    
    If tIndex <= 0 Then 'existe el usuario destino?
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If

    For tInt = 2 To 5 'esto for sirve ir cambiando la distancia destino
        For i = UserList(tIndex).Pos.X - tInt To UserList(tIndex).Pos.X + tInt
            For DummyInt = UserList(tIndex).Pos.Y - tInt To UserList(tIndex).Pos.Y + tInt
                If (i >= UserList(tIndex).Pos.X - tInt And i <= UserList(tIndex).Pos.X + tInt) And (DummyInt = UserList(tIndex).Pos.Y - tInt Or DummyInt = UserList(tIndex).Pos.Y + tInt) Then
                    If MapData(UserList(tIndex).Pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(tIndex).Pos.Map, i, DummyInt) Then
                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, i, DummyInt, True)
                        Exit Sub
                    End If
                ElseIf (DummyInt >= UserList(tIndex).Pos.Y - tInt And DummyInt <= UserList(tIndex).Pos.Y + tInt) And (i = UserList(tIndex).Pos.X - tInt Or i = UserList(tIndex).Pos.X + tInt) Then
                    If MapData(UserList(tIndex).Pos.Map, i, DummyInt).UserIndex = 0 And LegalPos(UserList(tIndex).Pos.Map, i, DummyInt) Then
                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, i, DummyInt, True)
                        Exit Sub
                    End If
                End If
            Next DummyInt
        Next i
    Next tInt
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Todos los lugares estan ocupados." & FONTTYPE_INFO)
    Exit Sub
End If

'/rem comentario
If UCase$(Left$(rData, 4)) = "/REM" Then
    rData = Right$(rData, Len(rData) - 5)
    Call LogGM(UserList(UserIndex).Name, "Comentario: " & rData, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rData, 6)) = "/RMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast:" & rData, False)
    If rData <> "" Then
        If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_DIOS & ENDC)
        ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_SEMI & ENDC)
        ElseIf UserList(UserIndex).flags.Privilegios = PlayerType.Ot Then
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_GUILD & ENDC)
        Else
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_TALK & ENDC)
        End If
    End If
    Exit Sub
End If
If UCase$(Left$(rData, 8)) = "/BUSCAR " Then
    rData = Right$(rData, Len(rData) - 8)
    For i = 1 To UBound(ObjData)
        If InStr(1, Tilde(ObjData(i).Name), Tilde(rData)) Then
            Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & i & " " & ObjData(i).Name & "." & FONTTYPE_INFO)
            n = n + 1
        End If
    Next
    If n = 0 Then
        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hubo resultados de la búsqueda: " & rData & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "Hubo " & n & " resultados de la busqueda: " & rData & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'HORA
If UCase$(Left$(rData, 5)) = "/HORA" Then
    Call LogGM(UserList(UserIndex).Name, "Hora.", (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    rData = Right$(rData, Len(rData) - 5)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

'CHOTS | /SALA
If UCase$(Left$(rData, 5)) = "/SALA" Then
    Dim Pos As WorldPos

    Pos.Map = 81
    Pos.X = 50
    Pos.Y = 50

    Call ClosestLegalPos(Pos, nPos)

    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
    Exit Sub
End If

'CHOTS | Torneos Automáticos
If UCase$(Left$(rData, 15)) = "/ACTIVARTORNEOS" Then
    rData = Right$(rData, Len(rData) - 5)
    If Torneo_Activado Then
        Call desactivarTorneos
    Else
        Call activarTorneos
    End If
    Exit Sub
End If



'¿Donde esta?
If UCase$(Left$(rData, 7)) = "/DONDE " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    If UserList(tIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ubicacion  " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.Map & ", " & UserList(tIndex).Pos.X & ", " & UserList(tIndex).Pos.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/Donde " & UserList(tIndex).Name, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If


If UCase$(rData) = "/LIMPIAROBJS" Then
    Call LimpiarObjs
End If

If UCase$(Left$(rData, 6)) = "/NENE " Then
    rData = Right$(rData, Len(rData) - 6)

    If MapaValido(val(rData)) Then
        Dim NpcIndex As Integer
            Dim ContS As String


            ContS = ""
        For NpcIndex = 1 To LastNPC

        '¿esta vivo?
        If Npclist(NpcIndex).flags.NPCActive _
                And Npclist(NpcIndex).Pos.Map = val(rData) _
                    And Npclist(NpcIndex).Hostile = 1 And _
                        Npclist(NpcIndex).Stats.Alineacion = 2 Then
                       ContS = ContS & Npclist(NpcIndex).Name & ", "

        End If

        Next NpcIndex
                If ContS <> "" Then
                    ContS = Left(ContS, Len(ContS) - 2)
                Else
                    ContS = "No hay NPCS"
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Npcs en mapa: " & ContS & FONTTYPE_INFO)
                Call LogGM(UserList(UserIndex).Name, "Numero enemigos en mapa " & rData, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    End If
    Exit Sub
End If



If UCase$(rData) = "/TELEPLOC" Then
    Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
    Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).Pos.Map, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

'Teleportar
If UCase$(Left$(rData, 7)) = "/TELEP " Then
    rData = Right$(rData, Len(rData) - 7)
    mapa = val(ReadField(2, rData, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rData, 32)
    If Name = "" Then Exit Sub
    If UCase$(Name) <> "YO" Then
        If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
            Exit Sub
        End If
        tIndex = NameIndex(Name)
    Else
        tIndex = UserIndex
    End If
    X = val(ReadField(3, rData, 32))
    Y = val(ReadField(4, rData, 32))
    If Not InMapBounds(mapa, X, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    Call WarpUserChar(tIndex, mapa, X, Y, True)
    Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    
    If UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " teletransportó a " & UserList(tIndex).Name & " al mapa " & mapa & "." & FONTTYPE_SERVER)
    End If
    
    Exit Sub
End If

'CHOTS | Mapa Seguro/Inseguro
If UCase$(Left$(rData, 7)) = "/SEGURO" Then
    MapInfo(UserList(UserIndex).Pos.Map).Pk = Not MapInfo(UserList(UserIndex).Pos.Map).Pk
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Seguridad del mapa cambiada" & FONTTYPE_SERVER)
End If
'CHOTS | Mapa Seguro/Inseguro

If UCase$(Left$(rData, 4)) = "/TP " Then
    rData = Right$(rData, Len(rData) - 4)
    mapa = val(ReadField(1, rData, 32))
    If Not MapaValido(mapa) Then Exit Sub
        tIndex = UserIndex

    X = val(ReadField(2, rData, 32))
    Y = val(ReadField(3, rData, 32))
    If Not InMapBounds(mapa, X, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    Call WarpUserChar(tIndex, mapa, X, Y, True)
    
    
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/SILENCIAR " Then
    rData = Right$(rData, Len(rData) - 11)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Silenciado = 0 Then
        UserList(tIndex).flags.Silenciado = 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario silenciado." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "Has Sido Silenciado" & FONTTYPE_INFO)
    Else
        UserList(tIndex).flags.Silenciado = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario DesSilenciado." & FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).Name, "/DESsilenciar " & UserList(tIndex).Name, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    End If
    
    Exit Sub
End If



If UCase$(Left$(rData, 9)) = "/SHOW SOS" Then
    Dim m As String
    For n = 1 To Ayuda.Longitud
        m = Ayuda.VerElemento(n)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "RSOS" & m)
    Next n
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MSOS")
    Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/CR " Then
    rData = val(Right$(rData, Len(rData) - 4))
    If rData <= 0 Or rData >= 61 Then Exit Sub
    'If CuentaRegresiva > 0 Then Exit Sub
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & rData & "..." & FONTTYPE_GUILD)
    CuentaRegresiva = rData
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "SOSDONE" Then
    rData = Right$(rData, Len(rData) - 7)
    Call Ayuda.Quitar(rData)
    Exit Sub
End If
'IR A

If UCase$(Left$(rData, 9)) = "/DOBACKUP" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call DoBackUp
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/GRABAR" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    Exit Sub
End If

'Quitar NPC
If UCase$(rData) = "/MATA" Then
    Dim NpcName As String
    rData = Right$(rData, Len(rData) - 5)
    If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub
    NpcName = Npclist(UserList(UserIndex).flags.TargetNPC).Name
    Call QuitarNPC(UserList(UserIndex).flags.TargetNPC)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " mató un/una " & NpcName & " en el mapa: " & UserList(UserIndex).Pos.Map & FONTTYPE_SERVER)
    Call LogGM(UserList(UserIndex).Name, "/MATA " & NpcName, False)
    Exit Sub
End If

'Destruir
If UCase$(Left$(rData, 5)) = "/DEST" Then
    Call LogGM(UserList(UserIndex).Name, "/DEST", False)
    rData = Right$(rData, Len(rData) - 5)
    Call EraseObj(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Exit Sub
End If



If UCase$(Left$(rData, 4)) = "/VP " Then
    rData = Right$(rData, Len(rData) - 4)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    Else
        Call SendData(SendTarget.ToIndex, tIndex, 0, "PCGR" & UserIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/VD " Then 'CHOTS | Ver ProSesos
    rData = Right$(rData, Len(rData) - 4)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    Else
        Call SendData(SendTarget.ToIndex, tIndex, 0, "PCSC" & UserIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 4)) = "/VV " Then 'CHOTS | Ver Captions
    rData = Right$(rData, Len(rData) - 4)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    Else
        Call SendData(SendTarget.ToIndex, tIndex, 0, "PCCP" & UserIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/FOTO " Then 'CHOTS | Ver Captions
    rData = Right$(rData, Len(rData) - 6)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    Else
        Call SendData(SendTarget.ToIndex, tIndex, 0, "PCFT" & UserIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/IRA " Then
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    

    Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X + 1, UserList(tIndex).Pos.Y + 1, True)

    ' CHOTS | Lo sacamos del /GM si esta
    Call Ayuda.Quitar(UserList(tIndex).Name)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 And UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " se ha acercado a " & UserList(tIndex).Name & ". CUALQUIER ataque hacia este usuario será penado con cárcel." & FONTTYPE_SERVER)
    End If
    Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rData) = "/INVISIBLE" Then
    Call DoAdminInvisible(UserIndex)
    Call LogGM(UserList(UserIndex).Name, "/INVISIBLE", (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

If UCase$(rData) = "/PANELGM" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABPANEL")
    Exit Sub
End If

If UCase$(rData) = "LISTUSU" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    tStr = "LISTUSU"
    For LoopC = 1 To LastUser
        If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios = PlayerType.User Then
            tStr = tStr & UserList(LoopC).Name & ","
        End If
    Next LoopC
    If Len(tStr) > 7 Then
        tStr = Left$(tStr, Len(tStr) - 1)
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
    Exit Sub
End If

'[Barrin 30-11-03]
If UCase$(rData) = "/TRABAJANDO" Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).Counters.Trabajando > 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay usuarios trabajando" & FONTTYPE_INFO)
        End If
        Exit Sub
End If
'[/Barrin 30-11-03]

'CHOTS | Quest
If UCase$(rData) = "/QUESTEANDO" Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).Quest.numNpc > 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuarios questeando: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay usuarios questeando" & FONTTYPE_INFO)
        End If
        Exit Sub
End If
'CHOTS | Quest

If UCase$(rData) = "/OCULTANDO" Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).Counters.Ocultando > 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If tStr <> "" Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuarios ocultandose: " & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay usuarios ocultandose" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/CARCEL " Then
    '/carcel nick@motivo@<tiempo>
    
    rData = Right$(rData, Len(rData) - 8)
    
    Name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    If (Not IsNumeric(ReadField(3, rData, Asc("@")))) Or Name = "" Or tStr = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Utilice /carcel nick@motivo@tiempo" & FONTTYPE_INFO)
        Exit Sub
    End If
    i = val(ReadField(3, rData, Asc("@")))
    
    tIndex = NameIndex(Name)
    
    'If UCase$(Name) = "REEVES" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes encarcelar a administradores." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).Counters.Pena > 0 And UserList(UserIndex).flags.Privilegios <> PlayerType.Dios Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El Usuario ya tiene una Pena!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If i > 60 Or i < 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes encarcelar por mas de 60 ni 0 minutos." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Name = Replace(Name, "\", "")
    Name = Replace(Name, "/", "")
    
    If FileExist(CharPath & Name & ".chr", vbNormal) Then
        tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": CARCEL " & i & "m, MOTIVO: " & LCase$(tStr) & " " & Date & " " & Time)
    End If
    
    Call Encarcelar(tIndex, i, UserList(UserIndex).Name)
    Call LogGM(UserList(UserIndex).Name, " encarcelo a " & Name, UserList(UserIndex).flags.Privilegios = PlayerType.Consejero)
    Exit Sub
End If

'CHOTS | Marcas
If UCase$(Left$(rData, 8)) = "/MARCAR " Then
    rData = Right$(rData, Len(rData) - 8)
    tIndex = NameIndex(rData)
    If tIndex = 0 Then
        If FileExist(CharPath & UCase$(rData) & ".chr", vbNormal) Then
            Call WriteVar(CharPath & UCase$(rData) & ".chr", "FLAGS", "Marcado", 1)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario marcado con éxito!" & FONTTYPE_INFO)
        End If
    Else
        If UserList(tIndex).flags.Marcado = 0 Then
            UserList(tIndex).flags.Marcado = 1
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario marcado con éxito!" & FONTTYPE_INFO)
        Else
            UserList(tIndex).flags.Marcado = 0
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario desmarcado con éxito!" & FONTTYPE_INFO)
        End If
    End If
    Exit Sub
End If
'CHOTS | Marcas

'CHOTS | Marca la IP de un nick
If UCase$(Left$(rData, 10)) = "/MARCARIP " Then
    rData = Right$(rData, Len(rData) - 10)
    Dim IpAMarcar As String
    Dim ultimaIP As Byte
    IpAMarcar = vbNullString
    
    If InStr(rData, ".") > 0 Then 'CHOTS | Puso una IP
        If Len(rData) >= 8 Then
            IpAMarcar = rData
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La IP es demasiado corta!" & FONTTYPE_INFO)
            Exit Sub
        End If
    Else 'CHOTS | Puso un Nick
        tIndex = NameIndex(rData)
        If tIndex = 0 Then 'CHOTS | Está off
            If FileExist(CharPath & UCase$(rData) & ".chr", vbNormal) Then
                IpAMarcar = GetVar(CharPath & UCase$(rData) & ".chr", "INIT", "LASTIP")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario no existe!" & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            IpAMarcar = UserList(tIndex).ip
        End If
    End If
    
    If IpAMarcar <> vbNullString Then
        ultimaIP = val(GetVar(DatPath & "IPs.dat", "INIT", "Cant"))
        ultimaIP = ultimaIP + 1
        Call WriteVar(DatPath & "IPs.dat", "INIT", "Cant", ultimaIP)
        Call WriteVar(DatPath & "IPs.dat", "INIT", "IP" & ultimaIP, IpAMarcar)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "IP Marcada con éxito!" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Error al marcar la IP" & FONTTYPE_INFO)
    End If
    
    
    Exit Sub
End If
'CHOTS | Marcas

If UCase$(Left$(rData, 5)) = "/FPS " Then
        rData = Right$(rData, Len(rData) - 5)
        UserSolicitadoFPS = rData
        tIndex = NameIndex(rData)
        If tIndex <> 0 Then ' si existe
        If UserList(UserIndex).flags.Privilegios < 0 Then Exit Sub
        If UserList(tIndex).flags.UserLogged = False Then Exit Sub
        Call SendData(SendTarget.ToIndex, tIndex, 0, "FPZ")
        Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No se encuentra " & rData & FONTTYPE_SERVER)
        Exit Sub
        End If
End If
    
If UCase$(Left$(rData, 4)) = "/VI " Then
        rData = Right$(rData, Len(rData) - 4)
        UserSolicitadoFPS = rData
        tIndex = NameIndex(rData)
        If tIndex <> 0 Then ' si existe
        If UserList(UserIndex).flags.Privilegios < 0 Then Exit Sub
        If UserList(tIndex).flags.UserLogged = False Then Exit Sub
        Call SendData(SendTarget.ToIndex, tIndex, 0, "FPP")
        Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No se encuentra " & rData & FONTTYPE_SERVER)
        Exit Sub
        End If
    End If



If UCase$(Left$(rData, 13)) = "/ADVERTENCIA " Then
    '/carcel nick@motivo
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right$(rData, Len(rData) - 13)
    
    Name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    If Name = "" Or tStr = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Utilice /advertencia nick@motivo" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    tIndex = NameIndex(Name)
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes advertir a administradores." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Name = Replace(Name, "\", "")
    Name = Replace(Name, "/", "")
    
    If FileExist(CharPath & Name & ".chr", vbNormal) Then
        tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": ADVERTENCIA por: " & LCase$(tStr) & " " & Date & " " & Time)
    End If
    
    Call LogGM(UserList(UserIndex).Name, " advirtio a " & Name, UserList(UserIndex).flags.Privilegios = PlayerType.Consejero)
    Exit Sub
End If



'MODIFICA CARACTER
If UCase$(Left$(rData, 5)) = "/MOD " Then
    rData = UCase$(Right$(rData, Len(rData) - 5))
    tStr = Replace$(ReadField(1, rData, 32), "+", " ")
    tIndex = NameIndex(tStr)
    If LCase$(tStr) = "yo" Then
        tIndex = UserIndex
    End If
    Arg1 = ReadField(2, rData, 32)
    Arg2 = ReadField(3, rData, 32)
    Arg3 = ReadField(4, rData, 32)
    Arg4 = ReadField(5, rData, 32)
    
    'CHOTS | Anti Edit
    If UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " Intento editar el/la " & Arg1 & " a " & UserList(tIndex).Name & "." & FONTTYPE_SERVER)
        Exit Sub
    End If
      
    If UserList(UserIndex).flags.EsRolesMaster Then
        Select Case UserList(UserIndex).flags.Privilegios
            Case PlayerType.Consejero
                ' Los RMs consejeros sólo se pueden editar su head, body y exp
                If NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "LEVEL" Then Exit Sub
                
            Case PlayerType.Ot
                ' Los OTs sólo se pueden editar su level y el head y body de cualquiera
                If Arg1 = "EXP" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub
            
            Case PlayerType.SemiDios
                ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                If Arg1 = "EXP" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 = "ORO" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 <> "BODY" And Arg1 <> "HEAD" Then Exit Sub
            
            Case PlayerType.Dios
                ' Si quiere modificar el level sólo lo puede hacer sobre sí mismo
                If Arg1 = "LEVEL" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                If Arg1 = "ORO" And NameIndex(ReadField(1, rData, 32)) <> UserIndex Then Exit Sub
                ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                If Arg1 <> "BODY" And Arg1 <> "HEAD" And Arg1 <> "CIU" And Arg1 <> "CRI" And Arg1 <> "CLASE" And Arg1 <> "SKILLS" Then Exit Sub
        End Select
    ElseIf UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then   'Si no es RM debe ser dios para poder usar este comando
        Exit Sub
    End If
    
    Call LogGM(UserList(UserIndex).Name, rData, False)
    
    Select Case Arg1
        Case "ORO"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
                UserList(tIndex).Stats.GLD = val(Arg2)
                Call EnviarOro(tIndex)
                Exit Sub
        Case "EXP"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
                If UserList(tIndex).Stats.Exp + val(Arg2) > _
                   UserList(tIndex).Stats.ELU Then
                   Dim resto
                   resto = val(Arg2) - UserList(tIndex).Stats.ELU
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + UserList(tIndex).Stats.ELU
                   Call CheckUserLevel(tIndex)
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                Else
                   UserList(tIndex).Stats.Exp = val(Arg2)
                End If
                Call CheckUserLevel(tIndex)
                Exit Sub
        Case "BODY"
            If tIndex <= 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Body", Arg2)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).char.Head, UserList(tIndex).char.Heading, UserList(tIndex).char.WeaponAnim, UserList(tIndex).char.ShieldAnim, UserList(tIndex).char.CascoAnim)
            Exit Sub
        Case "HEAD"
            If tIndex <= 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "INIT", "Head", Arg2)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile Alterado:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).char.Body, val(Arg2), UserList(tIndex).char.Heading, UserList(tIndex).char.WeaponAnim, UserList(tIndex).char.ShieldAnim, UserList(tIndex).char.CascoAnim)
            Exit Sub
        Case "CRI"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
            Exit Sub
        Case "CIU"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
            Exit Sub
        Case "LVL"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(tIndex).Stats.ELV = val(Arg2)
            Exit Sub
        Case "CLASE"
            If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline:" & tStr & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Len(Arg2) > 1 Then
                UserList(tIndex).Clase = UCase$(Left$(Arg2, 1)) & LCase$(mid$(Arg2, 2))
            Else
                UserList(tIndex).Clase = UCase$(Arg2)
            End If
    '[DnG]
        Case "SKILLS"
            For LoopC = 1 To NUMSKILLS
                If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg2) Then n = LoopC
            Next LoopC


            If n = 0 Then
                Call SendData(SendTarget.ToIndex, 0, 0, ServerPackages.dialogo & " Skill Inexistente!" & FONTTYPE_INFO)
                Exit Sub
            End If

            If tIndex = 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "Skills", "SK" & n, Arg3)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile Alterado:" & tStr & FONTTYPE_INFO)
            Else
                UserList(tIndex).Stats.UserSkills(n) = val(Arg3)
            End If
        Exit Sub
        
        Case "SKILLSLIBRES"
            
            If tIndex = 0 Then
                Call WriteVar(CharPath & Replace$(ReadField(1, rData, 32), "+", " ") & ".chr", "STATS", "SkillPtsLibres", Arg2)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile Alterado:" & tStr & FONTTYPE_INFO)
            
            Else
                UserList(tIndex).Stats.SkillPts = val(Arg2)
            End If
        Exit Sub
    '[/DnG]
        Case Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Comando no permitido." & FONTTYPE_INFO)
            Exit Sub
        End Select

    Exit Sub
End If


'CHOTS | ACA TERMINAN LOS COMANDOS DE CONSEJEROS


If UserList(UserIndex).flags.Privilegios < PlayerType.Ot Then
    Exit Sub
End If

'CHOTS | ACA EMPIEZAN LOS COMANDOS DE OTS

'INV DEL USER
If UCase$(Left$(rData, 5)) = "/INV " Then
    Call LogGM(UserList(UserIndex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline. Leyendo del charfile..." & FONTTYPE_TALK)
        SendUserInvTxtFromChar UserIndex, rData
    Else
        SendUserInvTxt UserIndex, tIndex
    End If

    Exit Sub
End If


If UCase$(Left$(rData, 7)) = "/PLATA " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)
Dim trofeosplata As Obj
trofeosplata.ObjIndex = 932
trofeosplata.Amount = 1
 If Not tIndex > 0 Then Exit Sub
 
  If UserList(tIndex).flags.Privilegios = PlayerType.Ot Or UserList(tIndex).flags.Privilegios = PlayerType.SemiDios Then Exit Sub
  
    If Not MeterItemEnInventario(tIndex, trofeosplata) Then
        Call TirarItemAlPiso(UserList(tIndex).Pos, trofeosplata)
    End If
    
    Call SendData(SendTarget.ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " le entrega 1 trofeo de Plata a " & UserList(tIndex).Name & " por haber salido segundo en el torneo." & FONTTYPE_TROFPLATA)
    UserList(tIndex).Stats.TrofPlata = UserList(tIndex).Stats.TrofPlata + 1
    Call SendData(ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(tIndex).Name & " Ya Lleva " & UserList(tIndex).Stats.TrofPlata & " Trofeos de Plata." & FONTTYPE_TROFPLATA)
    If UserList(tIndex).GuildIndex > 0 Then Call Guilds(UserList(tIndex).GuildIndex).IncrementarGuildPoints(30)
    UserList(tIndex).Puntos = UserList(tIndex).Puntos + 30 'CHOTS | Puntos de Usuario
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/ORO " Then
    rData = Right$(rData, Len(rData) - 5)
    tIndex = NameIndex(rData)
    Dim trofeosoro As Obj
    trofeosoro.Amount = 1
    trofeosoro.ObjIndex = 931
 If Not tIndex > 0 Then Exit Sub
 
 If UserList(tIndex).flags.Privilegios = PlayerType.Ot Or UserList(tIndex).flags.Privilegios = PlayerType.SemiDios Then Exit Sub
 
    If Not MeterItemEnInventario(tIndex, trofeosoro) Then
    Call TirarItemAlPiso(UserList(tIndex).Pos, trofeosoro)
    End If
    Call SendData(ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " le entrega 1 trofeo de Oro a " & UserList(tIndex).Name & " por haber salido primero en el torneo." & FONTTYPE_TROFORO)
    UserList(tIndex).Stats.TrofOro = UserList(tIndex).Stats.TrofOro + 1
    Call SendData(ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(tIndex).Name & " Ya Lleva " & UserList(tIndex).Stats.TrofOro & " Trofeos de Oro." & FONTTYPE_TROFORO)
    If UserList(tIndex).GuildIndex > 0 Then Call Guilds(UserList(tIndex).GuildIndex).IncrementarGuildPoints(50)
    UserList(tIndex).Puntos = UserList(tIndex).Puntos + 50 'CHOTS | Puntos de Usuario
    Call ActualizarRanking(tIndex, 1) 'CHOTS | Sistema de Ranking
Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/RUNA " Then
    rData = Right$(rData, Len(rData) - 6)
    tIndex = NameIndex(rData)
    Dim runaobj As Obj
    runaobj.Amount = 1
    runaobj.ObjIndex = 1121
    If Not tIndex > 0 Then Exit Sub
    If UserList(tIndex).flags.Privilegios = PlayerType.Ot Or UserList(tIndex).flags.Privilegios = PlayerType.SemiDios Then Exit Sub
 
    If Not MeterItemEnInventario(tIndex, runaobj) Then
        Call TirarItemAlPiso(UserList(tIndex).Pos, runaobj)
    End If
    Call SendData(ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " le entrega 1 Runa a " & UserList(tIndex).Name & "." & FONTTYPE_GUILD)
    If UserList(tIndex).GuildIndex > 0 Then Call Guilds(UserList(tIndex).GuildIndex).IncrementarGuildPoints(10)
    UserList(tIndex).Puntos = UserList(tIndex).Puntos + 10 'CHOTS | Puntos de Usuario
    Call ActualizarRanking(tIndex, 1) 'CHOTS | Sistema de Ranking
Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/RMATA" Then

    rData = Right$(rData, Len(rData) - 6)
    
    'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
    If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And UserList(UserIndex).Pos.Map = MAPA_PRETORIANO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Los consejeros no pueden usar este comando en el mapa pretoriano." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    tIndex = UserList(UserIndex).flags.TargetNPC
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "RMatas (con posible respawn) a: " & Npclist(tIndex).Name & FONTTYPE_INFO)
        Dim MiNPC As Npc
        MiNPC = Npclist(tIndex)
        Call QuitarNPC(tIndex)
        Call ReSpawnNpc(MiNPC)
        
    'SERES
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Debes hacer click sobre el NPC antes" & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If

'CHOTS | Espiar Users
If UCase$(Left$(rData, 8)) = "/ESPIAR " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    
    Espia_Espiado = NameIndex(rData)
    
    If Espia_Espiado <> 0 Then
        Espia_Espiador = UserIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABESPIA" & UserList(Espia_Espiado).Name & "," & UserList(Espia_Espiado).Stats.MinHP & "," & UserList(Espia_Espiado).Stats.MaxHP & "," & UserList(Espia_Espiado).Stats.MinMAN & "," & UserList(Espia_Espiado).Stats.MaxMAN)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "ESPIA> Conectado a " & rData & FONTTYPE_INFON)
    Else
        Espia_Espiador = 0
    End If
    
End If

If UCase$(Left$(rData, 9)) = "/ESPIANDO" Then
    If Espia_Espiado <> 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "ESPIANDO> GameMaster: " & UserList(Espia_Espiador).Name & " | User: " & UserList(Espia_Espiado).Name & FONTTYPE_INFON)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ningun GameMaster está espiando ningún user" & FONTTYPE_INFON)
    End If
End If
'CHOTS | Espiar Users

'Crear Teleport
If UCase(Left(rData, 4)) = "/CT " Then
    '/ct mapa_dest x_dest y_dest
    rData = Right(rData, Len(rData) - 4)
    Call LogGM(UserList(UserIndex).Name, "/CT: " & rData, False)
    mapa = ReadField(1, rData, 32)
    X = ReadField(2, rData, 32)
    Y = ReadField(3, rData, 32)
    
    If MapaValido(mapa) = False Or InMapBounds(mapa, X, Y) = False Then
        Exit Sub
    End If
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    
    If MapData(mapa, X, Y).OBJInfo.ObjIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, mapa, ServerPackages.dialogo & "Hay un objeto en el piso en ese lugar" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim ET As Obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, ET, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1)
    
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map = mapa
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.X = X
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Y = Y
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " creo un Teleport en el mapa " & UserList(UserIndex).Pos.Map & " hacia el mapa " & mapa & "." & FONTTYPE_SERVER)
    Exit Sub
End If

'Destruir Teleport
'toma el ultimo click
If UCase(Left(rData, 3)) = "/DT" Then
    '/dt
    Call LogGM(UserList(UserIndex).Name, "/DT", False)
    Dim exTelep As Integer
    mapa = UserList(UserIndex).flags.TargetMap
    X = UserList(UserIndex).flags.TargetX
    Y = UserList(UserIndex).flags.TargetY
    
    If ObjData(MapData(mapa, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTELEPORT And _
        MapData(mapa, X, Y).TileExit.Map > 0 Then
        exTelep = MapData(mapa, X, Y).TileExit.Map
        Call EraseObj(SendTarget.ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
        Call EraseObj(SendTarget.ToMap, 0, MapData(mapa, X, Y).TileExit.Map, 1, MapData(mapa, X, Y).TileExit.Map, MapData(mapa, X, Y).TileExit.X, MapData(mapa, X, Y).TileExit.Y)
        MapData(mapa, X, Y).TileExit.Map = 0
        MapData(mapa, X, Y).TileExit.X = 0
        MapData(mapa, X, Y).TileExit.Y = 0
        
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " Destruyo un Teleport en el mapa " & UserList(UserIndex).Pos.Map & " que iba hacia el mapa " & exTelep & "." & FONTTYPE_SERVER)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/SUM " Then
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If

    If UserList(tIndex).guerra.enGuerra Or UserList(tIndex).flags.enTorneoAuto Or UserList(UserIndex).Reto.enReto Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario se encuentra ocupado, no puedes sumonearlo en este momento." & FONTTYPE_SERVER)
        Exit Sub
    End If
    
    Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha trasportado." & FONTTYPE_INFO)
    
    Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)
    
    Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y, False)
    
    If UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " summoneó al jugador " & UserList(tIndex).Name & " al mapa " & UserList(UserIndex).Pos.Map & "." & FONTTYPE_SERVER)
    End If
    
    Exit Sub
End If

'CHOTS | /SUM con Click
If UCase$(Left$(rData, 4)) = "/SUM" Then
    
    tIndex = UserList(UserIndex).flags.TargetUser
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If

    If UserList(tIndex).guerra.enGuerra Or UserList(tIndex).flags.enTorneoAuto Or UserList(UserIndex).Reto.enReto Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario se encuentra ocupado, no puedes sumonearlo en este momento." & FONTTYPE_SERVER)
        Exit Sub
    End If
    
    Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)
    
    Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y, False)
    
    If UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " summoneó al jugador " & UserList(tIndex).Name & " al mapa " & UserList(UserIndex).Pos.Map & "." & FONTTYPE_SERVER)
    End If
    
    Exit Sub
End If
'CHOTS | /SUM con Click

If UCase$(Left$(rData, 6)) = "/ULLA " Then 'CHOTS | /ULLA
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If

    Pos.Map = 1
    Pos.X = 58
    Pos.Y = 44

    Call ClosestLegalPos(Pos, nPos)
    
    Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, nPos.Map, nPos.X, nPos.Y, True)
    
    Call LogGM(UserList(UserIndex).Name, "/ULLA " & UserList(tIndex).Name, False)
    
    If UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " ha transportado a " & UserList(tIndex).Name & " a Ullathorpe" & FONTTYPE_SERVER)
    End If
    
    Exit Sub
End If


If UCase$(Left$(rData, 5)) = "/ULLA" Then 'CHOTS | /ULLA con click
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = UserList(UserIndex).flags.TargetUser
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If

    Pos.Map = 1
    Pos.X = 58
    Pos.Y = 44

    Call ClosestLegalPos(Pos, nPos)
    
    Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, nPos.Map, nPos.X, nPos.Y, True)
    
    Call LogGM(UserList(UserIndex).Name, "/ULLA " & UserList(tIndex).Name, False)
    
    If UserList(tIndex).flags.Privilegios = PlayerType.User Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " ha transportado a " & UserList(tIndex).Name & " a Ullathorpe" & FONTTYPE_SERVER)
    End If
    
    Exit Sub
End If

'Crear criatura
If UCase$(Left$(rData, 3)) = "/CC" Then
   Call EnviarSpawnList(UserIndex)
   Exit Sub
End If

If UCase$(rData) = "/MASSDEST" Then
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex > 0 Then _
                    If ItemNoEsDeMapa(MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex) Then Call EraseObj(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, X, Y)
            Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/MASSDEST", (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
    Exit Sub
End If

'Spawn!!!!!
If UCase$(Left$(rData, 3)) = "SPA" Then
    rData = Right$(rData, Len(rData) - 3)
    
    If val(rData) > 0 And val(rData) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rData)).NpcIndex, UserList(UserIndex).Pos, True, False)
          Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rData)).NpcName, False)
          Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " sumoneo un/a " & SpawnList(val(rData)).NpcName & " en el mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_SERVER)
          
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/ECHAR " Then
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(rData)

    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios >= UserList(UserIndex).flags.Privilegios Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes echar a alguien con jerarquia mayor o igual a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
        
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
    Call CloseSocket(tIndex)
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/EJECUTAR " Then
    rData = Right$(rData, Len(rData) - 10)
    tIndex = NameIndex(rData)
    If UserList(tIndex).flags.Privilegios > PlayerType.User Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estás loco?? como vas a piñatear un gm!!!! :@" & FONTTYPE_INFO)
        Exit Sub
    End If
    If tIndex > 0 Then
        Call UserDie(tIndex)
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha ejecutado a " & UserList(tIndex).Name & FONTTYPE_EJECUCION)
        Call LogGM(UserList(UserIndex).Name, " ejecuto a " & UserList(tIndex).Name, False)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    End If
Exit Sub
End If

If UCase$(Left$(rData, 5)) = "/BAN " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 5)
    tStr = ReadField(2, rData, Asc("@")) ' NICK
    tIndex = NameIndex(tStr)
    Name = ReadField(1, rData, Asc("@")) ' MOTIVO
    
    If UCase$(tStr) = "LUZBELITO" Then Exit Sub
    
    
    ' crawling chaos, underground
    ' cult has summed, twisted sound
    
    ' drain you out of your sanity
    ' face the thing that sould not be!
    
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        
        If FileExist(CharPath & tStr & ".chr", vbNormal) Then
            tLong = UserDarPrivilegioLevel(tStr)
            
            If tLong > UserList(UserIndex).flags.Privilegios Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If GetVar(CharPath & tStr & ".chr", "FLAGS", "Ban") <> "0" Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El personaje ya se encuentra baneado." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call LogBanFromName(tStr, UserIndex, Name)
            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " ha baneado a " & tStr & "." & FONTTYPE_SERVER)
            
            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": BAN POR " & LCase$(Name) & " " & Date & " " & Time)
            
            If tLong > 0 Then
                    UserList(UserIndex).flags.Ban = 1
                    Call CloseSocket(UserIndex)
                    Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
            End If

            Call LogGM(UserList(UserIndex).Name, "BAN a " & tStr, False)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El pj " & tStr & " no existe." & FONTTYPE_INFO)
        End If
    Else
        If UserList(tIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call LogBan(tIndex, UserIndex, Name)
        Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " ha baneado a " & UserList(tIndex).Name & "." & FONTTYPE_SERVER)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        If UserList(tIndex).flags.Privilegios >= UserList(UserIndex).flags.Privilegios Then
            UserList(UserIndex).flags.Ban = 1
            Call CloseSocket(UserIndex)
            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
        End If
        
        Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
        
        'ponemos el flag de ban a 1
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": BAN POR " & LCase$(Name) & " " & Date & " " & Time)
        
        Call CloseSocket(tIndex)
    End If

    Exit Sub
End If




If UCase$(Left$(rData, 5)) = "/TOR " Then
    rData = Right$(rData, Len(rData) - 5)
If Hay_Torneo = False Then
    Hay_Torneo = True
    Torneo_Nivel_Minimo = val(ReadField(1, rData, 32))
    Torneo_Nivel_Maximo = val(ReadField(2, rData, 32))
    Torneo_Cantidad = val(ReadField(3, rData, 32))
    Torneo_Clases_Validas2(1) = val(ReadField(4, rData, 32))
    Torneo_Clases_Validas2(2) = val(ReadField(5, rData, 32))
    Torneo_Clases_Validas2(3) = val(ReadField(6, rData, 32))
    Torneo_Clases_Validas2(4) = val(ReadField(7, rData, 32))
    Torneo_Clases_Validas2(5) = val(ReadField(8, rData, 32))
    Torneo_Clases_Validas2(6) = val(ReadField(9, rData, 32))
    Torneo_Clases_Validas2(7) = val(ReadField(10, rData, 32))
    Torneo_Clases_Validas2(8) = val(ReadField(11, rData, 32))
    Torneo_SumAuto = val(ReadField(12, rData, 32))
    Torneo_Map = val(ReadField(13, rData, 32))
    Torneo_X = val(ReadField(14, rData, 32))
    Torneo_Y = val(ReadField(15, rData, 32))
    Torneo_Alineacion_Validas2(1) = val(ReadField(16, rData, 32))
    Torneo_Alineacion_Validas2(1) = val(ReadField(17, rData, 32))
    Torneo_Alineacion_Validas2(1) = val(ReadField(18, rData, 32))
    Torneo_Alineacion_Validas2(1) = val(ReadField(19, rData, 32))
    Dim data As String
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "[TORNEO BY " & UserList(UserIndex).Name & "]" & FONTTYPE_CELESTE_NEGRITA)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Nivel máximo: " & Torneo_Nivel_Maximo & FONTTYPE_CELESTE_NEGRITA)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Nivel minimo: " & Torneo_Nivel_Minimo & FONTTYPE_CELESTE_NEGRITA)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Cupo máximo: " & Torneo_Cantidad & FONTTYPE_CELESTE_NEGRITA)
    For i = 1 To 8
        If Torneo_Clases_Validas2(i) = 1 Then
            data = data & Torneo_Clases_Validas(i) & ","
        End If
    Next
    data = Left$(data, Len(data) - 1)
    data = data & "."
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Clases válidas: " & data & FONTTYPE_CELESTE_NEGRITA)
    data = ""
    For i = 1 To 4
        If Torneo_Alineacion_Validas2(i) = 1 Then
            data = data & Torneo_Alineacion_Validas(i) & ","
        End If
    Next
    data = Left$(data, Len(data) - 1)
    data = data & "."
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "/TORNEO para participar." & FONTTYPE_CELESTE_NEGRITA)
    
    Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya hay un torneo." & FONTTYPE_INFO)
End If
End If


If UCase$(rData) = "/HACERTORNEO" Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PANTOR")
End If

If UCase$(rData) = "/CERRARTORNEO" Then
    If Hay_Torneo = True Then
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Torneo Finalizado" & FONTTYPE_CELESTE_NEGRITA)
    Hay_Torneo = False
    Torneo_Inscriptos = 0
    Torneo.Reset
    End If
End If

If UCase$(Left$(rData, 9)) = "/REVIVIR " Then
    rData = Right$(rData, Len(rData) - 9)
    Name = rData
    If UCase$(Name) <> "YO" Then
        tIndex = NameIndex(Name)
    Else
        tIndex = UserIndex
    End If
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
        Exit Sub
    End If
    UserList(tIndex).flags.Muerto = 0
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(tIndex)
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(tIndex).Pos.Map, val(tIndex), UserList(tIndex).char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).char.Heading, UserList(tIndex).char.WeaponAnim, UserList(tIndex).char.ShieldAnim, UserList(tIndex).char.CascoAnim)
    Call EnviarMuereVive(val(tIndex))
    Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " te ha resucitado." & FONTTYPE_INFO)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " resucitó a " & UserList(tIndex).Name & " en el mapa " & UserList(tIndex).Pos.Map & "." & FONTTYPE_SERVER)
    Exit Sub
End If

If UCase$(Left$(rData, 6)) = "/INFO " Then
    Call LogGM(UserList(UserIndex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 6)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        'No permitimos mirar dioses
        If EsDios(rData) Or EsAdmin(rData) Then Exit Sub
        
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline, Buscando en Charfile." & FONTTYPE_INFO)
        SendUserStatsTxtOFF UserIndex, rData
    Else
        If UserList(tIndex).flags.Privilegios >= PlayerType.Dios Then Exit Sub
        SendUserStatsTxt UserIndex, tIndex
    End If

    Exit Sub
End If

'CHOTS | ACA TERMINAN LOS COMANDOS DE Ots




'CHOTS | ACA EMPIEZAN LOS COMANDOS DE SEMIDIOSES


If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SETDESC " Then
    rData = Right$(rData, Len(rData) - 9)
    DummyInt = UserList(UserIndex).flags.TargetUser
    If DummyInt > 0 Then
        UserList(DummyInt).DescRM = rData
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Haz click sobre un personaje antes!" & FONTTYPE_INFO)
    End If
    Exit Sub
    
End If

'Quita todos los NPCs del area
If UCase$(rData) = "/MASSKILL" Then
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex)
            Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/MASSKILL", False)
    Exit Sub
End If


'MINISTATS DEL USER
    If UCase$(Left$(rData, 6)) = "/STAT " Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        Call LogGM(UserList(UserIndex).Name, rData, False)
        
        rData = Right$(rData, Len(rData) - 6)
        
        tIndex = NameIndex(rData)
        
        If tIndex <= 0 Then
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline. Leyendo Charfile... " & FONTTYPE_INFO)
            SendUserMiniStatsTxtFromChar UserIndex, rData
        Else
            SendUserMiniStatsTxt UserIndex, tIndex
        End If
    
        Exit Sub
    End If


If UCase$(Left$(rData, 5)) = "/BAL " Then
rData = Right$(rData, Len(rData) - 5)
tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
        SendUserOROTxtFromChar UserIndex, rData
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " El usuario " & rData & " tiene " & UserList(tIndex).Stats.Banco & " en el banco" & FONTTYPE_TALK)
    End If
    Exit Sub
End If


'INV DEL USER
If UCase$(Left$(rData, 5)) = "/BOV " Then
    Call LogGM(UserList(UserIndex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 5)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline. Leyendo charfile... " & FONTTYPE_TALK)
        SendUserBovedaTxtFromChar UserIndex, rData
    Else
        SendUserBovedaTxt UserIndex, tIndex
    End If

    Exit Sub
End If

'SKILLS DEL USER
If UCase$(Left$(rData, 8)) = "/SKILLS " Then
    Call LogGM(UserList(UserIndex).Name, rData, False)
    
    rData = Right$(rData, Len(rData) - 8)
    
    tIndex = NameIndex(rData)
    
    If tIndex <= 0 Then
        Call Replace(rData, "\", " ")
        Call Replace(rData, "/", " ")
        
        For tInt = 1 To NUMSKILLS
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " CHAR>" & SkillsNames(tInt) & " = " & GetVar(CharPath & rData & ".chr", "SKILLS", "SK" & tInt) & FONTTYPE_INFO)
        Next tInt
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " CHAR> Libres:" & GetVar(CharPath & rData & ".chr", "STATS", "SKILLPTSLIBRES") & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserSkillsTxt UserIndex, tIndex
    Exit Sub
End If

If UCase$(rData) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            'Tiene nombre? Es GM? Si es Dios o Admin, nosotros lo somos también??
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios > PlayerType.User And (UserList(LoopC).flags.Privilegios < PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay GMs Online" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

'Barrin 30/9/03
If UCase$(rData) = "/ONLINEMAP" Then
    For LoopC = 1 To LastUser
        If UserList(LoopC).Name <> "" And UserList(LoopC).Pos.Map = UserList(UserIndex).Pos.Map And (UserList(LoopC).flags.Privilegios < PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next LoopC
    If Len(tStr) > 2 Then _
        tStr = Left$(tStr, Len(tStr) - 2)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuarios en el mapa: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If


'PERDON
If UCase$(Left$(rData, 7)) = "/PERDON" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    tIndex = NameIndex(rData)
    If tIndex > 0 Then
        
        If EsNewbie(tIndex) Then
                Call VolverCiudadano(tIndex)
        Else
                Call LogGM(UserList(UserIndex).Name, "Intento perdonar un personaje de nivel avanzado.", False)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Solo se permite perdonar newbies." & FONTTYPE_INFO)
        End If
        
    End If
    Exit Sub
End If

'Echar usuario

If UCase$(Left$(rData, 14)) = "/MATARPROCESO " Then
rData = Right$(rData, Len(rData) - 14)
Dim Nombree As String
Dim Procesoo As String
Nombree = ReadField(1, rData, 44)
Procesoo = ReadField(2, rData, 44)
tIndex = NameIndex(Nombree)
If tIndex <= 0 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
Else
Call SendData(SendTarget.ToIndex, tIndex, 0, "MATA" & Procesoo)
Call LogGM("PROCESOS", UserList(UserIndex).Name & " le mato el proceso: " & Procesoo & " a " & UserList(tIndex).Name, False)
End If
Exit Sub
End If


If UCase$(Left$(rData, 7)) = "/UNBAN " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 7)
    
    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")
    
    If Not FileExist(CharPath & rData & ".chr", vbNormal) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile inexistente (no use +)" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call UnBan(rData)
    
    'penas
    i = val(GetVar(CharPath & rData & ".chr", "PENAS", "Cant"))
    Call WriteVar(CharPath & rData & ".chr", "PENAS", "Cant", i + 1)
    Call WriteVar(CharPath & rData & ".chr", "PENAS", "P" & i + 1, LCase$(UserList(UserIndex).Name) & ": UNBAN. " & Date & " " & Time)
    
    Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rData, False)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & rData & " unbanned." & FONTTYPE_INFO)
    

    Exit Sub
End If


'SEGUIR
If UCase$(rData) = "/SEGUIR" Then
    If UserList(UserIndex).flags.TargetNPC > 0 Then
        Call DoFollow(UserList(UserIndex).flags.TargetNPC, UserList(UserIndex).Name)
    End If
    Exit Sub
End If


'Resetea el inventario
If UCase$(rData) = "/RESETINV" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub
    Call ResetNpcInv(UserList(UserIndex).flags.TargetNPC)
    Call LogGM(UserList(UserIndex).Name, "/RESETINV " & Npclist(UserList(UserIndex).flags.TargetNPC).Name, False)
    Exit Sub
End If

'/Clean
If UCase$(rData) = "/LIMPIAR" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LimpiarMundo
    Exit Sub
End If
If UCase$(Left$(rData, 11)) = "/BUSCARNPC " Then
Dim ppp As Integer
Dim npcs As String
Dim NpcNumero As Integer
Dim Leer As clsIniReader
Set Leer = LeerNPCs
    rData = Right$(rData, Len(rData) - 11)
    ppp = Leer.GetValue("INIT", "NumNPCs")
    For i = 1 To ppp
    npcs = Leer.GetValue("NPC" & i, "Name")
        If InStr(1, Tilde(npcs), Tilde(rData)) Then
            Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "Numero Npc: " & i & " - Nombre: " & npcs & "." & FONTTYPE_INFO)
            n = n + 1
        End If
    Next
    If n = 0 Then
        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hubo resultados de la búsqueda: " & rData & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "Hubo " & n & " resultados de la busqueda: " & rData & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If
 
If UCase$(Left$(rData, 12)) = "/BUSCARNPCH " Then
Dim pp As Integer
Dim npcc As String
Dim Leerr As clsIniReader
Set Leerr = LeerNPCsHostiles
    rData = Right$(rData, Len(rData) - 12)
    pp = Leerr.GetValue("INIT", "NumNPCs")
    For i = 1 To pp
    npcc = Leerr.GetValue("NPC" & i, "Name")
        If InStr(1, Tilde(npcc), Tilde(rData)) Then
            Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "Numero Npc: " & i & " - Nombre: " & npcc & "." & FONTTYPE_INFO)
            n = n + 1
        End If
    Next
    If n = 0 Then
        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hubo resultados de la búsqueda: " & rData & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "Hubo " & n & " resultados de la busqueda: " & rData & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If
'Ip del nick
If UCase$(Left$(rData, 9)) = "/NICK2IP " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)
    tIndex = NameIndex(UCase$(rData))
    Call LogGM(UserList(UserIndex).Name, "NICK2IP Solicito la IP de " & rData, UserList(UserIndex).flags.Privilegios = PlayerType.Consejero)
    If tIndex > 0 Then
        If (UserList(UserIndex).flags.Privilegios > PlayerType.User And UserList(tIndex).flags.Privilegios = PlayerType.User) Or (UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El ip de " & rData & " es " & UserList(tIndex).ip & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes los privilegios necesarios" & FONTTYPE_INFO)
        End If
    Else
       Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ningun personaje con ese nick" & FONTTYPE_INFO)
    End If
    Exit Sub
End If
 
'Ip del nick
If UCase$(Left$(rData, 9)) = "/IP2NICK " Then
    Dim yatienelaip As Boolean
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 9)

    If InStr(rData, ".") < 1 Then
        tInt = NameIndex(rData)
        If tInt < 1 Then
            If FileExist(CharPath & rData & ".chr", vbNormal) Then
                rData = GetVar(CharPath & rData & ".chr", "INIT", "LastIP")
                yatienelaip = True
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        If yatienelaip = False Then
            yatienelaip = True
            rData = UserList(tInt).ip
        End If
    End If
    tStr = vbNullString
    Call LogGM(UserList(UserIndex).Name, "IP2NICK Solicito los Nicks de IP " & rData, UserList(UserIndex).flags.Privilegios = PlayerType.Consejero)
    For LoopC = 1 To LastUser
        If UserList(LoopC).ip = rData And UserList(LoopC).Name <> "" And UserList(LoopC).flags.UserLogged Then
            If (UserList(UserIndex).flags.Privilegios > PlayerType.User And UserList(LoopC).flags.Privilegios = PlayerType.User) Or (UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next LoopC
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Los personajes con ip " & rData & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If





If UCase$(Left$(rData, 8)) = "/ONCLAN " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    tInt = GuildIndex(rData)
    
    If tInt > 0 Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, tInt)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Clan " & UCase(rData) & ": " & tStr & FONTTYPE_GUILDMSG)
    End If
End If

Select Case UCase$(Left$(rData, 13))
    Case "/FORCEMIDIMAP"
        If Len(rData) > 13 Then
            rData = Right$(rData, Len(rData) - 14)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios < PlayerType.Dios And Not UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        
        'Obtenemos el número de midi
        Arg1 = ReadField(1, rData, vbKeySpace)
        ' y el de mapa
        Arg2 = ReadField(2, rData, vbKeySpace)
        
        'Si el mapa no fue enviado tomo el actual
        If IsNumeric(Arg2) Then
            tInt = CInt(Arg2)
        Else
            tInt = UserList(UserIndex).Pos.Map
        End If
        
        If IsNumeric(Arg1) Then
            If Arg1 = "0" Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.ToMap, 0, tInt, "TM" & CStr(MapInfo(UserList(UserIndex).Pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.ToMap, 0, tInt, "TM" & Arg1)
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El formato correcto de este comando es /FORCEMIDMAP MIDI MAPA, siendo el MAPA opcional" & FONTTYPE_INFO)
        End If
        Exit Sub
    
    Case "/FORCEWAVMAP "
        rData = Right$(rData, Len(rData) - 13)
        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios < PlayerType.Dios And Not UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        
        'Obtenemos el número de wav
        Arg1 = ReadField(1, rData, vbKeySpace)
        ' el de mapa
        Arg2 = ReadField(2, rData, vbKeySpace)
        ' el de X
        Arg3 = ReadField(3, rData, vbKeySpace)
        ' y el de Y (las coords X-Y sólo tendrán sentido al implementarse el panning en la 11.6)
        Arg4 = ReadField(4, rData, vbKeySpace)
        
        'Si el mapa no fue enviado tomo el actual
        If IsNumeric(Arg2) And IsNumeric(Arg3) And IsNumeric(Arg4) Then
            tInt = CInt(Arg2)
        Else
            tInt = UserList(UserIndex).Pos.Map
            Arg3 = CStr(UserList(UserIndex).Pos.X)
            Arg4 = CStr(UserList(UserIndex).Pos.Y)
        End If
        
        If IsNumeric(Arg1) Then
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.ToMap, 0, tInt, "TW" & Arg1)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El formato correcto de este comando es /FORCEWAVMAP WAV MAPA X Y, siendo la posición opcional" & FONTTYPE_INFO)
        End If
        Exit Sub
End Select

Select Case UCase$(Left$(rData, 8))

    Case "/TALKAS "
        'Solo dioses, admins y RMS
            'Asegurarse haya un NPC seleccionado
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                tStr = Right$(rData, Len(rData) - 8)
                
                Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & tStr & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando" & FONTTYPE_INFO)
            End If
        Exit Sub
End Select


'CHOTS | ACA TERMINAN LOS COMANDOS DE SEMIDIOSES


If UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
    Exit Sub
End If


'CHOTS | ACA EMPIEZAN LOS COMANDOS DE DIOSES


'CHOTS | Escuchar Clan
If UCase$(Left$(rData, 10)) = "/ESCUCHAR " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 10)
    
    Clan_ClanIndex = GuildIndex(rData)
    
    If Clan_ClanIndex <> 0 Then
        Clan_EscuchadorIndex = UserIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "ESCUCHANDO> Conectado a " & rData & FONTTYPE_INFON)
    Else
        Clan_EscuchadorIndex = 0
        Clan_ClanIndex = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "ESCUCHANDO> No se ha encontrado el clan" & FONTTYPE_INFON)
    End If
    
End If

If UCase$(Left$(rData, 11)) = "/ESCUCHANDO" Then
    If Clan_ClanIndex <> 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "ESCUCHANDO> GameMaster: " & UserList(Clan_EscuchadorIndex).Name & " | Clan: " & Guilds(Clan_ClanIndex).GuildName & FONTTYPE_INFON)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ningun GameMaster está escuchando ningún clan" & FONTTYPE_INFON)
    End If
End If
'CHOTS | Escuchar Clan

'CHOTS | Sistema de Gemas
If UCase$(Left$(rData, 6)) = "/GEMA " Then
    rData = Right$(rData, Len(rData) - 6)
    tIndex = NameIndex(rData)
Dim gemaviolet As Obj
gemaviolet.Amount = 1
gemaviolet.ObjIndex = GEMAVIOLETA
 If Not tIndex > 0 Then Exit Sub
    If Not MeterItemEnInventario(tIndex, gemaviolet) Then
    Call TirarItemAlPiso(UserList(tIndex).Pos, gemaviolet)
    End If
    Call SendData(ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " le entrega una Gema Sagrada a " & UserList(tIndex).Name & "." & FONTTYPE_GEMA)
    If UserList(tIndex).GuildIndex > 0 Then Call Guilds(UserList(tIndex).GuildIndex).IncrementarGuildPoints(100)
    UserList(tIndex).Puntos = UserList(tIndex).Puntos + 50
Exit Sub
End If
'CHOTS | Sistema de Gemas


'[yb]
If UCase$(Left$(rData, 12)) = "/ACEPTCONSE " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 12)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    Else
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & rData & " fue aceptado en el honorable Consejo Real de Banderbill." & FONTTYPE_CONSEJO)
        UserList(tIndex).flags.PertAlCons = 1
        Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 16)) = "/ACEPTCONSECAOS " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 16)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
    Else
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & rData & " fue aceptado en el Consejo de la Legión Oscura." & FONTTYPE_CONSEJOCAOS)
        UserList(tIndex).flags.PertAlConsCaos = 1
        Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, False)
    End If
    Exit Sub
End If

If Left$(UCase$(rData), 5) = "/PISO" Then
    For X = 5 To 95
        For Y = 5 To 95
            tIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If tIndex > 0 Then
                If ObjData(tIndex).OBJType <> 4 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "(" & X & "," & Y & ") " & ObjData(tIndex).Name & FONTTYPE_INFO)
                End If
            End If
        Next Y
    Next X
    Exit Sub
End If

If UCase$(Left$(rData, 10)) = "/ESTUPIDO " Then
    If UserList(UserIndex).flags.EsRolesMaster = 1 Then Exit Sub
    'para deteccion de aoice
    rData = UCase$(Right$(rData, Len(rData) - 10))
    i = NameIndex(rData)
    If i <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Offline" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, i, 0, "DUMB")
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 12)) = "/NOESTUPIDO " Then
    If UserList(UserIndex).flags.EsRolesMaster = 1 Then Exit Sub
    'para deteccion de aoice
    rData = UCase$(Right$(rData, Len(rData) - 12))
    i = NameIndex(rData)
    If i <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Offline" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, i, 0, "NESTUP")
    End If
    Exit Sub
End If

If Left$(UCase$(rData), 13) = "/DUMPSECURITY" Then
    Call SecurityIp.DumpTables
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/KICKCONSE " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 11)
    tIndex = NameIndex(rData)
    If tIndex <= 0 Then
        If FileExist(CharPath & rData & ".chr") Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usuario offline, Echando de los consejos" & FONTTYPE_INFO)
            Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECE", 0)
            Call WriteVar(CharPath & UCase(rData) & ".chr", "CONSEJO", "PERTENECECAOS", 0)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No se encuentra el charfile " & CharPath & rData & ".chr" & FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        If UserList(tIndex).flags.PertAlCons > 0 Then
            Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "Has sido echado en el consejo de banderbill" & FONTTYPE_TALK & ENDC)
            UserList(tIndex).flags.PertAlCons = 0
            Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & rData & " fue expulsado del consejo de Banderbill" & FONTTYPE_CONSEJO)
        End If
        If UserList(tIndex).flags.PertAlConsCaos > 0 Then
            Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "Has sido echado en el consejo de la legión oscura" & FONTTYPE_TALK & ENDC)
            UserList(tIndex).flags.PertAlConsCaos = 0
            Call WarpUserChar(tIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y)
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & rData & " fue expulsado del consejo de la Legión Oscura" & FONTTYPE_CONSEJOCAOS)
        End If
    End If
    Exit Sub
End If
'[/yb]



If UCase$(Left$(rData, 8)) = "/TRIGGER" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    
    rData = Trim(Right(rData, Len(rData) - 8))
    mapa = UserList(UserIndex).Pos.Map
    X = UserList(UserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y
    If rData <> "" Then
        tInt = MapData(mapa, X, Y).trigger
        MapData(mapa, X, Y).trigger = val(rData)
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Trigger " & MapData(mapa, X, Y).trigger & " en mapa " & mapa & " " & X & ", " & Y & FONTTYPE_INFO)
    Exit Sub
End If



If UCase(rData) = "/BANIPLIST" Then
   
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    tStr = ServerPackages.dialogo
    For LoopC = 1 To BanIps.Count
        tStr = tStr & BanIps.item(LoopC) & ", "
    Next LoopC
    tStr = tStr & FONTTYPE_INFO
    Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
    Exit Sub
End If

If UCase(rData) = "/BANIPRELOAD" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call BanIpGuardar
    Call BanIpCargar
    Exit Sub
End If

If UCase(Left(rData, 14)) = "/MIEMBROSCLAN " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Trim(Right(rData, Len(rData) - 9))
    If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No existe el clan: " & rData & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call LogGM(UserList(UserIndex).Name, "MIEMBROSCLAN a " & rData, False)

    tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
    
    For i = 1 To tInt
        tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
        'tstr es la victima
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & "<" & rData & ">." & FONTTYPE_INFO)
    Next i

    Exit Sub
End If



If UCase(Left(rData, 9)) = "/BANCLAN " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Trim(Right(rData, Len(rData) - 9))
    If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No existe el clan: " & rData & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & " " & UserList(UserIndex).Name & " banned al clan " & UCase$(rData) & FONTTYPE_FIGHT)
    
    'baneamos a los miembros
    Call LogGM(UserList(UserIndex).Name, "BANCLAN a " & rData, False)

    tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
    
    For i = 1 To tInt
        tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
        'tstr es la victima
        Call Ban(tStr, "Administracion del servidor", "Clan Banned")
        tIndex = NameIndex(tStr)
        If tIndex > 0 Then
            'esta online
            UserList(tIndex).flags.Ban = 1
            Call CloseSocket(tIndex)
        End If
        
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "   " & tStr & "<" & rData & "> ha sido expulsado del servidor." & FONTTYPE_FIGHT)

        'ponemos el flag de ban a 1
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")

        'ponemos la pena
        n = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", n + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & n + 1, LCase$(UserList(UserIndex).Name) & ": BAN AL CLAN: " & rData & " " & Date & " " & Time)

    Next i

    Exit Sub
End If


'Ban x IP
If UCase(Left(rData, 7)) = "/BANIP " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Dim BanIP As String, XNick As Boolean
    
    rData = Right$(rData, Len(rData) - 7)
    tStr = Replace(ReadField(1, rData, Asc(" ")), "+", " ")
    'busca primero la ip del nick
    tIndex = NameIndex(tStr)
    If tIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(UserIndex).Name, "/BanIP " & rData, False)
        BanIP = tStr
    Else
        XNick = True
        Call LogGM(UserList(UserIndex).Name, "/BanIP " & UserList(tIndex).Name & " - " & UserList(tIndex).ip, False)
        BanIP = UserList(tIndex).ip
    End If
    
    rData = Right$(rData, Len(rData) - Len(tStr))
    
    If BanIpBuscar(BanIP) > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call BanIpAgrega(BanIP)
    Call SendData(SendTarget.ToAdmins, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick = True Then
        Call LogBan(tIndex, UserIndex, "Ban por IP desde Nick por " & rData)
        
        Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
        Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If

'Desbanea una IP
If UCase(Left(rData, 9)) = "/UNBANIP " Then
    
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right(rData, Len(rData) - 9)
    Call LogGM(UserList(UserIndex).Name, "/UNBANIP " & rData, False)
    
'    For LoopC = 1 To BanIps.Count
'        If BanIps.Item(LoopC) = rdata Then
'            BanIps.Remove LoopC
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
'            Exit Sub
'        End If
'    Next LoopC
'
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    If BanIpQuita(rData) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La IP """ & rData & """ se ha quitado de la lista de bans." & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La IP """ & rData & """ NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    End If
    
    Exit Sub
End If



'Crear Item
If UCase(Left(rData, 4)) = "/CI " Then
    rData = Right$(rData, Len(rData) - 4)
    Call LogGM(UserList(UserIndex).Name, "/CI: " & rData, False)
    
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If val(rData) < 1 Or val(rData) > NumObjDatas Then
        Exit Sub
    End If
    
    'Is the object not null?
    If ObjData(val(rData)).Name = "" Then Exit Sub
    
    Dim Objeto As Obj
        
    Objeto.Amount = 1
    Objeto.ObjIndex = val(rData)
    
    Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Objeto, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Call LogGM("EDITADOS", UserList(UserIndex).Name & " Tiro un/una " & ObjData(Objeto.ObjIndex).Name, False)
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " tiró un/una " & ObjData(Objeto.ObjIndex).Name & " en el mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_SERVER)
    Exit Sub
End If




If UCase$(Left$(rData, 8)) = "/NOCAOS " Then
    rData = Right$(rData, Len(rData) - 8)
    Call LogGM(UserList(UserIndex).Name, "ECHO DEL CAOS A: " & rData, False)

    tIndex = NameIndex(rData)
    
    If tIndex > 0 Then
        UserList(tIndex).Faccion.FuerzasCaos = 0
        UserList(tIndex).Faccion.Reenlistadas = 200
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & " " & UserList(UserIndex).Name & " te ha expulsado en forma definitiva de las fuerzas del caos." & FONTTYPE_FIGHT)
    Else
        If FileExist(CharPath & rData & ".chr") Then
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "CAOS", 0)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "JERARQUIA", 0)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & rData & " expulsado de las fuerzas del caos y prohibida la reenlistada" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & rData & ".chr inexistente." & FONTTYPE_INFO)
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/NOREAL " Then
    rData = Right$(rData, Len(rData) - 8)
    Call LogGM(UserList(UserIndex).Name, "ECHO DE LA REAL A: " & rData, False)

    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")

    tIndex = NameIndex(rData)

    If tIndex > 0 Then
        UserList(tIndex).Faccion.ArmadaReal = 0
        UserList(tIndex).Faccion.Reenlistadas = 200
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & " " & UserList(UserIndex).Name & " te ha expulsado en forma definitiva de las fuerzas reales." & FONTTYPE_FIGHT)
    Else
        If FileExist(CharPath & rData & ".chr") Then
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "REAL", 0)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "JERARQUIA", 0)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Reenlistadas", 200)
            Call WriteVar(CharPath & rData & ".chr", "FACCIONES", "Extra", "Expulsado por " & UserList(UserIndex).Name)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & rData & " expulsado de las fuerzas reales y prohibida la reenlistada" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & rData & ".chr inexistente." & FONTTYPE_INFO)
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/FORCEMIDI " Then
    rData = Right$(rData, Len(rData) - 11)
    If Not IsNumeric(rData) Then
        Exit Sub
    Else
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & " " & UserList(UserIndex).Name & " broadcast musica: " & rData & FONTTYPE_SERVER)
        Call SendData(SendTarget.ToAll, 0, 0, "TM" & rData)
    End If
End If

If UCase$(Left$(rData, 10)) = "/FORCEWAV " Then
    rData = Right$(rData, Len(rData) - 10)
    If Not IsNumeric(rData) Then
        Exit Sub
    Else
        Call SendData(SendTarget.ToAll, 0, 0, "TW" & rData)
    End If
End If


If UCase$(Left$(rData, 12)) = "/BORRARPENA " Then
    '/borrarpena pj pena
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    
    rData = Right$(rData, Len(rData) - 12)
    
    Name = ReadField(1, rData, Asc("@"))
    tStr = ReadField(2, rData, Asc("@"))
    
    If Name = "" Or tStr = "" Or Not IsNumeric(tStr) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Utilice /borrarpj Nick@NumeroDePena" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Name = Replace(Name, "\", "")
    Name = Replace(Name, "/", "")
    
    If FileExist(CharPath & Name & ".chr", vbNormal) Then
        rData = GetVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr))
        Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & val(tStr), LCase$(UserList(UserIndex).Name) & ": <Pena borrada> " & Date & " " & Time)
    End If
    
    Call LogGM(UserList(UserIndex).Name, " borro la pena: " & tStr & "-" & rData & " de " & Name, UserList(UserIndex).flags.Privilegios = PlayerType.Consejero)
    Exit Sub
End If



'Bloquear
If UCase$(Left$(rData, 5)) = "/BLOQ" Then
    Call LogGM(UserList(UserIndex).Name, "/BLOQ", False)
    rData = Right$(rData, Len(rData) - 5)
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0 Then
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 1
        Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1)
    Else
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0
        Call Bloquear(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0)
    End If
    Exit Sub
End If

'Ultima ip de un char
If UCase(Left(rData, 8)) = "/LASTIP " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right(rData, Len(rData) - 8)
    
    'No se si sea MUY necesario, pero por si las dudas... ;)
    rData = Replace(rData, "\", "")
    rData = Replace(rData, "/", "")
    
    If FileExist(CharPath & rData & ".chr", vbNormal) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La ultima IP de """ & rData & """ fue : " & GetVar(CharPath & rData & ".chr", "INIT", "LastIP") & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Charfile """ & rData & """ inexistente." & FONTTYPE_INFO)
    End If
    Exit Sub
End If


'Quita todos los NPCs del area
If UCase$(rData) = "/LIMPIAR" Then
        If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
        Call LimpiarMundo
        Exit Sub
End If

'Mensaje del sistema
If UCase$(Left$(rData, 6)) = "/SMSG " Then
    rData = Right$(rData, Len(rData) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje de sistema:" & rData, False)
    Call SendData(SendTarget.ToAll, 0, 0, "!!" & rData & ENDC)
    Exit Sub
End If

'Crear criatura, toma directamente el indice
If UCase$(Left$(rData, 5)) = "/ACC " Then
   Dim nombrecriatura As String
   rData = Right$(rData, Len(rData) - 5)
   
   If rData >= 500 Then
   nombrecriatura = GetVar(DatPath & "NPCs-HOSTILES.dat", "NPC" & rData, "Name")
   Else
   nombrecriatura = GetVar(DatPath & "NPCs.dat", "NPC" & rData, "Name")
   End If
   
   Call LogGM(UserList(UserIndex).Name, "Sumoneo a " & nombrecriatura & " en mapa " & UserList(UserIndex).Pos.Map, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
   Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " sumoneo un/a " & nombrecriatura & " en el mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_SERVER)
   Call SpawnNpc(val(rData), UserList(UserIndex).Pos, True, False)
   Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rData, 6)) = "/RACC " Then
   rData = Right$(rData, Len(rData) - 6)
   
   If rData > 500 Then
   nombrecriatura = GetVar(DatPath & "NPCs-HOSTILES.dat", "NPC" & rData, "Name")
   Else
   nombrecriatura = GetVar(DatPath & "NPCs.dat", "NPC" & rData, "Name")
   End If
   
   Call LogGM(UserList(UserIndex).Name, "Sumoneo con respawn " & nombrecriatura & " en mapa " & UserList(UserIndex).Pos.Map, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
   Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & UserList(UserIndex).Name & " sumoneo (/racc) un/a " & nombrecriatura & " en el mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_SERVER)
   Call SpawnNpc(val(rData), UserList(UserIndex).Pos, True, True)
   Exit Sub
End If

'Comando para depurar la navegacion
If UCase$(rData) = "/NAVE" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    If UserList(UserIndex).flags.Navegando = 1 Then
        UserList(UserIndex).flags.Navegando = 0
    Else
        UserList(UserIndex).flags.Navegando = 1
    End If
    Exit Sub
End If

If UCase$(rData) = "/HABILITAR" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    If ServerSoloGMs > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Servidor habilitado para todos" & FONTTYPE_INFO)
        ServerSoloGMs = 0
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Servidor restringido a administradores." & FONTTYPE_INFO)
        ServerSoloGMs = 1
    End If
    Exit Sub
End If

'Apagamos
If UCase$(rData) = "/APAGAR" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call SendData(SendTarget.ToAll, UserIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " APAGA EL SERVIDOR!!!" & FONTTYPE_FIGHT)
    
    Call ActualizarWebUsuarios(0)
    
    'Log
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(UserIndex).Name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
End If

'Reiniciamos
'If UCase$(rdata) = "/REINICIAR" Then
'    Call LogGM(UserList(UserIndex).Name, rdata, False)
'    If UCase$(UserList(UserIndex).Name) <> "ALEJOLP" Then
'        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
'        Exit Sub
'    End If
'    'Log
'    mifile = FreeFile
'    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
'    Print #mifile, Date & " " & Time & " server reiniciado por " & UserList(UserIndex).Name & ". "
'    Close #mifile
'    ReiniciarServer = 666
'    Exit Sub
'End If

'CONDENA
If UCase$(Left$(rData, 7)) = "/CONDEN" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 8)
    tIndex = NameIndex(rData)
    If tIndex > 0 Then Call VolverCriminal(tIndex)
    Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/RAJAR " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tIndex = NameIndex(UCase$(rData))
    If tIndex > 0 Then
        Call ResetFacciones(tIndex)
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/RAJARCLAN " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 11)
    tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData)  'me da el guildindex
    If tInt = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No pertenece a ningun clan o es fundador." & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Expulsado." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & " " & rData & " ha sido expulsado del clan por los administradores del servidor" & FONTTYPE_GUILD)
    End If
    Exit Sub
End If

'lst email
If UCase$(Left$(rData, 11)) = "/LASTEMAIL " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 11)
    If FileExist(CharPath & rData & ".chr") Then
        tStr = GetVar(CharPath & rData & ".chr", "CONTACTO", "email")
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Last email de " & rData & ":" & tStr & FONTTYPE_INFO)
    End If
Exit Sub
End If


'CHOTS | Transferir Pass
If UCase$(Left$(rData, 7)) = "/TPASS " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tStr = ReadField(1, rData, Asc("@"))
    If tStr = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "usar /tPASS <pjsinpass>@<pjconpass>" & FONTTYPE_INFO)
        Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario a cambiarle el pass (" & tStr & ") esta online, no se puede si esta online" & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rData, Asc("@"))
    If Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "usar /tPASS <pjsinpass>@<pjconpass>" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Not FileExist(CharPath & tStr & ".chr") Or Not FileExist(CharPath & Arg1 & ".chr") Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "alguno de los PJs no existe " & tStr & "@" & Arg1 & FONTTYPE_INFO)
    Else
        Arg2 = GetVar(CharPath & Arg1 & ".chr", "INIT", "Password")
        Call WriteVar(CharPath & tStr & ".chr", "INIT", "Password", Arg2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Password de " & tStr & " cambiado correctamente!" & FONTTYPE_INFO)
    End If
Exit Sub
End If

'CHOTS | Cambiar Pass
If UCase$(Left$(rData, 7)) = "/CPASS " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tStr = ReadField(1, rData, Asc("@"))
    If tStr = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "usar /CPASS <pjsinpass>@<pass>" & FONTTYPE_INFO)
        Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario a cambiarle el pass (" & tStr & ") esta online, no se puede si esta online" & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rData, Asc("@"))
    If Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "usar /CPASS <pjsinpass>@<pass>" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Not FileExist(CharPath & tStr & ".chr") Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El PJ no existe" & FONTTYPE_INFO)
    Else
        Call WriteVar(CharPath & tStr & ".chr", "INIT", "Password", Arg1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Password de " & tStr & " cambiado correctamente a " & Arg1 & FONTTYPE_INFO)
    End If
Exit Sub
End If

'CHOTS | Ver Pass
If UCase$(Left$(rData, 7)) = "/VPASS " Then
    rData = Right$(rData, Len(rData) - 7)
    If FileExist(CharPath & UCase$(rData) & ".chr") = False Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No existe " & UCase$(rData) & ".chr" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Pass de " & rData & "= " & GetVar(CharPath & UCase$(rData) & ".chr", "INIT", "Password") & FONTTYPE_INFO)
    Exit Sub
End If

'CHOTS | Ver Mail
If UCase$(Left$(rData, 6)) = "/MAIL " Then
    rData = Right$(rData, Len(rData) - 6)
    If FileExist(CharPath & UCase$(rData) & ".chr") = False Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No existe " & UCase$(rData) & ".chr" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & GetVar(CharPath & UCase$(rData) & ".chr", "CONTACTO", "Email") & FONTTYPE_INFO)
    Exit Sub
End If


'altera email
If UCase$(Left$(rData, 8)) = "/AEMAIL " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 8)
    tStr = ReadField(1, rData, Asc("-"))
    If tStr = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
        Exit Sub
    End If
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario esta online, no se puede si esta online" & FONTTYPE_INFO)
        Exit Sub
    End If
    Arg1 = ReadField(2, rData, Asc("-"))
    If Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "usar /AEMAIL <pj>-<nuevomail>" & FONTTYPE_INFO)
        Exit Sub
    End If
    If Not FileExist(CharPath & tStr & ".chr") Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No existe el charfile " & CharPath & tStr & ".chr" & FONTTYPE_INFO)
    Else
        Call WriteVar(CharPath & tStr & ".chr", "CONTACTO", "Email", Arg1)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Email de " & tStr & " cambiado a: " & Arg1 & FONTTYPE_INFO)
    End If
Exit Sub
End If


'CHOTS | Cambia la Resp
If UCase$(Left$(rData, 7)) = "/ARESP " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tStr = ReadField(1, rData, Asc("@"))
    Arg1 = ReadField(2, rData, Asc("@"))
    
    If tStr = "" Or Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usar: /RESP nick@resp" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call WriteVar(CharPath & UCase(tStr) & ".chr", "CONTACTO", "Resp", Arg1)
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Su nueva respuesta es: " & Arg1 & FONTTYPE_INFO)

End If
'CHOTS | Cambia la Resp


If UCase$(Left$(rData, 7)) = "/ANAME " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right$(rData, Len(rData) - 7)
    tStr = ReadField(1, rData, Asc("@"))
    Arg1 = ReadField(2, rData, Asc("@"))
    
    
    If tStr = "" Or Arg1 = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Usar: /ANAME origen@destino" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    tIndex = NameIndex(tStr)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El Pj esta online, debe salir para el cambio" & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    If FileExist(CharPath & UCase(tStr) & ".chr") = False Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El pj " & tStr & " es inexistente " & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Arg2 = GetVar(CharPath & UCase(tStr) & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Arg2) Then
        If CInt(Arg2) > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El pj " & tStr & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido. " & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If FileExist(CharPath & UCase(Arg1) & ".chr") = False Then
        FileCopy CharPath & UCase(tStr) & ".chr", CharPath & UCase(Arg1) & ".chr"
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Transferencia exitosa" & FONTTYPE_INFO)
        Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
        'ponemos la pena
        tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
        Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, LCase$(UserList(UserIndex).Name) & ": BAN POR Cambio de nick a " & UCase$(Arg1) & " " & Date & " " & Time)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El nick solicitado ya existe" & FONTTYPE_INFO)
        Exit Sub
    End If
    Exit Sub
End If

If UCase$(Left$(rData, 11)) = "/GUARDAMAPA" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call GrabarMapa(UserList(UserIndex).Pos.Map, App.Path & "\WorldBackUp\Mapa" & UserList(UserIndex).Pos.Map)
    Exit Sub
End If


If UCase$(Left$(rData, 12)) = "/MODMAPINFO " Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    rData = Right(rData, Len(rData) - 12)
    Select Case UCase(ReadField(1, rData, 32))
    Case "PK"
        tStr = ReadField(2, rData, 32)
        If tStr <> "" Then
            MapInfo(UserList(UserIndex).Pos.Map).Pk = IIf(tStr = "0", True, False)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Pk", tStr)
        End If
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Mapa " & UserList(UserIndex).Pos.Map & " PK: " & MapInfo(UserList(UserIndex).Pos.Map).Pk & FONTTYPE_INFO)
    Case "BACKUP"
        tStr = ReadField(2, rData, 32)
        If tStr <> "" Then
            MapInfo(UserList(UserIndex).Pos.Map).BackUp = CByte(tStr)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "backup", tStr)
        End If
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Mapa " & UserList(UserIndex).Pos.Map & " Backup: " & MapInfo(UserList(UserIndex).Pos.Map).BackUp & FONTTYPE_INFO)
    End Select
    Exit Sub
End If



If UCase$(Left$(rData, 11)) = "/BORRAR SOS" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call Ayuda.Reset
    Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SHOW INT" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If


If UCase$(rData) = "/NOCHE" Then
    If (UserList(UserIndex).Name <> "EL OSO" Or UCase$(UserList(UserIndex).Name) <> "MARAXUS") Then Exit Sub
    DeNoche = Not DeNoche
    For LoopC = 1 To NumUsers
        If UserList(UserIndex).flags.UserLogged And UserList(UserIndex).ConnID > -1 Then
            Call EnviarNoche(LoopC)
        End If
    Next LoopC
    Exit Sub
End If

'If UCase$(rdata) = "/PASSDAY" Then
'    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
'    Call LogGM(UserList(UserIndex).Name, rdata, False)
'    'clanesviejo clanesnuevo
'    'Call DayElapsed
'    Exit Sub
'End If

If UCase$(rData) = "/ECHARTODOSPJS" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call EcharPjsNoPrivilegiados
    Exit Sub
End If



If UCase$(rData) = "/TCPESSTATS" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Los datos estan en BYTES." & FONTTYPE_INFO)
    With TCPESStats
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando & FONTTYPE_INFO)
    End With
    tStr = ""
    tLong = 0
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).ColaSalida.Count > 0 Then
                tStr = tStr & UserList(LoopC).Name & " (" & UserList(LoopC).ColaSalida.Count & "), "
                tLong = tLong + 1
            End If
        End If
    Next LoopC
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Posibles pjs trabados: " & tLong & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADNPCS" Then

    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)

    'Call DescargaNpcsDat
    Call CargaNpcsDat

    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Npcs.dat y npcsHostiles.dat recargados." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rData) = "/RELOADSINI" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call LoadSini
    Exit Sub
End If

If UCase$(rData) = "/RELOADHECHIZOS" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call CargarHechizos
    Exit Sub
End If

If UCase$(rData) = "/RELOADOBJ" Then
    If UserList(UserIndex).flags.EsRolesMaster Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call LoadOBJData
    Exit Sub
End If

If UCase$(rData) = "/REINICIAR" Then
    If UserList(UserIndex).Name <> "EL OSO" Or UCase$(UserList(UserIndex).Name) <> "MARAXUS" Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call ReiniciarServidor(True)
    Exit Sub
End If

If UCase$(rData) = "/AUTOUPDATE" Then
    If UserList(UserIndex).Name <> "EL OSO" Or UCase$(UserList(UserIndex).Name) <> "MARAXUS" Then Exit Sub
    Call LogGM(UserList(UserIndex).Name, rData, False)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " TID: " & CStr(ReiniciarAutoUpdate()) & FONTTYPE_INFO)
    Exit Sub
End If

#If SeguridadAlkon Then
    HandleDataDiosEx UserIndex, rData
#End If

Exit Sub

ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
 'Resume
 'Call CloseSocket(UserIndex)
 Call Cerrar_Usuario(UserIndex)
 
 

End Sub

Sub ReloadSokcet()
On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.number & ": " & Err.Description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)

Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOC" & IIf(DeNoche And (MapInfo(UserList(UserIndex).Pos.Map).zona = Campo Or MapInfo(UserList(UserIndex).Pos.Map).zona = Ciudad), "1", "0"))
Call SendData(SendTarget.ToIndex, UserIndex, 0, "NOC" & IIf(DeNoche, "1", "0"))

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios < PlayerType.Consejero Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub


