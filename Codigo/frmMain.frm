VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   1785
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrAntiBoter 
      Interval        =   60000
      Left            =   960
      Top             =   120
   End
   Begin InetCtlsObjects.Inet InetRanking 
      Left            =   3360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetClanes 
      Left            =   4680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetGuerras 
      Left            =   4560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetCastillos 
      Left            =   3960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrOnline 
      Interval        =   60000
      Left            =   2880
      Top             =   0
   End
   Begin InetCtlsObjects.Inet InetUsers 
      Left            =   3360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrCastillos 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2400
      Top             =   120
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   540
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1020
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   945
      Top             =   540
   End
   Begin VB.Timer GameTimer 
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   1020
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   960
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.Timer tmrMinutosGuerra 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   2880
         Top             =   720
      End
      Begin VB.Timer tmrSegundosGuerra 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2520
         Top             =   120
      End
      Begin VB.Timer NuevoGameTimer 
         Interval        =   1000
         Left            =   360
         Top             =   -120
      End
      Begin InetCtlsObjects.Inet InetRankingUsers 
         Left            =   3960
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Timer tmrInvocaciones 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1800
         Top             =   360
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private ruta As String
Private str_archivo_temporal As String
Private lng_tamaño_archivo As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
Dim iUserIndex As Integer

For iUserIndex = 1 To MaxUsers
   
   'Conexion activa? y es un usuario loggeado?
   If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
            Call SendData(SendTarget.ToIndex, iUserIndex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
            'mato los comercios seguros
            If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(iUserIndex).ComUsu.DestUsu, 0, ServerPackages.dialogo & "Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call FinComerciarUsu(iUserIndex)
            End If
            Call Cerrar_Usuario(iUserIndex)
        End If
  End If
  
Next iUserIndex

End Sub



Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo

Exit Sub

errhand:
Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.number)
End Sub

Private Sub AutoSave_Timer()
'CHOTS | Agregado lo de torneos automáticos

On Error GoTo errhandler
'fired every minute
Static MinutosLatsClean As Long
Static MinsSocketReset As Long

Dim i As Integer
Dim num As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.segundos = 0
        DayStats.Promedio = 0
        
        Horas = 0
        
    End If
    MinsRunning = 0
End If
    
MinutosParaWs = MinutosParaWs + 1
MinutosParaGrabar = MinutosParaGrabar + 1

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call ModAreas.AreasOptimizacion
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

#If UsarQueSocket = 1 Then
' ok la cosa es asi, este cacho de codigo es para
' evitar los problemas de socket. a menos que estes
' seguro de lo que estas haciendo, te recomiendo
' que lo dejes tal cual está.
' alejo.
MinsSocketReset = MinsSocketReset + 1
If MinsSocketReset >= 5 Then
    MinsSocketReset = 0
    num = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then
            If UserList(i).Counters.IdleCount > ((IntervaloCerrarConexion * 2) / 3) Then
                Call CloseSocket(i)
            End If
        End If

        'CHOTS | Checkeamos los users aca
        If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged Then
            num = num + 1
        End If
    Next i

    'CHOTS | Checkeamos los users aca
    If num <> NumUsers Then
        NumUsers = num
    End If

End If
#End If

If MinutosParaWs = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Worldsave y Limpeza en 1 minuto ..." & FONTTYPE_SERVER)
End If

If MinutosParaWs >= MinutosWs Then
    Call DoBackUp
    MinutosParaWs = 0
End If

If MinutosParaGrabar = MinutosGrabar - 1 Then
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Grabado de personajes en 1 minuto ..." & FONTTYPE_SERVER)
End If

If MinutosParaGrabar >= MinutosGrabar Then
    Call GuardarUsuarios
    MinutosParaGrabar = 0
End If

'CHOTS | Torneos automáticos
If Torneo_Activado Then
  
    If Not Torneo_HAYTORNEO Then
      If MinutosParaWs = 1 And NumUsers >= 8  Then
          Call SendData(SendTarget.ToAll, 0, 0, "Z96")
      End If
      
      If MinutosParaWs = 5 And NumUsers >= 8 Then
          Call SendData(SendTarget.ToAll, 0, 0, "Z97")
      End If
      
      If MinutosParaWs = 6 And NumUsers >= 8 Then
          Call crearTorneo
      End If
    Else
      If MinutosParaWs = 7 And Torneo_CantidadInscriptos < Torneo_Cupo Then
          Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "El cupo no ha sido alcanzado! Ingrese /PARTICIPAR para entrar al torneo" & FONTTYPE_TORNEOAUTO)
      End If
      
      If MinutosParaWs = 8 And Torneo_CantidadInscriptos < Torneo_Cupo Then
          Call rearmarTorneo
      End If
    End If
End If


If Torneo_HAYTORNEO Then
    For i = 1 To LastUser
        If UserList(i).flags.enTorneoAuto = True Then
            If UserList(i).Counters.Torneo > 0 Then
                UserList(i).Counters.Torneo = UserList(i).Counters.Torneo - 1
                If UserList(i).Counters.Torneo = 0 Then Call terminarDuelo(i)
            End If
        End If
    Next i
End If

'CHOTS | Torneos automáticos

If MinutosLatsClean >= 15 Then
    MinutosLatsClean = 0
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Call LimpiarMundo
Else
    MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.number & ": " & Err.Description)

End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub


Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, 0, "!!" & BroadMsg.text & ENDC)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> " & BroadMsg.text & FONTTYPE_SERVER)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hWnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim n As Integer
n = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #n
Print #n, Date & " " & Time & " server cerrado."
Close #n

End

Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub NuevoGameTimer_Timer()
On Error GoTo hayerror
'CHOTS | Nuevo timer, intervalo de 1 segundo
Dim iUserIndex As Integer
For iUserIndex = 1 To MaxUsers

    'Conexion activa?
    If UserList(iUserIndex).ConnID <> -1 Then
  
        '¿User valido?
        If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
            If UserList(iUserIndex).flags.Desnudo And UserList(iUserIndex).flags.Muerto = 0 And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoFrio(iUserIndex)
            If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Muerto = 0 And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoVeneno(iUserIndex)
            If UserList(iUserIndex).flags.AdminInvisible <> 1 And UserList(iUserIndex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
            If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
            If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
            Call DuracionPociones(iUserIndex)
            Call hambreYSed(iUserIndex)
        End If
    End If
Next iUserIndex

Exit Sub
hayerror:
LogError ("Error en NuevoGameTimer: " & Err.Description & " UserIndex = " & iUserIndex)

End Sub

Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim bEnviarStats As Boolean

On Error GoTo hayerror

 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iUserIndex = 1 To MaxUsers

    'Conexion activa?
    If UserList(iUserIndex).ConnID <> -1 Then
    
        '¿User valido?
        If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
         
            '[Alejo-18-5]
            bEnviarStats = False

            'CHOTS | Removido de aca
            'Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.X, UserList(iUserIndex).Pos.Y)

            If UserList(iUserIndex).flags.Muerto = 0 Then
                   
                If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)

                If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                    If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                    If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                ElseIf UserList(iUserIndex).flags.Descansar Then
                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                    If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                    If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False

                    If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                        Call SendData(SendTarget.ToIndex, iUserIndex, 0, ServerPackages.dialogo & "Has terminado de descansar." & FONTTYPE_INFO)
                        UserList(iUserIndex).flags.Descansar = False
                    End If

                End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.hambre = 0 And UserList(UserIndex).Flags.Sed = 0)

            End If 'Muerto
        Else
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
                UserList(iUserIndex).Counters.IdleCount = 0
                Call CloseSocket(iUserIndex)
            End If
        End If 'UserLogged
    End If

   Next iUserIndex

Exit Sub
hayerror:
LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuCerrar_Click()

Call ActualizarWebUsuarios(0)

If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

 

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim Npc As Integer

For Npc = 1 To LastNPC
    Npclist(Npc).CanAttack = 1
Next Npc

End Sub



Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                ''ia comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                     'Usamos AI si hay algun user en el mapa
                     If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                     End If
                     mapa = Npclist(NpcIndex).Pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Or Npclist(NpcIndex).guerra.enGuerra = True Then
                                        Call NPCAI(NpcIndex)
                                  End If
                          End If
                     End If
                     
                End If
        End If
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
Next i

End Sub

Private Sub tmrAntiBoter_Timer()
'CHOTS, BysNacK | Anti Bots
Erase ArrayIps
NumIps = 0
End Sub

Private Sub tmrCastillos_Timer()
On Error GoTo ErrorHandler

'CHOTS | Sistema de Castillos

 Call DarPremioCastillos
 Exit Sub

ErrorHandler:
 Call LogError("Error en tmrCastillos " & Err.number & ": " & Err.Description)
 Exit Sub
End Sub

Private Sub tmrInvocaciones_Timer()
On Error GoTo ErrorHandler
'CHOTS | Sistema de Invocaciones
If PuedeInvocar Then InvocarInvocacion

If MapData(INVOCACIONTELEP_MAPA, INVOCACIONTELEP_X1, INVOCACIONTELEP_Y1).UserIndex <> 0 And MapData(INVOCACIONTELEP_MAPA, INVOCACIONTELEP_X2, INVOCACIONTELEP_Y2).UserIndex <> 0 Then
        
    Dim ViejaPos As WorldPos
    Dim NuevaPos As WorldPos
            
    If LegalPos(INVOCACION_MAPA, INVOCACIONTELEP_X3, INVOCACIONTELEP_Y3) Then
        Call WarpUserChar(MapData(INVOCACIONTELEP_MAPA, INVOCACIONTELEP_X1, INVOCACIONTELEP_Y1).UserIndex, INVOCACION_MAPA, INVOCACIONTELEP_X3, INVOCACIONTELEP_Y3, True)
    Else
        ViejaPos.Map = INVOCACION_MAPA
        ViejaPos.X = INVOCACIONTELEP_X3
        ViejaPos.Y = INVOCACIONTELEP_Y3
        Call ClosestLegalPos(ViejaPos, NuevaPos)
        Call WarpUserChar(MapData(INVOCACIONTELEP_MAPA, INVOCACIONTELEP_X1, INVOCACIONTELEP_Y1).UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    End If
        
    If LegalPos(INVOCACION_MAPA, INVOCACIONTELEP_X4, INVOCACIONTELEP_Y4) Then
        Call WarpUserChar(MapData(INVOCACIONTELEP_MAPA, INVOCACIONTELEP_X2, INVOCACIONTELEP_Y2).UserIndex, INVOCACION_MAPA, INVOCACIONTELEP_X4, INVOCACIONTELEP_Y4, True)
    Else
        ViejaPos.Map = INVOCACION_MAPA
        ViejaPos.X = INVOCACIONTELEP_X4
        ViejaPos.Y = INVOCACIONTELEP_Y4
        Call ClosestLegalPos(ViejaPos, NuevaPos)
        Call WarpUserChar(MapData(INVOCACIONTELEP_MAPA, INVOCACIONTELEP_X2, INVOCACIONTELEP_Y2).UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
    End If

End If
Exit Sub

ErrorHandler:
 Call LogError("Error en tmrInvocaciones " & Err.number & ": " & Err.Description)
 Exit Sub
'CHOTS | Sistema de Invocaciones
End Sub

Private Sub tmrMinutosGuerra_Timer()
Call TimerMinutosGuerra
End Sub

Private Sub tmrOnline_Timer()
  Call ActualizarWebUsuarios
End Sub

Private Sub tmrSegundosGuerra_Timer()
Call TimerSegundosGuerra
End Sub

Private Sub tPiqueteC_Timer()
On Error GoTo errhandler
Static segundos As Integer

segundos = segundos + 6

Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
            
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                Call SendData(SendTarget.ToIndex, i, 0, "Z39")
                If UserList(i).Counters.PiqueteC > 23 Then
                  UserList(i).Counters.PiqueteC = 0
                  Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                End If
            Else
                If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If

            If segundos >= 18 Then
                UserList(i).Counters.Pasos = 0
            End If
    End If
Next i

If segundos >= 18 Then segundos = 0
   
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.number & ": " & Err.Description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex
        
        'CHOTS | Seguridad by DyE
        UserList(NewIndex).ClavePublica = SecurityParameters.keyA
        UserList(NewIndex).ClavePrivada = SecurityParameters.keyB
        'CHOTS | Seguridad by DyE
        
        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
Dim t() As String
Dim LoopC As Long
Dim RD As String
On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If

RD = StrConv(Datos, vbUnicode)

'call logindex(MiDato, "Read. ConnId: " & ID & " Midato: " & MiDato & " Dato: " & RD)

UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD

t = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(MiDato).RDBuffer = t(UBound(t))
    
    For LoopC = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If t(LoopC) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(t(LoopC)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then
                Call HandleData(MiDato, t(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If
Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub

#End If
