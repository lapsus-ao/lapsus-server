Attribute VB_Name = "Invocaciones"
'Modulo de Invocaciones
'Programado por Juan Andrés Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'Para Lapsus AO 2.0
'05/09/2010

'CHOTS | Mapa donde está la invocación
Public Const INVOCACION_MAPA As Byte = 63

'CHOTS | Mapa anterior
Public Const INVOCACIONTELEP_MAPA As Byte = 60
Public Const INVOCACIONTELEP_X1 As Byte = 62
Public Const INVOCACIONTELEP_Y1 As Byte = 29
Public Const INVOCACIONTELEP_X2 As Byte = 68
Public Const INVOCACIONTELEP_Y2 As Byte = 29

'CHOTS | Mapa de invocacion donde caen los teleportados
Public Const INVOCACIONTELEP_X3 As Byte = 64
Public Const INVOCACIONTELEP_Y3 As Byte = 54
Public Const INVOCACIONTELEP_X4 As Byte = 68
Public Const INVOCACIONTELEP_Y4 As Byte = 54

'CHOTS | Mapa donde se tienen q parar para que salga el chobi
Public Const INVOCACION_X1 As Byte = 61
Public Const INVOCACION_Y1 As Byte = 33
Public Const INVOCACION_X2 As Byte = 67
Public Const INVOCACION_Y2 As Byte = 33
Public Const INVOCACION_X3 As Byte = 61
Public Const INVOCACION_Y3 As Byte = 39
Public Const INVOCACION_X4 As Byte = 67
Public Const INVOCACION_Y4 As Byte = 39

'CHOTS | Datos del chobi
Public Const INVOCACION_RESPAWNX As Byte = 64
Public Const INVOCACION_RESPWANY As Byte = 36
Public Const INVOCACION_NPC As Integer = 584

Public INVOCACION_INVOCADO As Boolean

Public Sub InvocarInvocacion()
    Dim Pos As WorldPos
    Pos.Map = INVOCACION_MAPA
    Pos.X = INVOCACION_RESPAWNX
    Pos.Y = INVOCACION_RESPWANY
    Call SpawnNpc(INVOCACION_NPC, Pos, True, False)
    Call SendData(SendTarget.ToAll, 0, 0, "Z78")
    INVOCACION_INVOCADO = True
End Sub

Public Sub MuereInvocacion(ByVal Npc As Integer)
    Call QuitarNPC(Npc)
    INVOCACION_INVOCADO = False
End Sub

Public Function PuedeInvocar() As Boolean

PuedeInvocar = False

If (MapData(INVOCACION_MAPA, INVOCACION_X1, INVOCACION_Y1).UserIndex <> 0) And (MapData(INVOCACION_MAPA, INVOCACION_X2, INVOCACION_Y2).UserIndex <> 0) And (MapData(INVOCACION_MAPA, INVOCACION_X3, INVOCACION_Y3).UserIndex <> 0) And (MapData(INVOCACION_MAPA, INVOCACION_X4, INVOCACION_Y4).UserIndex <> 0) Then
    If UserList(MapData(INVOCACION_MAPA, INVOCACION_X1, INVOCACION_Y1).UserIndex).flags.Muerto = 0 And UserList(MapData(INVOCACION_MAPA, INVOCACION_X2, INVOCACION_Y2).UserIndex).flags.Muerto = 0 And UserList(MapData(INVOCACION_MAPA, INVOCACION_X3, INVOCACION_Y3).UserIndex).flags.Muerto = 0 And UserList(MapData(INVOCACION_MAPA, INVOCACION_X4, INVOCACION_Y4).UserIndex).flags.Muerto = 0 Then
        If Not INVOCACION_INVOCADO Then PuedeInvocar = True
    End If
End If

End Function
