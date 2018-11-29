Attribute VB_Name = "Monturas"
'MÓDULO DE MONTURAS
'CREADO POR JUAN ANDRÉS DALMASSO (CHOTS)
'EL 02/11/11
'PARA LAPSUS 3.0
'
'Mascotas
'1=Tigre
'2=Pony
'3=Efelante
'4=Draco Rojo
'5=Draco Dorado
'6=Caballo
Private Const PrimerNpc = 586 'CHOTS | Primer Mascota (Tigre)
Private Const PrimerItem = 1139 'CHOTS | Primer Mascota (Tigre)
Public MapasMontura() As Integer

Public Sub CrearMontura(UserIndex As Integer, Tipo As Integer)
    Tipo = (Tipo - PrimerNpc) + 1
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = (PrimerItem + Tipo) - 1
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End Sub

Public Sub cargarMapasMonturas()
    Dim i As Integer
    Open (App.Path & "\Dat\MapasMontura.txt") For Input As #1
        Do While Not EOF(1)
            i = i + 1
            ReDim Preserve MapasMontura(i)
            Input #1, MapasMontura(i)
        Loop
    Close #1
End Sub

Public Sub RespawnearMonturas()
Dim i As Byte
Dim wpaux As WorldPos

'Tigre
wpaux.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
wpaux.X = RandomNumber(20, 80)
wpaux.Y = RandomNumber(20, 80)
Call SpawnNpc(PrimerNpc, wpaux, False, True)

'Unicornio
wpaux.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
wpaux.X = RandomNumber(20, 80)
wpaux.Y = RandomNumber(20, 80)
Call SpawnNpc(PrimerNpc + 1, wpaux, False, True)

'Efelante
wpaux.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
wpaux.X = RandomNumber(20, 80)
wpaux.Y = RandomNumber(20, 80)
Call SpawnNpc(PrimerNpc + 2, wpaux, False, True)


'Draco Rojo
wpaux.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
wpaux.X = RandomNumber(20, 80)
wpaux.Y = RandomNumber(20, 80)
Call SpawnNpc(PrimerNpc + 3, wpaux, False, True)

'Draco Dorado
wpaux.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
wpaux.X = RandomNumber(20, 80)
wpaux.Y = RandomNumber(20, 80)
Call SpawnNpc(PrimerNpc + 4, wpaux, False, True)

'Caballo
For i = 1 To 3
    wpaux.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
    wpaux.X = RandomNumber(20, 80)
    wpaux.Y = RandomNumber(20, 80)
    Call SpawnNpc(PrimerNpc + 5, wpaux, False, True)
Next i

End Sub

Public Sub respawnMontura(ByVal Numero As Integer)
    If Numero < 586 Then Exit Sub
    Dim respawnPos As WorldPos
    Dim nPos as WorldPos
    respawnPos.Map = MapasMontura(RandomNumber(1, UBound(MapasMontura)))
    respawnPos.X = RandomNumber(20, 80)
    respawnPos.Y = RandomNumber(20, 80)
    Call ClosestLegalPos(respawnPos, nPos)
    Call CrearNPC(Numero, nPos.Map, nPos)
End Sub

Public Sub DesMascotar(UserIndex As Integer)
    Dim lBody As Integer

    UserList(UserIndex).flags.montandoMascota = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_CAPTURA)
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex = 0 Then
        Call DarCuerpoDesnudo(UserIndex)
        lBody = UserList(UserIndex).char.Body
    Else
        lBody = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
    End If
    
    
    Call ChangeUserChar(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserIndex, lBody, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & FXIDs.FXCAPTURA & ",0")
  
End Sub
