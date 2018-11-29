Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal Userindex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(Userindex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(Userindex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal Userindex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(Userindex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(Userindex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(Userindex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, Userindex, j)
        
        End If
Next

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UserList(Userindex).Pos.Map = 37 Then
    
    Dim DeDonde As WorldPos
    
    Select Case UCase$(UserList(Userindex).Hogar)
        Case "LINDOS" 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case "ULLATHORPE"
            DeDonde = Ullathorpe
        Case "BANDERBILL"
            DeDonde = Banderbill
        Case Else
            DeDonde = Nix
    End Select
       
    Call WarpUserChar(Userindex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal Userindex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(Userindex).Invent.Object(j).ObjIndex = 0
        UserList(Userindex).Invent.Object(j).Amount = 0
        UserList(Userindex).Invent.Object(j).Equipped = 0
        
Next

UserList(Userindex).Invent.NroItems = 0

UserList(Userindex).Invent.ArmourEqpObjIndex = 0
UserList(Userindex).Invent.ArmourEqpSlot = 0

UserList(Userindex).Invent.WeaponEqpObjIndex = 0
UserList(Userindex).Invent.WeaponEqpSlot = 0

UserList(Userindex).Invent.CascoEqpObjIndex = 0
UserList(Userindex).Invent.CascoEqpSlot = 0

UserList(Userindex).Invent.EscudoEqpObjIndex = 0
UserList(Userindex).Invent.EscudoEqpSlot = 0

UserList(Userindex).Invent.HerramientaEqpObjIndex = 0
UserList(Userindex).Invent.HerramientaEqpSlot = 0

UserList(Userindex).Invent.MunicionEqpObjIndex = 0
UserList(Userindex).Invent.MunicionEqpSlot = 0

UserList(Userindex).Invent.BarcoObjIndex = 0
UserList(Userindex).Invent.BarcoSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal Userindex As Integer)
On Error GoTo errhandler

If Cantidad > 10000 Then Exit Sub

'SI EL NPC TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(Userindex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon
        If Cantidad > 39999 Then
            Dim j As Integer
            Dim k As Integer
            Dim m As Integer
            Dim Cercanos As String
            m = UserList(Userindex).Pos.Map
            For j = UserList(Userindex).Pos.X - 10 To UserList(Userindex).Pos.X + 10
                For k = UserList(Userindex).Pos.Y - 10 To UserList(Userindex).Pos.Y + 10
                    If InMapBounds(m, j, k) Then
                        If MapData(m, j, k).Userindex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(m, j, k).Userindex).Name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(Userindex).Name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        
        Do While (Cantidad > 0) And (UserList(Userindex).Stats.GLD > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(Userindex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            
            If UserList(Userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(Userindex).Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
            
            Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
    
End If

Exit Sub

errhandler:

End Sub

Sub QuitarUserInvItem(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(Userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(Userindex, Slot)

'CHOTS | Monturas
If ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex).OBJType = eOBJType.otMontura And UserList(Userindex).flags.montandoMascota = 1 Then Call DesMascotar(Userindex)

'Quita un objeto
UserList(Userindex).Invent.Object(Slot).Amount = UserList(Userindex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(Userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
    UserList(Userindex).Invent.Object(Slot).ObjIndex = 0
    UserList(Userindex).Invent.Object(Slot).Amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte, Optional Conecta As Boolean = False)

Dim NullObj As UserOBJ
Dim LoopC As Byte

If Conecta Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call ChangeUserInvConecta(Userindex, LoopC, UserList(Userindex).Invent.Object(LoopC))
    Next LoopC
    Exit Sub
End If

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(Userindex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(Userindex, Slot, UserList(Userindex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(Userindex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(Userindex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(Userindex, LoopC, UserList(Userindex).Invent.Object(LoopC))
        Else
            
            Call ChangeUserInv(Userindex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub DropObj(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(Userindex).Invent.Object(Slot).Amount Then num = UserList(Userindex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = 0 Or MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex Then
        If UserList(Userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(Userindex, Slot)
        Obj.ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        If num + MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.Amount
        End If
        
        Obj.Amount = num
        
        Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
        Call QuitarUserInvItem(Userindex, Slot, num)
        Call UpdateUserInv(False, Userindex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!" & FONTTYPE_TALK)
        End If
        If ObjData(Obj.ObjIndex).Caos = 1 Or ObjData(Obj.ObjIndex).Real = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "¡ATENCION!! ¡¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!!" & FONTTYPE_TALK)
        End If
        
        If UserList(Userindex).flags.Privilegios > PlayerType.User Then Call LogGM("EDITADOS", UserList(Userindex).Name & " tiró " & num & " " & ObjData(Obj.ObjIndex).Name, False)
  Else
    Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No hay espacio en el piso." & FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToAreaByPos(Map, X, Y, "BO" & X & "," & Y)
   Else
        Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
    End If
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, X, Y).OBJInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount + Obj.Amount
    Else
        MapData(Map, X, Y).OBJInfo = Obj
        
        If sndRoute = SendTarget.ToMap Then
            Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
        Else
            Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y)
        End If
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal Userindex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(Userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(Userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(Userindex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call SendData(SendTarget.ToIndex, Userindex, 0, "Z24")
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(Userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(Userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(Userindex).Invent.Object(Slot).Amount = UserList(Userindex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(Userindex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, Userindex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal Userindex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(Userindex).Pos.X
        Y = UserList(Userindex).Pos.Y
        Obj = ObjData(MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.ObjIndex
        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z24")
        Else
            'Quitamos el objeto
            Call EraseObj(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, MapData(UserList(Userindex).Pos.Map, X, Y).OBJInfo.Amount, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
            If UserList(Userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(Userindex).Name, "Agarro: " & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name, False)
            If UserList(Userindex).flags.Privilegios > PlayerType.User Then Call LogGM("EDITADOS", UserList(Userindex).Name & " Agarro " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, False)
        End If
        
    End If
Else
End If

End Sub

Sub Desequipar(ByVal Userindex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(Userindex).Invent.Object)) Or (Slot > UBound(UserList(Userindex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(Userindex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        UserList(Userindex).Invent.Object(Slot).Equipped = 0
        UserList(Userindex).Invent.WeaponEqpObjIndex = 0
        UserList(Userindex).Invent.WeaponEqpSlot = 0
        Call SendUserArma(Userindex)
        If Not UserList(Userindex).flags.Mimetizado = 1 Then
            UserList(Userindex).char.WeaponAnim = NingunArma
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
        End If
    
    Case eOBJType.otFlechas
        UserList(Userindex).Invent.Object(Slot).Equipped = 0
        UserList(Userindex).Invent.MunicionEqpObjIndex = 0
        UserList(Userindex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otHerramientas
        UserList(Userindex).Invent.Object(Slot).Equipped = 0
        UserList(Userindex).Invent.HerramientaEqpObjIndex = 0
        UserList(Userindex).Invent.HerramientaEqpSlot = 0
    
    Case eOBJType.otArmadura
        If UserList(Userindex).flags.montandoMascota = 1 Then Call DesMascotar(Userindex) 'CHOTS | Monturas
        UserList(Userindex).Invent.Object(Slot).Equipped = 0
        UserList(Userindex).Invent.ArmourEqpObjIndex = 0
        UserList(Userindex).Invent.ArmourEqpSlot = 0
        Call SendUserRopa(Userindex)
        Call SendUserDefMag(Userindex)
        Call DarCuerpoDesnudo(Userindex, UserList(Userindex).flags.Mimetizado = 1)
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
            
    Case eOBJType.otCASCO
        UserList(Userindex).Invent.Object(Slot).Equipped = 0
        UserList(Userindex).Invent.CascoEqpObjIndex = 0
        UserList(Userindex).Invent.CascoEqpSlot = 0
        Call SendUserCasco(Userindex)
        Call SendUserDefMag(Userindex)
        If Not UserList(Userindex).flags.Mimetizado = 1 Then
            UserList(Userindex).char.CascoAnim = NingunCasco
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
        UserList(Userindex).Invent.Object(Slot).Equipped = 0
        UserList(Userindex).Invent.EscudoEqpObjIndex = 0
        UserList(Userindex).Invent.EscudoEqpSlot = 0
        Call SendUserEscu(Userindex)
        If Not UserList(Userindex).flags.Mimetizado = 1 Then
            UserList(Userindex).char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
        End If
End Select

Call EnviarSta(Userindex)
Call UpdateUserInv(False, Userindex, Slot)
'Call SendUserHitBox(UserIndex)

End Sub

Function SexoPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(Userindex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(Userindex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(Userindex) Then
        If UserList(Userindex).Faccion.Jerarquia >= ObjData(ObjIndex).Jerarquia Then
            FaccionPuedeUsarItem = (UserList(Userindex).Faccion.ArmadaReal = 1)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(Userindex) Then
        If UserList(Userindex).Faccion.Jerarquia >= ObjData(ObjIndex).Jerarquia Then
            FaccionPuedeUsarItem = (UserList(Userindex).Faccion.FuerzasCaos = 1)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal Userindex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(Userindex) Then
     Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.OBJType
    Case eOBJType.otWeapon
    If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then
If ObjData(ObjIndex).DosManos = 1 Then
Call SendData(SendTarget.ToIndex, Userindex, 0, "Z63")
Exit Sub
End If
End If
       If ClasePuedeUsarItem(Userindex, ObjIndex) And _
          FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(Userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, Slot)
                    'Animacion por defecto
                    If UserList(Userindex).flags.Mimetizado = 1 Then
                        UserList(Userindex).CharMimetizado.WeaponAnim = NingunArma
                    Else
                        UserList(Userindex).char.WeaponAnim = NingunArma
                        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
                    End If
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, UserList(Userindex).Invent.WeaponEqpSlot)
                End If
        
                UserList(Userindex).Invent.Object(Slot).Equipped = 1
                UserList(Userindex).Invent.WeaponEqpObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
                UserList(Userindex).Invent.WeaponEqpSlot = Slot
                Call SendUserArma(Userindex)
                
                'Sonido
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_SACARARMA)
        
                If UserList(Userindex).flags.Mimetizado = 1 Then
                    UserList(Userindex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
                Else
                    UserList(Userindex).char.WeaponAnim = Obj.WeaponAnim
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
                End If
       Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z42")
       End If
    
    Case eOBJType.otHerramientas
       If ClasePuedeUsarItem(Userindex, ObjIndex) And _
          FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(Userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(Userindex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, UserList(Userindex).Invent.HerramientaEqpSlot)
                End If
        
                UserList(Userindex).Invent.Object(Slot).Equipped = 1
                UserList(Userindex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(Userindex).Invent.HerramientaEqpSlot = Slot
                
       Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z42")
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(Userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(Userindex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, UserList(Userindex).Invent.MunicionEqpSlot)
                End If
        
                UserList(Userindex).Invent.Object(Slot).Equipped = 1
                UserList(Userindex).Invent.MunicionEqpObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
                UserList(Userindex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z42")
       End If
    
    Case eOBJType.otArmadura
        If UserList(Userindex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) Then
           
           'Si esta equipado lo quita
            If UserList(Userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(Userindex, Slot)
                Call DarCuerpoDesnudo(Userindex, UserList(Userindex).flags.Mimetizado = 1)
                If Not UserList(Userindex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(Userindex, UserList(Userindex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            If UserList(Userindex).flags.montandoMascota = 1 Then Call DesMascotar(Userindex)
            UserList(Userindex).Invent.Object(Slot).Equipped = 1
            UserList(Userindex).Invent.ArmourEqpObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
            UserList(Userindex).Invent.ArmourEqpSlot = Slot
            Call SendUserRopa(Userindex)
            Call SendUserDefMag(Userindex)
                
            If UserList(Userindex).flags.Mimetizado = 1 Then
                UserList(Userindex).CharMimetizado.Body = Obj.Ropaje
            Else
                UserList(Userindex).char.Body = Obj.Ropaje
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
            End If
            UserList(Userindex).flags.Desnudo = 0
            

        Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z42")
        End If
    
    Case eOBJType.otCASCO
        If UserList(Userindex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(Userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(Userindex, Slot)
                If UserList(Userindex).flags.Mimetizado = 1 Then
                    UserList(Userindex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(Userindex).char.CascoAnim = NingunCasco
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(Userindex, UserList(Userindex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(Userindex).Invent.Object(Slot).Equipped = 1
            UserList(Userindex).Invent.CascoEqpObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
            UserList(Userindex).Invent.CascoEqpSlot = Slot
            Call SendUserCasco(Userindex)
            Call SendUserDefMag(Userindex)
            If UserList(Userindex).flags.Mimetizado = 1 Then
                UserList(Userindex).CharMimetizado.CascoAnim = Obj.CascoAnim
            Else
                UserList(Userindex).char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
            End If
        Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z42")
        End If
    
    Case eOBJType.otESCUDO
    If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).DosManos = 1 Then
Call SendData(SendTarget.ToIndex, Userindex, 0, "Z64")
Exit Sub
End If
End If

        If UserList(Userindex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(Slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(Userindex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(Userindex, Slot)
                 If UserList(Userindex).flags.Mimetizado = 1 Then
                     UserList(Userindex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(Userindex).char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(Userindex, UserList(Userindex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(Userindex).Invent.Object(Slot).Equipped = 1
             UserList(Userindex).Invent.EscudoEqpObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
             UserList(Userindex).Invent.EscudoEqpSlot = Slot
             Call SendUserEscu(Userindex)
             
             If UserList(Userindex).flags.Mimetizado = 1 Then
                 UserList(Userindex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
             Else
                 UserList(Userindex).char.ShieldAnim = Obj.ShieldAnim
                 
                 Call ChangeUserChar(SendTarget.ToMap, 0, UserList(Userindex).Pos.Map, Userindex, UserList(Userindex).char.Body, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
             End If
         Else
             Call SendData(SendTarget.ToIndex, Userindex, 0, "Z42")
         End If
End Select

'Actualiza
Call UpdateUserInv(False, Userindex, Slot)
'Call SendUserHitBox(UserIndex)

Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

'Verifica si la raza puede usar la ropa 'CHOTS |Raza Orco
If UserList(Userindex).Raza = "Humano" Or _
   UserList(Userindex).Raza = "Elfo" Or _
   UserList(Userindex).Raza = "Orco" Or _
   UserList(Userindex).Raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal Userindex As Integer, ByVal Slot As Byte, Optional Numero As Byte)
On Local Error GoTo errh
'Usa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(Userindex).Invent.Object(Slot).Amount = 0 Then Exit Sub

Obj = ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex)
If Obj.Newbie = 1 And Not EsNewbie(Userindex) Then
    Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Solo los newbies pueden usar estos objetos." & FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
    If Obj.proyectil = 1 Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsarArcos(Userindex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
End If

ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
UserList(Userindex).flags.TargetObjInvIndex = ObjIndex
UserList(Userindex).flags.TargetObjInvSlot = Slot
Select Case Obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If

        'Usa el item
        UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam + Obj.MinHam
        If UserList(Userindex).Stats.MinHam > UserList(Userindex).Stats.MaxHam Then _
            UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam
        UserList(Userindex).flags.Hambre = 0
        Call EnviarhambreYsed(Userindex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(Userindex, Slot, 1)
        
        Call UpdateUserInv(False, Userindex, Slot)

    Case eOBJType.otGuita
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If
        
        UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + UserList(Userindex).Invent.Object(Slot).Amount
        UserList(Userindex).Invent.Object(Slot).Amount = 0
        UserList(Userindex).Invent.Object(Slot).ObjIndex = 0
        UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, Userindex, Slot)
        Call EnviarOro(Userindex)
        
    Case eOBJType.otWeapon
        If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
        End If

        If ObjData(ObjIndex).proyectil = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Proyectiles)
        Else
            If UserList(Userindex).flags.TargetObj = 0 Then Exit Sub
            
            '¿El target-objeto es leña?
            If UserList(Userindex).flags.TargetObj = Leña Then
                If UserList(Userindex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(Userindex).flags.TargetObjMap, _
                         UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY, Userindex)
                End If
            End If
        End If
    
    Case eOBJType.otPociones
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If
        
        UserList(Userindex).flags.TomoPocion = True
        UserList(Userindex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(Userindex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(Userindex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                
                If UCase$(UserList(Userindex).Raza) = "HUMANO" Then
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) >= STAT_MAXATRIBUTOS_HUMANO Then _
                        UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = STAT_MAXATRIBUTOS_HUMANO
                Else
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) >= STAT_MAXATRIBUTOS Then _
                        UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = STAT_MAXATRIBUTOS
                End If

                If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) >= (UserList(Userindex).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2) Then _
                    UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = (UserList(Userindex).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call EnviarDopa(Userindex)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
        
            Case 2 'Modif la fuerza
                UserList(Userindex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                
                If UCase$(UserList(Userindex).Raza) = "HUMANO" Then
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) >= STAT_MAXATRIBUTOS_HUMANO Then _
                        UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = STAT_MAXATRIBUTOS_HUMANO
                Else
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) >= STAT_MAXATRIBUTOS Then _
                        UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = STAT_MAXATRIBUTOS
                End If

                If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) >= (UserList(Userindex).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2) Then _
                    UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = (UserList(Userindex).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call EnviarDopa(Userindex)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
                
            Case 3 'Pocion roja, restaura HP

                'Usa el item
                UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(Userindex).Stats.MinHP > UserList(Userindex).Stats.MaxHP Then _
                    UserList(Userindex).Stats.MinHP = UserList(Userindex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
                
                'CHOTS | Si es la pota roja envia q suba 30 sino lo q dice el modificador
                If ObjIndex = 38 Then
                    Call EnviarPocionRoja(Userindex)
                Else
                    Call EnviarHP(Userindex)
                End If
                
            Case 4 'Pocion azul, restaura MANA

                'Usa el item
                UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN + Porcentaje(UserList(Userindex).Stats.MaxMAN, 5)
                If UserList(Userindex).Stats.MinMAN > UserList(Userindex).Stats.MaxMAN Then _
                    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
                
                Call EnviarPocionAzul(Userindex)
                
            Case 5 ' Pocion violeta
                If UserList(Userindex).flags.Envenenado = 1 Then
                    UserList(Userindex).flags.Envenenado = 0
                    Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Te has curado del envenenamiento." & FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
            
            Case 6  ' CHOTS | Pocion Azul Druida
                'Usa el item
                UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN + Porcentaje(UserList(Userindex).Stats.MaxMAN, 7)
                If UserList(Userindex).Stats.MinMAN > UserList(Userindex).Stats.MaxMAN Then _
                    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
                Call EnviarMn(Userindex)
                
            Case 7 ' CHOTS | Pocion Removedora
            
                If UserList(Userindex).flags.Paralizado = 1 Then
                    UserList(Userindex).flags.Paralizado = 0
                    Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Te has removido la parálisis!" & FONTTYPE_FIGHT)
                    Call QuitarUserInvItem(Userindex, Slot, 1)
                    Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "DOK")
                End If
                
            Case 8 ' CHOTS | Pocion de Energia
        
                UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(Userindex).Stats.MinSta > UserList(Userindex).Stats.MaxSta Then UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
                Call EnviarSta(Userindex)


       End Select
       
       Call UpdateUserInv(False, Userindex, Slot)

     Case eOBJType.otBebidas
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If
        UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + Obj.MinSed
        If UserList(Userindex).Stats.MinAGU > UserList(Userindex).Stats.MaxAGU Then _
            UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
        UserList(Userindex).flags.Sed = 0
        Call EnviarhambreYsed(Userindex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(Userindex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_BEBER)
        
        Call UpdateUserInv(False, Userindex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If
        
        If UserList(Userindex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(Userindex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(Userindex).flags.TargetObj = MapData(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Has abierto la puerta." & FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Has cerrado con llave la puerta." & FONTTYPE_INFO)
                        UserList(Userindex).flags.TargetObj = MapData(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No esta cerrada." & FONTTYPE_INFO)
                  Exit Sub
            End If
            
        End If
    
        Case eOBJType.otBotellaVacia
            If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
            End If
            If Not HayAgua(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY) Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No hay agua allí." & FONTTYPE_INFO)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(Userindex, Slot, 1)
            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
            End If
            
            Call UpdateUserInv(False, Userindex, Slot)
    
        Case eOBJType.otBotellaLlena
            If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
            End If
            UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + Obj.MinSed
            If UserList(Userindex).Stats.MinAGU > UserList(Userindex).Stats.MaxAGU Then _
                UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
            UserList(Userindex).flags.Sed = 0
            Call EnviarhambreYsed(Userindex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(Userindex, Slot, 1)
            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
            End If
            
        Case eOBJType.otHerramientas
            If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
            End If
            If Not UserList(Userindex).Stats.MinSta > 0 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(Userindex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
            If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(Userindex).Reputacion.PlebeRep = MAXREP
            Select Case ObjIndex
                Case CAÑA_PESCA, RED_PESCA
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & Talar)
                Case PIQUETE_MINERO
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & Mineria)
                Case TIJERA_DRUIDA
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & Botanica)
                Case MARTILLO_HERRERO
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(Userindex)
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "SFC")
                Case HILO_SASTRE
                    Call EnivarRopasConstruibles(Userindex)
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "SAS")
                Case OLLA
                    Call EnivarObjPocionesConstruibles(Userindex)
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "ALQ")

            End Select
        
    Case eOBJType.otPergaminos
                If UserList(Userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                    Exit Sub
                End If
               
                If UserList(Userindex).flags.Hambre = 0 And _
                   UserList(Userindex).flags.Sed = 0 Then
                If ClasePuedeUsarItem(Userindex, ObjIndex) And _
                    FaccionPuedeUsarItem(Userindex, ObjIndex) Then

                        Call AgregarHechizo(Userindex, Slot)
                        Call UpdateUserInv(False, Userindex, Slot)
                    Else
                        Call SendData(ToIndex, Userindex, 0, ServerPackages.dialogo & "Tú clase no puede aprender este hechizo." & FONTTYPE_INFO)
                    End If
                Else
                   Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Estas demasiado hambriento y sediento." & FONTTYPE_INFO)
                End If
       
       Case eOBJType.otMinerales
           If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
           End If
           Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & FundirMetal)
           
        'CHOTS | Monturas
       Case eOBJType.otGemaCaptura
           If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
           End If
           If UCase$(UserList(Userindex).Clase) <> "DRUIDA" Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z100")
                Exit Sub
           End If
           Call SendData(SendTarget.ToIndex, Userindex, 0, "T01" & CapturarNpc)
        'CHOTS | Monturas
       
       Case eOBJType.otInstrumentos
            If UserList(Userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
                Exit Sub
            End If
            Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & Obj.Snd1)
       
       Case eOBJType.otBarcos
    'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(Userindex).Stats.ELV < 25 Then
            If UCase$(UserList(Userindex).Clase) <> "PESCADOR" And UCase$(UserList(Userindex).Clase) <> "PIRATA" Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "Para recorrer los mares debes ser nivel 25 o superior." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If UserList(Userindex).flags.montandoMascota = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No puedes navegar estando montado" & FONTTYPE_MONTURA)
            Exit Sub
        End If
        If ((LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X - 1, UserList(Userindex).Pos.Y, True) Or _
            LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1, True) Or _
            LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X + 1, UserList(Userindex).Pos.Y, True) Or _
            LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, True)) And _
            UserList(Userindex).flags.Navegando = 0) _
            Or UserList(Userindex).flags.Navegando = 1 Then
           Call DoNavega(Userindex, Obj, Slot)
        Else
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "¡Debes aproximarte al agua para usar el barco!" & FONTTYPE_INFO)
        End If
        
    'CHOTS | Monturas
Case eOBJType.otMontura

        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If
        
        If UserList(Userindex).flags.Navegando = 1 Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No puedes montar si estás navegando!" & FONTTYPE_MONTURA)
            Exit Sub
        End If
        
        If MapInfo(UserList(Userindex).Pos.Map).Zona = "DUNGEON" Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No puedes montar en un Dungeon!" & FONTTYPE_MONTURA)
            Exit Sub
        End If
        
        If UserList(Userindex).Pos.Map > 59 And (UserList(Userindex).Pos.Map <> 80 Or UserList(Userindex).Pos.Map <> 81) Then
            Call SendData(SendTarget.ToIndex, Userindex, 0, ServerPackages.dialogo & "No puedes montar aquí!" & FONTTYPE_MONTURA)
            Exit Sub
        End If

        If UserList(Userindex).flags.montandoMascota = 0 Then

            If TieneObjetos(ROPAMONTURA, 1, Userindex) = False Then
                Call SendData(SendTarget.ToIndex, Userindex, 0, "Z98")
                Exit Sub
            End If

            UserList(Userindex).flags.montandoMascota = 1
            UserList(Userindex).Invent.MonturaObjIndex = ObjIndex
            UserList(Userindex).Invent.MonturaSlot = Slot
            Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "TW" & SND_CAPTURA)
            Call ChangeUserChar(ToMap, 0, UserList(Userindex).Pos.Map, Userindex, Obj.BodyMontura, UserList(Userindex).char.Head, UserList(Userindex).char.Heading, UserList(Userindex).char.WeaponAnim, UserList(Userindex).char.ShieldAnim, UserList(Userindex).char.CascoAnim)
            Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.Map, "CXF" & UserList(Userindex).char.CharIndex & "," & FXIDs.FXCAPTURA & ",0")

        Else
            Call DesMascotar(Userindex)
            UserList(Userindex).Invent.MonturaObjIndex = 0
            UserList(Userindex).Invent.MonturaSlot = 0
        End If
        Exit Sub
        
        
'CHOTS | Pasajes
Case eOBJType.otPasajes
     'Se asegura que el target es un npc
         If UserList(Userindex).flags.TargetNPC = 0 Then
             Call SendData(ToIndex, Userindex, 0, "Z30")
             Exit Sub
         End If
         ' Verificamos que sea el pirata
        If Npclist(UserList(Userindex).flags.TargetNPC).NPCtype <> eNPCType.Pasajes Then Exit Sub

         ' Si esta muy lejos no actua
         If Distancia(UserList(Userindex).Pos, Npclist(UserList(Userindex).flags.TargetNPC).Pos) > 5 Then
             Call SendData(ToIndex, Userindex, 0, "Z27")
             Exit Sub
         End If

         ' Si esta muerto no puede usar el pasaje.
        If UserList(Userindex).flags.Muerto = 1 Then
            Call SendData(ToIndex, Userindex, 0, "Z12")
            Exit Sub
        End If
        
        If UserList(Userindex).Stats.MinHam = 0 Or UserList(Userindex).Stats.MinAGU = 0 Then
            Call SendData(ToIndex, Userindex, 0, ServerPackages.dialogo & "Estas demasiado Hambriento y/o Sediento!" & FONTTYPE_INFO)
            Exit Sub
        End If

    ' Para prevenir que no se quede trabado el pj verificamos si el mapa de destino es un mapa valido.
            If MapaValido(Obj.mapa) Then
            Dim ViejaPos As WorldPos
            ViejaPos.Map = Obj.mapa
            ViejaPos.X = Obj.X
            ViejaPos.Y = Obj.Y
            
            Dim NuevaPos As WorldPos
        ' Transportamos al usuario
                
                If LegalPos(Obj.mapa, Obj.X, Obj.Y) Then
                    Call WarpUserChar(Userindex, Obj.mapa, Obj.X, Obj.Y, True)
                Else
                    Call ClosestLegalPos(ViejaPos, NuevaPos)
                    Call WarpUserChar(Userindex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                End If
                
                Call SendData(ToIndex, Userindex, 0, "Z81")
                UserList(Userindex).Stats.MinHam = 10
                UserList(Userindex).Stats.MinAGU = 10
                Call EnviarhambreYsed(Userindex)
                'Quitamos el Item del inventario.
                Call QuitarUserInvItem(Userindex, Slot, 1)
            Else
' Si no es un mapa valido se lo informamos al usuario.
                Call SendData(ToIndex, Userindex, 0, ServerPackages.dialogo & " El mapa al que se dirije el pase no es un mapa valido." & FONTTYPE_INFO)
         Exit Sub
     End If
           
                Call UpdateUserInv(False, Userindex, Slot)
       
        Exit Sub
           
           
           
           
End Select

Exit Sub
errh: Call LogError("Error en Usa")
'Actualiza
'call scenduserstatsbox(UserIndex)
'Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub EnivarArmasConstruibles(ByVal Userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(Userindex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(Userindex).Clase) Then
        If ObjData(ArmasHerrero(i)).OBJType = eOBJType.otWeapon Or ObjData(ArmasHerrero(i)).OBJType = eOBJType.otGemaCaptura Then
        '[DnG!]
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & " (" & ObjData(ArmasHerrero(i)).LingH & "-" & ObjData(ArmasHerrero(i)).LingP & "-" & ObjData(ArmasHerrero(i)).LingO & ")" & "," & ArmasHerrero(i) & ","
        '[/DnG!]
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).Name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData(SendTarget.ToIndex, Userindex, 0, "LAH" & cad$)

End Sub
Sub EnivarRopasConstruibles(ByVal Userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjSastre)
    If ObjData(ObjSastre(i)).SkSastreria <= UserList(Userindex).Stats.UserSkills(eSkill.Sastreria) / ModSastreria(UserList(Userindex).Clase) Then _
        cad$ = cad$ & ObjData(ObjSastre(i)).Name & " (" & ObjData(ObjSastre(i)).PielLobo & "-" & ObjData(ObjSastre(i)).PielOsoPardo & "-" & ObjData(ObjSastre(i)).PielOsoPolar & ")" & "," & ObjSastre(i) & ","
Next i

Call SendData(SendTarget.ToIndex, Userindex, 0, "OBS" & cad$)

End Sub
Sub EnivarObjPocionesConstruibles(ByVal Userindex As Integer)

Dim i As Integer, cad$
For i = 1 To UBound(ObjDruida)
    If ObjData(ObjDruida(i)).SkAlquimia <= UserList(Userindex).Stats.UserSkills(eSkill.Alquimia) / ModAlquimia(UserList(Userindex).Clase) Then _
        cad$ = cad$ & ObjData(ObjDruida(i)).Name & " (" & ObjData(ObjDruida(i)).Chala & ")" & "," & ObjDruida(i) & ","
Next i
Call SendData(SendTarget.ToIndex, Userindex, 0, "LGL" & cad$)

End Sub
Sub EnivarObjConstruibles(ByVal Userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) / ModCarpinteria(UserList(Userindex).Clase) Then _
        cad$ = cad$ & ObjData(ObjCarpintero(i)).Name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & ObjCarpintero(i) & ","
Next i

Call SendData(SendTarget.ToIndex, Userindex, 0, "OBR" & cad$)

End Sub

Sub EnivarArmadurasConstruibles(ByVal Userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).Clase) Then
        '[DnG!]
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).Name & " (" & ObjData(ArmadurasHerrero(i)).LingH & "-" & ObjData(ArmadurasHerrero(i)).LingP & "-" & ObjData(ArmadurasHerrero(i)).LingO & ")" & "," & ArmadurasHerrero(i) & ","
        '[/DnG!]
    End If
Next i

Call SendData(SendTarget.ToIndex, Userindex, 0, "LAR" & cad$)

End Sub


                   

Sub TirarTodo(ByVal Userindex As Integer)
On Error Resume Next

Call TirarTodosLosItems(Userindex)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And _
            (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And _
            ObjData(Index).OBJType <> eOBJType.otLlaves And _
            ObjData(Index).OBJType <> eOBJType.otBarcos And _
            ObjData(Index).NoSeCae = 0


End Function
Sub TirarBandera(ByVal Userindex As Integer)
    On Local Error Resume Next
    Dim ItemIndex As Integer
    Dim i As Byte
    Dim MiObj As Obj
    Dim NuevaPos As WorldPos
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(Userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
            If ObjData(ItemIndex).Bandera = 1 Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                MiObj.Amount = UserList(Userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                Tilelibre UserList(Userindex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
                End If
            End If
        End If
    Next i
End Sub

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    
    If EsNewbie(UserIndex) Then Exit Sub 'CHOTS | Los nws no pierden items
    If Userlist(UserIndex).guerra.enGuerra = True Then Exit Sub 'CHOTS | Guerras

    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(Userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                
                MiObj.Amount = UserList(Userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(Userindex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub
Sub DropItems(ByVal Userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
   
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(Userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
            If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
               
                'Creo el Obj
                MiObj.Amount = UserList(Userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
               
                Tilelibre UserList(Userindex).Pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
            End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal Userindex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(Userindex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(Userindex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            
            Tilelibre UserList(Userindex).Pos, NuevaPos, MiObj
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
