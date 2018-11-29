Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 10
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex <> 906 Then Suerte = Suerte + 20

'cazador con armadura de cazador oculto no se hace visible
If UCase$(UserList(UserIndex).Clase) = "CAZADOR" And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
    If UserList(UserIndex).Invent.ArmourEqpObjIndex = 774 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
        Exit Sub
    End If
End If

res = RandomNumber(1, Suerte)

If res > 9 Or EstaEnCastillo(UserIndex) Or UserList(UserIndex).flags.enTorneoAuto Or UserList(UserIndex).Reto.enReto Or UserList(UserIndex).guerra.enGuerra Then
    UserList(UserIndex).flags.Oculto = 0
    If UserList(UserIndex).flags.Invisible = 0 Then
        Dim ChotsNover As String
        ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z11")
    End If
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 55
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 50
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 48
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 44
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 42
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 40
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 38
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 25
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex = 906 Or UCase$(UserList(UserIndex).Clase) = "CAZADOR" Then Suerte = Suerte - 20

If EstaEnCastillo(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRE79")
    Exit Sub
End If

'CHOTS | Guerras
If UserList(UserIndex).guerra.enGuerra = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ocultarte en una Guerra!" & FONTTYPE_GUERRA)
    Exit Sub
End If

If UserList(UserIndex).flags.enTorneoAuto = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ocultarte en un Torneo!" & FONTTYPE_TORNEOAUTO)
    Exit Sub
End If

If UserList(UserIndex).Reto.enReto = True Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ocultarte en un Reto!" & FONTTYPE_DUELO)
    Exit Sub
End If

res = RandomNumber(1, Suerte)

If res <= 5 Then
    UserList(UserIndex).flags.Oculto = 1
    Dim ChotsNover As String
    ChotsNover = UserList(UserIndex).char.CharIndex & ",1"
    'ChotsNover = Encriptar(ChotsNover)
    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z52")
    Call SubirSkill(UserIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z53")
        UserList(UserIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).Clase)

If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).Invent.BarcoSlot = Slot

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).char.Body = Barco.Ropaje
    Else
        UserList(UserIndex).char.Body = iFragataFantasmal
    End If
    
    UserList(UserIndex).char.ShieldAnim = NingunEscudo
    UserList(UserIndex).char.WeaponAnim = NingunArma
    UserList(UserIndex).char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
    
Else
    
    UserList(UserIndex).flags.Navegando = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).char.Body = iCuerpoMuerto
        UserList(UserIndex).char.Head = iCabezaMuerto
        UserList(UserIndex).char.ShieldAnim = NingunEscudo
        UserList(UserIndex).char.WeaponAnim = NingunArma
        UserList(UserIndex).char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "NAVEG")

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(UserIndex).Clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes conocimientos de mineria suficientes para trabajar este mineral." & FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal itemIndex As Integer, ByVal Cant As Long, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = itemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal itemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = itemIndex Then
        
        'CHOTS | Monturas
        If ObjData(itemIndex).OBJType = eOBJType.otMontura And UserList(UserIndex).flags.montandoMascota = 1 Then Call DesMascotar(UserIndex)
            
        Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - Cant
        If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
            UserList(UserIndex).Invent.Object(i).Amount = 0
            UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, Cantidad)
    If ObjData(itemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(itemIndex).LingH * Cantidad, UserIndex)
    If ObjData(itemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(itemIndex).LingP * Cantidad, UserIndex)
    If ObjData(itemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(itemIndex).LingO * Cantidad, UserIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)
    If ObjData(itemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(itemIndex).Madera * Cantidad, UserIndex)
End Sub
Sub DruidaQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)
    If ObjData(itemIndex).Chala > 0 Then Call QuitarObjetos(Chala, ObjData(itemIndex).Chala * Cantidad, UserIndex)
End Sub
Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)
    Dim CantPielesLobo As Integer
    Dim CantPielesOsoPardo As Integer
    Dim CantPielesOsoPolar As Integer
    CantPielesLobo = ObjData(itemIndex).PielLobo * Cantidad
    CantPielesOsoPardo = ObjData(itemIndex).PielOsoPardo * Cantidad
    CantPielesOsoPolar = ObjData(itemIndex).PielOsoPolar * Cantidad
    If ObjData(itemIndex).PielLobo > 0 Then Call QuitarObjetos(PielLobo, CantPielesLobo, UserIndex)
    If ObjData(itemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PielOsoPardo, CantPielesOsoPardo, UserIndex)
    If ObjData(itemIndex).PielOsoPolar > 0 Then Call QuitarObjetos(PielOsoPolar, CantPielesOsoPolar, UserIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single) As Boolean
    
    If ObjData(itemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(itemIndex).Madera * Cantidad, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes madera. Precisas de " & ObjData(itemIndex).Madera * Cantidad & " leños." & FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
Function DruidaTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single) As Boolean
    
    If ObjData(itemIndex).Chala > 0 Then
            If Not TieneObjetos(Chala, ObjData(itemIndex).Chala * Cantidad, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).Chala * Cantidad & " raíces." & FONTTYPE_INFO)
                    DruidaTieneMateriales = False
                    Exit Function
            End If
    End If
    
    DruidaTieneMateriales = True

End Function
 'sastreria
 Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single) As Boolean
    If ObjData(itemIndex).PielLobo > 0 Then
            If Not TieneObjetos(PielLobo, ObjData(itemIndex).PielLobo * Cantidad, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).PielLobo * Cantidad & " pieles de lobo." & FONTTYPE_INFO)
                    SastreTieneMateriales = False
                    Exit Function
            End If
    End If
    
    If ObjData(itemIndex).PielOsoPardo > 0 Then
            If Not TieneObjetos(PielOsoPardo, ObjData(itemIndex).PielOsoPardo * Cantidad, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).PielOsoPardo * Cantidad & " pieles de oso pardo." & FONTTYPE_INFO)
                    SastreTieneMateriales = False
                    Exit Function
            End If
    End If
    
    If ObjData(itemIndex).PielOsoPolar > 0 Then
            If Not TieneObjetos(PielOsoPolar, ObjData(itemIndex).PielOsoPolar * Cantidad, UserIndex) Then
                    Call SendData(ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).PielOsoPolar * Cantidad & " pieles de oso polar." & FONTTYPE_INFO)
                    SastreTieneMateriales = False
                    Exit Function
            End If
    End If
    
    SastreTieneMateriales = True

End Function
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single) As Boolean
    If ObjData(itemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(itemIndex).LingH * Cantidad, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).LingH * Cantidad & " lingotes de hierro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(itemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(itemIndex).LingP * Cantidad, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).LingP * Cantidad & " lingotes de plata." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(itemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(itemIndex).LingO * Cantidad, UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes suficientes materiales. Precisas de " & ObjData(itemIndex).LingO * Cantidad & " lingotes de oro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single) As Boolean
PuedeConstruir = HerreroTieneMateriales(UserIndex, itemIndex, Cantidad) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= _
 ObjData(itemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal itemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = itemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = itemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(UserIndex, itemIndex, Cantidad) And PuedeConstruirHerreria(itemIndex) Then
    Call HerreroQuitarMateriales(UserIndex, itemIndex, Cantidad)
    ' AGREGAR FX
    If ObjData(itemIndex).OBJType = eOBJType.otWeapon Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    ElseIf ObjData(itemIndex).OBJType = eOBJType.otESCUDO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    ElseIf ObjData(itemIndex).OBJType = eOBJType.otCASCO Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    ElseIf ObjData(itemIndex).OBJType = eOBJType.otArmadura Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    ElseIf ObjData(itemIndex).OBJType = eOBJType.otPociones Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    End If
    Dim MiObj As Obj
    MiObj.Amount = Cantidad
    MiObj.ObjIndex = itemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call SubirSkill(UserIndex, Herreria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MARTILLOHERRERO)
    
End If

Call Trabajando(UserIndex)

End Sub


Public Function PuedeConstruirCarpintero(ByVal itemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = itemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)

If CarpinteroTieneMateriales(UserIndex, itemIndex, Cantidad) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(itemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(itemIndex) And _
   UserList(UserIndex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(UserIndex, itemIndex, Cantidad)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    
    Dim MiObj As Obj
    MiObj.Amount = Cantidad
    MiObj.ObjIndex = itemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

Call Trabajando(UserIndex)

End Sub
Public Function PuedeConstruirDruida(ByVal itemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjDruida)
    If ObjDruida(i) = itemIndex Then
        PuedeConstruirDruida = True
        Exit Function
    End If
Next i
PuedeConstruirDruida = False

End Function

Public Sub DruidaConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

If DruidaTieneMateriales(UserIndex, itemIndex, Cantidad) And _
   UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) >= _
   ObjData(itemIndex).SkAlquimia And _
   PuedeConstruirDruida(itemIndex) And _
   UserList(UserIndex).Invent.HerramientaEqpObjIndex = OLLA Then

    Call DruidaQuitarMateriales(UserIndex, itemIndex, Cantidad)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z54")
    
    Dim MiObj As Obj
    MiObj.Amount = Cantidad
    MiObj.ObjIndex = itemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Alquimia)
    Call UpdateUserInv(True, UserIndex, 0)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

Call Trabajando(UserIndex)

End Sub
'sastreria
Public Function PuedeConstruirSastre(ByVal itemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjSastre)
    If ObjSastre(i) = itemIndex Then
        PuedeConstruirSastre = True
        Exit Function
    End If
Next i
PuedeConstruirSastre = False

End Function
Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal itemIndex As Integer, ByVal Cantidad As Single)

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

If SastreTieneMateriales(UserIndex, itemIndex, Cantidad) And _
   UserList(UserIndex).Stats.UserSkills(Sastreria) >= _
   ObjData(itemIndex).SkSastreria And _
   PuedeConstruirSastre(itemIndex) And _
   UserList(UserIndex).Invent.HerramientaEqpObjIndex = HILO_SASTRE Then
    Call SastreQuitarMateriales(UserIndex, itemIndex, Cantidad)
    
   Call SendData(ToIndex, UserIndex, 0, "Z54")
    
    Dim MiObj As Obj
    MiObj.Amount = Cantidad
    MiObj.ObjIndex = itemIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SubirSkill(UserIndex, Sastreria)
    Call UpdateUserInv(True, UserIndex, 0)
  '  Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & LABUROCARPINTERO)
End If

Call Trabajando(UserIndex)

End Sub
Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 15
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20
        Case iMinerales.OroCrudo
            MineralesParaLingote = 25
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal UserIndex As Integer)
'    Call LogTarea("Sub DoLingotes")
Dim Slot As Integer
Dim obji As Integer

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(UserIndex).Invent.Object(Slot).Amount < (MineralesParaLingote(obji) * 10) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes suficientes minerales para hacer un lingote." & FONTTYPE_INFO)
            Exit Sub
    End If
    
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount - (MineralesParaLingote(obji) * 10)
    If UserList(UserIndex).Invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(Slot).Amount = 0
        UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z55")
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = 10
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    


Call Trabajando(UserIndex)

End Sub

Function ModNavegacion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "MINERO"
        ModFundicion = 1
    Case "HERRERO"
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function
Function ModAlquimia(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "DRUIDA"
        ModAlquimia = 1
    Case Else
        ModAlquimia = 3
End Select

End Function
Function ModSastreria(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case Else
        ModSastreria = 1
End Select

End Function
Function ModCarpinteria(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "CARPINTERO"
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "HERRERO"
        ModHerreriA = 1
    Case "MINERO"
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer
    Select Case UCase$(Clase)
        Case "DRUIDA"
            ModDomar = 6
        Case Else
            ModDomar = 10
    End Select
End Function

Function CalcularPoderDomador(ByVal UserIndex As Integer) As Long
    With UserList(UserIndex).Stats
        CalcularPoderDomador = .UserAtributos(eAtributos.Carisma) _
            * (.UserSkills(eSkill.Domar) / ModDomar(UserList(UserIndex).Clase)) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3)
    End With
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La criatura ya tiene amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SubirSkill(UserIndex, Domar)
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(UserIndex) Then
        Dim Index As Integer
        UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
        Index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(Index) = NpcIndex
        UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No has logrado domar la criatura." & FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 5
        End If
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes controlar mas criaturas." & FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)

    Dim ChotsNover As String
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
            UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
            UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
            UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
            UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
            UserList(UserIndex).Counters.Mimetismo = 0
            UserList(UserIndex).flags.Mimetizado = 0
        End If
        
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.Invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).char.Body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).char.Head
        UserList(UserIndex).char.Body = 0
        UserList(UserIndex).char.Head = 0
        ChotsNover = UserList(UserIndex).char.CharIndex & ",1"
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.Invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).char.Body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).char.Head = UserList(UserIndex).flags.OldHead
        ChotsNover = UserList(UserIndex).char.CharIndex & ",0"
        
    End If
    
    'CHOTS | Envia el invi
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).char.Body, UserList(UserIndex).char.Head, UserList(UserIndex).char.Heading, UserList(UserIndex).char.WeaponAnim, UserList(UserIndex).char.ShieldAnim, UserList(UserIndex).char.CascoAnim)
    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, Nover(5) & ChotsNover)
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(Map, X, Y) Then Exit Sub

With posMadera
    .Map = Map
    .X = X
    .Y = Y
End With

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes hacer fogatas estando muerto." & FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount \ 3
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    
    Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(UserIndex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No has podido hacer la fogata." & FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo errhandler

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

Dim Suerte As Integer
Dim res As Integer


If UCase$(UserList(UserIndex).Clase) = "PESCADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    MiObj.Amount = 1
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z88")
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z89")
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Pesca)

Call Trabajando(UserIndex)

Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UCase(UserList(UserIndex).Clase) = "PESCADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡No se puede pescar con Red en zonas seguras!" & FONTTYPE_INFO)
    Exit Sub
End If

iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)


Select Case iSkill
Case 0:         Suerte = 0
Case 1 To 10:   Suerte = 60
Case 11 To 20:  Suerte = 54
Case 21 To 30:  Suerte = 49
Case 31 To 40:  Suerte = 43
Case 41 To 50:  Suerte = 38
Case 51 To 60:  Suerte = 32
Case 61 To 70:  Suerte = 27
Case 71 To 80:  Suerte = 21
Case 81 To 90:  Suerte = 16
Case 91 To 100: Suerte = 11
Case Else:      Suerte = 0
End Select

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 10) As Integer
        
        'CHOTS | Posibilidades de pescar cada pescado
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO1
        PecesPosibles(3) = PESCADO1
        PecesPosibles(4) = PESCADO1
        PecesPosibles(5) = PESCADO2
        PecesPosibles(6) = PESCADO2
        PecesPosibles(7) = PESCADO2
        PecesPosibles(8) = PESCADO3
        PecesPosibles(9) = PESCADO3
        PecesPosibles(10) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(MinTrabajo + UserList(UserIndex).Stats.ELV, MaxTrabajo + UserList(UserIndex).Stats.ELV)
        Else
            MiObj.Amount = 1
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(1, UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z88")
        
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z89")
    End If
    
    Call SubirSkill(UserIndex, Pesca)
End If

Call Trabajando(UserIndex)

Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "Debes quitar el seguro para robar" & FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "No puedes robar a otros miembros de las fuerzas del caos" & FONTTYPE_FIGHT)
    Exit Sub
End If

If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "Los miembros de la Armada Real no tienen permitido robar!" & FONTTYPE_WARNING)
    Exit Sub
End If


If UserList(VictimaIndex).flags.Privilegios = PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
    
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 2) = 1) And (UCase$(UserList(LadrOnIndex).Clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & UserList(VictimaIndex).Name & " no tiene objetos." & FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim n As Integer
                
                If UCase$(UserList(LadrOnIndex).Clase) = "LADRON" Then
                    n = RandomNumber(10 * UserList(LadrOnIndex).Stats.ELV, 100 * UserList(LadrOnIndex).Stats.ELV)
                Else
                    n = RandomNumber(1, 100)
                End If
                If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO

                Call EnviarOro(LadrOnIndex)
                Call EnviarOro(VictimaIndex)
                
                Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & UserList(VictimaIndex).Name & " no tiene oro." & FONTTYPE_INFO)
            End If
        End If
    Else
        Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "¡No has logrado robar nada!" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, ServerPackages.dialogo & "¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, VictimaIndex, 0, ServerPackages.dialogo & "¡" & UserList(LadrOnIndex).Name & " es un criminal!" & FONTTYPE_INFO)
    End If

    If Not Criminal(LadrOnIndex) Then
        Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron
    If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then _
        UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
    Call SubirSkill(LadrOnIndex, Robar)
End If


End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 2) = 1 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                If RandomNumber(1, 5) = 1 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
        If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 5) = 1 Then flag = True
        End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, UserList(LadrOnIndex).Stats.ELV * 10)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
        num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
        Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.ToIndex, LadrOnIndex, 0, ServerPackages.dialogo & "No has logrado robar un objetos." & FONTTYPE_INFO)
End If

End Sub
Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= -1 Then
                    Suerte = 200
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 11 Then
                    Suerte = 190
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 21 Then
                    Suerte = 180
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 31 Then
                    Suerte = 170
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 41 Then
                    Suerte = 160
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 51 Then
                    Suerte = 150
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 61 Then
                    Suerte = 140
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 71 Then
                    Suerte = 130
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 81 Then
                    Suerte = 120
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= 91 Then
                    Suerte = 110
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) = 100 Then
                    Suerte = 100
End If

res = RandomNumber(0, Suerte)

If UCase$(UserList(UserIndex).Clase) = "ASESINO" Then
    If EstaLiberado(UserIndex) And UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" Then
        If res < 40 Then res = 0 'CHOTS | Gemas (Ase EO)
    Else
        If res < 25 Then res = 0
    End If
ElseIf UCase$(UserList(UserIndex).Clase) = "PALADIN" Then
    If EstaLiberado(UserIndex) And UCase$(UserList(UserIndex).Raza) = "HUMANO" Then
        If res < 25 Then res = 0 'CHOTS | Gemas (Pala Humano)
    Else
        If res < 15 Then res = 0
    End If
ElseIf UCase$(UserList(UserIndex).Clase) = "GUERRERO" Then
    If EstaLiberado(UserIndex) And (UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" Or UCase$(UserList(UserIndex).Raza) = "HUMANO") Then
        If res < 25 Then res = 0 'CHOTS | Gemas (Guerre EO, Guerre Humano)
    Else
        If res < 15 Then res = 0
    End If
ElseIf UCase$(UserList(UserIndex).Clase) = "BARDO" Then
    If EstaLiberado(UserIndex) And UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        If res < 25 Then res = 0 'CHOTS | Gemas (Bardo Gnomo)
    Else
        If res < 12 Then res = 0
    End If
ElseIf UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
    If res < 14 Then res = 0
Else
    If res < 10 Then res = 0
End If

If res = 0 Then
    If VictimUserIndex <> 0 Then
        If EstaLiberado(UserIndex) And UCase$(UserList(UserIndex).Raza) = "GNOMO" And UCase$(UserList(UserIndex).Clase) = "GUERRERO" Then
            UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(daño * 2) 'CHOTS | Gemas (Guerre Gnomo)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & Int(daño * 2) & FONTTYPE_APU)
            Call SendData(SendTarget.ToIndex, VictimUserIndex, 0, ServerPackages.dialogo & "Te ha apuñalado " & UserList(UserIndex).Name & " por " & Int(daño * 2) & FONTTYPE_APU)
        Else
            UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(daño * 1.5)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & Int(daño * 1.5) & FONTTYPE_APU)
            Call SendData(SendTarget.ToIndex, VictimUserIndex, 0, ServerPackages.dialogo & "Te ha apuñalado " & UserList(UserIndex).Name & " por " & Int(daño * 1.5) & FONTTYPE_APU)
        End If
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has apuñalado la criatura por " & Int(daño * 2) & FONTTYPE_APU)
        Call SubirSkill(UserIndex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, Int(daño * 2))
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z56")
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
End Sub
Public Sub DoSacarChala(ByVal UserIndex As Integer)
On Error GoTo errhandler

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

Dim Suerte As Integer
Dim res As Integer

If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡No se puede sacar raíces en zonas seguras!" & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Botanica) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UCase$(UserList(UserIndex).Clase) = "DRUIDA" Then
        MiObj.Amount = RandomNumber(MinTrabajo + Fix(UserList(UserIndex).Stats.ELV / 2), MaxTrabajo + Fix(UserList(UserIndex).Stats.ELV / 2))
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.ObjIndex = Chala
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "CHL" & MiObj.Amount)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z57")
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Botanica)

Call Trabajando(UserIndex)

Exit Sub

errhandler:
    Call LogError("Error en DoSacarChala")

End Sub
Public Sub DoTalar(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡No se puede talar en zonas seguras!" & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(UserList(UserIndex).Clase) = "LEÑADOR" Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Talar) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Talar) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UCase$(UserList(UserIndex).Clase) = "LEÑADOR" Then
        MiObj.Amount = RandomNumber(MinTrabajo + UserList(UserIndex).Stats.ELV, MaxTrabajo + UserList(UserIndex).Stats.ELV)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "LEÑ" & MiObj.Amount)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z58")
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Talar)

Call Trabajando(UserIndex)

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
If UserList(UserIndex).flags.enTorneoAuto Then Exit Sub
If EstaEnCastillo(UserIndex) And UserList(UserIndex).Faccion.ArmadaReal = 1 Then Exit Sub

If UserList(UserIndex).flags.Privilegios < PlayerType.SemiDios Then
    UserList(UserIndex).Reputacion.BurguesRep = 0
    UserList(UserIndex).Reputacion.NobleRep = 0
    UserList(UserIndex).Reputacion.PlebeRep = 0
    UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep + vlASALTO
    If UserList(UserIndex).Reputacion.BandidoRep > MAXREP Then _
        UserList(UserIndex).Reputacion.BandidoRep = MAXREP
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
End If

Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, False)

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlASALTO
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP
    
Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, False)

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo errhandler

If UserList(UserIndex).flags.Privilegios > PlayerType.User Then Exit Sub

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UCase$(UserList(UserIndex).Clase) = "MINERO" Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
    If UCase$(UserList(UserIndex).Clase) = "MINERO" Then
        MiObj.Amount = RandomNumber(MinTrabajo + UserList(UserIndex).Stats.ELV, MaxTrabajo + UserList(UserIndex).Stats.ELV)
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MIN" & MiObj.Amount)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z59")
        UserList(UserIndex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Mineria)

Call Trabajando(UserIndex)

Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub



Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(UserIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If

If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z16")
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).char.FX = 0
    UserList(UserIndex).char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If

Suerte = Fix(40 - (UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 3))
res = RandomNumber(1, Suerte)

If res = 1 Then
    Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "RMN" & Cant)
        UserList(UserIndex).flags.UltimoMensaje = 22
    End If
    
    Call EnviarMn(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub



Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= -1 Then
                    Suerte = 35
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 11 Then
                    Suerte = 30
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 21 Then
                    Suerte = 28
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 31 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 41 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 51 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 61 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 71 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) <= 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Wresterling) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has logrado desarmar a tu oponente!" & FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then Call SendData(SendTarget.ToIndex, VictimIndex, 0, ServerPackages.dialogo & "Tu oponente te ha desarmado!" & FONTTYPE_FIGHT)
    End If
End Sub

'CHOTS | Sube el counter de trabajo y checkea Centinela
Public Sub Trabajando(ByVal UserIndex As Integer)
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    If UserList(UserIndex).Counters.Trabajando >= MAX_TRABAJO_CENTINELA Then
        'CHOTS | No lo pongo en cero asi sale en el /trabajando
        UserList(UserIndex).Counters.Trabajando = 1
        Dim CodigoCentinela As String
        CodigoCentinela = Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90)) & Chr(RandomNumber(65, 90))

        Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABCENTI" & CodigoCentinela)
    End If
End Sub
