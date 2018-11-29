Attribute VB_Name = "ModFacciones"
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

'MODULO OBTENIDO DE LAND OF DRAGONS AO
'REPROGRAMADO Y ADAPTADO POR CHOTS PARA LAPSUS AO 2009
'DOMINGO 8 DE NOVIEMBRE DE 2009



'CHOTS | Faccion Armada

Public ATunicaMBDAlta1ra As Integer
Public ATunicaMBDAlta2da As Integer
Public ATunicaMBDAlta3ra As Integer

Public AArmaduraACAlta1ra As Integer
Public AArmaduraACAlta2da As Integer
Public AArmaduraACAlta3ra As Integer

Public AArmaduraPGKAlta1ra As Integer
Public AArmaduraPGKAlta2da As Integer
Public AArmaduraPGKAlta3ra As Integer

Public ATunicaMBDBaja1ra As Integer
Public ATunicaMBDBaja2da As Integer
Public ATunicaMBDBaja3ra As Integer

Public AArmaduraACBaja1ra As Integer
Public AArmaduraACBaja2da As Integer
Public AArmaduraACBaja3ra As Integer

Public AArmaduraPGKBaja1ra As Integer
Public AArmaduraPGKBaja2da As Integer
Public AArmaduraPGKBaja3ra As Integer

'CHOTS | Faccion Armada



'CHOTS | Faccion Caos

Public CTunicaMBDAlta1ra As Integer
Public CTunicaMBDAlta2da As Integer
Public CTunicaMBDAlta3ra As Integer

Public CArmaduraACAlta1ra As Integer
Public CArmaduraACAlta2da As Integer
Public CArmaduraACAlta3ra As Integer

Public CArmaduraPGKAlta1ra As Integer
Public CArmaduraPGKAlta2da As Integer
Public CArmaduraPGKAlta3ra As Integer

Public CTunicaMBDBaja1ra As Integer
Public CTunicaMBDBaja2da As Integer
Public CTunicaMBDBaja3ra As Integer

Public CArmaduraACBaja1ra As Integer
Public CArmaduraACBaja2da As Integer
Public CArmaduraACBaja3ra As Integer

Public CArmaduraPGKBaja1ra As Integer
Public CArmaduraPGKBaja2da As Integer
Public CArmaduraPGKBaja3ra As Integer

'CHOTS | Faccion Caos

Public Const ExpAlUnirse As Long = 500000
Public Const ExpX100 As Long = 50000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No se permiten criminales en el ejercito imperial!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 50 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 50 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FueCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Has sido expulsado de las fuerzas reales demasiadas veces!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Jerarquia = 1
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
UserList(UserIndex).Faccion.Amatar = 200
UserList(UserIndex).Faccion.FueReal = 1
UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbCyan & "°" & "¡¡¡Bienvenido a al Ejército Imperial!!!, aquí tienes tus vestimentas. Por cada centena de criminales que acabes te daré un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmadura = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
            
    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = ATunicaMBDBaja1ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = AArmaduraACBaja1ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = AArmaduraPGKBaja1ra
        End Select
    
    
    ElseIf UCase$(UserList(UserIndex).Raza) = "HUMANO" Or UCase$(UserList(UserIndex).Raza) = "ELFO" Or _
           UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" Or UCase$(UserList(UserIndex).Raza) = "ORCO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = ATunicaMBDAlta1ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = AArmaduraACAlta1ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = AArmaduraPGKAlta1ra
        End Select
              
    End If
    
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmadura = 1
    
    
End If

If UserList(UserIndex).Faccion.RecibioExpInicial = 0 Then
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpAlUnirse
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_GUILD)
    UserList(UserIndex).Faccion.RecibioExpInicial = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoReal(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CriminalesMatados < UserList(UserIndex).Faccion.Amatar Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbCyan & "°" & "Ya has recibido tu recompensa.. mata " & UserList(UserIndex).Faccion.Amatar & " criminales y obtendrás la próxima!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbBlue & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_GUILD)
    UserList(UserIndex).Faccion.Jerarquia = UserList(UserIndex).Faccion.Jerarquia + 1
    Call CheckUserLevel(UserIndex)
End If

If UserList(UserIndex).Faccion.Jerarquia = 5 And UserList(UserIndex).Faccion.RecibioArmadura < 3 Then
    Dim MiObj As Obj
    MiObj.Amount = 1

    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = ATunicaMBDBaja3ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = AArmaduraACBaja3ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = AArmaduraPGKBaja3ra
        End Select
    
    
    ElseIf UCase$(UserList(UserIndex).Raza) = "HUMANO" Or UCase$(UserList(UserIndex).Raza) = "ELFO" Or _
           UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" Or UCase$(UserList(UserIndex).Raza) = "ORCO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = ATunicaMBDAlta3ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = AArmaduraACAlta3ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = AArmaduraPGKAlta3ra
        End Select
              
    End If
    
    UserList(UserIndex).Faccion.RecibioArmadura = 3
    
End If

If UserList(UserIndex).Faccion.Jerarquia = 3 And UserList(UserIndex).Faccion.RecibioArmadura < 2 Then
    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = ATunicaMBDBaja2da
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = AArmaduraACBaja2da
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = AArmaduraPGKBaja2da
        End Select
    
    
    ElseIf UCase$(UserList(UserIndex).Raza) = "HUMANO" Or UCase$(UserList(UserIndex).Raza) = "ELFO" Or _
           UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" Or UCase$(UserList(UserIndex).Raza) = "ORCO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = ATunicaMBDAlta2da
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = AArmaduraACAlta2da
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = AArmaduraPGKAlta2da
        End Select
              
    End If
    
    UserList(UserIndex).Faccion.RecibioArmadura = 2
    
End If


Select Case UserList(UserIndex).Faccion.Jerarquia
    Case 2
        UserList(UserIndex).Faccion.Amatar = 300
    Case 3
        UserList(UserIndex).Faccion.Amatar = 400
    Case 4
        UserList(UserIndex).Faccion.Amatar = 500
    Case Else
        UserList(UserIndex).Faccion.Amatar = UserList(UserIndex).Faccion.Amatar * 2
End Select



MiObj.Amount = 1

If MiObj.ObjIndex <> 0 Then
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.ArmadaReal = 0
    UserList(UserIndex).Faccion.Jerarquia = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.FuerzasCaos = 0
    UserList(UserIndex).Faccion.Jerarquia = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has sido expulsado de la legión oscura!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.Jerarquia
    Case 1
        TituloReal = "Aprendiz"
    Case 2
        TituloReal = "Escudero"
    Case 3
        TituloReal = "Caballero"
    Case 4
        TituloReal = "Teniente"
    Case Else
        TituloReal = "Campeón de la Luz"
End Select

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)

If Not Criminal(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Ya perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(UserIndex).Faccion.FueReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No permitiré que ningún insecto real ingrese ¡Traidor del Rey!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If
'[/Barrin]

If UserList(UserIndex).Faccion.CiudadanosMatados < 100 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 100 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If


If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Has sido expulsado de las fuerzas oscuras demasiadas veces!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.Jerarquia = 1
UserList(UserIndex).Faccion.Amatar = 200
UserList(UserIndex).Faccion.FueCaos = 1

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Si asesinas más te daré un recompensa, buena suerte soldado!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmadura = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
            
            
    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = CTunicaMBDBaja1ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = CArmaduraACBaja1ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = CArmaduraPGKBaja1ra
        End Select
    
    
    ElseIf UCase$(UserList(UserIndex).Raza) = "HUMANO" Or UCase$(UserList(UserIndex).Raza) = "ELFO" Or _
           UCase$(UserList(UserIndex).Raza) = "ELFO OSCURO" Or UCase$(UserList(UserIndex).Raza) = "ORCO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = CTunicaMBDAlta1ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = CArmaduraACAlta1ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = CArmaduraPGKAlta1ra
        End Select
              
    End If


    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmadura = 1
    
End If

If UserList(UserIndex).Faccion.RecibioExpInicial = 0 Then
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpAlUnirse
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_GUILD)
    UserList(UserIndex).Faccion.RecibioExpInicial = 1
    Call CheckUserLevel(UserIndex)
End If


Call LogEjercitoCaos(UserList(UserIndex).Name)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CiudadanosMatados < UserList(UserIndex).Faccion.Amatar Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Ya has recibido tu recompensa, mata " & UserList(UserIndex).Faccion.Amatar & " ciudadanos para recibir la proxima!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    Exit Sub
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_GUILD)
    UserList(UserIndex).Faccion.Jerarquia = UserList(UserIndex).Faccion.Jerarquia + 1
    Call CheckUserLevel(UserIndex)

If UserList(UserIndex).Faccion.Jerarquia = 5 And UserList(UserIndex).Faccion.RecibioArmadura < 3 Then
    Dim MiObj As Obj
    MiObj.Amount = 1

    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = CTunicaMBDBaja3ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = CArmaduraACBaja3ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = CArmaduraPGKBaja3ra
        End Select
    
    
    Else
    
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = CTunicaMBDAlta3ra
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = CArmaduraACAlta3ra
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = CArmaduraPGKAlta3ra
        End Select
              
    End If
    
    UserList(UserIndex).Faccion.RecibioArmadura = 3
    
End If

If UserList(UserIndex).Faccion.Jerarquia = 3 And UserList(UserIndex).Faccion.RecibioArmadura < 2 Then
    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = CTunicaMBDBaja2da
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = CArmaduraACBaja2da
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = CArmaduraPGKBaja2da
        End Select
    
    
    Else
    
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "DRUIDA", "BARDO"
                MiObj.ObjIndex = CTunicaMBDAlta2da
            Case "CLERIGO", "ASESINO"
                MiObj.ObjIndex = CArmaduraACAlta2da
            Case "PALADIN", "GUERRERO", "CAZADOR"
                MiObj.ObjIndex = CArmaduraPGKAlta2da
        End Select
              
    End If
    
    UserList(UserIndex).Faccion.RecibioArmadura = 2
    
End If


Select Case UserList(UserIndex).Faccion.Jerarquia
    Case 2
        UserList(UserIndex).Faccion.Amatar = 300
    Case 3
        UserList(UserIndex).Faccion.Amatar = 400
    Case 4
        UserList(UserIndex).Faccion.Amatar = 500
    Case Else
        UserList(UserIndex).Faccion.Amatar = UserList(UserIndex).Faccion.Amatar * 2
End Select


MiObj.Amount = 1

If MiObj.ObjIndex <> 0 Then
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End If

End If


End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.Jerarquia
    Case 1
        TituloCaos = "Esbirro"
    Case 2
        TituloCaos = "Servidor de las Sombras"
    Case 3
        TituloCaos = "Acólito"
    Case 4
        TituloCaos = "Guerrero Sombrío"
    Case Else
        TituloCaos = "Devorador de Almas"
End Select


End Function
