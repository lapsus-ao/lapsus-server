Attribute VB_Name = "modGemas"
'Modulo Gemas
'Silv AO 2008
'Programado por Juan Andres Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'Se empezo a programar el 03/09/08 a las 22:15 hs
'Se Termino de programar el 03/09/08 a las 22:26 hs
'REPROGRAMADO Y ADAPTADO POR CHOTS PARA LAPSUS AO 2009 EL 14/02/09 A LAS 11:27 hs


Public Sub LiberarHabilidad(ByVal UserIndex As Integer)

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La gema se ha desintegrado y sientes un nuevo poder fluyendo dentro de ti" & FONTTYPE_GEMA)
'CHOTS | a esta Linea se la debo a Bruno (Andareal)

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " ha liberado su Habilidad con la Gema Sagrada!." & FONTTYPE_GEMA)
                
UserList(UserIndex).flags.Liberado = 1 'CHOTS | Libera la Habilidad

Call LogGM("GEMAS", Date & " - " & Time & " - " & UserList(UserIndex).Name & " ha Liberado su Habilidad con la Gema", False)

Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & ",35,0") 'CHOTS | Enviamos el Tornadito

Select Case UCase$(UserList(UserIndex).Raza)

    Case "HUMANO"
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO" 'CHOTS | Gemas (Mago Humano)
                UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 20
            Case "BARDO" 'CHOTS | Gemas (Bardo Humano)
                UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 20
            Case "ASESINO" 'CHOTS | Gemas (Ase Humano)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 50
        End Select
        
    Case "ELFO"
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO", "BARDO" 'CHOTS | Gemas (Mago Elfo, Bardo Elfo)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 100
        End Select
        
    Case "ENANO"
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "PALADIN" 'CHOTS | Gemas (Pala Enano)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 70
            Case "ASESINO" 'CHOTS | Gemas (Ase Enano)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 70
            Case "GUERRERO" 'CHOTS | Gemas (Guerre Enano)
                UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 20
            Case "CAZADOR" 'CHOTS | Gemas (Caza Enano)
                UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 20
            Case "CLERIGO" 'CHOTS | Gemas (Clero Enano)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 130
            Case "DRUIDA" 'CHOTS | Gemas (Druida Enano)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 200

        End Select
        
    Case "ELFO OSCURO"
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "CLERIGO" 'CHOTS | Gemas (Clero EO)
                UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 20
        End Select
        
    Case "GNOMO"
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "CAZADOR" 'CHOTS | Gemas (Caza Gnomo)
                UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 20
        End Select
        
    Case "ORCO"
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "BARDO" 'CHOTS | Gemas (Bardo Orco)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 100
            Case "PALADIN" 'CHOTS | Gemas (Pala Orco)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 50
            Case "ASESINO" 'CHOTS | Gemas (Ase Orco)
                UserList(UserIndex).Stats.MaxMAN = UserList(UserIndex).Stats.MaxMAN + 70
            End Select
        
End Select

Call EnviarMn(UserIndex)
Call EnviarHP(UserIndex)

End Sub

Public Function EstaLiberado(ByVal UserIndex As Integer) As Boolean
    EstaLiberado = UserList(UserIndex).flags.Liberado = 1
End Function
