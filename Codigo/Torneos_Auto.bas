Attribute VB_Name = "Torneos_Auto"
'Módulo de Torneos Automáticos
'Creado por Juan Andrés Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'Para Lapsus 3
'05/10/2011

Public Torneo_Activado As Boolean
Public Torneo_Cupo As Byte
Public Torneo_CantidadInscriptos As Byte
Public Torneo_HAYTORNEO As Boolean
Public Torneo_Fixture As String
Public Const Torneo_MAPATORNEO As Byte = 78
Public Const Torneo_MAPAMUERTE As Byte = 62

Public Type tAreasTorneo
    mapa As Byte
    minX As Byte
    maxX As Byte
    minY As Byte
    maxY As Byte
End Type

Public Type tCuenta
    segundos As Byte
    razon As String
    next As eRonda
End Type

Public Type tDuelo
    usuario1 As Integer
    usuario2 As Integer
    ganador As Integer
End Type

Public Enum eRonda
    Ronda_Dieciseisavos = 1
    Ronda_Octavos = 2
    Ronda_Cuartos = 3
    Ronda_Semi = 4
    Ronda_Final = 5
End Enum

Public Torneo_AreaDescanso As tAreasTorneo
Public Torneo_AreasDuelo() As tAreasTorneo

Public Torneo_CR As tCuenta
Public Torneo_CuentaPelea As Byte

Public Torneo_RondaActual As eRonda

Public Torneo_UsuariosInscriptos() As Integer
Public Torneo_Final As tDuelo
Public Torneo_Semifinal() As tDuelo
Public Torneo_Cuartos() As tDuelo
Public Torneo_Octavos() As tDuelo
Public Torneo_Dieciseisavos() As tDuelo
'

Public Sub activarTorneos()
    Torneo_Activado = True
    Call SendData(SendTarget.ToAll, 0, 0, "Z92")
End Sub

Public Sub desactivarTorneos()
    Torneo_Activado = False
    Torneo_HAYTORNEO = False
    Call SendData(SendTarget.ToAll, 0, 0, "Z91")
End Sub

Public Sub crearTorneo()
    Dim i As Byte
    If Not Torneo_Activado Then Exit Sub

    Torneo_Cupo = calcularCantidad()
    Torneo_CuentaPelea = 0
    
    If Torneo_Cupo = 0 Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "No hay la suficiente cantidad de usuarios para iniciar un torneo!" & FONTTYPE_TORNEOAUTO)
        Exit Sub
    End If
    
    ReDim Torneo_UsuariosInscriptos(1 To Torneo_Cupo) As Integer
    
    For i = 1 To Torneo_Cupo
        Torneo_UsuariosInscriptos(i) = 0
    Next i
    
    Call inicializarAreas
    
    Torneo_CantidadInscriptos = 0
    
    Torneo_HAYTORNEO = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "TAU" & Torneo_Cupo)

    Call LogGM("TORNEOAUTO", "Se abrio un torneo auto para " & Torneo_Cupo & " participantes.", False)
    
End Sub
Private Sub inicializarAreas()
    ReDim Torneo_AreasDuelo(1 To 4) As tAreasTorneo
    
    Torneo_AreasDuelo(1).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(1).minX = 13
    Torneo_AreasDuelo(1).maxX = 26
    Torneo_AreasDuelo(1).minY = 11
    Torneo_AreasDuelo(1).maxY = 20
    
    Torneo_AreasDuelo(2).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(2).minX = 75
    Torneo_AreasDuelo(2).maxX = 88
    Torneo_AreasDuelo(2).minY = 11
    Torneo_AreasDuelo(2).maxY = 20
    
    Torneo_AreasDuelo(3).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(3).minX = 13
    Torneo_AreasDuelo(3).maxX = 26
    Torneo_AreasDuelo(3).minY = 81
    Torneo_AreasDuelo(3).maxY = 90
    
    Torneo_AreasDuelo(4).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(4).minX = 75
    Torneo_AreasDuelo(4).maxX = 88
    Torneo_AreasDuelo(4).minY = 81
    Torneo_AreasDuelo(4).maxY = 90
    
    Torneo_AreaDescanso.mapa = Torneo_MAPATORNEO
    Torneo_AreaDescanso.minX = 41
    Torneo_AreaDescanso.maxX = 59
    Torneo_AreaDescanso.minY = 43
    Torneo_AreaDescanso.maxY = 59
    
End Sub

Private Function calcularCantidad() As Byte
    Dim n As Integer
    Dim LoopC As Integer
    calcularCantidad = 0

    If Not NumUsers >= 8 Then Exit Function
    
    If NumUsers < 30 Then
        calcularCantidad = 8
    ElseIf NumUsers < 100 Then
        calcularCantidad = 16
    Else
        calcularCantidad = 32
    End If
    
End Function

Public Function inscribirseTorneo(ByVal UserIndex As Integer) As Boolean
    Dim i As Integer
    
    If Torneo_CantidadInscriptos >= Torneo_Cupo Then
        inscribirseTorneo = False
        Exit Function
    End If
    
    For i = 1 To Torneo_Cupo
        If Torneo_UsuariosInscriptos(i) = 0 Then
            Torneo_UsuariosInscriptos(i) = UserIndex
            UserList(UserIndex).flags.enTorneoAuto = True
            Torneo_CantidadInscriptos = Torneo_CantidadInscriptos + 1
            inscribirseTorneo = True
            Call telepToAreaDescanso(UserIndex)
            Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " se inscribio a un torneo Auto.", False)
            If Torneo_CantidadInscriptos = Torneo_Cupo Then armarFixture
            Exit Function
        End If
    Next i
End Function

Public Sub telepToAreaDescanso(ByVal UserIndex As Integer)
    On Error GoTo chotserror
    Dim dPos2 As WorldPos
    Dim nPos2 As WorldPos
    dPos2.Map = Torneo_AreaDescanso.mapa
    dPos2.X = RandomNumber(Torneo_AreaDescanso.minX, Torneo_AreaDescanso.maxX)
    dPos2.Y = RandomNumber(Torneo_AreaDescanso.minY, Torneo_AreaDescanso.maxY)
    Call ClosestLegalPos(dPos2, nPos2)
    Call WarpUserChar(UserIndex, nPos2.Map, nPos2.X, nPos2.Y, True)
    UserList(UserIndex).flags.enDueloTorneoAuto = False
    Exit Sub
    
chotserror:
    Call LogError("Error en TelepAreaDescanso " & Err.number & " " & Err.Description)
End Sub

Public Sub telepToAreaDuelo(ByVal UserIndex As Integer, ByVal area As Byte)
    On Error GoTo chotserror
    Dim dPos As WorldPos
    Dim nPos As WorldPos
    dPos.Map = Torneo_AreasDuelo(area).mapa
    dPos.X = RandomNumber(Torneo_AreasDuelo(area).minX, Torneo_AreasDuelo(area).maxX)
    dPos.Y = RandomNumber(Torneo_AreasDuelo(area).minY, Torneo_AreasDuelo(area).maxY)
    Call ClosestLegalPos(dPos, nPos)
    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
    UserList(UserIndex).flags.enDueloTorneoAuto = True
    Exit Sub
    
chotserror:
    Call LogError("Error en TelepAreaDuelo " & Err.number & " " & Err.Description)
End Sub

Private Sub armarFixture()
    Call SendData(SendTarget.ToAll, 0, 0, "Z94")
    Call SendData(SendTarget.ToMap, 0, Torneo_MAPATORNEO, "Z95")
    
    Dim i As Integer
    
    If Torneo_CantidadInscriptos = 32 Then
        ReDim Torneo_Dieciseisavos(1 To 16) As tDuelo
        For i = 1 To 16
            Torneo_Dieciseisavos(i).usuario1 = 0
            Torneo_Dieciseisavos(i).usuario2 = 0
            Torneo_Dieciseisavos(i).ganador = 0
        Next i
    End If
    
    If Torneo_CantidadInscriptos >= 16 Then
        ReDim Torneo_Octavos(1 To 8) As tDuelo
        For i = 1 To 8
            Torneo_Octavos(i).usuario1 = 0
            Torneo_Octavos(i).usuario2 = 0
            Torneo_Octavos(i).ganador = 0
        Next i
    End If
    
    ReDim Torneo_Cuartos(1 To 4) As tDuelo
    For i = 1 To 4
        Torneo_Cuartos(i).usuario1 = 0
        Torneo_Cuartos(i).usuario2 = 0
        Torneo_Cuartos(i).ganador = 0
    Next i
        
    ReDim Torneo_Semifinal(1 To 2) As tDuelo
    For i = 1 To 2
        Torneo_Semifinal(i).usuario1 = 0
        Torneo_Semifinal(i).usuario2 = 0
        Torneo_Semifinal(i).ganador = 0
    Next i

    Torneo_Final.ganador = 0
    Torneo_Final.usuario1 = 0
    Torneo_Final.usuario2 = 0

    Dim contador As Byte
    contador = 1

    Torneo_Fixture = Torneo_CantidadInscriptos & "@"

    For i = 1 To Torneo_CantidadInscriptos

        Select Case Torneo_CantidadInscriptos

            Case 32:
                Torneo_Dieciseisavos(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                Torneo_Dieciseisavos(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                Torneo_Fixture = Torneo_Fixture & Torneo_Dieciseisavos(contador).usuario1 & "," & Torneo_Dieciseisavos(contador).usuario2 & ","
                contador = contador + 1

            Case 16:
                Torneo_Octavos(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                Torneo_Octavos(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                Torneo_Fixture = Torneo_Fixture & Torneo_Octavos(contador).usuario1 & "," & Torneo_Octavos(contador).usuario2 & ","
                contador = contador + 1

            Case 8:
                Torneo_Cuartos(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                Torneo_Cuartos(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                Torneo_Fixture = Torneo_Fixture & Torneo_Cuartos(contador).usuario1 & "," & Torneo_Cuartos(contador).usuario2 & ","
                contador = contador + 1

            Case 4:
                Torneo_Semifinal(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                Torneo_Semifinal(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                Torneo_Fixture = Torneo_Fixture & Torneo_Semifinal(contador).usuario1 & "," & Torneo_Semifinal(contador).usuario2 & ","
                contador = contador + 1
                
        End Select
        
        i = i + 1
        
    Next i
    
    Call comenzarTorneo
            
End Sub

Public Sub comenzarTorneo()

    MapInfo(Torneo_MAPATORNEO).Pk = False

    Select Case Torneo_CantidadInscriptos
        
        Case 32:
            Call setearCuenta(10, eRonda.Ronda_Dieciseisavos)
            Exit Sub
            
        Case 16:
            Call setearCuenta(10, eRonda.Ronda_Octavos)
            Exit Sub
            
        Case 8:
            Call setearCuenta(10, eRonda.Ronda_Cuartos)
            Exit Sub
            
        Case 4:
            Call setearCuenta(10, eRonda.Ronda_Semi)
            Exit Sub
              
    End Select
    
End Sub

Public Sub comenzarDieciseisavos()
    Torneo_RondaActual = eRonda.Ronda_Dieciseisavos
    Dim i As Byte
    Dim j As Byte
    j = 0

    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 16
        If Torneo_Dieciseisavos(i).ganador = 0 Then
        
            j = j + 1
            
            If Torneo_Dieciseisavos(i).usuario1 = 0 And Torneo_Dieciseisavos(i).usuario2 > 0 Then
                Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
            ElseIf Torneo_Dieciseisavos(i).usuario2 = 0 And Torneo_Dieciseisavos(i).usuario1 > 0 Then
                Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
            ElseIf Torneo_Dieciseisavos(i).usuario1 > 0 And Torneo_Dieciseisavos(i).usuario2 > 0 Then
                If UserList(Torneo_Dieciseisavos(i).usuario1).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
                ElseIf UserList(Torneo_Dieciseisavos(i).usuario2).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
                Else
                    Call telepToAreaDuelo(Torneo_Dieciseisavos(i).usuario1, j)
                    UserList(Torneo_Dieciseisavos(i).usuario1).Counters.Torneo = 5
                    Call telepToAreaDuelo(Torneo_Dieciseisavos(i).usuario2, j)
                End If
            End If
        End If
        
        If j = 4 Then Exit For
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarOctavos()
    Torneo_RondaActual = eRonda.Ronda_Octavos
    Dim i As Byte
    Dim j As Byte
    j = 0

    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 8
        If Torneo_Octavos(i).ganador = 0 Then
        
            j = j + 1
            
            If Torneo_Octavos(i).usuario1 = 0 And Torneo_Octavos(i).usuario2 > 0 Then
                Call ganaUsuario(Torneo_Octavos(i).usuario2)
            ElseIf Torneo_Octavos(i).usuario2 = 0 And Torneo_Octavos(i).usuario1 > 0 Then
                Call ganaUsuario(Torneo_Octavos(i).usuario1)
            ElseIf Torneo_Octavos(i).usuario1 > 0 And Torneo_Octavos(i).usuario2 > 0 Then
                If UserList(Torneo_Octavos(i).usuario1).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Octavos(i).usuario2)
                ElseIf UserList(Torneo_Octavos(i).usuario1).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Octavos(i).usuario1)
                Else
                    Call telepToAreaDuelo(Torneo_Octavos(i).usuario1, j)
                    UserList(Torneo_Octavos(i).usuario1).Counters.Torneo = 6
                    Call telepToAreaDuelo(Torneo_Octavos(i).usuario2, j)
                End If
            End If
        End If

        If j = 4 Then Exit For
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarCuartos()
    Torneo_RondaActual = eRonda.Ronda_Cuartos
    Dim i As Byte
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 4
        If Torneo_Cuartos(i).usuario1 = 0 And Torneo_Cuartos(i).usuario2 > 0 Then
            Call ganaUsuario(Torneo_Cuartos(i).usuario2)
        ElseIf Torneo_Cuartos(i).usuario2 = 0 And Torneo_Cuartos(i).usuario1 > 0 Then
            Call ganaUsuario(Torneo_Cuartos(i).usuario1)
        ElseIf Torneo_Cuartos(i).usuario1 > 0 And Torneo_Cuartos(i).usuario2 > 0 Then
            If UserList(Torneo_Cuartos(i).usuario1).flags.enTorneoAuto = False Then
                Call ganaUsuario(Torneo_Cuartos(i).usuario2)
            ElseIf UserList(Torneo_Cuartos(i).usuario2).flags.enTorneoAuto = False Then
                Call ganaUsuario(Torneo_Cuartos(i).usuario1)
            Else
                Call telepToAreaDuelo(Torneo_Cuartos(i).usuario1, i)
                UserList(Torneo_Cuartos(i).usuario1).Counters.Torneo = 7
                Call telepToAreaDuelo(Torneo_Cuartos(i).usuario2, i)
            End If
        End If
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub
Public Sub comenzarSemifinal()
    Torneo_RondaActual = eRonda.Ronda_Semi
    Dim i As Byte
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 2
        If Torneo_Semifinal(i).usuario1 = 0 And Torneo_Semifinal(i).usuario2 > 0 Then
            Call ganaUsuario(Torneo_Semifinal(i).usuario2)
        ElseIf Torneo_Semifinal(i).usuario2 = 0 And Torneo_Semifinal(i).usuario1 > 0 Then
            Call ganaUsuario(Torneo_Semifinal(i).usuario1)
        ElseIf Torneo_Semifinal(i).usuario1 > 0 And Torneo_Semifinal(i).usuario2 > 0 Then
            If UserList(Torneo_Semifinal(i).usuario1).flags.enTorneoAuto = False Then
                 Call ganaUsuario(Torneo_Semifinal(i).usuario2)
            ElseIf UserList(Torneo_Semifinal(i).usuario2).flags.enTorneoAuto = False Then
                 Call ganaUsuario(Torneo_Semifinal(i).usuario1)
            Else
                Call telepToAreaDuelo(Torneo_Semifinal(i).usuario1, i)
                UserList(Torneo_Semifinal(i).usuario1).Counters.Torneo = 8
                Call telepToAreaDuelo(Torneo_Semifinal(i).usuario2, i)
            End If
        End If
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub
Public Sub comenzarFinal()
    Torneo_RondaActual = eRonda.Ronda_Final
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
        
    If Torneo_Final.usuario1 = 0 And Torneo_Final.usuario2 > 0 Then
        Call ganaUsuario(Torneo_Final.usuario2)
    ElseIf Torneo_Final.usuario2 = 0 And Torneo_Final.usuario1 > 0 Then
        Call ganaUsuario(Torneo_Final.usuario1)
    ElseIf Torneo_Final.usuario1 > 0 And Torneo_Final.usuario2 > 0 Then
        If UserList(Torneo_Final.usuario1).flags.enTorneoAuto = False Then
            Call ganaUsuario(Torneo_Final.usuario2)
        ElseIf UserList(Torneo_Final.usuario2).flags.enTorneoAuto = False Then
            Call ganaUsuario(Torneo_Final.usuario1)
        Else
            Call telepToAreaDuelo(Torneo_Final.usuario1, 1)
            Call telepToAreaDuelo(Torneo_Final.usuario2, 1)
        End If
    Else
        Call cerrarTorneo
    End If
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub ganaUsuario(ByVal UserIndex As Integer)
    Dim i As Byte
    
    UserList(UserIndex).Counters.Torneo = 0
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
        
            For i = 1 To 8
            
                If Torneo_Dieciseisavos(i).usuario1 = UserIndex Then
                    Torneo_Dieciseisavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Octavos)
                    Call vuelveUlla(Torneo_Dieciseisavos(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Dieciseisavos(i).usuario2 = UserIndex Then
                    Torneo_Dieciseisavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Octavos)
                    Call vuelveUlla(Torneo_Dieciseisavos(i).usuario1)
                    Exit Sub
                End If
                
            Next i
    
    
        Case eRonda.Ronda_Octavos:
        
            For i = 1 To 8
            
                If Torneo_Octavos(i).usuario1 = UserIndex Then
                    Torneo_Octavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Cuartos)
                    Call vuelveUlla(Torneo_Octavos(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Octavos(i).usuario2 = UserIndex Then
                    Torneo_Octavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Cuartos)
                    Call vuelveUlla(Torneo_Octavos(i).usuario1)
                    Exit Sub
                End If
                
            Next i
    
        Case eRonda.Ronda_Cuartos:
        
            For i = 1 To 4
            
                If Torneo_Cuartos(i).usuario1 = UserIndex Then
                    Torneo_Cuartos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Semi)
                    Call vuelveUlla(Torneo_Cuartos(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Cuartos(i).usuario2 = UserIndex Then
                    Torneo_Cuartos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Semi)
                    Call vuelveUlla(Torneo_Cuartos(i).usuario1)
                    Exit Sub
                End If
                
            Next i
            
            
        Case eRonda.Ronda_Semi
        
            For i = 1 To 2
            
                If Torneo_Semifinal(i).usuario1 = UserIndex Then
                    Torneo_Semifinal(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Final)
                    Call vuelveUlla(Torneo_Semifinal(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Semifinal(i).usuario2 = UserIndex Then
                    Torneo_Semifinal(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Final)
                    Call vuelveUlla(Torneo_Semifinal(i).usuario1)
                    Exit Sub
                End If
                
            Next i
            
            
        Case eRonda.Ronda_Final
            
            If Torneo_Final.usuario1 = UserIndex Then
                Call vuelveUlla(Torneo_Final.usuario2)
                Torneo_Final.ganador = UserIndex
                Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto>" & UserList(UserIndex).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                Call ganaTorneo(UserIndex)
            End If
            
            If Torneo_Final.usuario2 = UserIndex Then
                Call vuelveUlla(Torneo_Final.usuario1)
                Torneo_Final.ganador = UserIndex
                Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto>" & UserList(UserIndex).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                Call ganaTorneo(UserIndex)
            End If
            
        End Select
        
        Torneo_Fixture = Torneo_Fixture & UserList(UserIndex).Name & ","
        
End Sub

Public Sub pasaDeRonda(ByVal UserIndex As Integer, ByVal ronda As eRonda)
    Dim i As Byte
    
    Select Case ronda

        Case eRonda.Ronda_Octavos:
            
            For i = 1 To 8
            
                If Torneo_Octavos(i).usuario1 = 0 Then
                    Torneo_Octavos(i).usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Octavos de final.", False)
                    Call telepToAreaDescanso(UserIndex)
                    Exit Sub
                End If
                
                If Torneo_Octavos(i).usuario2 = 0 Then
                    Torneo_Octavos(i).usuario2 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será: " & UserList(Torneo_Octavos(i).usuario1).Name & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Octavos de final. Juega vs " & UserList(Torneo_Octavos(i).usuario1).Name, False)
                    Call telepToAreaDescanso(UserIndex)
                    If i = 8 Then
                        Call setearCuenta(10, eRonda.Ronda_Octavos)
                    ElseIf (i = 6 Or i = 4 Or i = 2) Then
                        Call comenzarDieciseisavos
                    End If
                    Exit Sub
                End If
            
            Next i
            
    
        Case eRonda.Ronda_Cuartos:
            
            For i = 1 To 4
            
                If Torneo_Cuartos(i).usuario1 = 0 Then
                    Torneo_Cuartos(i).usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Cuartos de final.", False)
                    Call telepToAreaDescanso(UserIndex)
                    Exit Sub
                End If
                
                If Torneo_Cuartos(i).usuario2 = 0 Then
                    Torneo_Cuartos(i).usuario2 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será: " & UserList(Torneo_Cuartos(i).usuario1).Name & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Cuartos de final. Juega vs " & UserList(Torneo_Cuartos(i).usuario1).Name, False)
                    Call telepToAreaDescanso(UserIndex)
                    If i = 4 Then
                        Call setearCuenta(10, eRonda.Ronda_Cuartos)
                    ElseIf i = 2 Then
                        Call comenzarOctavos
                    End If
                    Exit Sub
                End If
            
            Next i
    
        Case eRonda.Ronda_Semi:
            
            For i = 1 To 2
            
                If Torneo_Semifinal(i).usuario1 = 0 Then
                    Torneo_Semifinal(i).usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Semifinal.", False)
                    Call telepToAreaDescanso(UserIndex)
                    Exit Sub
                End If
                
                If Torneo_Semifinal(i).usuario2 = 0 Then
                    Torneo_Semifinal(i).usuario2 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será: " & UserList(Torneo_Semifinal(i).usuario1).Name & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Semifinal. Juega vs " & UserList(Torneo_Semifinal(i).usuario1).Name, False)
                    Call telepToAreaDescanso(UserIndex)
                    If i = 2 Then Call setearCuenta(10, eRonda.Ronda_Semi)
                    Exit Sub
                End If
            
            Next i
            
        Case eRonda.Ronda_Final:
            
                If Torneo_Final.usuario1 = 0 Then
                    Torneo_Final.usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado a la final del torneo!" & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a la Final.", False)
                    Call telepToAreaDescanso(UserIndex)
                    Exit Sub
                End If
                
                If Torneo_Final.usuario2 = 0 Then
                    Torneo_Final.usuario2 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado a la final del torneo! Tu rival será: " & UserList(Torneo_Final.usuario1).Name & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Final.usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a la Final. Juega vs " & UserList(Torneo_Final.usuario1).Name, False)
                    Call telepToAreaDescanso(UserIndex)
                    Call setearCuenta(10, eRonda.Ronda_Final)
                    Exit Sub
                End If
            
    End Select
                
        
End Sub

Public Sub vuelveUlla(ByVal UserIndex As Integer)
    On Error GoTo chotserror
    If UserIndex = 0 Or UserList(UserIndex).ConnID = -1 Or UserList(UserIndex).ConnIDValida = False Or UserList(UserIndex).flags.enTorneoAuto = False Then Exit Sub
    Dim Pos As WorldPos
    Dim nPos As WorldPos
    Pos.Map = Torneo_MAPAMUERTE
    Pos.X = 50
    Pos.Y = 50
    Call ClosestLegalPos(Pos, nPos)
    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
    Call UserDie(UserIndex)
    UserList(UserIndex).flags.enTorneoAuto = False
    UserList(UserIndex).flags.enDueloTorneoAuto = False
    UserList(UserIndex).Counters.Torneo = 0
    Exit Sub
    
chotserror:
    Call LogError("Error en Vuelveulla " & Err.number & " " & Err.Description)
End Sub

Public Sub ganaTorneo(ByVal UserIndex As Integer)
    If UserIndex > 0 Then
        If UserList(UserIndex).flags.enTorneoAuto = True Then
            Call WarpUserChar(UserIndex, Torneo_MAPAMUERTE, 50, 51, True)
            UserList(UserIndex).Stats.TorneosAuto = UserList(UserIndex).Stats.TorneosAuto + 1
            UserList(UserIndex).flags.enTorneoAuto = False
            UserList(UserIndex).flags.enDueloTorneoAuto = False
            Call darPremioTorneo(UserIndex)
            Call ActualizarRanking(UserIndex, 4) 'CHOTS | Ranking de Torneos Automaticos
            Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " ha ganado el torneo.", False)
        End If
    End If
    Torneo_HAYTORNEO = False
End Sub

Public Sub setearCuenta(ByVal segundos As Byte, ByRef ronda As eRonda)
    Dim razon As String

    Torneo_RondaActual = ronda
    
    Select Case ronda
        Case eRonda.Ronda_Dieciseisavos:
            razon = "Los diesiseisavos de final"
        Case eRonda.Ronda_Octavos:
            razon = "Los octavos de final"
        Case eRonda.Ronda_Cuartos:
            razon = "Los cuartos de final"
        Case eRonda.Ronda_Semi
            razon = "La semifinal"
        Case eRonda.Ronda_Final
            razon = "La Final"
        Case Else
            razon = "El siguiente evento"
    End Select
            
    Call SendData(SendTarget.ToMap, 0, Torneo_MAPATORNEO, ServerPackages.dialogo & "En " & segundos & " segundos comenzará " & razon & FONTTYPE_TORNEOAUTO)
    Torneo_CR.razon = razon
    Torneo_CR.segundos = segundos
    Torneo_CR.next = ronda
    
End Sub

Public Sub finalizarCuenta()
    Select Case Torneo_CR.next
        Case eRonda.Ronda_Dieciseisavos:
            Call comenzarDieciseisavos
            Exit Sub
            
        Case eRonda.Ronda_Octavos:
            Call comenzarOctavos
            Exit Sub
            
        Case eRonda.Ronda_Cuartos:
            Call comenzarCuartos
            Exit Sub
            
        Case eRonda.Ronda_Semi:
            Call comenzarSemifinal
            Exit Sub
            
        Case eRonda.Ronda_Final:
            Call comenzarFinal
            Exit Sub
            
    End Select
        
End Sub

Public Sub irseTorneo(ByVal UserIndex As Integer)
    Dim i As Integer
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
            For i = 1 To 16
                If Torneo_Dieciseisavos(i).ganador <> UserIndex Then
                    If Torneo_Dieciseisavos(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
                        Exit Sub
                    End If
                    
                    If Torneo_Dieciseisavos(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
                        Exit Sub
                    End If
                Else
                    For j = 1 To 8
                        If Torneo_Octavos(j).usuario1 = UserIndex Then
                            Torneo_Octavos(j).usuario1 = 0
                            Exit Sub
                        End If
                        
                        If Torneo_Octavos(j).usuario2 = UserIndex Then
                            Torneo_Octavos(j).usuario2 = 0
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
        Case eRonda.Ronda_Octavos:
            For i = 1 To 8
                If Torneo_Octavos(i).ganador <> UserIndex Then
                    If Torneo_Octavos(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Octavos(i).usuario2)
                        Exit Sub
                    End If
                    
                    If Torneo_Octavos(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Octavos(i).usuario1)
                        Exit Sub
                    End If
                Else
                    For j = 1 To 4
                        If Torneo_Cuartos(j).usuario1 = UserIndex Then
                            Torneo_Cuartos(j).usuario1 = 0
                            Exit Sub
                        End If
                        
                        If Torneo_Cuartos(i).usuario2 = UserIndex Then
                            Torneo_Cuartos(j).usuario2 = 0
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
        Case eRonda.Ronda_Cuartos:
            For i = 1 To 4
                If Torneo_Cuartos(i).ganador <> UserIndex Then
                    If Torneo_Cuartos(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Cuartos(i).usuario2)
                        Exit Sub
                    End If
                    
                    If Torneo_Cuartos(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Cuartos(i).usuario1)
                        Exit Sub
                    End If
                Else
                    For j = 1 To 2
                        If Torneo_Semifinal(j).usuario1 = UserIndex Then
                            Torneo_Semifinal(j).usuario1 = 0
                            Exit Sub
                        End If
                        
                        If Torneo_Semifinal(i).usuario2 = UserIndex Then
                            Torneo_Semifinal(j).usuario2 = 0
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
        Case eRonda.Ronda_Semi:
            For i = 1 To 2
                If Torneo_Semifinal(i).ganador <> UserIndex Then
                    If Torneo_Semifinal(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Semifinal(i).usuario2)
                        Exit Sub
                    End If
                    
                    If Torneo_Semifinal(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Semifinal(i).usuario1)
                        Exit Sub
                    End If
                Else
                    If Torneo_Final.usuario1 = UserIndex Then
                        Torneo_Final.usuario1 = 0
                        Exit Sub
                    End If
                    
                    If Torneo_Final.usuario2 = UserIndex Then
                        Torneo_Final.usuario2 = 0
                        Exit Sub
                    End If
                End If
            Next i
            
        Case eRonda.Ronda_Final:
            If Torneo_Final.usuario1 = UserIndex Then
                Call SendData(SendTarget.ToIndex, Torneo_Final.usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                Call ganaUsuario(Torneo_Final.usuario2)
                Exit Sub
            End If
            
            If Torneo_Final.usuario2 = UserIndex Then
                Call SendData(SendTarget.ToIndex, Torneo_Final.usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                Call ganaUsuario(Torneo_Final.usuario1)
                Exit Sub
            End If
            
        Case Else:
            For i = 1 To Torneo_CantidadInscriptos
            
                If Torneo_UsuariosInscriptos(i) = UserIndex Then
                    Torneo_UsuariosInscriptos(i) = 0
                    Exit Sub
                End If
                
            Next i
    
    End Select
    
End Sub

Public Sub cerrarTorneo()
    Dim i As Integer
    Dim Pos As WorldPos
    Dim nPos As WorldPos
    
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> El torneo ha sido cancelado." & FONTTYPE_TORNEOAUTO)
    
    For i = 1 To LastUser
        If UserList(i).flags.enTorneoAuto = True Then
            Pos.Map = 1
            Pos.X = 58
            Pos.Y = 45
            Call ClosestLegalPos(Pos, nPos)
            Call WarpUserChar(i, nPos.Map, nPos.X, nPos.Y, True)
            UserList(i).flags.enTorneoAuto = False
        End If
    Next i

    Torneo_HAYTORNEO = False

End Sub

Public Sub terminarDuelo(ByVal UserIndex As Integer)
Dim i As Integer
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
            For i = 1 To 16
            
                If Torneo_Dieciseisavos(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
                    End If
                    Exit Sub
                End If

            Next i
            
        Case eRonda.Ronda_Octavos:
            For i = 1 To 8
            
                If Torneo_Octavos(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Octavos(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Octavos(i).usuario1)
                    End If
                    Exit Sub
                End If
            Next i
            
        Case eRonda.Ronda_Cuartos:
            For i = 1 To 4
            
                If Torneo_Cuartos(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Cuartos(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Cuartos(i).usuario1)
                    End If
                    
                    Exit Sub
                End If

            Next i
            
        Case eRonda.Ronda_Semi:
            For i = 1 To 2
            
                If Torneo_Semifinal(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Semifinal(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Semifinal(i).usuario1)
                    End If
                    
                    Exit Sub
                End If
                
            Next i
    
    End Select
End Sub

Public Sub rearmarTorneo()
Dim Pos As WorldPos
Dim nPos As WorldPos

If Torneo_CantidadInscriptos >= (Torneo_Cupo / 2) Then

    Call SendData(SendTarget.ToMap, 0, Torneo_MAPATORNEO, ServerPackages.dialogo & "El torneo se ha reorganizado para " & Torneo_Cupo / 2 & " usuarios. Los sobrantes serán enviados a Ullatorphe" & FONTTYPE_TORNEOAUTO)

    'CHOTS | Los envía a Ulla a los sobrantes
    If Torneo_CantidadInscriptos > (Torneo_Cupo / 2) Then
        For i = ((Torneo_Cupo / 2) + 1) To Torneo_Cupo
            If Torneo_UsuariosInscriptos(i) > 0 Then
                Pos.Map = 1
                Pos.X = 58
                Pos.Y = 45
                Call ClosestLegalPos(Pos, nPos)
                Call WarpUserChar(Torneo_UsuariosInscriptos(i), nPos.Map, nPos.X, nPos.Y, True)
                UserList(Torneo_UsuariosInscriptos(i)).flags.enTorneoAuto = False
                Torneo_UsuariosInscriptos(i) = 0
            End If
        Next i
    End If
    
    Torneo_Cupo = Torneo_Cupo / 2
    Torneo_CantidadInscriptos = Torneo_Cupo
    
    Call armarFixture

    Call LogGM("TORNEOAUTO", "El torneo se reorganizo para " & Torneo_Cupo & " participantes.", False)

Else
    Call LogGM("TORNEOAUTO", "El torneo se cancelo por falta de participantes.", False)
    Call cerrarTorneo
End If

End Sub

Public Function minutosProxTorneo() As String
    On Error GoTo CHOTSERR
    Dim minutelis As Long
    If MinutosParaWs > 6 Then
        minutelis = (MinutosWs - MinutosParaWs) + 6
    Else
        minutelis = 6 - MinutosParaWs
    End If
    
    minutosProxTorneo = Trim$(str$(minutelis))
    Exit Function
    
CHOTSERR:
    Call LogError("Error en Proxtorneos " & Err.number & " " & Err.Description)
    
End Function

Private Sub darPremioTorneo(ByVal UserIndex As Integer)
Dim item As Integer
Dim MiObj As Obj
item = 0

Select Case Torneo_Cupo
    Case 32:
        UserList(UserIndex).Puntos = UserList(UserIndex).Puntos + 30
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z99")
        item = 931
        
    Case 16:
        UserList(UserIndex).Puntos = UserList(UserIndex).Puntos + 30
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z99")
        item = 1121
        
    Case 8:
        UserList(UserIndex).Puntos = UserList(UserIndex).Puntos + 30
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z99")
    
    Case Else:
        Exit Sub
        
End Select

If item > 0 Then
    MiObj.Amount = 1
    MiObj.ObjIndex = item
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
    
End Sub
