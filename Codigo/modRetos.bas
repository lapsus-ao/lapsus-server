Attribute VB_Name = "modRetos"
'Modulo para retos 1v1 - 2v2 - Clan vs Clan por Items/Oro/Puntos
'Programado por Andrés Nicolini alias BysNacK para lapsus 2017
Option Explicit
Private X, Y, Sala As Byte
Private Const CANTIDAD_MAPAS_RETO = 4
Private MapasReto(1 To CANTIDAD_MAPAS_RETO) As Byte


Public Sub InicializarMapasReto()
    MapasReto(1) = 84
    MapasReto(2) = 87
    MapasReto(3) = 88
    MapasReto(4) = 56
End Sub

Public Sub InitiateReto(participantes() As String)
    On Error GoTo errhandler

    Dim warpPos As WorldPos
    Dim nPos As WorldPos
    Dim cantpart As Byte
    Dim i As Byte
    Dim posy1, posx1, posx2, posy2 As Integer
    Dim numeroMapa, SalaVacia As Byte
    Dim haySalaVacia As Boolean
    cantpart = UserList(NameIndex(participantes(1))).Reto.TipoReto * 2
    
    'Reto en Curso
    haySalaVacia = False
    For numeroMapa = 1 To CANTIDAD_MAPAS_RETO
        If MapInfo(MapasReto(numeroMapa)).NumUsers = 0 Then
            haySalaVacia = True
            Exit For
        End If
    Next numeroMapa
        
    If haySalaVacia = False Then
        For i = 1 To cantpart
            SendData SendTarget.ToIndex, NameIndex(participantes(i)), 0, ServerPackages.dialogo & "Todas las salas de retos se encuentran ocupadas!." & FONTTYPE_DUELO
        Next i
    Else

        For i = 1 To cantpart
            UserList(NameIndex(participantes(i))).Reto.enReto = True
        Next i

        Select Case numeroMapa
            Case 1
                For i = 1 To cantpart
                    If i Mod 2 <> 0 Then
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 30
                        warpPos.Y = 30
                    Else
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 61
                        warpPos.Y = 49
                    End If
                    
                    Call ClosestLegalPos(warpPos, nPos)
                    Call WarpUserChar(NameIndex(participantes(i)), nPos.Map, nPos.X, nPos.Y, True)
                Next i
            Case 2
                For i = 1 To cantpart
                    If i Mod 2 <> 0 Then
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 26
                        warpPos.Y = 30
                    Else
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 51
                        warpPos.Y = 49
                    End If
                    
                    Call ClosestLegalPos(warpPos, nPos)
                    Call WarpUserChar(NameIndex(participantes(i)), nPos.Map, nPos.X, nPos.Y, True)
                Next i
                
            Case 3
                For i = 1 To cantpart
                    If i Mod 2 <> 0 Then
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 27
                        warpPos.Y = 32
                    Else
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 61
                        warpPos.Y = 53
                    End If
                    
                    Call ClosestLegalPos(warpPos, nPos)
                    Call WarpUserChar(NameIndex(participantes(i)), nPos.Map, nPos.X, nPos.Y, True)
                Next i
                
            Case 4
                For i = 1 To cantpart
                    If i Mod 2 <> 0 Then
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 20
                        warpPos.Y = 20
                    Else
                        warpPos.Map = MapasReto(numeroMapa)
                        warpPos.X = 40
                        warpPos.Y = 40
                    End If
                    
                    Call ClosestLegalPos(warpPos, nPos)
                    Call WarpUserChar(NameIndex(participantes(i)), nPos.Map, nPos.X, nPos.Y, True)
                Next i
        End Select
        
        Call EnviarMensajesGlobales(participantes, cantpart)
    End If

    Exit Sub

errhandler:
    Call LogError("InitiateReto - Error = " & Err.number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub
Public Sub EndReto(ByVal UserIndex As Integer)

    Dim warpPos As WorldPos
    Dim nPos As WorldPos
    Dim cantpart, i, mapaReto As Byte
    Dim participantes() As String
    
    cantpart = UserList(UserIndex).Reto.TipoReto * 2
    
    ReDim participantes(1 To cantpart)
    
    participantes(1) = UserList(UserIndex).Name
    participantes(2) = UserList(UserIndex).Reto.Oponente
    If cantpart = 4 Then
        participantes(3) = UserList(UserIndex).Reto.Pareja
        participantes(4) = UserList(NameIndex(UserList(UserIndex).Reto.Oponente)).Reto.Pareja
    End If
    
    mapaReto = UserList(UserIndex).Pos.Map
    
            For i = 1 To cantpart
                If i Mod 2 <> 0 Then
                    UserList(NameIndex(participantes(i))).Reto.EsperandoReto = False
                    UserList(NameIndex(participantes(i))).Reto.Oponente = ""
                    UserList(NameIndex(participantes(i))).Reto.enReto = False
                    UserList(NameIndex(participantes(i))).Reto.EnvioRequest = False
                    UserList(NameIndex(participantes(i))).Reto.TimeReto = 2
                    UserList(NameIndex(participantes(i))).Reto.PerdioReto = True
                    Select Case mapaReto
                        Case 84
                            warpPos.Map = 85
                            warpPos.X = 61
                            warpPos.Y = 23
                        Case 87
                            warpPos.Map = 85
                            warpPos.X = 23
                            warpPos.Y = 78
                        Case 88
                            warpPos.Map = 85
                            warpPos.X = 24
                            warpPos.Y = 20
                        Case 56
                            warpPos.Map = 85
                            warpPos.X = 78
                            warpPos.Y = 80
                    End Select
                    
                    Call ClosestLegalPos(warpPos, nPos)
                    Call WarpUserChar(NameIndex(participantes(i)), nPos.Map, nPos.X, nPos.Y, True) 'Este va a dropear
                Else
                    UserList(NameIndex(participantes(i))).Reto.EsperandoReto = False
                    UserList(NameIndex(participantes(i))).Reto.Oponente = ""
                    UserList(NameIndex(participantes(i))).Reto.EnvioRequest = False
                    UserList(NameIndex(participantes(i))).Reto.enReto = False
                    UserList(NameIndex(participantes(i))).Reto.TimeReto = 8
                    UserList(NameIndex(participantes(i))).Reto.GanoReto = True
                    If UserList(NameIndex(participantes(i))).flags.Muerto = 1 Then Call Resucitar(NameIndex(participantes(i)))
                    
                    Select Case mapaReto
                        Case 84
                            warpPos.Map = 85
                            warpPos.X = 62
                            warpPos.Y = 24
                        Case 87
                            warpPos.Map = 85
                            warpPos.X = 24
                            warpPos.Y = 76
                        Case 88
                            warpPos.Map = 85
                            warpPos.X = 25
                            warpPos.Y = 22
                        Case 56
                            warpPos.Map = 85
                            warpPos.X = 79
                            warpPos.Y = 81
                    End Select
                    
                    Call ClosestLegalPos(warpPos, nPos)
                    Call WarpUserChar(NameIndex(participantes(i)), nPos.Map, nPos.X, nPos.Y, True) 'Este va a lukear
                    Call SendData(SendTarget.ToIndex, NameIndex(participantes(i)), 0, ServerPackages.dialogo & "Has ganado el reto, tienes 10 segundos para obtener tus recompensas y luego seras llevado a la ciudad." & FONTTYPE_DUELO)
                End If
            Next i

    Call HandlePremios(participantes, cantpart)
    
End Sub
Sub HandlePremios(participantes() As String, ByVal cantpart As Byte)

Dim i As Byte

    For i = 1 To cantpart
        If i Mod 2 = 0 Then
            UserList(NameIndex(participantes(i))).Puntos = UserList(NameIndex(participantes(i))).Puntos + UserList(NameIndex(participantes(i))).Reto.pricePuntos
            UserList(NameIndex(participantes(i))).Stats.GLD = UserList(NameIndex(participantes(i))).Stats.GLD + UserList(NameIndex(participantes(i))).Reto.priceGold
            Call EnviarOro(NameIndex(participantes(i)))
        Else
            UserList(NameIndex(participantes(i))).Puntos = UserList(NameIndex(participantes(i))).Puntos - UserList(NameIndex(participantes(i))).Reto.pricePuntos
            UserList(NameIndex(participantes(i))).Stats.GLD = UserList(NameIndex(participantes(i))).Stats.GLD - UserList(NameIndex(participantes(i))).Reto.priceGold
            Call EnviarOro(NameIndex(participantes(i)))
            If UserList(NameIndex(participantes(i))).Reto.priceItems = 1 Then Call DropItems(NameIndex(participantes(i)))
        End If
    Next i
    
End Sub
Public Sub InitiateFlagsReto(participantes() As String, ByVal Puntos As Integer, ByVal Gold As Long, ByVal poritems As Byte, ByVal RetType As Byte)
On Error GoTo errhandler
Dim n, i As Byte
    n = RetType
    
    UserList(NameIndex(participantes(1))).Reto.EnvioRequest = True
    
    For i = 1 To RetType
        'Tipo reto
        If RetType = 2 Then
            UserList(NameIndex(participantes(i))).Reto.TipoReto = 1
        Else
            'Set pareja
            If i <= 2 Then
                UserList(NameIndex(participantes(i))).Reto.Pareja = participantes(i + 2)
            Else
                UserList(NameIndex(participantes(i))).Reto.Pareja = participantes(i - 2)
            End If
            UserList(NameIndex(participantes(i))).Reto.TipoReto = 2
        End If
        'Caen items
        UserList(NameIndex(participantes(i))).Reto.priceItems = poritems
        'Por oro
        UserList(NameIndex(participantes(i))).Reto.priceGold = Gold
        'Por puntos
        UserList(NameIndex(participantes(i))).Reto.pricePuntos = Puntos
        'Esperando reto ON
        UserList(NameIndex(participantes(i))).Reto.EsperandoReto = True
        'Set oponente
        UserList(NameIndex(participantes(i))).Reto.Oponente = participantes(n)
        n = n - 1
        'No acepto aún
        UserList(NameIndex(participantes(i))).Reto.AceptoReto = False
    Next i
        UserList(NameIndex(participantes(1))).Reto.AceptoReto = True

    Exit Sub
errhandler:
    Call LogError("InitiateFlagsReto - Error = " & Err.number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub
Public Function PuedeEnviarReto(ByRef msgError As String, participantes() As String, ByVal dsPuntos As Integer, ByVal dsOro As Long, ByVal chkItems As Byte, ByVal TypeReto As Byte) As Boolean
                
    PuedeEnviarReto = False
    
    Dim i As Byte
    
 Select Case TypeReto
    Case 1
        If NameIndex(participantes(2)) <= 0 Then
            msgError = "Rival invalido." & FONTTYPE_DUELO
            PuedeEnviarReto = False
            Exit Function
        End If
        'Check Mapa
        For i = 1 To 2
            If Not UserList(NameIndex(participantes(i))).Pos.Map = 1 Then
                msgError = participantes(i) & " no se encuentra la ciudad del NPC." & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If

            'Check puntos
            If UserList(NameIndex(participantes(i))).Puntos < dsPuntos Then
                msgError = participantes(i) & " no tiene suficientes puntos para este desafio" & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
                
            'Check oro
            If UserList(NameIndex(participantes(i))).Stats.GLD < dsOro Then
                msgError = participantes(i) & " no tiene  suficiente oro para este desafio" & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
                
            'Check Muerto
            If UserList(NameIndex(participantes(i))).flags.Muerto = 1 Then
                msgError = participantes(i) & " está muerto." & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
        
            'Check pedido en curso
            If UserList(NameIndex(participantes(i))).Reto.EsperandoReto Then
                msgError = "El usuario " & participantes(i) & " tiene otra solicitud de reto en curso " & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
        
    Next i
                
        'Check Vos mismo
        If participantes(1) = participantes(2) Then
            msgError = "No podes retarte a vos mismo." & FONTTYPE_DUELO
            PuedeEnviarReto = False
            Exit Function
        End If

                
        msgError = "Has retado a " & participantes(2) & FONTTYPE_DUELO
        PuedeEnviarReto = True
        
    Case 2
    
        If NameIndex(participantes(2)) <= 0 Or NameIndex(participantes(3)) <= 0 Or NameIndex(participantes(4)) <= 0 Then
            msgError = "Rival invalido." & FONTTYPE_DUELO
            PuedeEnviarReto = False
            Exit Function
        End If
        
        'Check Mapa
        For i = 1 To 4
            If Not UserList(NameIndex(participantes(i))).Pos.Map = 1 Then
                msgError = participantes(i) & " no se encuentra la ciudad del NPC." & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If

            'Check puntos
            If UserList(NameIndex(participantes(i))).Puntos < dsPuntos Then
                msgError = participantes(i) & "No tiene suficientes puntos para este desafio" & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
                    
            'Check oro
            If UserList(NameIndex(participantes(i))).Stats.GLD < dsOro Then
                msgError = participantes(i) & "No tiene  suficiente oro para este desafio" & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
                    
            'Check Muerto
            If UserList(NameIndex(participantes(i))).flags.Muerto = 1 Then
                msgError = participantes(i) & " está muerto." & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
            
            'Check pedido en curso
            If UserList(NameIndex(participantes(i))).Reto.EsperandoReto Then
                msgError = "El usuario " & participantes(i) & " tiene otra solicitud de reto en curso " & FONTTYPE_DUELO
                PuedeEnviarReto = False
                Exit Function
            End If
        
    Next i
                
        'Check Vos mismo
        If participantes(1) = participantes(2) Or participantes(1) = participantes(3) Or participantes(1) = participantes(4) Then
            msgError = "No podes dueliar con vos mismo." & FONTTYPE_DUELO
            PuedeEnviarReto = False
            Exit Function
        End If
    
                
        msgError = "Has retado a " & participantes(2) & " y " & participantes(4) & FONTTYPE_DUELO
        PuedeEnviarReto = True
End Select

End Function
Public Function PuedeAceptarReto(ByRef msgError As String, participantes() As String, ByVal dsPuntos As Integer, ByVal dsOro As Long, ByVal chkItems As Byte, ByVal TypeReto As Byte) As Boolean
                
    PuedeAceptarReto = False
    
    Dim i As Byte
    
 Select Case TypeReto
    Case 1
    
        If NameIndex(participantes(2)) <= 0 Then
            msgError = "Rival invalido." & FONTTYPE_DUELO
            PuedeAceptarReto = False
            Exit Function
        End If
        'Check Mapa
        For i = 1 To 2
            If Not UserList(NameIndex(participantes(i))).Pos.Map = 1 Then
                msgError = participantes(i) & " no se encuentra la ciudad del NPC." & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If

            'Check puntos
            If UserList(NameIndex(participantes(i))).Puntos < dsPuntos Then
                msgError = participantes(i) & "No tiene suficientes puntos para este desafio" & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If
                
            'Check oro
            If UserList(NameIndex(participantes(i))).Stats.GLD < dsOro Then
                msgError = participantes(i) & "No tiene  suficiente oro para este desafio" & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If
                
            'Check Muerto
            If UserList(NameIndex(participantes(i))).flags.Muerto = 1 Then
                msgError = participantes(i) & " está muerto." & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If
        
    Next i

        PuedeAceptarReto = True
        
    Case 2
    
        If NameIndex(participantes(2)) <= 0 Or NameIndex(participantes(3)) <= 0 Or NameIndex(participantes(4)) <= 0 Then
            msgError = "Rival invalido." & FONTTYPE_DUELO
            PuedeAceptarReto = False
            Exit Function
        End If
        
    For i = 1 To 4
    
        If Not UserList(NameIndex(participantes(i))).Pos.Map = 1 Then
                msgError = participantes(i) & " no se encuentra la ciudad del NPC." & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If

            'Check puntos
            If UserList(NameIndex(participantes(i))).Puntos < dsPuntos Then
                msgError = participantes(i) & "No tiene suficientes puntos para este desafio" & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If
                
            'Check oro
            If UserList(NameIndex(participantes(i))).Stats.GLD < dsOro Then
                msgError = participantes(i) & "No tiene  suficiente oro para este desafio" & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If
                
            'Check Muerto
            If UserList(NameIndex(participantes(i))).flags.Muerto = 1 Then
                msgError = participantes(i) & " está muerto." & FONTTYPE_DUELO
                PuedeAceptarReto = False
                Exit Function
            End If
        
    Next i
    
                
        PuedeAceptarReto = True
End Select
End Function
Public Sub ResetFlagsReto(ByVal UserIndex As Integer)
    On Error GoTo errhandler

    Dim cantpart, i As Byte
    Dim participantes() As String
    
    cantpart = UserList(UserIndex).Reto.TipoReto * 2
    
    ReDim participantes(1 To cantpart)
    
    participantes(1) = UserList(UserIndex).Name
    participantes(2) = UserList(UserIndex).Reto.Oponente
    If cantpart = 4 Then
        participantes(3) = UserList(UserIndex).Reto.Pareja
        participantes(4) = UserList(NameIndex(participantes(2))).Reto.Pareja
    End If
    
    For i = 1 To cantpart
        Dim CurrentUser as Integer
        CurrentUser = NameIndex(participantes(i))
        If CurrentUser > 0 Then
            UserList(CurrentUser).Reto.EnvioRequest = False
            UserList(CurrentUser).Reto.AceptoReto = False
            UserList(CurrentUser).Reto.enReto = False
            UserList(CurrentUser).Reto.EsperandoReto = False
            UserList(CurrentUser).Reto.Pareja = ""
            UserList(CurrentUser).Reto.Oponente = ""
            SendData SendTarget.ToIndex, CurrentUser, 0, ServerPackages.dialogo & "El reto ha sido rechazado por " & participantes(1) & FONTTYPE_DUELO
        End if
    Next i

    Exit Sub

errhandler:
        Call LogError("ResetFlagsReto - Error = " & Err.number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub
Public Sub EnviarRequestReto(participantes() As String, Cant As Byte)
Select Case Cant
    Case 2 ' 1v1
        Select Case UserList(NameIndex(participantes(1))).Reto.priceItems
            Case 1 'Caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'Y'No hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'Y'N'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'Y'N'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & " y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                    Case Else 'Y' Hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'Y'Y'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'Y'Y'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                End Select
            Case Else 'No caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'N'No hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'N'N'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'N'N'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & ", /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                    Case Else 'N' Hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'N'Y'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'N'Y'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " te ha retado a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                End Select
        End Select
    Case Else '2v2
        Select Case UserList(NameIndex(participantes(1))).Reto.priceItems
            Case 1 'Caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'Y'No hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'Y'N'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por todos sus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'Y'N'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & " y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & " y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & " y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                    Case Else 'Y' Hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'Y'Y'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'Y'Y'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro, " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro, " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro, " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos y todos tus items, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                End Select
            Case Else 'No caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'N'No hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'N'N'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & " y " & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & " y " & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'N'N'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & ", /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & ", /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro " & ", /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                    Case Else 'N' Hay puntos en juego
                        Select Case UserList(NameIndex(participantes(1))).Reto.priceGold
                            Case 0 'N'Y'No hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                            Case Else 'N'Y'Si hay oro en juego
                                SendData SendTarget.ToIndex, NameIndex(participantes(4)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(2) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(2)), 0, ServerPackages.dialogo & participantes(1) & participantes(3) & " te han retado a ti y a " & participantes(4) & " a un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                                SendData SendTarget.ToIndex, NameIndex(participantes(3)), 0, ServerPackages.dialogo & participantes(1) & " te ha nombrado como su pareja para retar a " & participantes(2) & " y " & participantes(4) & " en un duelo por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos, /ACEPTAR para aceptar o /CANCELAR para rechazar." & FONTTYPE_DUELO
                        End Select
                End Select
        End Select
End Select
End Sub
Sub EnviarMensajesGlobales(participantes() As String, Cant As Byte)
Select Case Cant
    Case 2 ' 1v1
        Select Case UserList(NameIndex(participantes(1))).Reto.priceItems
            Case 1 'Caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'Y'No hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "El jugador " & participantes(1) & " enfrentará en un reto a " & participantes(2) & " por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y todos sus items." & FONTTYPE_DUELO)
                    Case Else 'Y' Hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "El jugador " & participantes(1) & " enfrentará en un reto a " & participantes(2) & " por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos." & FONTTYPE_DUELO)
                End Select
            Case Else 'No caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'N'No hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "El jugador " & participantes(1) & " enfrentará en un reto a " & participantes(2) & " por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro." & FONTTYPE_DUELO)
                    Case Else 'N' Hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "El jugador " & participantes(1) & " enfrentará en un reto a " & participantes(2) & " por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos." & FONTTYPE_DUELO)
                End Select
        End Select
    Case Else '2v2
        Select Case UserList(NameIndex(participantes(1))).Reto.priceItems
            Case 1 'Caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'Y'No hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Los jugadores " & participantes(1) & " y " & participantes(3) & " retaron a un duelo a los jugadores " & participantes(2) & " y " & participantes(4) & " por todos sus items y " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro." & FONTTYPE_DUELO)
                    Case Else 'Y' Hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Los jugadores " & participantes(1) & " y " & participantes(3) & " retaron a un duelo a los jugadores " & participantes(2) & " y " & participantes(4) & " por todos sus items, " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos." & FONTTYPE_DUELO)
                End Select
            Case Else 'No caen items
                Select Case UserList(NameIndex(participantes(1))).Reto.pricePuntos
                    Case 0 'N'No hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Los jugadores " & participantes(1) & " y " & participantes(3) & " retaron a un duelo a los jugadores " & participantes(2) & " y " & participantes(4) & " por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro." & FONTTYPE_DUELO)
                    Case Else 'N' Hay puntos en juego
                        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Los jugadores " & participantes(1) & " y " & participantes(3) & " retaron a un duelo a los jugadores " & participantes(2) & " y " & participantes(4) & " por " & UserList(NameIndex(participantes(1))).Reto.priceGold & " monedas de oro y " & UserList(NameIndex(participantes(1))).Reto.pricePuntos & " puntos." & FONTTYPE_DUELO)
                End Select
        End Select
End Select
End Sub
