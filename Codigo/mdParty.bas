Attribute VB_Name = "mdParty"
Option Explicit

''
' SOPORTES PARA LAS PARTIES
' (Ver este modulo como una clase abstracta "PartyManager")
'


''
'cantidad maxima de parties en el servidor
Public Const MAX_PARTIES As Integer = 100

''
'nivel minimo para crear party
Public Const MINPARTYLEVEL As Byte = 15

''
'Cantidad maxima de gente en la party
Public Const PARTY_MAXMEMBERS As Byte = 5

''
'Si esto esta en True, la exp sale por cada golpe que le da
'Si no, la exp la recibe al salirse de la party (pq las partys, floodean)
Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = True

''
'maxima diferencia de niveles permitida en una party
Public Const MAXPARTYDELTALEVEL As Byte = 10

''
'distancia al leader para que este acepte el ingreso
Public Const MAXDISTANCIAINGRESOPARTY As Byte = 5

''
'maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MAXDISTANCIA As Byte = 20

''
'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

''
'tPartyMember
'
' @param UserIndex UserIndex
' @param Experiencia Experiencia
'
Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Long
End Type


Public Function NextParty() As Integer
Dim i As Integer
NextParty = -1
For i = 1 To MAX_PARTIES
    If Parties(i) Is Nothing Then
        NextParty = i
        Exit Function
    End If
Next i
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
    PuedeCrearParty = True
'    If UserList(UserIndex).Stats.ELV < MINPARTYLEVEL Then
    If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 10 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu liderazgo no es suficientes para liderar una party." & FONTTYPE_PARTY)
        PuedeCrearParty = False
    ElseIf UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estás muerto!" & FONTTYPE_PARTY)
        PuedeCrearParty = False
    End If
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
Dim tInt As Integer
If UserList(UserIndex).PartyIndex = 0 Then
    If UserList(UserIndex).flags.Muerto = 0 Then
        If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
            tInt = mdParty.NextParty
            If tInt = -1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Por el momento no se pueden crear mas parties" & FONTTYPE_PARTY)
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " La party está llena, no puedes entrar" & FONTTYPE_PARTY)
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " ¡ Has formado una party !" & FONTTYPE_PARTY)
                    UserList(UserIndex).PartyIndex = tInt
                    UserList(UserIndex).PartySolicitud = 0
                    If Not Parties(tInt).HacerLeader(UserIndex) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No puedes hacerte líder." & FONTTYPE_PARTY)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " ¡ Te has convertido en líder de la party !" & FONTTYPE_PARTY)
                    End If
                End If
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No tienes suficientes puntos de liderazgo para liderar una party." & FONTTYPE_PARTY)
        End If
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Estás muerto!" & FONTTYPE_PARTY)
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Ya perteneces a una party." & FONTTYPE_PARTY)
End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
Dim tInt As Integer

    If UserList(UserIndex).PartyIndex > 0 Then
        'si ya esta en una party
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Ya perteneces a una party, escribe /SALIRPARTY para abandonarla" & FONTTYPE_PARTY)
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " ¡Estás muerto!" & FONTTYPE_INFO)
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    tInt = UserList(UserIndex).flags.TargetUser
    If tInt > 0 Then
        If UserList(tInt).PartyIndex > 0 Then
            UserList(UserIndex).PartySolicitud = UserList(tInt).PartyIndex
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " El fundador decidirá si te acepta en la party" & FONTTYPE_PARTY)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & UserList(tInt).Name & " no es fundador de ninguna party." & FONTTYPE_INFO)
            UserList(UserIndex).PartySolicitud = 0
            Exit Sub
        End If
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY" & FONTTYPE_PARTY)
        UserList(UserIndex).PartySolicitud = 0
    End If

End Sub
Public Sub SalirDeParty(ByVal UserIndex As Integer)
Dim PI As Integer
PI = UserList(UserIndex).PartyIndex
If PI > 0 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "ENP")
    If Parties(PI).SaleMiembro(UserIndex) Then
        'sale el leader
        Set Parties(PI) = Nothing
    Else
        UserList(UserIndex).PartyIndex = 0
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No eres miembro de ninguna party." & FONTTYPE_INFO)
End If

End Sub


Public Sub ExpulsarDeParty(ByVal Leader As Integer, ByVal OldMember As Integer)
Dim PI As Integer
Dim razon As String
PI = UserList(Leader).PartyIndex
If PI > 0 Then
    If PI = UserList(OldMember).PartyIndex Then
        If Parties(PI).EsPartyLeader(Leader) Then
            If Parties(PI).SaleMiembro(OldMember) Then
                'si la funcion me da true, entonces la party se disolvio
                'y los partyindex fueron reseteados a 0
                Set Parties(PI) = Nothing
            Else
                UserList(OldMember).PartyIndex = 0
            End If
        Else
            Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " Solo el fundador puede expulsar miembros de una party." & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " " & UserList(OldMember).Name & " no pertenece a tu party." & FONTTYPE_INFO)
    End If
Else
    Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " No eres miembro de ninguna party." & FONTTYPE_INFO)
End If



End Sub


Public Sub AprobarIngresoAParty(ByVal Leader As Integer, ByVal NewMember As Integer)
'el UI es el leader
Dim PI As Integer
Dim razon As String

PI = UserList(Leader).PartyIndex

If PI > 0 Then
    If Parties(PI).EsPartyLeader(Leader) Then
        If UserList(NewMember).PartyIndex = 0 Then
            If Not UserList(Leader).flags.Muerto = 1 Then
                If Not UserList(NewMember).flags.Muerto = 1 Then
                    If UserList(NewMember).PartySolicitud = PI Then
                        If Parties(PI).PuedeEntrar(NewMember, razon) Then
                            If Parties(PI).NuevoMiembro(NewMember) Then
                                Call Parties(PI).MandarMensajeAConsola(UserList(Leader).Name & " ha aceptado a " & UserList(NewMember).Name & " en la party.", "Servidor")
                                UserList(NewMember).PartyIndex = PI
                                UserList(NewMember).PartySolicitud = 0
                            Else
                                'no pudo entrar
                                'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                                Call SendData(SendTarget.ToAdmins, Leader, 0, ServerPackages.dialogo & " Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S " & FONTTYPE_PARTY)
                            End If
                        Else
                            'no debe entrar
                            Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " " & razon & FONTTYPE_PARTY)
                        End If
                    Else
                        Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " " & UserList(NewMember).Name & " no ha solicitado ingresar a tu party." & FONTTYPE_PARTY)
                        Exit Sub
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " ¡Está muerto, no puedes aceptar miembros en ese estado!" & FONTTYPE_PARTY)
                    Exit Sub
                End If
            Else
                Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " ¡Estás muerto, no puedes aceptar miembros en ese estado!" & FONTTYPE_PARTY)
                Exit Sub
            End If
        Else
            Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & UserList(NewMember).Name & " ya es miembro de otra party." & FONTTYPE_PARTY)
            ' ya tiene party el otro tipo
        End If
    Else
        Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " No eres líder, no puedes aceptar miembros." & FONTTYPE_PARTY)
        Exit Sub
    End If
Else
    Call SendData(SendTarget.ToIndex, Leader, 0, ServerPackages.dialogo & " No eres miembro de ninguna party." & FONTTYPE_INFO)
    Exit Sub
End If

End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
Dim PI As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
    End If

End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
Dim PI As Integer
Dim texto As String

    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(texto)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & texto & FONTTYPE_PARTY)
    End If
    

End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub

PI = UserList(OldLeader).PartyIndex

If PI > 0 Then
    If PI = UserList(NewLeader).PartyIndex Then
        If UserList(NewLeader).flags.Muerto = 0 Then
            If Parties(PI).EsPartyLeader(OldLeader) Then
                If Parties(PI).HacerLeader(NewLeader) Then
                    Call Parties(PI).MandarMensajeAConsola("El nuevo líder de la party es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
                Else
                    Call SendData(SendTarget.ToIndex, OldLeader, 0, ServerPackages.dialogo & "¡No se ha hecho el cambio de mando!" & FONTTYPE_PARTY)
                End If
            Else
                Call SendData(SendTarget.ToIndex, OldLeader, 0, ServerPackages.dialogo & "¡No eres el líder!" & FONTTYPE_PARTY)
            End If
        Else
            Call SendData(SendTarget.ToIndex, OldLeader, 0, ServerPackages.dialogo & "¡Está muerto!" & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.ToIndex, OldLeader, 0, ServerPackages.dialogo & UserList(NewLeader).Name & " no pertenece a tu party." & FONTTYPE_INFO)
    End If
End If

End Sub


Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim i As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    
    haciendoBK = True

    
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Distribuyendo experiencia en parties." & FONTTYPE_SERVER)
    For i = 1 To MAX_PARTIES
        If Not Parties(i) Is Nothing Then
            Call Parties(i).FlushExperiencia
        End If
    Next i
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Experiencia distribuida." & FONTTYPE_SERVER)
    haciendoBK = False

End If

End Sub
Public Sub DividirOro(ByVal UserIndex As Integer, ByVal oro As Long, mapa As Integer, X As Integer, Y As Integer)
    If oro <= 0 Then Exit Sub
    
    Call Parties(UserList(UserIndex).PartyIndex).DividirOro(oro, mapa, X, Y)

End Sub
Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Double, mapa As Integer, X As Integer, Y As Integer)
    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, X, Y)

End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
CantMiembros = 0
If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If

End Function
