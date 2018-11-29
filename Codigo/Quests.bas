Attribute VB_Name = "Quests"
'Modulo Quests
'Creado por Juan Andrés Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'04/09/2010
'Para LapsusAO 2.0

Public Sub TerminarQuest(ByVal UserIndex As Integer)

UserList(UserIndex).Quest.cantNpc = 0
UserList(UserIndex).Quest.cantMatados = 0
UserList(UserIndex).Quest.numNpc = 0

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Has Terminado la Quest! Felicidades! Aquí tienes tu premio..." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)

Call EntregarPremioQuest(UserIndex)

End Sub
Public Sub EnviarListaQuest(ByVal UserIndex As Integer)
  Call SendData(SendTarget.ToIndex, UserIndex, 0, "LSTQUE")
End Sub

Public Sub EntregarPremioQuest(ByVal UserIndex As Integer)

Dim item As Obj
Dim oro As Long
Dim Puntos As Integer

oro = val(GetVar(DatPath & "QUEST.dat", "QUEST" & UserList(UserIndex).Quest.nroQuest, "Oro"))
Puntos = val(GetVar(DatPath & "QUEST.dat", "QUEST" & UserList(UserIndex).Quest.nroQuest, "Puntos"))
item.ObjIndex = val(GetVar(DatPath & "QUEST.dat", "QUEST" & UserList(UserIndex).Quest.nroQuest, "Item"))
item.Amount = 1

UserList(UserIndex).Quest.nroQuest = 0

If oro <> 0 Then
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + oro
End If

If Puntos <> 0 Then
    UserList(UserIndex).Puntos = UserList(UserIndex).Puntos + Puntos
End If

If item.ObjIndex <> 0 Then
    If Not MeterItemEnInventario(UserIndex, item) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, item)
    End If
End If

Call UpdateUserInv(True, UserIndex, 0)
Call EnviarOro(UserIndex)

End Sub

Public Sub AceptarQuest(ByVal UserIndex As Integer, ByVal Numero As Byte)

UserList(UserIndex).Quest.nroQuest = Numero

UserList(UserIndex).Quest.cantNpc = val(GetVar(DatPath & "QUEST.dat", "QUEST" & UserList(UserIndex).Quest.nroQuest, "Cant"))
UserList(UserIndex).Quest.numNpc = val(GetVar(DatPath & "QUEST.dat", "QUEST" & UserList(UserIndex).Quest.nroQuest, "Npc"))

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Has Aceptado la Quest! Cuando finalizes puedes venir hacia mi y tipear /FINQUEST. Utiliza /INFOQUEST para ver el estado de la Quest actual o /CANCELARQUEST si cambias de opinión." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
End Sub

Public Sub CancelarQuest(ByVal UserIndex As Integer)

UserList(UserIndex).Quest.nroQuest = 0
UserList(UserIndex).Quest.cantNpc = 0
UserList(UserIndex).Quest.numNpc = 0
UserList(UserIndex).Quest.cantMatados = 0

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Has cancelado la Quest. Puedes realizar otra siempre que desees." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
End Sub

Public Sub EnviarInfoQuest(ByVal UserIndex As Integer)

Dim Criatura As String
Criatura = GetVar(DatPath & "NPCs-HOSTILES.dat", "NPC" & UserList(UserIndex).Quest.numNpc, "Name")

Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & FONTTYPE_INFO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Información Quest:" & FONTTYPE_INFON)
Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Criatura a matar: " & Criatura & FONTTYPE_INFO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Cantidad a matar: " & UserList(UserIndex).Quest.cantNpc & FONTTYPE_INFO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Cantidad de matados: " & UserList(UserIndex).Quest.cantMatados & FONTTYPE_INFO)
Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Acercate al Quester y tipea /FINQUEST para recibir tu recompensa. O tipea /CANCELARQUEST si deseas realizar otra." & FONTTYPE_INFO)

End Sub

