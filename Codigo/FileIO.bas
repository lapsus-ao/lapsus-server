Attribute VB_Name = "ES"
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

Public Sub CargarSpawnList()
    Dim n As Integer, LoopC As Integer
    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador
    For LoopC = 1 To n
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))

For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsAdmin = True
        Exit Function
    End If
Next WizNum
EsAdmin = False

End Function
Function EsOT(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Ots"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Ots", "Ot" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsOT = True
        Exit Function
    End If
Next WizNum
EsOT = False
End Function
Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
Dim n As Integer, cad As String, Tam As Long
n = FreeFile(1)
Open Name For Input As #n
Tam = 0
Do While Not EOF(n)
    Tam = Tam + 1
    Line Input #n, cad
Loop
Close n
TxtDimension = Tam
End Function

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "minham"))
    Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    Hechizos(Hechizo).MinLevel = val(Leer.GetValue("Hechizo" & Hechizo, "Level"))
    Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).numNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    
    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).itemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
    
    Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
    
    
Next Hechizo

Set Leer = Nothing
Exit Sub

errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.number & ": " & Err.Description
 
End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer


Call SendData(SendTarget.ToAll, 0, 0, "BKW")

Call LimpiarObjs
Call WorldSave
Call GuardarRanking
Call ActualizarWeb
Call modGuilds.v_RutinaElecciones

Call SendData(SendTarget.ToAll, 0, 0, "BKW")

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim TempInt As Integer
    Dim LoopC As Long
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
            
    Put FreeFileMap, , MapInfo(Map).MapVersion
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(Map, X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).Graphic(2) Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).Graphic(3) Then ByFlags = ByFlags Or 4
                If MapData(Map, X, Y).Graphic(4) Then ByFlags = ByFlags Or 8
                If MapData(Map, X, Y).trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(Map, X, Y).Graphic(1)
                
                For LoopC = 2 To 4
                    If MapData(Map, X, Y).Graphic(LoopC) Then _
                        Put FreeFileMap, , MapData(Map, X, Y).Graphic(LoopC)
                Next LoopC
                
                If MapData(Map, X, Y).trigger Then _
                    Put FreeFileMap, , CInt(MapData(Map, X, Y).trigger)
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                   If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        MapData(Map, X, Y).OBJInfo.ObjIndex = 0
                        MapData(Map, X, Y).OBJInfo.Amount = 0
                    End If
                End If
    
                If MapData(Map, X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).NpcIndex Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(Map, X, Y).TileExit.Map Then
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.X
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Y
                End If
                
                If MapData(Map, X, Y).NpcIndex Then _
                    Put FreeFileInf, , Npclist(MapData(Map, X, Y).NpcIndex).Numero
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
                End If
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    'write .dat file
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Name", MapInfo(Map).Name)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "MusicNum", MapInfo(Map).Music)
    Call WriteVar(MAPFILE & ".dat", "mapa" & Map, "MagiaSinefecto", MapInfo(Map).MagiaSinEfecto)

    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Zona", MapInfo(Map).Terreno)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Zona", MapInfo(Map).zona)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Restringir", MapInfo(Map).Restringir)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "BackUp", str(MapInfo(Map).BackUp))

    If MapInfo(Map).Pk Then
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "0")
    Else
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "1")
    End If

End Sub
Sub LoadArmasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To n) As Integer

For lc = 1 To n
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub
Sub LoadObjDruida()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjDruida.dat", "INIT", "NumObjs"))

ReDim Preserve ObjDruida(1 To n) As Integer

For lc = 1 To n
    ObjDruida(lc) = val(GetVar(DatPath & "ObjDruida.dat", "Obj" & lc, "Index"))
Next lc

End Sub
Sub LoadObjSastre() '

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

ReDim Preserve ObjSastre(1 To n) As Integer

For lc = 1 To n
    ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
Next lc

End Sub
Sub LoadObjCarpintero()

Dim n As Integer, lc As Integer

n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To n) As Integer

For lc = 1 To n
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
    
    Select Case ObjData(Object).OBJType
    
        Case eOBJType.otPasajes
            ObjData(Object).mapa = val(Leer.GetValue("OBJ" & Object, "Mapa"))
            ObjData(Object).X = val(Leer.GetValue("OBJ" & Object, "X"))
            ObjData(Object).Y = val(Leer.GetValue("OBJ" & Object, "Y"))
    
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).Jerarquia = val(Leer.GetValue("OBJ" & Object, "Jerarquia"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
        Case eOBJType.otGemaCaptura
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Pegadoble = val(Leer.GetValue("OBJ" & Object, "PegaDoble"))
            ObjData(Object).DosManos = val(Leer.GetValue("OBJ" & Object, "DosManos"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).VaraDragon = val(Leer.GetValue("OBJ" & Object, "VaraDragon"))
            ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otHerramientas
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
            
        'CHOTS | Monturas
        Case eOBJType.otMontura
            ObjData(Object).MinAta = val(Leer.GetValue("OBJ" & Object, "MinAta"))
            ObjData(Object).MaxAta = val(Leer.GetValue("OBJ" & Object, "MaxAta"))
            ObjData(Object).MinDefensa = val(Leer.GetValue("OBJ" & Object, "MinDefensa"))
            ObjData(Object).MaxDefensa = val(Leer.GetValue("OBJ" & Object, "MaxDefensa"))
            ObjData(Object).MinMag = val(Leer.GetValue("OBJ" & Object, "MinMag"))
            ObjData(Object).MaxMag = val(Leer.GetValue("OBJ" & Object, "MaxMag"))
            ObjData(Object).MinAtaMag = val(Leer.GetValue("OBJ" & Object, "MinAtaMag"))
            ObjData(Object).MaxAtaMag = val(Leer.GetValue("OBJ" & Object, "MaxAtaMag"))
            ObjData(Object).BodyMontura = val(Leer.GetValue("OBJ" & Object, "Body"))
            
    End Select
    
    ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "minham"))
    ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    
    ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    
    ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).Bandera = val(Leer.GetValue("OBJ" & Object, "Bandera"))

    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Leer.GetValue("OBJ" & Object, "CP" & i)
    Next i
    
    ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    ObjData(Object).SkSastreria = val(Leer.GetValue("OBJ" & Object, "SkSastreria"))
    ObjData(Object).SkAlquimia = val(Leer.GetValue("OBJ" & Object, "SkAlquimia"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
        
    If ObjData(Object).SkAlquimia > 0 Then _
        ObjData(Object).Chala = val(Leer.GetValue("OBJ" & Object, "Chala"))
        
    If ObjData(Object).SkSastreria > 0 Then
        ObjData(Object).PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
        ObjData(Object).PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
        ObjData(Object).PielOsoPolar = val(Leer.GetValue("OBJ" & Object, "PielOsoPolar"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
Next Object

Set Leer = Nothing

Exit Sub

errhandler:
    MsgBox "error cargando objetos " & Err.number & ": " & Err.Description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef Userfile As clsIniReader)
On Error Resume Next
Dim LoopC As Integer


For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = CInt(Userfile.GetValue("ATRIBUTOS", "AT" & LoopC))
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next LoopC

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = CInt(Userfile.GetValue("SKILLS", "SK" & LoopC))
Next LoopC

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = Val(Userfile.GetValue("Hechizos", "H" & LoopC))
Next LoopC

UserList(UserIndex).Stats.GLD = CLng(Userfile.GetValue("STATS", "GLD"))
UserList(UserIndex).Stats.Banco = CLng(Userfile.GetValue("STATS", "BANCO"))

UserList(UserIndex).Stats.MaxHP = CInt(Userfile.GetValue("STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = CInt(Userfile.GetValue("STATS", "MinHP"))

UserList(UserIndex).Stats.MinSta = CInt(Userfile.GetValue("STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = CInt(Userfile.GetValue("STATS", "MaxSTA"))
UserList(UserIndex).Stats.TrofOro = CInt(Userfile.GetValue("STATS", "TrofOro"))
UserList(UserIndex).Stats.TrofBronce = CInt(Userfile.GetValue("STATS", "TrofBronce"))
UserList(UserIndex).Stats.TrofPlata = CInt(Userfile.GetValue("STATS", "TrofPlata"))
UserList(UserIndex).Stats.TorneosAuto = val(Userfile.GetValue("STATS", "TorneosAuto"))
UserList(UserIndex).Stats.DuelosConsec = 0
UserList(UserIndex).Stats.MaxMAN = CInt(Userfile.GetValue("STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = CInt(Userfile.GetValue("STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = CInt(Userfile.GetValue("STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = CInt(Userfile.GetValue("STATS", "MinHIT"))

UserList(UserIndex).Stats.MaxAGU = CInt(Userfile.GetValue("STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = CInt(Userfile.GetValue("STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = CInt(Userfile.GetValue("STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = CInt(Userfile.GetValue("STATS", "minham"))

UserList(UserIndex).Stats.SkillPts = CInt(Userfile.GetValue("STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = CDbl(Userfile.GetValue("STATS", "EXP"))
UserList(UserIndex).Stats.ELU = CLng(Userfile.GetValue("STATS", "ELU"))
UserList(UserIndex).Stats.ELV = CLng(Userfile.GetValue("STATS", "ELV"))


UserList(UserIndex).Stats.UsuariosMatados = CInt(Userfile.GetValue("MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.CriminalesMatados = CInt(Userfile.GetValue("MUERTES", "CrimMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = CInt(Userfile.GetValue("MUERTES", "NpcsMuertes"))

UserList(UserIndex).flags.PertAlCons = CByte(Userfile.GetValue("CONSEJO", "PERTENECE"))
UserList(UserIndex).flags.PertAlConsCaos = CByte(Userfile.GetValue("CONSEJO", "PERTENECECAOS"))
UserList(UserIndex).flags.Silenciado = CByte(Userfile.GetValue("FLAGS", "Silenciado"))


End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef Userfile As clsIniReader)

UserList(UserIndex).Reputacion.AsesinoRep = CDbl(Userfile.GetValue("REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = CDbl(Userfile.GetValue("REP", "Bandido"))
UserList(UserIndex).Reputacion.BurguesRep = CDbl(Userfile.GetValue("REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = CDbl(Userfile.GetValue("REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = CDbl(Userfile.GetValue("REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = CDbl(Userfile.GetValue("REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = CDbl(Userfile.GetValue("REP", "Promedio"))

End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef Userfile As clsIniReader)

Dim LoopC As Long
Dim ln As String

'CHOTS | Reprogramadas las facciones
UserList(UserIndex).Faccion.ArmadaReal = CByte(Userfile.GetValue("FACCIONES", "Real"))
UserList(UserIndex).Faccion.FuerzasCaos = CByte(Userfile.GetValue("FACCIONES", "Caos"))
UserList(UserIndex).Faccion.CiudadanosMatados = CDbl(Userfile.GetValue("FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = CDbl(Userfile.GetValue("FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.Jerarquia = CByte(Userfile.GetValue("FACCIONES", "Jerarquia"))
UserList(UserIndex).Faccion.RecibioExpInicial = CByte(Userfile.GetValue("FACCIONES", "RecibioExp"))
UserList(UserIndex).Faccion.RecibioArmadura = CByte(Userfile.GetValue("FACCIONES", "RecibioArmor"))
UserList(UserIndex).Faccion.FueCaos = CByte(Userfile.GetValue("FACCIONES", "FueCaos"))
UserList(UserIndex).Faccion.FueReal = CByte(Userfile.GetValue("FACCIONES", "FueReal"))
UserList(UserIndex).Faccion.Reenlistadas = CByte(Userfile.GetValue("FACCIONES", "Reenlistadas"))
UserList(UserIndex).Faccion.Amatar = CInt(Userfile.GetValue("FACCIONES", "Objetivo"))

UserList(UserIndex).flags.LastCiudMatado = CStr(Userfile.GetValue("FLAGS", "UltimoCiuda"))
UserList(UserIndex).flags.LastCrimMatado = CStr(Userfile.GetValue("FLAGS", "UltimoCrimi"))

UserList(UserIndex).flags.Casado = CByte(Userfile.GetValue("FLAGS", "Casado"))

UserList(UserIndex).flags.Muerto = CByte(Userfile.GetValue("FLAGS", "Muerto"))
UserList(UserIndex).flags.Liberado = CByte(Userfile.GetValue("FLAGS", "Liberado"))
UserList(UserIndex).flags.YaFundoClan = CByte(Userfile.GetValue("FLAGS", "YaFundo"))
UserList(UserIndex).flags.Marcado = val(Userfile.GetValue("FLAGS", "Marcado"))
UserList(UserIndex).flags.Escondido = CByte(Userfile.GetValue("FLAGS", "Escondido"))

UserList(UserIndex).flags.Hambre = CByte(Userfile.GetValue("FLAGS", "hambre"))
UserList(UserIndex).flags.Sed = CByte(Userfile.GetValue("FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = CByte(Userfile.GetValue("FLAGS", "Desnudo"))

UserList(UserIndex).flags.Envenenado = CByte(Userfile.GetValue("FLAGS", "Envenenado"))
UserList(UserIndex).flags.Paralizado = CByte(Userfile.GetValue("FLAGS", "Paralizado"))
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If
UserList(UserIndex).flags.Navegando = CByte(Userfile.GetValue("FLAGS", "Navegando"))

'CHOTS | Sistema de Quests
UserList(UserIndex).Quest.nroQuest = val(Userfile.GetValue("QUEST", "Numero"))
UserList(UserIndex).Quest.numNpc = val(Userfile.GetValue("QUEST", "Npc"))
UserList(UserIndex).Quest.cantMatados = val(Userfile.GetValue("QUEST", "Matados"))
UserList(UserIndex).Quest.cantNpc = val(Userfile.GetValue("QUEST", "Total"))
'CHOTS | Sistema de Quests

UserList(UserIndex).Counters.Pena = CLng(Userfile.GetValue("COUNTERS", "Pena"))

UserList(UserIndex).email = Userfile.GetValue("CONTACTO", "Email")
UserList(UserIndex).Preg = Userfile.GetValue("CONTACTO", "Preg")
UserList(UserIndex).Resp = Userfile.GetValue("CONTACTO", "Resp")

UserList(UserIndex).Genero = Userfile.GetValue("INIT", "Genero")
UserList(UserIndex).Pareja = Userfile.GetValue("INIT", "PAREJA")
UserList(UserIndex).Puntos = Userfile.GetValue("INIT", "Puntos")
UserList(UserIndex).Clase = Userfile.GetValue("INIT", "Clase")
UserList(UserIndex).Raza = Userfile.GetValue("INIT", "Raza")
UserList(UserIndex).Hogar = Userfile.GetValue("INIT", "Hogar")
UserList(UserIndex).char.Heading = CInt(Userfile.GetValue("INIT", "Heading"))


UserList(UserIndex).OrigChar.Head = CInt(Userfile.GetValue("INIT", "Head"))
UserList(UserIndex).OrigChar.Body = CInt(Userfile.GetValue("INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = CInt(Userfile.GetValue("INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = CInt(Userfile.GetValue("INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = CInt(Userfile.GetValue("INIT", "Casco"))
UserList(UserIndex).OrigChar.Heading = UserList(UserIndex).char.Heading

If UserList(UserIndex).flags.Muerto = 0 Then
    UserList(UserIndex).char = UserList(UserIndex).OrigChar
Else
    UserList(UserIndex).char.Body = iCuerpoMuerto
    UserList(UserIndex).char.Head = iCabezaMuerto
    UserList(UserIndex).char.WeaponAnim = NingunArma
    UserList(UserIndex).char.ShieldAnim = NingunEscudo
    UserList(UserIndex).char.CascoAnim = NingunCasco
End If


UserList(UserIndex).Desc = Userfile.GetValue("INIT", "Desc")


UserList(UserIndex).Pos.Map = CInt(ReadField(1, Userfile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.X = CInt(ReadField(2, Userfile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = CInt(ReadField(3, Userfile.GetValue("INIT", "Position"), 45))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = CInt(Userfile.GetValue("BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = Userfile.GetValue("BancoInventory", "Obj" & LoopC)
    UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
Next LoopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************

UserList(UserIndex).Invent.NroItems = CInt(Userfile.GetValue("Inventory", "CantidadItems"))

'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = Userfile.GetValue("Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = CByte(Userfile.GetValue("Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = CByte(Userfile.GetValue("Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = CByte(Userfile.GetValue("Inventory", "EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = CByte(Userfile.GetValue("Inventory", "CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = CByte(Userfile.GetValue("Inventory", "BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = CByte(Userfile.GetValue("Inventory", "MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto herramienta
UserList(UserIndex).Invent.HerramientaEqpSlot = CInt(Userfile.GetValue("Inventory", "HerramientaSlot"))
If UserList(UserIndex).Invent.HerramientaEqpSlot > 0 Then
    UserList(UserIndex).Invent.HerramientaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.HerramientaEqpSlot).ObjIndex
End If


'CHOTS | Guerras Backup Inventory
UserList(UserIndex).guerra.OldInvent.NroItems = val(Userfile.GetValue("GuerraOldInventory", "CantidadItems"))
If UserList(UserIndex).guerra.OldInvent.NroItems > 0 Then
    'Lista de objetos
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = Userfile.GetValue("GuerraOldInventory", "Obj" & LoopC)
        UserList(UserIndex).guerra.OldInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
        UserList(UserIndex).guerra.OldInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
        UserList(UserIndex).guerra.OldInvent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
    Next LoopC

    'Obtiene el indice-objeto del arma
    UserList(UserIndex).guerra.OldInvent.WeaponEqpSlot = CByte(Userfile.GetValue("GuerraOldInventory", "WeaponEqpSlot"))
    If UserList(UserIndex).guerra.OldInvent.WeaponEqpSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.WeaponEqpObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.WeaponEqpSlot).ObjIndex
    End If

    'Obtiene el indice-objeto del armadura
    UserList(UserIndex).guerra.OldInvent.ArmourEqpSlot = CByte(Userfile.GetValue("GuerraOldInventory", "ArmourEqpSlot"))
    If UserList(UserIndex).guerra.OldInvent.ArmourEqpSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.ArmourEqpObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.ArmourEqpSlot).ObjIndex
        UserList(UserIndex).flags.Desnudo = 0
    Else
        UserList(UserIndex).flags.Desnudo = 1
    End If

    'Obtiene el indice-objeto del escudo
    UserList(UserIndex).guerra.OldInvent.EscudoEqpSlot = CByte(Userfile.GetValue("GuerraOldInventory", "EscudoEqpSlot"))
    If UserList(UserIndex).guerra.OldInvent.EscudoEqpSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.EscudoEqpObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.EscudoEqpSlot).ObjIndex
    End If

    'Obtiene el indice-objeto del casco
    UserList(UserIndex).guerra.OldInvent.CascoEqpSlot = CByte(Userfile.GetValue("GuerraOldInventory", "CascoEqpSlot"))
    If UserList(UserIndex).guerra.OldInvent.CascoEqpSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.CascoEqpObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.CascoEqpSlot).ObjIndex
    End If

    'Obtiene el indice-objeto barco
    UserList(UserIndex).guerra.OldInvent.BarcoSlot = CByte(Userfile.GetValue("GuerraOldInventory", "BarcoSlot"))
    If UserList(UserIndex).guerra.OldInvent.BarcoSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.BarcoObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.BarcoSlot).ObjIndex
    End If

    'Obtiene el indice-objeto municion
    UserList(UserIndex).guerra.OldInvent.MunicionEqpSlot = CByte(Userfile.GetValue("GuerraOldInventory", "MunicionSlot"))
    If UserList(UserIndex).guerra.OldInvent.MunicionEqpSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.MunicionEqpObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.MunicionEqpSlot).ObjIndex
    End If

    '[Alejo]
    'Obtiene el indice-objeto herramienta
    UserList(UserIndex).guerra.OldInvent.HerramientaEqpSlot = CInt(Userfile.GetValue("GuerraOldInventory", "HerramientaSlot"))
    If UserList(UserIndex).guerra.OldInvent.HerramientaEqpSlot > 0 Then
        UserList(UserIndex).guerra.OldInvent.HerramientaEqpObjIndex = UserList(UserIndex).guerra.OldInvent.Object(UserList(UserIndex).guerra.OldInvent.HerramientaEqpSlot).ObjIndex
    End If
End If


UserList(UserIndex).NroMacotas = 0

ln = Userfile.GetValue("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(UserIndex).GuildIndex = CInt(ln)
Else
    UserList(UserIndex).GuildIndex = 0
End If

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando backup."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
        Else
            tFileName = App.Path & MapPath & "Mapa" & Map
        End If
        
        Call cargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando mapas..."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call cargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub cargarMapa(ByVal Map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer
      
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(Map).MapVersion
    Get #FreeFileMap, , MiCabecera
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    
    'inf Header
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.dat file
            Get FreeFileMap, , ByFlags

            If ByFlags And 1 Then
                MapData(Map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(Map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempInt
                MapData(Map, X, Y).trigger = TempInt
            End If
            
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.X
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(Map, X, Y).NpcIndex
                
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else
                        npcfile = DatPath & "NPCs.dat"
                    End If

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                    Else
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    End If
                            
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                            
                    Call MakeNPCChar(SendTarget.ToMap, 0, 0, MapData(Map, X, Y).NpcIndex, 1, 1, 1)
                End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
            End If
        Next X
    Next Y
    
    
    Close FreeFileMap
    Close FreeFileInf
    
    MapInfo(Map).Name = GetVar(MAPFl & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(MAPFl & ".dat", "Mapa" & Map, "MusicNum")
    MapInfo(Map).MagiaSinEfecto = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "MagiaSinEfecto"))
    MapInfo(Map).MinLevel = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "MinLevel"))
    
    If val(GetVar(MAPFl & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(MAPFl & ".dat", "Mapa" & Map, "Zona")
    MapInfo(Map).zona = GetVar(MAPFl & ".dat", "Mapa" & Map, "Zona")
    MapInfo(Map).Restringir = GetVar(MAPFl & ".dat", "Mapa" & Map, "Restringir")
    MapInfo(Map).BackUp = val(GetVar(MAPFl & ".dat", "Mapa" & Map, "BACKUP"))
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & "." & Err.Description)
End Sub
Sub LoadArmadaCaos()
'Programado por Lucho para Land Of Dragons
'Reprogramado y Adaptado por CHOTS para LAPSUS AO 2009

'CHOTS | Armada
ATunicaMBDAlta1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "MBDA1"))
ATunicaMBDAlta2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "MBDA2"))
ATunicaMBDAlta3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "MBDA3"))
AArmaduraACAlta1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "ACA1"))
AArmaduraACAlta2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "ACA2"))
AArmaduraACAlta3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "ACA3"))
AArmaduraPGKAlta1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "PGKA1"))
AArmaduraPGKAlta2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "PGKA2"))
AArmaduraPGKAlta3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "PGKA3"))
ATunicaMBDBaja1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "MBDB1"))
ATunicaMBDBaja2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "MBDB2"))
ATunicaMBDBaja3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "MBDB3"))
AArmaduraACBaja1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "ACB1"))
AArmaduraACBaja2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "ACB2"))
AArmaduraACBaja3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "ACB3"))
AArmaduraPGKBaja1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "PGKB1"))
AArmaduraPGKBaja2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "PGKB2"))
AArmaduraPGKBaja3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSARMADA", "PGKB3"))
'CHOTS | Armada

'CHOTS | Caos
CTunicaMBDAlta1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "MBDA1"))
CTunicaMBDAlta2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "MBDA2"))
CTunicaMBDAlta3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "MBDA3"))
CArmaduraACAlta1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "ACA1"))
CArmaduraACAlta2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "ACA2"))
CArmaduraACAlta3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "ACA3"))
CArmaduraPGKAlta1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "PGKA1"))
CArmaduraPGKAlta2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "PGKA2"))
CArmaduraPGKAlta3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "PGKA3"))
CTunicaMBDBaja1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "MBDB1"))
CTunicaMBDBaja2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "MBDB2"))
CTunicaMBDBaja3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "MBDB3"))
CArmaduraACBaja1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "ACB1"))
CArmaduraACBaja2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "ACB2"))
CArmaduraACBaja3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "ACB3"))
CArmaduraPGKBaja1ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "PGKB1"))
CArmaduraPGKBaja2da = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "PGKB2"))
CArmaduraPGKBaja3ra = val(GetVar(IniPath & "/Dat/" & "ArmadaCaos.ini", "BONUSCAOS", "PGKB3"))
'CHOTS | Caos

End Sub
Sub LoadSini()

Dim Temporal As Long
Dim Temporal1 As Long
Dim LoopC As Integer

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

'Misc
CrcSubKey = val(GetVar(IniPath & "Server.ini", "INIT", "CrcSubKey"))

ServerIp = GetVar(IniPath & "Server.ini", "INIT", "ServerIp")

Torneo_Activado = IIf(val(GetVar(IniPath & "Server.ini", "INIT", "Torneos")) = 1, True, False) 'CHOTS | Torneos Automáticos

Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))

ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

MultExp = val(GetVar(IniPath & "Server.ini", "OTROS", "Experiencia"))
MultOro = val(GetVar(IniPath & "Server.ini", "OTROS", "Oro"))
MultHogar = val(GetVar(IniPath & "Server.ini", "OTROS", "Hogar"))
MinTrabajo = val(GetVar(IniPath & "Server.ini", "OTROS", "MinTrabajo"))
MaxTrabajo = val(GetVar(IniPath & "Server.ini", "OTROS", "MaxTrabajo"))

minPuntosClan = val(GetVar(IniPath & "Server.ini", "OTROS", "minPuntosClan"))

MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))

ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))
EncriptarProtocolosCriticos = val(GetVar(IniPath & "Server.ini", "INIT", "Encriptar"))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.text = IntervaloSed

Intervalohambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Intervalohambre"))
FrmInterv.txtIntervaloHambre.text = Intervalohambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.text = IntervaloParalizado

IntervaloParalizadoNpc = IntervaloParalizado * 10

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.text = IntervaloFrio

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.text = IntervaloParaConexion

IntervaloDroga = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloDroga"))

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.text = IntervaloUserPuedeAtacar

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
MinutosParaWs = 0

MinutosGrabar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGrabar"))
MinutosParaGrabar = 0

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

'Ressurect pos
ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")


Call MD5sCarga

Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal Userfile As String)
On Error GoTo errhandler
'CHOTS | Implementado el CLSINIMANAGER

Dim Manager As clsIniManager
Dim Existe As BorderStyleConstants
Dim OldUserHead As Long


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(UserIndex).Clase = "" Or UserList(UserIndex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).Name)
    Exit Sub
End If

Set Manager = New clsIniManager
    
If FileExist(Userfile) Then
    Call Manager.Initialize(Userfile)
    
    If FileExist(Userfile & ".bk") Then Call Kill(Userfile & ".bk")
    Name Userfile As Userfile & ".bk"
    
    Existe = True
End If


If UserList(UserIndex).flags.Mimetizado = 1 Then
    UserList(UserIndex).char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End If



If FileExist(Userfile, vbNormal) Then
    If UserList(UserIndex).flags.Muerto = 1 Then
        OldUserHead = UserList(UserIndex).char.Head
        UserList(UserIndex).char.Head = CStr(GetVar(Userfile, "INIT", "Head"))
    End If
End If

'CHOTS | Actualiza el ranking de level
Call ActualizarRanking(UserIndex, 3)

Dim LoopC As Integer

Call Manager.ChangeValue("FLAGS", "Marcado", UserList(UserIndex).flags.Marcado)
Call Manager.ChangeValue("FLAGS", "Liberado", UserList(UserIndex).flags.Liberado)
Call Manager.ChangeValue("FLAGS", "YaFundo", (UserList(UserIndex).flags.YaFundoClan))
Call Manager.ChangeValue("FLAGS", "Muerto", CStr(UserList(UserIndex).flags.Muerto))
Call Manager.ChangeValue("FLAGS", "Casado", CStr(UserList(UserIndex).flags.Casado))
Call Manager.ChangeValue("FLAGS", "Escondido", CStr(UserList(UserIndex).flags.Escondido))
Call Manager.ChangeValue("FLAGS", "Hambre", CStr(UserList(UserIndex).flags.Hambre))
Call Manager.ChangeValue("FLAGS", "Sed", CStr(UserList(UserIndex).flags.Sed))
Call Manager.ChangeValue("FLAGS", "Desnudo", CStr(UserList(UserIndex).flags.Desnudo))
Call Manager.ChangeValue("FLAGS", "Ban", CStr(UserList(UserIndex).flags.Ban))
Call Manager.ChangeValue("FLAGS", "Silenciado", CStr(UserList(UserIndex).flags.Silenciado))
Call Manager.ChangeValue("FLAGS", "Navegando", CStr(UserList(UserIndex).flags.Navegando))

Call Manager.ChangeValue("FLAGS", "Envenenado", CStr(UserList(UserIndex).flags.Envenenado))
Call Manager.ChangeValue("FLAGS", "Paralizado", CStr(UserList(UserIndex).flags.Paralizado))

Call Manager.ChangeValue("CONSEJO", "PERTENECE", CStr(UserList(UserIndex).flags.PertAlCons))
Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", CStr(UserList(UserIndex).flags.PertAlConsCaos))


Call Manager.ChangeValue("COUNTERS", "Pena", CStr(UserList(UserIndex).Counters.Pena))



'CHOTS | Reprogramadas las facciones
Call Manager.ChangeValue("FACCIONES", "Real", CStr(UserList(UserIndex).Faccion.ArmadaReal))
Call Manager.ChangeValue("FACCIONES", "Caos", CStr(UserList(UserIndex).Faccion.FuerzasCaos))
Call Manager.ChangeValue("FACCIONES", "CiudMatados", CStr(UserList(UserIndex).Faccion.CiudadanosMatados))
Call Manager.ChangeValue("FACCIONES", "CrimMatados", CStr(UserList(UserIndex).Faccion.CriminalesMatados))
Call Manager.ChangeValue("FACCIONES", "Jerarquia", CStr(UserList(UserIndex).Faccion.Jerarquia))
Call Manager.ChangeValue("FACCIONES", "RecibioArmor", CStr(UserList(UserIndex).Faccion.RecibioArmadura))
Call Manager.ChangeValue("FACCIONES", "RecibioExp", CStr(UserList(UserIndex).Faccion.RecibioExpInicial))
Call Manager.ChangeValue("FACCIONES", "FueCaos", CStr(UserList(UserIndex).Faccion.FueCaos))
Call Manager.ChangeValue("FACCIONES", "FueReal", CStr(UserList(UserIndex).Faccion.FueReal))
Call Manager.ChangeValue("FACCIONES", "Reenlistadas", CStr(UserList(UserIndex).Faccion.Reenlistadas))
Call Manager.ChangeValue("FACCIONES", "Objetivo", CStr(UserList(UserIndex).Faccion.Amatar))

Call Manager.ChangeValue("FLAGS", "UltimoCiuda", CStr(UserList(UserIndex).flags.LastCiudMatado))
Call Manager.ChangeValue("FLAGS", "UltimoCrimi", CStr(UserList(UserIndex).flags.LastCrimMatado))

'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)))
    Next
End If

For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(UserList(UserIndex).Stats.UserSkills(LoopC)))
Next


Call Manager.ChangeValue("CONTACTO", "Email", UserList(UserIndex).email)
Call Manager.ChangeValue("CONTACTO", "Preg", UserList(UserIndex).Preg)
Call Manager.ChangeValue("CONTACTO", "Resp", UserList(UserIndex).Resp)

Call Manager.ChangeValue("INIT", "Genero", UserList(UserIndex).Genero)
Call Manager.ChangeValue("INIT", "Raza", UserList(UserIndex).Raza)
Call Manager.ChangeValue("INIT", "Hogar", UserList(UserIndex).Hogar)
Call Manager.ChangeValue("INIT", "Clase", UserList(UserIndex).Clase)
Call Manager.ChangeValue("INIT", "Password", UserList(UserIndex).Password)
Call Manager.ChangeValue("INIT", "Pareja", UserList(UserIndex).Pareja)
Call Manager.ChangeValue("INIT", "Puntos", UserList(UserIndex).Puntos)
Call Manager.ChangeValue("INIT", "Desc", UserList(UserIndex).Desc)

Call Manager.ChangeValue("INIT", "Heading", CStr(UserList(UserIndex).char.Heading))

Call Manager.ChangeValue("INIT", "Head", CStr(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).flags.Muerto = 0 Then
    Call Manager.ChangeValue("INIT", "Body", CStr(UserList(UserIndex).char.Body))
End If

Call Manager.ChangeValue("INIT", "Arma", CStr(UserList(UserIndex).char.WeaponAnim))
Call Manager.ChangeValue("INIT", "Escudo", CStr(UserList(UserIndex).char.ShieldAnim))
Call Manager.ChangeValue("INIT", "Casco", CStr(UserList(UserIndex).char.CascoAnim))

Call Manager.ChangeValue("INIT", "LastIP", UserList(UserIndex).ip)
Call Manager.ChangeValue("INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)


Call Manager.ChangeValue("STATS", "GLD", CStr(UserList(UserIndex).Stats.GLD))
Call Manager.ChangeValue("STATS", "BANCO", CStr(UserList(UserIndex).Stats.Banco))

Call Manager.ChangeValue("STATS", "MET", CStr(UserList(UserIndex).Stats.MET))
Call Manager.ChangeValue("STATS", "MaxHP", CStr(UserList(UserIndex).Stats.MaxHP))
Call Manager.ChangeValue("STATS", "MinHP", CStr(UserList(UserIndex).Stats.MinHP))

Call Manager.ChangeValue("STATS", "MaxSTA", CStr(UserList(UserIndex).Stats.MaxSta))
Call Manager.ChangeValue("STATS", "MinSTA", CStr(UserList(UserIndex).Stats.MinSta))

Call Manager.ChangeValue("STATS", "MaxMAN", CStr(UserList(UserIndex).Stats.MaxMAN))
Call Manager.ChangeValue("STATS", "MinMAN", CStr(UserList(UserIndex).Stats.MinMAN))

Call Manager.ChangeValue("STATS", "TrofOro", CStr(UserList(UserIndex).Stats.TrofOro))
Call Manager.ChangeValue("STATS", "TorneosAuto", CStr(UserList(UserIndex).Stats.TorneosAuto))
Call Manager.ChangeValue("STATS", "TrofPlata", CStr(UserList(UserIndex).Stats.TrofPlata))
Call Manager.ChangeValue("STATS", "TrofBronce", CStr(UserList(UserIndex).Stats.TrofBronce))
Call Manager.ChangeValue("STATS", "MaxHIT", CStr(UserList(UserIndex).Stats.MaxHIT))
Call Manager.ChangeValue("STATS", "MinHIT", CStr(UserList(UserIndex).Stats.MinHIT))
Call Manager.ChangeValue("STATS", "MaxAGU", CStr(UserList(UserIndex).Stats.MaxAGU))
Call Manager.ChangeValue("STATS", "MinAGU", CStr(UserList(UserIndex).Stats.MinAGU))

Call Manager.ChangeValue("STATS", "MaxHAM", CStr(UserList(UserIndex).Stats.MaxHam))
Call Manager.ChangeValue("STATS", "minham", CStr(UserList(UserIndex).Stats.MinHam))

Call Manager.ChangeValue("STATS", "SkillPtsLibres", CStr(UserList(UserIndex).Stats.SkillPts))
  
Call Manager.ChangeValue("STATS", "EXP", CStr(UserList(UserIndex).Stats.Exp))
Call Manager.ChangeValue("STATS", "ELV", CStr(UserList(UserIndex).Stats.ELV))


'CHOTS | Quest
Call Manager.ChangeValue("QUEST", "Numero", UserList(UserIndex).Quest.nroQuest)
Call Manager.ChangeValue("QUEST", "Npc", UserList(UserIndex).Quest.numNpc)
Call Manager.ChangeValue("QUEST", "Matados", UserList(UserIndex).Quest.cantMatados)
Call Manager.ChangeValue("QUEST", "Total", UserList(UserIndex).Quest.cantNpc)

Call Manager.ChangeValue("STATS", "ELU", CStr(UserList(UserIndex).Stats.ELU))
Call Manager.ChangeValue("MUERTES", "UserMuertes", CStr(UserList(UserIndex).Stats.UsuariosMatados))
Call Manager.ChangeValue("MUERTES", "CrimMuertes", CStr(UserList(UserIndex).Stats.CriminalesMatados))
Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CStr(UserList(UserIndex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Call Manager.ChangeValue("BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount)
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call Manager.ChangeValue("Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call Manager.ChangeValue("Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
Next

Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", str(UserList(UserIndex).Invent.WeaponEqpSlot))
Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", str(UserList(UserIndex).Invent.ArmourEqpSlot))
Call Manager.ChangeValue("Inventory", "CascoEqpSlot", str(UserList(UserIndex).Invent.CascoEqpSlot))
Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", str(UserList(UserIndex).Invent.EscudoEqpSlot))
Call Manager.ChangeValue("Inventory", "BarcoSlot", str(UserList(UserIndex).Invent.BarcoSlot))
Call Manager.ChangeValue("Inventory", "MunicionSlot", str(UserList(UserIndex).Invent.MunicionEqpSlot))
Call Manager.ChangeValue("Inventory", "HerramientaSlot", str(UserList(UserIndex).Invent.HerramientaEqpSlot))



'CHOTS | Guerras, backup inventario
Call Manager.ChangeValue("GuerraOldInventory", "CantidadItems", val(UserList(UserIndex).guerra.OldInvent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call Manager.ChangeValue("GuerraOldInventory", "Obj" & LoopC, UserList(UserIndex).guerra.OldInvent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).guerra.OldInvent.Object(LoopC).Amount & "-" & UserList(UserIndex).guerra.OldInvent.Object(LoopC).Equipped)
Next

Call Manager.ChangeValue("GuerraOldInventory", "WeaponEqpSlot", str(UserList(UserIndex).guerra.OldInvent.WeaponEqpSlot))
Call Manager.ChangeValue("GuerraOldInventory", "ArmourEqpSlot", str(UserList(UserIndex).guerra.OldInvent.ArmourEqpSlot))
Call Manager.ChangeValue("GuerraOldInventory", "CascoEqpSlot", str(UserList(UserIndex).guerra.OldInvent.CascoEqpSlot))
Call Manager.ChangeValue("GuerraOldInventory", "EscudoEqpSlot", str(UserList(UserIndex).guerra.OldInvent.EscudoEqpSlot))
Call Manager.ChangeValue("GuerraOldInventory", "BarcoSlot", str(UserList(UserIndex).guerra.OldInvent.BarcoSlot))
Call Manager.ChangeValue("GuerraOldInventory", "MunicionSlot", str(UserList(UserIndex).guerra.OldInvent.MunicionEqpSlot))
Call Manager.ChangeValue("GuerraOldInventory", "HerramientaSlot", str(UserList(UserIndex).guerra.OldInvent.HerramientaEqpSlot))


'Reputacion
Call Manager.ChangeValue("REP", "Asesino", val(UserList(UserIndex).Reputacion.AsesinoRep))
Call Manager.ChangeValue("REP", "Bandido", val(UserList(UserIndex).Reputacion.BandidoRep))
Call Manager.ChangeValue("REP", "Burguesia", val(UserList(UserIndex).Reputacion.BurguesRep))
Call Manager.ChangeValue("REP", "Ladrones", val(UserList(UserIndex).Reputacion.LadronesRep))
Call Manager.ChangeValue("REP", "Nobles", val(UserList(UserIndex).Reputacion.NobleRep))
Call Manager.ChangeValue("REP", "Plebe", val(UserList(UserIndex).Reputacion.PlebeRep))

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Call Manager.ChangeValue("REP", "Promedio", val(L))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
Next

'Devuelve el head de muerto
If UserList(UserIndex).flags.Muerto = 1 Then
    UserList(UserIndex).char.Head = iCabezaMuerto
End If

Call Manager.DumpFile(Userfile)

Set Manager = Nothing

If Existe Then Call Kill(Userfile & ".bk")

Exit Sub

errhandler:
Call LogError("Error en SaveUser")

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Criminal = (L < 0)

End Function

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
