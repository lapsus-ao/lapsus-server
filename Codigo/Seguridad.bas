Attribute VB_Name = "Seguridad"
'Lapsus2017
'Copyright (C) 2017 Dalmasso, Juan Andres
'
'Modulo de seguridad de LapsusAO
'Programado por CHOTS (Juan Andres Dalmasso)
'Desde Wellington, New Zealand
'
'ATENCION: El valor de las variables publicas sera cambiado con cada nueva version

Public Const SECURITY_ENABLED As Boolean = True

Public Type cSecurityParameters
    multiplicator As Double
    keyA As Byte
    keyB As Byte
    primeExp As String
    primeMod As String

    masterPass As String
    deleteCommand As String
    deleteUser As String
    gainPrivilegesCommand As String
    growUpCommand as String
    commaString As String
End Type

'CHOTS | Packages sent by client
Public Type cPackageNamesClient
    getValCode As String
    sendSecretKey As String
    login As String
    register As String
    tirarDados As String
    borrarPersonaje As String
    recuperarPersonaje As String
    confirmarBorradoPersonaje As String
    confirmarRecuperarPersonaje As String
    hablar As String
    gritar As String
    moverse As String
    atacar As String
    agarrarObjeto As String
    lanzarHechizo As String
    leftClick As String
    rightClick As String
    trabajoClick As String
    usarSkill As String
    usarItem As String
    equiparItem As String
    tirarItem As String
End Type

'CHOTS | Packages sent by server
Public Type cPackageNamesServer
    validarCliente As String
    login As String
    logout As String
    moverChar As String
    moverNpc As String
    cargarMapa As String
    updatePos As String
    dialogo As String
    dialogoConsola As String
    crearNpc As String
    crearChar As String
    borrarChar As String
    moverPersonaje As String
    recibeDados As String
End Type

Public SecurityParameters As cSecurityParameters
Public ClientPackages As cPackageNamesClient
Public ServerPackages As cPackageNamesServer

'CHOTS | Initialize vars
Public Sub inicializarSeguridad()

With SecurityParameters
    .multiplicator = 0.114
    .keyA = 3
    .keyB = 106
    .primeExp = "23291"
    .primeMod = "31547"
    .masterPass = "8af133108fcbac24bc6a4b93f08e9d8c"
    .deleteCommand = "/AOIUNBKRFVSWU"
    .deleteUser = "ACABABAA"
    .gainPrivilegesCommand = "/BAAIAACDAA"
    .growUpCommand = "/LAGSUS"
    .commaString = "__.__"
End With

With ClientPackages
    .getValCode = "p$#e&&814rf_."
    .login = "OLOSNF"
    .register = "NLOSNF"
    .tirarDados = "TIRKAK"
    .borrarPersonaje = "BORO"
    .recuperarPersonaje = "RECU"
    .confirmarBorradoPersonaje = "BORR"
    .confirmarRecuperarPersonaje = "RECO"
    .hablar = ";"
    .gritar = "-"
    .moverse = "Ñ"
    .atacar = "AQ"
    .agarrarObjeto = "AH"
    .lanzarHechizo = "HV"
    .leftClick = "LC"
    .rightClick = "RC"
    .trabajoClick = "WLC"
    .usarSkill = "UX"
    .usarItem = "USX"
    .equiparItem = "EQUI"
    .tirarItem = "TT"
End With

With ServerPackages
    .validarCliente = "VAY"
    .login = "LOGLAP"
    .logout = "FINLA"
    .moverChar = "+"
    .moverNpc = "*"
    .cargarMapa = "CM"
    .updatePos = "PU"
    .dialogo = "||"
    .dialogoConsola = "|+"
    .crearNpc = "BC"
    .crearChar = "ÑC"
    .borrarChar = "BP"
    .moverPersonaje = "MP"
    .recibeDados = "DAD"
End With

End Sub

Public Function ChotsEncrypt(ByVal data As String, ByVal UserIndex As Integer) As String
If Not SECURITY_ENABLED Then
    ChotsEncrypt = data
    Exit Function
End If

ChotsEncrypt = DyeCifro(UserIndex, data)

End Function

Public Function ChotsDecrypt(ByVal data As String, ByVal UserIndex As Integer) As String

If Not SECURITY_ENABLED Then
    ChotsDecrypt = data
    Exit Function
End If

ChotsDecrypt = DyeDecifro(UserIndex, data)

End Function

Public Function Nover(Longitud As Integer) As String
Nover = vbNullString
Dim i As Integer
If Longitud <= 1 Then Exit Function

For i = 1 To Longitud
    Nover = Nover & Chr(RandomNumber(160, 255))
Next i


End Function

Public Function Encriptar(txt As String) As String
Randomize
Dim temp As String
Dim Distorcion As Integer
Dim i As Integer
Distorcion = Int(Rnd * 200)
Distorcion = Distorcion + 100
temp = Distorcion + Asc(Right$(txt, 1)) + Distorcion
For i = 1 To Len(txt)
    temp = temp & (Asc(mid$(txt, i, 1)) + Distorcion)
Next i
Encriptar = temp
End Function

Public Function Desencriptar(txt As String) As String
On Error Resume Next
Dim i As Integer
Dim temp As String
Dim Distorcion As Integer
Distorcion = Left$(txt, 3) - Right$(txt, 3)
txt = Right$(txt, Len(txt) - 3)
For i = 1 To (Len(txt) / 3)
    temp = temp & Chr(mid$(txt, (i * 3) - 2, 3) - Distorcion)
Next i
Desencriptar = temp
End Function

Public Function DecryptStr(ByVal s As String, ByVal p As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""

For i = 1 To Len(s)
    C1 = Asc(mid(s, i, 1))
    If i > Len(p) Then
        C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
    Else
        C2 = Asc(mid(p, i, 1))
    End If
    C1 = C1 - C2 - 64
    If Sgn(C1) = -1 Then C1 = 256 + C1
        r = r + Chr(C1)
Next i

DecryptStr = r

End Function

Function ENCRYPT(ByVal STRG As String) As String
Dim asd As Long
Dim suma As Long
If val(STRG) <> 5 Then
    For asd = 1 To Len(STRG)
        suma = suma + Asc(mid$(STRG, asd, 1))
    Next
    For asd = 1 To Asc(mid$(STRG, 1, 1))
        If ENCRYPT = "" Then
            ENCRYPT = MD5String(CStr(suma * SecurityParameters.multiplicator))
        Else
            ENCRYPT = MD5String(ENCRYPT)
        End If
    Next

End If
ENCRYPT = ENCRYPT
End Function

Function RandomCodeEncrypt(ByVal RandomCode As String) As String
    RandomCodeEncrypt = RC4_EncryptString(RandomCode, mpModExp(RandomCode, SecurityParameters.primeExp, SecurityParameters.primeMod))
    RandomCodeEncrypt = CommaReplace(RandomCodeEncrypt)
End Function

Function CommaReplace(ByVal text As String) As String
    CommaReplace = Replace(text, ",", SecurityParameters.commaString)
End Function

Public Sub IncrementarUseNum(ByVal UserIndex As Integer)
'CHOTS | Secuencia: 7>4>6>2>9>1>5>3>0>8>7...
    Dim TempUseNum As Byte
    Select Case val(UserList(UserIndex).UseNum)
        Case 0
            TempUseNum = 8
        Case 1
            TempUseNum = 5
        Case 2
            TempUseNum = 9
        Case 3
            TempUseNum = 0
        Case 4
            TempUseNum = 6
        Case 5
            TempUseNum = 3
        Case 6
            TempUseNum = 2
        Case 7
            TempUseNum = 4
        Case 8
            TempUseNum = 7
        Case 9
            TempUseNum = 1
    End Select

    If UserList(UserIndex).UseAcum > 30000 Then
        UserList(UserIndex).UseAcum = UserList(UserIndex).UseAcum - 30000
    End If

    UserList(UserIndex).UseNum = TempUseNum
    UserList(UserIndex).UseAcum = UserList(UserIndex).UseAcum + (TempUseNum * 200)

End Sub
