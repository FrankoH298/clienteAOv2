Attribute VB_Name = "Mod_General"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
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

Public iplst As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long
Private timerFlush As Long

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(path(INIT) & "" & "armas.dat")
    
    NumWeaponAnims = Val(Lector.GetValue("INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(Lector.GetValue("ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(Lector.GetValue("ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(Lector.GetValue("ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(Lector.GetValue("ARMA" & loopc, "Dir4")), 0
    Next loopc
    
End Sub

Sub CargarColores()
On Error Resume Next
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(path(INIT) & "colores.dat")
    
    If Not FileExist(path(INIT) & "colores.dat", vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(Lector.GetValue(CStr(i), "R"))
        ColoresPJ(i).g = CByte(Lector.GetValue(CStr(i), "G"))
        ColoresPJ(i).b = CByte(Lector.GetValue(CStr(i), "B"))
    Next i
    
    ' Crimi
    ColoresPJ(50).r = CByte(Lector.GetValue("CR", "R"))
    ColoresPJ(50).g = CByte(Lector.GetValue("CR", "G"))
    ColoresPJ(50).b = CByte(Lector.GetValue("CR", "B"))
    
    ' Ciuda
    ColoresPJ(49).r = CByte(Lector.GetValue("CI", "R"))
    ColoresPJ(49).g = CByte(Lector.GetValue("CI", "G"))
    ColoresPJ(49).b = CByte(Lector.GetValue("CI", "B"))
    
    ' Atacable
    ColoresPJ(48).r = CByte(Lector.GetValue("AT", "R"))
    ColoresPJ(48).g = CByte(Lector.GetValue("AT", "G"))
    ColoresPJ(48).b = CByte(Lector.GetValue("AT", "B"))
    
End Sub

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(path(INIT) & "" & "escudos.dat")
    
    NumEscudosAnims = Val(Lector.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(Lector.GetValue("ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(Lector.GetValue("ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(Lector.GetValue("ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(Lector.GetValue("ESC" & loopc, "Dir4")), 0
    Next loopc
    
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    frmMain.lblName.Caption = UserName
    'Load main form
    frmMain.Visible = True
    
    Call frmMain.ControlSM(eSMType.mSpells, False)
    Call frmMain.ControlSM(eSMType.mWork, False)

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    Static lastMovement As Long
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If (charlist(UserCharIndex).Heading <> Direccion) And (timeGetTime > lastMovement) Then
              Call WriteChangeHeading(Direccion)
              lastMovement = timeGetTime + 96
        End If
    End If
    
    If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Remueve la tecla que teniamos presionada
            End If
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Remueve la tecla que teniamos presionada
            End If
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Remueve la tecla que teniamos presionada
            End If
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) = False Then lastKeys.Add (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Agrega la tecla al arraylist
            Else
                If lastKeys.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then lastKeys.Remove (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Remueve la tecla que teniamos presionada
            End If
            
            'Move Up
            If lastKeys.Count() = 38 Then
                Call MoveTo(NORTH)
                Exit Sub
            End If
            
            'Move Right
            If lastKeys.Count() = 39 Then
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If lastKeys.Count() = 40 Then
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If lastKeys.Count() = 37 Then
                Call MoveTo(WEST)
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            'frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
            frmMain.Coord.Caption = "X: " & UserPos.X & " Y: " & UserPos.Y
        End If
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer)
      '**************************************************************
      'Formato de mapas optimizado para reducir el espacio que ocupan.
      'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
      '**************************************************************

      ' @@ Modificacion realizada por Facundo (GodKer), robada por Marcos.. ?
      ' @@ 06/11/2014
      ' @@ Facu cabe aporte (?)

      Dim Y         As Long
      Dim X         As Long
      Dim ByFlags   As Byte
      Dim handle    As Integer
      Dim fileBuff  As clsByteBuffer
   
      Dim dData()   As Byte
      Dim dLen      As Long
      
      Call Particle_Group_Remove_All
        
      Set fileBuff = New clsByteBuffer
   
      dLen = FileLen(path(Mapas) & "Mapa" & Map & ".map")
      ReDim dData(dLen - 1)
      
      handle = FreeFile()
   
      Open path(Mapas) & "Mapa" & Map & ".map" For Binary As handle
      'Seek handle, 1
      Get handle, , dData
      Close handle
     
      fileBuff.initializeReader dData

      MapInfo.MapVersion = fileBuff.getInteger
   
      MiCabecera.Desc = fileBuff.getString(Len(MiCabecera.Desc))
      MiCabecera.CRC = fileBuff.getLong
      MiCabecera.MagicWord = fileBuff.getLong
   
      fileBuff.getDouble
   
      'Load arrays

      For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
                  'Get handle, , ByFlags
                  ByFlags = fileBuff.getByte()
           
                  MapData(X, Y).Blocked = (ByFlags And 1)
           
                  'Get handle, , MapData(X, Y).Graphic(1).GrhIndex
                  MapData(X, Y).Graphic(1).GrhIndex = fileBuff.getInteger()
                  InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
           
                  'Layer 2 used?

                  If ByFlags And 2 Then
                        'Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                        MapData(X, Y).Graphic(2).GrhIndex = fileBuff.getInteger()
                        InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
                  Else
                        MapData(X, Y).Graphic(2).GrhIndex = 0
                  End If
               
                  'Layer 3 used?

                  If ByFlags And 4 Then
                        'Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                        MapData(X, Y).Graphic(3).GrhIndex = fileBuff.getInteger()
                        InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
                  Else
                        MapData(X, Y).Graphic(3).GrhIndex = 0
                  End If
               
                  'Layer 4 used?

                  If ByFlags And 8 Then
                        'Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                        MapData(X, Y).Graphic(4).GrhIndex = fileBuff.getInteger()
                        InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
                  Else
                        MapData(X, Y).Graphic(4).GrhIndex = 0
                  End If
           
                  'Trigger used?

                  If ByFlags And 16 Then
                        'Get handle, , MapData(X, Y).Trigger
                        MapData(X, Y).Trigger = fileBuff.getInteger()
                  Else
                        MapData(X, Y).Trigger = 0
                  End If
           
                  'Erase NPCs

                  If MapData(X, Y).CharIndex > 0 Then
                        Call EraseChar(MapData(X, Y).CharIndex)
                  End If
           
                  'Erase OBJs
                  MapData(X, Y).ObjGrh.GrhIndex = 0

            Next X
      Next Y

      'Close handle
     
      Set fileBuff = Nothing ' @@ Tanto te costaba Destruir el buff una ves que se termino de usar?
      
      MapInfo.name = vbNullString
      MapInfo.Music = vbNullString
   
      CurMap = Map
      Call General_Particle_Create(97, 60, 45)
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open path(INIT) & "Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
'********************************
'Author: Unknown
'Last Modification: 07/26/07
'Last Modified by: Rapsodius
'Added Instruction "CloseClient" before End so the mutex is cleared
'********************************
On Error GoTo errorH
    Dim c As Integer
    Dim i As Long
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(path(INIT) & "sinfo.dat")
    
    c = Val(Lector.GetValue("INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For i = 1 To c
        ServersLst(i).Desc = Lector.GetValue("S" & i, "Desc")
        ServersLst(i).Ip = Trim$(Lector.GetValue("S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(Lector.GetValue("S" & i, "P2"))
        ServersLst(i).Puerto = CInt(Lector.GetValue("S" & i, "PJ"))
    Next i
    CurServer = 1
    
Exit Sub

errorH:
    
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    
    Call CloseClient
End Sub

Public Sub InitServersList()
On Error Resume Next
    Dim NumServers As Integer
    Dim i As Integer
    Dim Cont As Integer
    
    i = 1
    
    Do While (ReadField(i, RawServersList, Asc(";")) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont
        Dim cur$
        cur$ = ReadField(i, RawServersList, Asc(";"))
        ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(i).Desc = ReadField(4, cur$, Asc(":"))
        ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next i
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
End Function

Public Function CurServerIp() As String
    If CurServer <> 0 Then
        CurServerIp = ServersLst(CurServer).Ip
    Else
        CurServerIp = frmConnect.IPTxt
    End If
End Function

Public Function CurServerPort() As Integer
    If CurServer <> 0 Then
        CurServerPort = ServersLst(CurServer).Puerto
    Else
        CurServerPort = Val(frmConnect.PortTxt)
    End If
End Function

Sub Main()

    Call WriteClientVer
    
    Call LeerConfiguracion
    
    ' Contraseña del archivo Graphics.AO
    Call modCompression.GenerateContra("", 0) ' 0 = Graficos.AO
    
    DirectXInit
    
    If FindPreviousInstance Then
        'Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        'End
    End If
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(path(EXTRAS) & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(path(EXTRAS) & "Hand.ico")
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    Call AddtoRichTextBox(frmCargando.status, "Buscando servidores... ", 255, 255, 255, True, False, True)

    Call CargarServidores
'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmCargando.status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
    
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    With frmConnect
        .txtNombre = "BetaTester"
        .txtNombre.SelStart = 0
        .txtNombre.SelLength = Len(.txtNombre)
    End With
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando motor gráfico... ", 255, 255, 255, True, False, True)
    
    If Not InitTileEngine(frmMain.hWnd, 149, 13, 32, 32, 13, 17, 9, 8, 8, 0.018) Then
        Call CloseClient
    End If
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    
UserMap = 1
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Set lastKeys = New clsArrayList
    Call lastKeys.Initialize(1, 4)
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
    
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hWnd, path(WAV), path(MIDI))
    'Enable / Disable audio
    Audio.MusicActivated = ClientSetup.bMusic
    Audio.SoundActivated = ClientSetup.bSound
    Audio.SoundEffectsActivated = ClientSetup.bSoundEffects
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS)
    
    Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "                    ¡Bienvenido a Argentum Online!", 255, 255, 255, True, False, True)
    
    'Give the user enough time to read the welcome text
    Call Sleep(500)
    
    Unload frmCargando

#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    frmConnect.Visible = True
    
    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    'Set the dialog's font
    Dialogos.Font = frmMain.Font
    DialogosClanes.Font = frmMain.Font
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
        
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call CheckKeys
            Else
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
        End If
        'FPS Counter - mostramos las FPS
        If timeGetTime >= lFrameTimer Then
            If frmMain.Visible Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            lFrameTimer = timeGetTime + 1000
        End If
        
        ' If there is anything to be sent, we send it
        If timeGetTime >= timerFlush Then
            Call FlushBuffer
            timerFlush = timeGetTime + 12
        End If
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    frmCargando.Show
    Call AddtoRichTextBox(frmCargando.status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set lastKeys = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    End
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function

Public Sub checkText(ByVal Text As String)
Dim Nivel As Integer
If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
    Call ScreenCapture(True)
    Exit Sub
End If
If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
    EsperandoLevel = True
    Exit Sub
End If
If EsperandoLevel Then
    If Right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
        If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > ClientSetup.byMurderedLevel Then
            Call ScreenCapture(True)
        End If
    End If
End If
EsperandoLevel = False
End Sub

Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).Nombre = name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function
