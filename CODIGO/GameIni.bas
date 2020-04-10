Attribute VB_Name = "GameIni"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc        As String * 255
    CRC         As Long
    MagicWord   As Long
End Type

Public Enum ePath
    INIT
    Graficos
    MIDI
    WAV
    Mapas
    EXTRAS
End Enum

Public Type tSetupMods
    byMemory            As Byte
    bNoRes              As Boolean  ' 24/06/2006 - ^[GS]^
    bVertexProcessing   As Byte
    bVSync              As Boolean
    bMusic              As Boolean
    bSound              As Boolean
    bSoundEffects       As Boolean
    bGuildNews          As Boolean  ' 11/19/09
    bDie                As Boolean  ' 11/23/09 - FragShooter
    bKill               As Boolean  ' 11/23/09 - FragShooter
    byMurderedLevel     As Byte     ' 11/23/09 - FragShooter
    bActive             As Boolean
    bGldMsgConsole      As Boolean
    bCantMsgs           As Byte
    bNombres            As Byte
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Private Lector As clsIniManager
Private Const CLIENT_FILE As String = "Config.ini"

Public Sub IniciarCabecera()
    With MiCabecera
        .Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
        .CRC = Rnd * 100
        .MagicWord = Rnd * 10
    End With
End Sub

Public Function path(ByVal PathType As ePath) As String

    Select Case PathType
        
        Case ePath.INIT
            path = App.path & "\INIT\"
        
        Case ePath.Graficos
            path = App.path & "\GRAFICOS\"
            
        Case ePath.Mapas
            path = App.path & "\MAPAS\"
            
        Case ePath.MIDI
            path = App.path & "\MIDI\"
            
        Case ePath.WAV
            path = App.path & "\WAV\"
            
        Case ePath.EXTRAS
            path = App.path & "\Extras\"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    
    Call IniciarCabecera
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(path(INIT) & CLIENT_FILE)
    
    With ClientSetup
        
        ' VIDEO
        .byMemory = Lector.GetValue("VIDEO", "DINAMIC_MEMORY")
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DISABLE_RESOLUTION_CHANGE"))
        .bNombres = CByte(Lector.GetValue("VIDEO", "NOMBRES"))
        If .bNombres > 2 Or .bNombres < 0 Then .bNombres = 0
        .bVSync = CBool(Lector.GetValue("VIDEO", "VSYNC"))
        .bVertexProcessing = CByte(Lector.GetValue("VIDEO", "VERTEX_PROCESSING"))
        
        ' AUDIO
        .bMusic = CBool(Lector.GetValue("AUDIO", "MIDI"))
        .bSound = CBool(Lector.GetValue("AUDIO", "WAV"))
        .bSoundEffects = CBool(Lector.GetValue("AUDIO", "SOUND_EFFECTS"))
        
        ' GUILD
        .bGuildNews = CBool(Lector.GetValue("GUILD", "NEWS"))
        .bGldMsgConsole = CBool(Lector.GetValue("GUILD", "MESSAGES"))
        .bCantMsgs = CByte(Lector.GetValue("GUILD", "MAX_MESSAGES"))
        
        ' FRAGSHOOTER
        .bDie = CBool(Lector.GetValue("FRAGSHOOTER", "DIE"))
        .bKill = CBool(Lector.GetValue("FRAGSHOOTER", "KILL"))
        .byMurderedLevel = CBool(Lector.GetValue("FRAGSHOOTER", "MURDERED_LEVEL"))
        .bActive = CBool(Lector.GetValue("FRAGSHOOTER", "ACTIVE"))
        
    End With
End Sub

Public Sub GuardarConfiguracion()
    On Local Error GoTo fileErr:
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(path(INIT) & CLIENT_FILE)
    
    With ClientSetup
        
        ' VIDEO
        Call Lector.ChangeValue("VIDEO", "DINAMIC_MEMORY", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DISABLE_RESOLUTION_CHANGE", CInt(.bNoRes))
        Call Lector.ChangeValue("VIDEO", "NOMBRES", .bNombres)
        Call Lector.ChangeValue("VIDEO", "VSYNC", CInt(.bVSync))
        Call Lector.ChangeValue("VIDEO", "VERTEX_PROCESSING", .bVertexProcessing)
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MIDI", CInt(.bMusic))
        Call Lector.ChangeValue("AUDIO", "WAV", CInt(.bSound))
        Call Lector.ChangeValue("AUDIO", "SOUND_EFFECTS", CInt(.bSoundEffects))
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "NEWS", CInt(.bGuildNews))
        Call Lector.ChangeValue("GUILD", "MESSAGES", CInt(.bGldMsgConsole))
        Call Lector.ChangeValue("GUILD", "MAX_MESSAGES", CInt(.bCantMsgs))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "DIE", CInt(.bDie))
        Call Lector.ChangeValue("FRAGSHOOTER", "KILL", CInt(.bKill))
        Call Lector.ChangeValue("FRAGSHOOTER", "MURDERED_LEVEL", CInt(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "ACTIVE", CInt(.bActive))

    End With
    
    Call Lector.DumpFile(path(INIT) & CLIENT_FILE)
    
fileErr:
    
    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
        End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub

