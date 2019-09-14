Attribute VB_Name = "modParticlesORE"
'ImperiumAO 1.4.6
'Modulo Particles

Option Explicit

'******Particulas******
'RGB Type
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    alphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
End Type

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
End Type

Private Type Particle_Group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    Particle_Count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alphaBlend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Mart�n Sotuyo Dodero
    speed As Single
    life_counter As Long
End Type

Dim particle_group_list() As Particle_Group
Dim particle_group_count As Long
Dim particle_group_last As Long
Public TotalStreams As Integer
Public StreamData() As Stream

Public Const PI As Single = 3.14159265358979
'******Particulas******

Public Sub CargarParticulas()
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim Leer As New clsIniManager

    Dim StreamFile As String

    StreamFile = path(INIT) & "particulas.ini"

    Leer.Initialize StreamFile

    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = Leer.GetValue(Val(loopc), "X1")
        StreamData(loopc).y1 = Leer.GetValue(Val(loopc), "Y1")
        StreamData(loopc).x2 = Leer.GetValue(Val(loopc), "X2")
        StreamData(loopc).y2 = Leer.GetValue(Val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
        StreamData(loopc).alphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(Leer.GetValue(Val(loopc), "Speed"))
        
        StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = ReadField(i, GrhListing, Asc(","))
        Next i
        
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = ReadField(1, TempSet, Asc(","))
            StreamData(loopc).colortint(ColorSet - 1).g = ReadField(2, TempSet, Asc(","))
            StreamData(loopc).colortint(ColorSet - 1).b = ReadField(3, TempSet, Asc(","))
        Next ColorSet

                
    Next loopc
    
    Set Leer = Nothing

End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal Particle_Group_Index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If Particle_Group_Index > 0 And Particle_Group_Index <= particle_group_last Then
        If particle_group_list(Particle_Group_Index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Private Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal Particle_Count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Mart�n Sotuyo Dodero
'Modified by Augusto Jos� Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
    End If

End Function

Public Function Particle_Group_Remove(ByVal Particle_Group_Index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(Particle_Group_Index) Then
        Particle_Group_Destroy Particle_Group_Index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Private Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
Private Function Particle_Get_Type(ByVal Particle_Group_Index As Long) As Byte
On Error GoTo ErrorHandler:
    Particle_Get_Type = particle_group_list(Particle_Group_Index).stream_type
Exit Function
ErrorHandler:
    Particle_Get_Type = 0
End Function
Private Sub Particle_Group_Destroy(ByVal Particle_Group_Index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    Dim temp As Particle_Group
    Dim i As Integer
    
    If particle_group_list(Particle_Group_Index).map_x > 0 And particle_group_list(Particle_Group_Index).map_y > 0 Then
        MapData(particle_group_list(Particle_Group_Index).map_x, particle_group_list(Particle_Group_Index).map_y).Particle_Group_Index = 0
    ElseIf particle_group_list(Particle_Group_Index).char_index Then
        If Char_Check(particle_group_list(Particle_Group_Index).char_index) Then
            For i = 1 To charlist(particle_group_list(Particle_Group_Index).char_index).Particle_Count
                If charlist(particle_group_list(Particle_Group_Index).char_index).Particle_Group(i) = Particle_Group_Index Then
                    charlist(particle_group_list(Particle_Group_Index).char_index).Particle_Group(i) = 0
        
                    'We don't resize arrays by now, it's really a waste...
                    'If i = UBound(CharList(particle_group_list(particle_group_index).char_index).particle_group) Then
                    '    CharList(particle_group_list(particle_group_index).char_index).particle_count = i - 1
                    '    ReDim Preserve CharList(particle_group_list(particle_group_index).char_index).particle_group(1 To (i - 1)) As Long
                    'End If
        
                    Exit For
                End If
            Next i
        End If
    End If
    
    particle_group_list(Particle_Group_Index) = temp
    
    'Update array size
    If Particle_Group_Index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        Debug.Print particle_group_last & "," & UBound(particle_group_list)
        ReDim Preserve particle_group_list(1 To particle_group_last) As Particle_Group
    End If
    particle_group_count = particle_group_count - 1
End Sub
Public Sub Particle_Group_Render(ByVal Particle_Group_Index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    If Particle_Group_Index > UBound(particle_group_list) Then Exit Sub
    
    If GetTickCount - particle_group_list(Particle_Group_Index).live > (particle_group_list(Particle_Group_Index).liv1 * 25) And Not particle_group_list(Particle_Group_Index).liv1 = -1 Then
        Particle_Group_Destroy Particle_Group_Index
        Exit Sub
    End If
        
    With particle_group_list(Particle_Group_Index)
        'Set colors
        temp_rgb(0) = .rgb_list(0)
        temp_rgb(1) = .rgb_list(1)
        temp_rgb(2) = .rgb_list(2)
        temp_rgb(3) = .rgb_list(3)

        'See if it is time to move a particle
        .frame_counter = .frame_counter + timerTicksPerFrame
        If .frame_counter > .frame_speed Then
            .frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
            
        'If it's still alive render all the particles inside
        For loopc = 1 To .Particle_Count
                
        'Render particle
            Particle_Render .particle_stream(loopc), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(RandomNumber(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alphaBlend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin
        Next loopc
                
        If no_move = False Then
            'Update the group alive counter
            If .never_die = False Then
                .alive_counter = .alive_counter - 1
            End If
        End If
    End With
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alphaBlend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index
            temp_particle.X = RandomNumber(x1, x2) - 16
            temp_particle.Y = RandomNumber(y1, y2) - 16
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            'temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin Then temp_particle.angle = temp_particle.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    
    If temp_particle.Grh.GrhIndex Then
        Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, rgb_list(), 1, True, temp_particle.angle
    End If
End Sub
Private Sub Particle_Group_Make(ByVal Particle_Group_Index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal Particle_Count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Mart�n Sotuyo Dodero
'*****************************************************************
    'Update array size
    If Particle_Group_Index > particle_group_last Then
        particle_group_last = Particle_Group_Index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(Particle_Group_Index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(Particle_Group_Index).map_x = map_x
        particle_group_list(Particle_Group_Index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(Particle_Group_Index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(Particle_Group_Index).grh_index_list() = grh_index_list()
    particle_group_list(Particle_Group_Index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(Particle_Group_Index).alive_counter = -1
        particle_group_list(Particle_Group_Index).liv1 = -1
        particle_group_list(Particle_Group_Index).never_die = True
    Else
        particle_group_list(Particle_Group_Index).alive_counter = alive_counter
        particle_group_list(Particle_Group_Index).liv1 = alive_counter
        particle_group_list(Particle_Group_Index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(Particle_Group_Index).alphaBlend = alphaBlend
    
    'stream type
    particle_group_list(Particle_Group_Index).stream_type = stream_type
    
    'speed
    particle_group_list(Particle_Group_Index).frame_speed = frame_speed
    
    particle_group_list(Particle_Group_Index).x1 = x1
    particle_group_list(Particle_Group_Index).y1 = y1
    particle_group_list(Particle_Group_Index).x2 = x2
    particle_group_list(Particle_Group_Index).y2 = y2
    particle_group_list(Particle_Group_Index).angle = angle
    particle_group_list(Particle_Group_Index).vecx1 = vecx1
    particle_group_list(Particle_Group_Index).vecx2 = vecx2
    particle_group_list(Particle_Group_Index).vecy1 = vecy1
    particle_group_list(Particle_Group_Index).vecy2 = vecy2
    particle_group_list(Particle_Group_Index).life1 = life1
    particle_group_list(Particle_Group_Index).life2 = life2
    particle_group_list(Particle_Group_Index).fric = fric
    particle_group_list(Particle_Group_Index).spin = spin
    particle_group_list(Particle_Group_Index).spin_speedL = spin_speedL
    particle_group_list(Particle_Group_Index).spin_speedH = spin_speedH
    particle_group_list(Particle_Group_Index).gravity = gravity
    particle_group_list(Particle_Group_Index).grav_strength = grav_strength
    particle_group_list(Particle_Group_Index).bounce_strength = bounce_strength
    particle_group_list(Particle_Group_Index).XMove = XMove
    particle_group_list(Particle_Group_Index).YMove = YMove
    particle_group_list(Particle_Group_Index).move_x1 = move_x1
    particle_group_list(Particle_Group_Index).move_x2 = move_x2
    particle_group_list(Particle_Group_Index).move_y1 = move_y1
    particle_group_list(Particle_Group_Index).move_y2 = move_y2
    
    particle_group_list(Particle_Group_Index).rgb_list(0) = rgb_list(0)
    particle_group_list(Particle_Group_Index).rgb_list(1) = rgb_list(1)
    particle_group_list(Particle_Group_Index).rgb_list(2) = rgb_list(2)
    particle_group_list(Particle_Group_Index).rgb_list(3) = rgb_list(3)
    
    'handle
    particle_group_list(Particle_Group_Index).id = id
    
    particle_group_list(Particle_Group_Index).live = GetTickCount()
    
    'create particle stream
    particle_group_list(Particle_Group_Index).Particle_Count = Particle_Count
    ReDim particle_group_list(Particle_Group_Index).particle_stream(1 To Particle_Count)
    
    'plot particle group on map
    If (map_x <> -1 And map_x <> 0) And (map_y <> -1 And map_x <> 0) Then
        MapData(map_x, map_y).Particle_Group_Index = Particle_Group_Index
    End If
    
End Sub
Private Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).Particle_Group_Index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).alphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).alphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function
Private Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal Particle_Count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, Particle_Count, stream_type, grh_index_list(), rgb_list(), alphaBlend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
    End If

End Function


Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto Jos� Rando
'**************************************************************
    Dim char_part_index As Integer

    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto Jos� Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) And Not charlist(char_index).Particle_Count = 0 Then
        For i = 1 To UBound(charlist(char_index).Particle_Group)
            If charlist(char_index).Particle_Group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).Particle_Group(i))
        Next i
        Erase charlist(char_index).Particle_Group
        charlist(char_index).Particle_Count = 0
    End If
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto Jos� Rando
'Modified: returns slot or -1
'*****************************************************************
On Error Resume Next
Dim i As Integer

For i = 1 To charlist(char_index).Particle_Count
    If particle_group_list(charlist(char_index).Particle_Group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).Particle_Group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto Jos� Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If charlist(char_index).Particle_Count = 0 Then
        Char_Particle_Group_Next_Open = charlist(char_index).Particle_Count + 1
        charlist(char_index).Particle_Count = Char_Particle_Group_Next_Open
        ReDim Preserve charlist(char_index).Particle_Group(1 To Char_Particle_Group_Next_Open) As Long
        Exit Function
    End If
    
    loopc = 1
    Do Until charlist(char_index).Particle_Group(loopc) = 0
        If loopc = charlist(char_index).Particle_Count Then
            Char_Particle_Group_Next_Open = charlist(char_index).Particle_Count + 1
            charlist(char_index).Particle_Count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).Particle_Group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).Particle_Count = 1
    ReDim charlist(char_index).Particle_Group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function
Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Mart�n Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
    
End Function
Private Sub Char_Particle_Group_Make(ByVal Particle_Group_Index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal Particle_Count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alphaBlend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Mart�n Sotuyo Dodero
'*****************************************************************
    'Update array size
    If Particle_Group_Index > particle_group_last Then
        particle_group_last = Particle_Group_Index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(Particle_Group_Index).active = True
    
    'Char index
    particle_group_list(Particle_Group_Index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(Particle_Group_Index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(Particle_Group_Index).grh_index_list() = grh_index_list()
    particle_group_list(Particle_Group_Index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(Particle_Group_Index).alive_counter = -1
        particle_group_list(Particle_Group_Index).liv1 = -1
        particle_group_list(Particle_Group_Index).never_die = True
    Else
        particle_group_list(Particle_Group_Index).alive_counter = alive_counter
        particle_group_list(Particle_Group_Index).liv1 = alive_counter
        particle_group_list(Particle_Group_Index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(Particle_Group_Index).alphaBlend = alphaBlend
    
    'stream type
    particle_group_list(Particle_Group_Index).stream_type = stream_type
    
    'speed
    particle_group_list(Particle_Group_Index).frame_speed = frame_speed
    
    particle_group_list(Particle_Group_Index).x1 = x1
    particle_group_list(Particle_Group_Index).y1 = y1
    particle_group_list(Particle_Group_Index).x2 = x2
    particle_group_list(Particle_Group_Index).y2 = y2
    particle_group_list(Particle_Group_Index).angle = angle
    particle_group_list(Particle_Group_Index).vecx1 = vecx1
    particle_group_list(Particle_Group_Index).vecx2 = vecx2
    particle_group_list(Particle_Group_Index).vecy1 = vecy1
    particle_group_list(Particle_Group_Index).vecy2 = vecy2
    particle_group_list(Particle_Group_Index).life1 = life1
    particle_group_list(Particle_Group_Index).life2 = life2
    particle_group_list(Particle_Group_Index).fric = fric
    particle_group_list(Particle_Group_Index).spin = spin
    particle_group_list(Particle_Group_Index).spin_speedL = spin_speedL
    particle_group_list(Particle_Group_Index).spin_speedH = spin_speedH
    particle_group_list(Particle_Group_Index).gravity = gravity
    particle_group_list(Particle_Group_Index).grav_strength = grav_strength
    particle_group_list(Particle_Group_Index).bounce_strength = bounce_strength
    particle_group_list(Particle_Group_Index).XMove = XMove
    particle_group_list(Particle_Group_Index).YMove = YMove
    particle_group_list(Particle_Group_Index).move_x1 = move_x1
    particle_group_list(Particle_Group_Index).move_x2 = move_x2
    particle_group_list(Particle_Group_Index).move_y1 = move_y1
    particle_group_list(Particle_Group_Index).move_y2 = move_y2
    
    particle_group_list(Particle_Group_Index).rgb_list(0) = rgb_list(0)
    particle_group_list(Particle_Group_Index).rgb_list(1) = rgb_list(1)
    particle_group_list(Particle_Group_Index).rgb_list(2) = rgb_list(2)
    particle_group_list(Particle_Group_Index).rgb_list(3) = rgb_list(3)
    
    'handle
    particle_group_list(Particle_Group_Index).id = id
    particle_group_list(Particle_Group_Index).live = GetTickCount()
    
    'create particle stream
    particle_group_list(Particle_Group_Index).Particle_Count = Particle_Count
    ReDim particle_group_list(Particle_Group_Index).particle_stream(1 To Particle_Count)
    
    'plot particle group on char
    charlist(char_index).Particle_Group(particle_char_index) = Particle_Group_Index
End Sub
