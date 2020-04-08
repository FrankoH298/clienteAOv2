Attribute VB_Name = "Mod_MessagesUp"

Public Type textUp
    
    Text            As String
    Alpha           As Byte
    R               As Byte
    G               As Byte
    B               As Byte
    startTickCount  As Long
    Sube            As Long
    active          As Byte
End Type

Private Enum TipoMsgUp
    Damage = 1
    Gold = 2
    Trabajo = 3
End Enum

Public Sub createMessageUp(ByVal Text As String, ByVal tipo As Byte, ByVal CharIndex As Integer)
    With charlist(CharIndex).messageUp
        Select Case tipo
        
            Case TipoMsgUp.Damage
                .R = 255
                .G = 0
                .B = 0
                .Alpha = 255
                .startTickCount = timeGetTime + 20
                .Sube = 0
                
            Case TipoMsgUp.Gold
                .R = 220
                .G = 250
                .B = 5
                .Alpha = 255
                .startTickCount = timeGetTime + 20
                .Sube = 0
    
            Case TipoMsgUp.Trabajo
                .R = 255
                .G = 0
                .B = 0
                .Alpha = 255
                .startTickCount = timeGetTime + 20
                .Sube = 0
            Case Else
                
        End Select
        
        .Text = Text
        .active = 1
    End With
End Sub

Public Sub renderMessageUp(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    With charlist(CharIndex)
        If .messageUp.active = 1 Then
            Call DrawText(PixelOffsetX + 10, PixelOffsetY - 20 - .messageUp.Sube, .messageUp.Text, D3DColorARGB(.messageUp.Alpha, .messageUp.R, .messageUp.G, .messageUp.B))
            If .messageUp.Sube < 20 Then
                If timeGetTime > .messageUp.startTickCount Then
                    .messageUp.Alpha = .messageUp.Alpha - 12
                    .messageUp.Sube = .messageUp.Sube + 1
                    .messageUp.startTickCount = timeGetTime + 20
                End If
            Else
                .messageUp.active = 0
            End If
        End If
    End With
    
End Sub
