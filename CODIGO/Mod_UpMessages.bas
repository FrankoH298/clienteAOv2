Attribute VB_Name = "Mod_MessagesUp"
'***************************************************************
'Component        : Mod_MessagesUp
'Author           : FrankoH298
'Description      : Utilizado para renderizar mensajes que suben al pegar, agarrar oro, o al trabajar.
'***************************************************************

Option Explicit

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

Private Const delayTime As Long = 40

Public Sub createMessageUp(ByVal Text As String, ByVal tipo As Byte, ByVal CharIndex As Integer)
'***************************************************************
'Author           : FrankoH298
'Description      : Sirve para crear el mensaje.
'Last Modification: 08/04/2020
'***************************************************************

    With charlist(CharIndex).messageUp
        Select Case tipo
        
            Case TipoMsgUp.Damage
                .R = 220
                .G = 0
                .B = 0
                .Alpha = 255
                .startTickCount = timeGetTime + delayTime
                .Sube = 0
                
            Case TipoMsgUp.Gold
                .R = 250
                .G = 240
                .B = 5
                .Alpha = 255
                .startTickCount = timeGetTime + delayTime
                .Sube = 0
    
            Case TipoMsgUp.Trabajo
                .R = 10
                .G = 190
                .B = 190
                .Alpha = 210
                .startTickCount = timeGetTime + delayTime
                .Sube = 0
            Case Else
                
        End Select
        
        .Text = Text
        .active = 1
    End With
End Sub

Public Sub renderMessageUp(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************************
'Author           : FrankoH298
'Description      : Metodo que renderiza el mensaje, le aumenta su posicion en eje Y, y le baja su alpha.
'Last Modification: 08/04/2020
'***************************************************************

    With charlist(CharIndex)
        If .messageUp.active = 1 Then
            Call DrawText(PixelOffsetX + 10, PixelOffsetY - 20 - .messageUp.Sube, .messageUp.Text, D3DColorARGB(.messageUp.Alpha, .messageUp.R, .messageUp.G, .messageUp.B), , 2)
            If .messageUp.Sube < 20 Then
                If timeGetTime > .messageUp.startTickCount Then
                    .messageUp.Alpha = .messageUp.Alpha - 10
                    .messageUp.Sube = .messageUp.Sube + 1
                    .messageUp.startTickCount = timeGetTime + delayTime
                End If
            Else
                .messageUp.active = 0
            End If
        End If
    End With
    
End Sub
