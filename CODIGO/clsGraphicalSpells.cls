VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private listOffset As Long             'Number of lines we scrolled down from topmost
Private SpellSelected As Long             'Currently selected item

Private Pressed As Boolean

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer
Private maxList As Long

Private mouseX As Single
Private mouseY As Single

Private ImgContainer As Boolean

Private inventorySpells() As Spell

Private WithEvents spellWindow As PictureBox    'Placeholder where to render the inventory
Attribute spellWindow.VB_VarHelpID = -1

Public Sub SetSpell(ByVal slot As Byte, _
                    ByVal eGrhIndex As Integer, _
                    ByVal eName As String)
'***************************************************
'Author: FrankoH298
'Last Modify Date: 26/03/20
'***************************************************
    If slot < 1 Or slot > UBound(inventorySpells) Then Exit Sub
    
    With inventorySpells(slot)
        .name = eName
        .GrhIndex = eGrhIndex
    End With
End Sub

Private Function ClickSpell(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: FrankoH298
'Last Modify Date: 26/03/20
'***************************************************
    Dim TempSpell As Long
        If Y >= spellWindow.Height Then
            ClickSpell = (spellWindow.Height / ownTilePixelHeight) + listOffset
        ElseIf Y < 0 Then
            ClickSpell = 1 + listOffset
        Else
            TempSpell = Y \ ownTilePixelHeight + 1 + listOffset
            'Make sure it's within limits
            If TempSpell <= UBound(inventorySpells) And TempSpell > 0 Then
                ClickSpell = TempSpell
            End If
        End If
End Function

Private Sub DrawInventory()
'***************************************************
'Author: FrankoH298
'Last Modify Date: 26/03/20
'***************************************************
    Dim loopc As Long
    Dim TempRect As RECT
    Dim listMax As Long
    Dim listMin As Long
    
    With TempRect
        .bottom = spellWindow.Height
        .Right = spellWindow.Width
    End With
    
    listMin = listOffset + 1
    listMax = maxList + listOffset
    
    'Clear the inventory window
        Call Engine_BeginScene
        startdY = 0
        For loopc = listMin To listMax
            With inventorySpells(loopc)
            If .GrhIndex Then Call Draw_GrhIndex(.GrhIndex, startdX, startdY, 1, Normal_RGBList)
                
            'Render a box around the selected item
            If SpellSelected = loopc Then
                If ImgContainer Then
                    Call Draw_GrhIndex(23980, startdX, startdY, 1, Normal_RGBList)
                End If
            End If

            'Render the name
            DrawText startdX + ownTilePixelWidth, startdY + (ownTilePixelHeight * 0.5) - 8, inventorySpells(loopc).name, -1
            
            End With
            startdY = startdY + ownTilePixelHeight
        Next loopc
        
        Call Engine_EndScene(TempRect, spellWindow.hWnd)
    
End Sub
Public Sub DrawInv()
    Call DrawInventory
    DoEvents
    Call DrawInventory
    DoEvents
    Call DrawInventory
End Sub
Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, _
                        Optional ByVal TileWidth As Integer = 32, _
                        Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                        Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True)
'***************************************************
'Author: FrankoH298
'Last Modify Date: 26/03/20
'***************************************************
    Set spellWindow = InvPic
    
    'Make sure auto-redraw is set to true
    spellWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    spellWindow.ScaleMode = 3
    
    ReDim inventorySpells(1 To MAXHECHI) As Spell

    'Make sure DirectD3D was correctly initialized
    If DirectD3D Is Nothing Then Exit Sub

    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    Pressed = False
    
    maxList = (spellWindow.Height / ownTilePixelHeight)
    
    ImgContainer = bImgContainer
End Sub


Private Sub spellWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: FrankoH298
'Last Modify Date: 26/03/20
'***************************************************
    
    If Button = 1 Then
        mouseX = X
        mouseY = Y
        'Store previously selected item
        Dim prevSelSpell As Long
        
        'Exit if it got outside the control's area
        
        If X < 0 Or X > spellWindow.Width Then Exit Sub

        prevSelSpell = SpellSelected
        
        'Get the currently clickced item
        SpellSelected = ClickSpell(CInt(X), CInt(Y))
        If SpellSelected = -1 Then SpellSelected = prevSelSpell
    
        'Update needed inventory slots
        If prevSelSpell <> SpellSelected Then
            If prevSelSpell <> SpellSelected Then Call DrawInventory
        End If
    End If
End Sub

Private Sub spellWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: FrankoH298
'Last Modify Date: 26/03/20
'***************************************************
    Pressed = True
    mouseX = X
    mouseY = Y
    
    'Store previously selected item
    Dim prevSelSpell As Long

    'Exit if it got outside the control's area
    If mouseX < 0 Or mouseY < 0 Or mouseX > spellWindow.Width Or mouseY > spellWindow.Height Then _
        Exit Sub
    
    prevSelSpell = SpellSelected
    
    'Get the currently clickced item
    SpellSelected = ClickSpell(CInt(mouseX), CInt(mouseY))
    If SpellSelected = -1 Then SpellSelected = prevSelSpell
    
    'Update needed inventory slots
    If prevSelSpell <> SpellSelected Then
        If prevSelSpell <> SpellSelected Then Call DrawInventory
    End If
End Sub

Private Sub spellWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pressed = False
    mouseX = X
    mouseY = Y
End Sub

Public Sub Update()
    Static lastMovement As Long
    If Pressed Then
        If timeGetTime > lastMovement Then
            If mouseY < 0 Then
                If listOffset > 0 Then listOffset = listOffset - 1
            ElseIf mouseY >= spellWindow.Height Then
                If listOffset < MAXHECHI - maxList Then listOffset = listOffset + 1
            End If
            Call spellWindow_MouseMove(1, 0, mouseX, mouseY)
            lastMovement = timeGetTime + 32
        End If
    End If
End Sub
