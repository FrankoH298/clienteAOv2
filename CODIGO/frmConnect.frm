VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkRecordarPassword 
      Caption         =   "Recordar Contrase�a"
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   3435
      Width           =   1935
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3720
      Width           =   2460
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4905
      TabIndex        =   0
      Top             =   3210
      Width           =   2460
   End
   Begin SHDocVwCtl.WebBrowser WebAuxiliar 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   330
      ExtentX         =   582
      ExtentY         =   635
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgTeclas 
      Height          =   375
      Left            =   6120
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image imgConectarse 
      Height          =   375
      Left            =   4800
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image imgSalir 
      Height          =   375
      Left            =   9960
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgBorrarPj 
      Height          =   375
      Left            =   8400
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgCodigoFuente 
      Height          =   375
      Left            =   6840
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgReglamento 
      Height          =   375
      Left            =   5280
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgManual 
      Height          =   375
      Left            =   3720
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgRecuperar 
      Height          =   375
      Left            =   2160
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgCrearPj 
      Height          =   375
      Left            =   600
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Mat�as Fernando Peque�o
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'C�digo Postal 1405

Option Explicit

Private cBotonCrearPj As clsGraphicalButton
Private cBotonRecuperarPass As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonReglamento As clsGraphicalButton
Private cBotonCodigoFuente As clsGraphicalButton
Private cBotonBorrarPj As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton
Private cBotonLeerMas As clsGraphicalButton
Private cBotonForo As clsGraphicalButton
Private cBotonConectarse As clsGraphicalButton
Private cBotonTeclas As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_Load()
    EngineRun = False

    version.Caption = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
    Me.Picture = LoadPicture(path(Graficos) & "VentanaConectar.jpg")
    
    Call LoadButtons

    Call CheckLicenseAgreement
        
End Sub

Private Sub CheckLicenseAgreement()
    'Recordatorio para cumplir la licencia, por si borr�s el Boton sin leer el code...
    Dim i As Long
    
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls(i).Name = "imgCodigoFuente" Then
            Exit For
        End If
    Next i
    
    If i = Me.Controls.Count Then
        MsgBox "No debe eliminarse la posibilidad de bajar el c�digo de sus servidor. Caso contrario estar�an violando la licencia Affero GPL y con ella derechos de autor, incurriendo de esta forma en un delito punible por ley." & vbCrLf & vbCrLf & vbCrLf & _
                "Argentum Online es libre, es de todos. Mantengamoslo as�. Si tanto te gusta el juego y quer�s los cambios que hacemos nosotros, compart� los tuyos. Es un cambio justo. Si no est�s de acuerdo, no uses nuestro c�digo, pues nadie te obliga o bien utiliza una versi�n anterior a la 0.12.0.", vbCritical Or vbApplicationModal
    End If

End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = path(Graficos)
    
    Set cBotonCrearPj = New clsGraphicalButton
    Set cBotonRecuperarPass = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonReglamento = New clsGraphicalButton
    Set cBotonCodigoFuente = New clsGraphicalButton
    Set cBotonBorrarPj = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonLeerMas = New clsGraphicalButton
    Set cBotonForo = New clsGraphicalButton
    Set cBotonConectarse = New clsGraphicalButton
    Set cBotonTeclas = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

        
    Call cBotonCrearPj.Initialize(imgCrearPj, GrhPath & "BotonCrearPersonajeConectar.jpg", _
                                    GrhPath & "BotonCrearPersonajeRolloverConectar.jpg", _
                                    GrhPath & "BotonCrearPersonajeClickConectar.jpg", Me)
                                    
    Call cBotonRecuperarPass.Initialize(imgRecuperar, GrhPath & "BotonRecuperarPass.jpg", _
                                    GrhPath & "BotonRecuperarPassRollover.jpg", _
                                    GrhPath & "BotonRecuperarPassClick.jpg", Me)
                                    
    Call cBotonManual.Initialize(imgManual, GrhPath & "BotonManual.jpg", _
                                    GrhPath & "BotonManualRollover.jpg", _
                                    GrhPath & "BotonManualClick.jpg", Me)
                                    
    Call cBotonReglamento.Initialize(imgReglamento, GrhPath & "BotonReglamento.jpg", _
                                    GrhPath & "BotonReglamentoRollover.jpg", _
                                    GrhPath & "BotonReglamentoClick.jpg", Me)
                                    
    Call cBotonCodigoFuente.Initialize(imgCodigoFuente, GrhPath & "BotonCodigoFuente.jpg", _
                                    GrhPath & "BotonCodigoFuenteRollover.jpg", _
                                    GrhPath & "BotonCodigoFuenteClick.jpg", Me)
                                    
    Call cBotonBorrarPj.Initialize(imgBorrarPj, GrhPath & "BotonBorrarPersonaje.jpg", _
                                    GrhPath & "BotonBorrarPersonajeRollover.jpg", _
                                    GrhPath & "BotonBorrarPersonajeClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirConnect.jpg", _
                                    GrhPath & "BotonBotonSalirRolloverConnect.jpg", _
                                    GrhPath & "BotonSalirClickConnect.jpg", Me)
                                    
    Call cBotonConectarse.Initialize(imgConectarse, GrhPath & "BotonConectarse.jpg", _
                                    GrhPath & "BotonConectarseRollover.jpg", _
                                    GrhPath & "BotonConectarseClick.jpg", Me)
                                    
    Call cBotonTeclas.Initialize(imgTeclas, GrhPath & "BotonTeclas.jpg", _
                                    GrhPath & "BotonTeclasRollover.jpg", _
                                    GrhPath & "BotonTeclasClick.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgBorrarPj_Click()

On Error GoTo errH
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)

    Exit Sub

errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgCodigoFuente_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el c�digo de sus servidor de esta forma.
'Caso contrario estar�an violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es libre, es de todos. Mantengamoslo as�. Si tanto te gusta el juego y quer�s los
'cambios que hacemos nosotros, compart� los tuyos. Es un cambio justo. Si no est�s de acuerdo,
'no uses nuestro c�digo, pues nadie te obliga o bien utiliza una versi�n anterior a la 0.12.0.
'***********************************
    Call ShellExecute(0, "Open", "https://sourceforge.net/project/downloading.php?group_id=67718&filename=AOServerSrc0.12.2.zip&a=42868900", "", App.path, SW_SHOWNORMAL)

End Sub

Private Sub imgConectarse_Click()
    
    Call Mod_RecordarPassword.savePassword(frmConnect.txtNombre, frmConnect.txtPasswd, frmConnect.chkRecordarPassword)

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
#End If
    
    'update user info
    UserName = txtNombre.Text
    
    Dim aux As String
    aux = txtPasswd.Text
    userPassword = aux
    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = IPdelServer
    frmMain.Socket1.RemotePort = PORTdelServer
    frmMain.Socket1.Connect
#Else
    frmMain.Winsock1.Connect IPdelServer, PORTdelServer
#End If

    End If
    
End Sub

Private Sub imgCrearPj_Click()
    
    EstadoLogin = E_MODO.Dados
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = IPdelServer
    frmMain.Socket1.RemotePort = PORTdelServer
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect IPdelServer, PORTdelServer
#End If

End Sub

Private Sub imgLeerMas_Click()
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgManual_Click()
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/manual/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgRecuperar_Click()
On Error GoTo errH

    Call Audio.PlayWave(SND_CLICK)
    Call Shell(App.path & "\RECUPERAR.EXE", vbNormalFocus)
    Exit Sub
errH:
    Call MsgBox("No se encuentra el programa recuperar.exe", vbCritical, "Argentum Online")
End Sub

Private Sub imgReglamento_Click()
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/reglamento.html", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgSalir_Click()
    prgRun = False
End Sub

Private Sub imgTeclas_Click()
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    txtPasswd.SetFocus
End Sub

Private Sub imgVerForo_Click()
    Call ShellExecute(0, "Open", "http://www.alkon.com.ar/foro/argentum-online.53/", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub txtNombre_Change()
    frmConnect.txtPasswd = Mod_RecordarPassword.loadPassword(frmConnect.txtNombre)
End Sub

Private Sub txtPasswd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then imgConectarse_Click
End Sub
