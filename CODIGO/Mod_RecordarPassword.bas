Attribute VB_Name = "Mod_RecordarPassword"
Option Explicit

Private Const FilePassword As String = "Passwd" ' CAMBIAR ESTO
Private FilePath As String

Public Function loadPassword(ByVal Name As String) As String
    Dim userPassword As String
    FilePath = App.path & "\INIT\Personajes.ini"
    If Trim$(Name) <> "" Then
        userPassword = CStr(GetVar(FilePath, "PJS", Encryption_RC4_EncryptString(Name, FilePassword)))
        Debug.Print (userPassword)
        If userPassword <> "" Then
            loadPassword = Encryption_RC4_DecryptString(userPassword, FilePassword)
        End If
    End If
End Function

Public Sub savePassword(ByVal Name As String, ByVal Password As String, ByVal CheckPassword As Boolean)
    If CheckPassword Then
        FilePath = App.path & "\INIT\Personajes.ini"
        If Trim$(Name) <> "" Then
            If Trim$(Password) <> "" Then
                Call WriteVar(FilePath, "PJS", Encryption_RC4_EncryptString(Name, FilePassword), Encryption_RC4_EncryptString(Password, FilePassword))
            End If
        End If
    Else
        FilePath = App.path & "\INIT\Personajes.ini"
        If Trim$(Name) <> "" Then
            If Trim$(Password) <> "" Then
                Call WriteVar(FilePath, "PJS", Encryption_RC4_EncryptString(Name, FilePassword), "")
            End If
        End If
    End If
End Sub

