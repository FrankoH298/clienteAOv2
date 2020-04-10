Attribute VB_Name = "Mod_EncryptRC4"
Option Explicit
 
Private m_sBoxRC4(0 To 255) As Integer
 
'Key-dependant
Private m_KeyS As String
 
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 
Public Sub Encryption_RC4_DecryptByte(ByteArray() As Byte, Optional Key As String)
'*****************************************************************
'Decryptes a byte array with RC4 encryption
'More info: [url=http://www.vbgore.com/CommonCode.Encryptions.Encryption_RC4_DecryptByte]http://www.vbgore.com/CommonCode.Encryp ... ecryptByte[/url]
'*****************************************************************
 
    Call Encryption_RC4_EncryptByte(ByteArray(), Key)
 
End Sub
 
 
Public Function Encryption_RC4_DecryptString(Text As String, Optional Key As String) As String
'*****************************************************************
'Decrypts a string array with RC4 encryption
'More info: [url=http://www.vbgore.com/CommonCode.Encryptions.Encryption_RC4_DecryptString]http://www.vbgore.com/CommonCode.Encryp ... ryptString[/url]
'*****************************************************************
Dim ByteArray() As Byte
 
'Convert the data into a byte array
 
    ByteArray() = StrConv(Text, vbFromUnicode)
 
    'Decrypt the byte array
    Call Encryption_RC4_DecryptByte(ByteArray(), Key)
 
    'Convert the byte array back into a string
    Encryption_RC4_DecryptString = StrConv(ByteArray(), vbUnicode)
 
End Function
 
Public Sub Encryption_RC4_EncryptByte(ByteArray() As Byte, Optional Key As String)
'*****************************************************************
'Encrypts a byte array with RC4 encryption
'More info: [url=http://www.vbgore.com/CommonCode.Encryptions.Encryption_RC4_EncryptByte]http://www.vbgore.com/CommonCode.Encryp ... ncryptByte[/url]
'*****************************************************************
Dim i As Long
Dim J As Long
Dim temp As Byte
Dim Offset As Long
Dim OrigLen As Long
Dim sBox(0 To 255) As Integer
 
    'Set the new key (optional)
    If (Len(Key) > 0) Then Encryption_RC4_SetKey Key
 
    'Create a local copy of the sboxes, this
    'is much more elegant than recreating
    'before encrypting/decrypting anything
    Call CopyMem(sBox(0), m_sBoxRC4(0), 512)
 
    'Get the size of the source array
    OrigLen = UBound(ByteArray) + 1
 
    'Encrypt the data
    For Offset = 0 To (OrigLen - 1)
        i = (i + 1) Mod 256
        J = (J + sBox(i)) Mod 256
        temp = sBox(i)
        sBox(i) = sBox(J)
        sBox(J) = temp
        ByteArray(Offset) = ByteArray(Offset) Xor (sBox((sBox(i) + sBox(J)) Mod 256))
    Next
 
End Sub
 
 
Public Function Encryption_RC4_EncryptString(Text As String, Optional Key As String) As String
'*****************************************************************
'Encrypts a string with RC4 encryption
'More info: [url=http://www.vbgore.com/CommonCode.Encryptions.Encryption_RC4_EncryptString]http://www.vbgore.com/CommonCode.Encryp ... ryptString[/url]
'*****************************************************************
Dim ByteArray() As Byte
 
    'Convert the data into a byte array
    ByteArray() = StrConv(Text, vbFromUnicode)
 
    'Encrypt the byte array
    Call Encryption_RC4_EncryptByte(ByteArray(), Key)
 
    'Convert the byte array back into a string
    Encryption_RC4_EncryptString = StrConv(ByteArray(), vbUnicode)
 
End Function
 
Public Sub Encryption_RC4_SetKey(New_Value As String)
'*****************************************************************
'Sets the encryption key for RC4 encryption
'More info: [url=http://www.vbgore.com/CommonCode.Encryptions.Encryption_RC4_SetKey]http://www.vbgore.com/CommonCode.Encryp ... RC4_SetKey[/url]
'*****************************************************************
Dim a As Long
Dim b As Long
Dim temp As Byte
Dim Key() As Byte
Dim KeyLen As Long
 
    'Do nothing if the key is buffered
    If (m_KeyS = New_Value) Then Exit Sub
 
    'Set the new key
    m_KeyS = New_Value
 
    'Save the password in a byte array
    Key() = StrConv(m_KeyS, vbFromUnicode)
    KeyLen = Len(m_KeyS)
 
    'Initialize s-boxes
    For a = 0 To 255
        m_sBoxRC4(a) = a
    Next a
    For a = 0 To 255
        b = (b + m_sBoxRC4(a) + Key(a Mod KeyLen)) Mod 256
        temp = m_sBoxRC4(a)
        m_sBoxRC4(a) = m_sBoxRC4(b)
        m_sBoxRC4(b) = temp
    Next
 
End Sub
