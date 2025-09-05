Attribute VB_Name = "ACBrComum"
Option Explicit

' APIs do Windows para conversão UTF-8
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long
    
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

' Constantes para conversão UTF-8
Private Const CP_UTF8 As Long = 65001
Private Const CP_ACP As Long = 0

' Função para converter string UTF-8 para ANSI (VB6)
Public Function FromUTF8(ByVal sUtf As String) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    
    If Len(sUtf) = 0 Then
        FromUTF8 = ""
        Exit Function
    End If
    
    ' Remove caracteres nulos do final
    sUtf = RTrim$(sUtf)
    If Len(sUtf) = 0 Then
        FromUTF8 = ""
        Exit Function
    End If
    
    ' Primeiro passo: converter UTF-8 para Unicode
    lLength = Len(sUtf)
    lRet = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sUtf), lLength, 0, 0)
    
    If lRet > 0 Then
        Dim sUnicode As String
        sUnicode = String$(lRet, Chr$(0))
        lRet = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sUtf), lLength, StrPtr(sUnicode), lRet)
        
        If lRet > 0 Then
            ' Segundo passo: converter Unicode para ANSI
            lBufferSize = WideCharToMultiByte(CP_ACP, 0, StrPtr(sUnicode), lRet, 0, 0, 0, 0)
            
            If lBufferSize > 0 Then
                Dim sResult As String
                sResult = String$(lBufferSize, Chr$(0))
                lRet = WideCharToMultiByte(CP_ACP, 0, StrPtr(sUnicode), Len(sUnicode), StrPtr(sResult), lBufferSize, 0, 0)
                
                If lRet > 0 Then
                    FromUTF8 = Left$(sResult, lRet)
                Else
                    FromUTF8 = sUtf
                End If
            Else
                FromUTF8 = sUtf
            End If
        Else
            FromUTF8 = sUtf
        End If
    Else
        FromUTF8 = sUtf
    End If
End Function

' Função para converter string ANSI (VB6) para UTF-8
Public Function ToUTF8(ByVal sAnsi As String) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    
    If Len(sAnsi) = 0 Then
        ToUTF8 = ""
        Exit Function
    End If
    
    ' Primeiro passo: converter ANSI para Unicode
    lLength = Len(sAnsi)
    lRet = MultiByteToWideChar(CP_ACP, 0, StrPtr(sAnsi), lLength, 0, 0)
    
    If lRet > 0 Then
        Dim sUnicode As String
        sUnicode = String$(lRet, Chr$(0))
        lRet = MultiByteToWideChar(CP_ACP, 0, StrPtr(sAnsi), lLength, StrPtr(sUnicode), lRet)
        
        If lRet > 0 Then
            ' Segundo passo: converter Unicode para UTF-8
            lBufferSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sUnicode), lRet, 0, 0, 0, 0)
            
            If lBufferSize > 0 Then
                Dim sResult As String
                sResult = String$(lBufferSize, Chr$(0))
                lRet = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sUnicode), Len(sUnicode), StrPtr(sResult), lBufferSize, 0, 0)
                
                If lRet > 0 Then
                    ToUTF8 = Left$(sResult, lRet)
                Else
                    ToUTF8 = sAnsi
                End If
            Else
                ToUTF8 = sAnsi
            End If
        Else
            ToUTF8 = sAnsi
        End If
    Else
        ToUTF8 = sAnsi
    End If
End Function

' Função auxiliar para verificar se uma string contém caracteres UTF-8
Public Function IsUTF8(ByVal sText As String) As Boolean
    Dim i As Long
    Dim bChar As Byte
    
    IsUTF8 = False
    
    For i = 1 To Len(sText)
        bChar = Asc(Mid$(sText, i, 1))
        If bChar > 127 Then
            IsUTF8 = True
            Exit For
        End If
    Next i
End Function
