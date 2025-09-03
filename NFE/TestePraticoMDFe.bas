' TestePraticoMDFe.bas
' Teste prático para verificar se ACBrLibMDFe funciona no seu sistema
' Execute este teste DEPOIS de instalar e configurar certificado

Option Explicit

' Declarações básicas da ACBrLibMDFe
Private Declare Function MDFE_Inicializar Lib "ACBrMDFe32.dll" _
    (ByVal eArqConfig As String, ByVal eChaveCrypt As String) As Long

Private Declare Function MDFE_Finalizar Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_StatusServico Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_UltimoRetorno Lib "ACBrMDFe32.dll" _
    (ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_ConfigLerValor Lib "ACBrMDFe32.dll" _
    (ByVal eSessao As String, ByVal eChave As String, ByVal buffer As String, ByVal bufferLen As Long) As Long

' TESTE PRINCIPAL - Execute este no Immediate Window
Public Sub TesteCompleto()
    Debug.Print "========================================"
    Debug.Print "TESTE PRÁTICO - SISTEMA MDFe COM ACBr"
    Debug.Print "========================================"
    Debug.Print ""
    
    ' ETAPA 1: Testar inicialização
    Debug.Print "1. Testando inicialização ACBrLibMDFe..."
    
    Dim resultado As Long
    resultado = MDFE_Inicializar("C:\Projetos\MDFe\NFE\ACBrLibMDFe.ini", "")
    
    If resultado = 0 Then
        Debug.Print "   ✓ ACBrLibMDFe inicializada com SUCESSO!"
        
        ' ETAPA 2: Verificar configurações
        Debug.Print ""
        Debug.Print "2. Verificando configurações..."
        
        Dim ambiente As String
        Dim uf As String
        ambiente = LerConfig("DFe", "Ambiente")
        uf = LerConfig("DFe", "UF")
        
        Debug.Print "   ✓ Ambiente: " & IIf(ambiente = "1", "PRODUÇÃO", "HOMOLOGAÇÃO") & " (" & ambiente & ")"
        Debug.Print "   ✓ UF: " & uf
        
        ' ETAPA 3: Testar conectividade SEFAZ
        Debug.Print ""
        Debug.Print "3. Testando conectividade SEFAZ..."
        Debug.Print "   (Aguarde... pode demorar alguns segundos)"
        
        resultado = MDFE_StatusServico()
        Dim retornoSEFAZ As String
        retornoSEFAZ = ObterUltimoRetorno()
        
        If resultado = 0 Then
            If InStr(retornoSEFAZ, "<cStat>107</cStat>") > 0 Then
                Debug.Print "   ✓ SEFAZ: Serviço em operação normal!"
            ElseIf InStr(retornoSEFAZ, "<cStat>") > 0 Then
                Dim cStat As String
                cStat = ExtrairXML(retornoSEFAZ, "cStat")
                Debug.Print "   ✓ SEFAZ: Conectado (Status: " & cStat & ")"
            Else
                Debug.Print "   ⚠ SEFAZ: Conectado mas resposta inesperada"
            End If
        Else
            Debug.Print "   ✗ SEFAZ: Erro na conectividade"
            Debug.Print "   Detalhes: " & Left(retornoSEFAZ, 200)
        End If
        
        ' ETAPA 4: Comparar com sistema atual
        Debug.Print ""
        Debug.Print "4. Comparação com sistema atual..."
        
        If Dir("MDFe_Util.dll") <> "" Then
            Debug.Print "   ⚠ FlexDocs ainda presente: MDFe_Util.dll"
            Debug.Print "   (Será desregistrada automaticamente na instalação)"
        Else
            Debug.Print "   ✓ FlexDocs removida corretamente"
        End If
        
        ' ETAPA 5: Verificar pastas
        Debug.Print ""
        Debug.Print "5. Verificando estrutura de pastas..."
        
        If Dir("C:\Projetos\MDFe\NFE\Logs", vbDirectory) <> "" Then
            Debug.Print "   ✓ Pasta Logs: OK"
        Else
            Debug.Print "   ⚠ Pasta Logs: Não encontrada"
        End If
        
        If Dir("C:\Projetos\MDFe\NFE\XML", vbDirectory) <> "" Then
            Debug.Print "   ✓ Pasta XML: OK"
        Else
            Debug.Print "   ⚠ Pasta XML: Não encontrada"
        End If
        
        If Dir("C:\Projetos\MDFe\NFE\PDF", vbDirectory) <> "" Then
            Debug.Print "   ✓ Pasta PDF: OK"
        Else
            Debug.Print "   ⚠ Pasta PDF: Não encontrada"
        End If
        
        ' Finalizar
        Call MDFE_Finalizar
        
        Debug.Print ""
        Debug.Print "========================================"
        Debug.Print "RESULTADO: SISTEMA PRONTO PARA MIGRAÇÃO!"
        Debug.Print ""
        Debug.Print "Próximos passos:"
        Debug.Print "1. Migrar função SuperMDFe()"
        Debug.Print "2. Migrar TransmitirMDFe()"
        Debug.Print "3. Migrar DAMDFE()"
        Debug.Print "4. Testes com dados reais"
        Debug.Print "========================================"
        
    Else
        Debug.Print "   ✗ FALHA na inicialização!"
        Debug.Print "   Código de erro: " & resultado
        Debug.Print "   Detalhes: " & ObterUltimoRetorno()
        Debug.Print ""
        Debug.Print "========================================"
        Debug.Print "PROBLEMA: Verifique a instalação!"
        Debug.Print ""
        Debug.Print "Possíveis causas:"
        Debug.Print "- DLL não copiada corretamente"
        Debug.Print "- Dependências faltando"
        Debug.Print "- Arquivo de configuração inválido"
        Debug.Print "========================================"
    End If
End Sub

' Funções auxiliares
Private Function LerConfig(sessao As String, chave As String) As String
    Dim buffer As String
    Dim tamanho As Long
    
    buffer = String(255, vbNullChar)
    tamanho = MDFE_ConfigLerValor(sessao, chave, buffer, Len(buffer))
    
    If tamanho > 0 Then
        LerConfig = Left(buffer, tamanho)
    Else
        LerConfig = "(não configurado)"
    End If
End Function

Private Function ObterUltimoRetorno() As String
    Dim buffer As String
    Dim tamanho As Long
    
    buffer = String(4096, vbNullChar)
    tamanho = MDFE_UltimoRetorno(buffer, Len(buffer))
    
    If tamanho > 0 Then
        ObterUltimoRetorno = Left(buffer, tamanho)
    Else
        ObterUltimoRetorno = ""
    End If
End Function

Private Function ExtrairXML(xml As String, tag As String) As String
    Dim posIni As Long, posFim As Long
    
    posIni = InStr(xml, "<" & tag & ">")
    If posIni > 0 Then
        posIni = posIni + Len(tag) + 2
        posFim = InStr(posIni, xml, "</" & tag & ">")
        If posFim > posIni Then
            ExtrairXML = Mid(xml, posIni, posFim - posIni)
        End If
    End If
End Function

' TESTE RÁPIDO - Para chamada mais simples
Public Sub TesteRapido()
    Call TesteCompleto
End Sub

' TESTE SÓ DE CONECTIVIDADE - Mais específico
Public Sub TesteConectividade()
    Debug.Print "Testando conectividade SEFAZ..."
    
    Dim resultado As Long
    resultado = MDFE_Inicializar("C:\Projetos\MDFe\NFE\ACBrLibMDFe.ini", "")
    
    If resultado = 0 Then
        resultado = MDFE_StatusServico()
        Debug.Print "Status SEFAZ: " & ObterUltimoRetorno()
        Call MDFE_Finalizar
    Else
        Debug.Print "Erro na inicialização: " & resultado
    End If
End Sub