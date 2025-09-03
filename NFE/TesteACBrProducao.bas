' TesteACBrProducao.bas
' Modulo de teste para ACBrLibMDFe no ambiente de producao
' Pasta: C:\Projetos\MDFe\NFE

Option Explicit

' ============================================================================
' DECLARACOES DAS FUNCOES ACBrLibMDFe
' ============================================================================

' Funcoes de controle da biblioteca
Private Declare Function MDFE_Inicializar Lib "ACBrMDFe32.dll" _
    (ByVal eArqConfig As String, ByVal eChaveCrypt As String) As Long

Private Declare Function MDFE_Finalizar Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_Nome Lib "ACBrMDFe32.dll" _
    (ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_Versao Lib "ACBrMDFe32.dll" _
    (ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_UltimoRetorno Lib "ACBrMDFe32.dll" _
    (ByVal buffer As String, ByVal bufferLen As Long) As Long

' Funcoes de configuracao
Private Declare Function MDFE_ConfigLerValor Lib "ACBrMDFe32.dll" _
    (ByVal eSessao As String, ByVal eChave As String, ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_ConfigGravarValor Lib "ACBrMDFe32.dll" _
    (ByVal eSessao As String, ByVal eChave As String, ByVal eValor As String) As Long

' Funcoes de status
Private Declare Function MDFE_StatusServico Lib "ACBrMDFe32.dll" () As Long

' ============================================================================
' FUNCOES DE TESTE
' ============================================================================

Public Function TestarInstalacaoCompleta() As Boolean
    On Error GoTo ErroTeste
    
    Dim resultado As Long
    Dim buffer As String
    Dim tamanho As Long
    
    Debug.Print "========================================"
    Debug.Print "TESTE COMPLETO - ACBrLibMDFe PRODUCAO"
    Debug.Print "Pasta: C:\Projetos\MDFe\NFE"
    Debug.Print "========================================"
    Debug.Print ""
    
    ' =============================
    ' TESTE 1: Inicializacao
    ' =============================
    Debug.Print "1. Testando inicializacao..."
    resultado = MDFE_Inicializar("C:\Projetos\MDFe\NFE\ACBrLibMDFe.ini", "")
    
    If resultado = 0 Then
        Debug.Print "   ✓ Inicializacao: SUCESSO"
        
        ' =============================
        ' TESTE 2: Informacoes da biblioteca
        ' =============================
        Debug.Print ""
        Debug.Print "2. Obtendo informacoes da biblioteca..."
        
        ' Nome da biblioteca
        buffer = String(256, vbNullChar)
        tamanho = MDFE_Nome(buffer, Len(buffer))
        If tamanho > 0 Then
            Debug.Print "   ✓ Nome: " & Left(buffer, tamanho)
        End If
        
        ' Versao
        buffer = String(256, vbNullChar)
        tamanho = MDFE_Versao(buffer, Len(buffer))
        If tamanho > 0 Then
            Debug.Print "   ✓ Versao: " & Left(buffer, tamanho)
        End If
        
        ' =============================
        ' TESTE 3: Leitura de configuracoes
        ' =============================
        Debug.Print ""
        Debug.Print "3. Testando leitura de configuracoes..."
        
        ' Ambiente configurado
        Dim ambiente As String
        ambiente = LerConfiguracao("DFe", "Ambiente")
        Debug.Print "   ✓ Ambiente configurado: " & IIf(ambiente = "1", "PRODUCAO", "HOMOLOGACAO") & " (" & ambiente & ")"
        
        ' UF configurada
        Dim uf As String
        uf = LerConfiguracao("DFe", "UF")
        Debug.Print "   ✓ UF configurada: " & uf
        
        ' Pasta de logs
        Dim pastaLogs As String
        pastaLogs = LerConfiguracao("Principal", "LogPath")
        Debug.Print "   ✓ Pasta de logs: " & pastaLogs
        
        ' =============================
        ' TESTE 4: Conectividade (opcional)
        ' =============================
        Debug.Print ""
        Debug.Print "4. Testando conectividade com SEFAZ..."
        Debug.Print "   (Este teste pode demorar alguns segundos...)"
        
        resultado = MDFE_StatusServico()
        
        If resultado = 0 Then
            Dim retornoStatus As String
            retornoStatus = ObterUltimoRetorno()
            
            If InStr(retornoStatus, "<cStat>107</cStat>") > 0 Then
                Debug.Print "   ✓ Conectividade: SUCESSO - Servico em operacao"
            ElseIf InStr(retornoStatus, "<cStat>") > 0 Then
                Debug.Print "   ✓ Conectividade: OK - Status: " & ExtrairEntreXML(retornoStatus, "cStat")
                Debug.Print "   ✓ Motivo: " & ExtrairEntreXML(retornoStatus, "xMotivo")
            Else
                Debug.Print "   ⚠ Conectividade: Resposta inesperada"
            End If
        Else
            Debug.Print "   ⚠ Conectividade: FALHOU - " & ObterUltimoRetorno()
            Debug.Print "   (Isso pode ser normal se nao houver certificado configurado)"
        End If
        
        ' =============================
        ' TESTE 5: Verificacao de arquivos
        ' =============================
        Debug.Print ""
        Debug.Print "5. Verificando estrutura de arquivos..."
        
        If Dir("C:\Projetos\MDFe\NFE\ACBrMDFe32.dll") <> "" Then
            Debug.Print "   ✓ ACBrMDFe32.dll: Encontrado"
        ElseIf Dir("C:\Projetos\MDFe\NFE\ACBrMDFe64.dll") <> "" Then
            Debug.Print "   ✓ ACBrMDFe64.dll: Encontrado"
        Else
            Debug.Print "   ✗ DLL principal: NAO ENCONTRADA"
        End If
        
        If Dir("C:\Projetos\MDFe\NFE\libxml2.dll") <> "" Then
            Debug.Print "   ✓ libxml2.dll: Encontrado"
        Else
            Debug.Print "   ⚠ libxml2.dll: Nao encontrado"
        End If
        
        If Dir("C:\Projetos\MDFe\NFE\Logs", vbDirectory) <> "" Then
            Debug.Print "   ✓ Pasta Logs: Existe"
        Else
            Debug.Print "   ⚠ Pasta Logs: Nao encontrada"
        End If
        
        If Dir("C:\Projetos\MDFe\NFE\XML", vbDirectory) <> "" Then
            Debug.Print "   ✓ Pasta XML: Existe"
        Else
            Debug.Print "   ⚠ Pasta XML: Nao encontrada"
        End If
        
        If Dir("C:\Projetos\MDFe\NFE\PDF", vbDirectory) <> "" Then
            Debug.Print "   ✓ Pasta PDF: Existe"
        Else
            Debug.Print "   ⚠ Pasta PDF: Nao encontrada"
        End If
        
        ' =============================
        ' FINALIZACAO
        ' =============================
        Call MDFE_Finalizar
        Debug.Print ""
        Debug.Print "6. Finalizacao: SUCESSO"
        
        Debug.Print ""
        Debug.Print "========================================"
        Debug.Print "*** TESTE CONCLUIDO COM SUCESSO! ***"
        Debug.Print "ACBrLibMDFe esta instalado e operacional"
        Debug.Print "========================================"
        
        TestarInstalacaoCompleta = True
        
    Else
        Debug.Print "   ✗ ERRO na inicializacao!"
        Debug.Print "   Codigo: " & resultado
        Debug.Print "   Detalhe: " & ObterUltimoRetorno()
        Debug.Print ""
        Debug.Print "========================================"
        Debug.Print "*** TESTE FALHOU! ***"
        Debug.Print "Verifique a instalacao da ACBrLibMDFe"
        Debug.Print "========================================"
        
        TestarInstalacaoCompleta = False
    End If
    
    Exit Function
    
ErroTeste:
    Debug.Print "   ✗ ERRO CRITICO no teste: " & Err.Description
    Debug.Print "   Numero: " & Err.Number
    TestarInstalacaoCompleta = False
End Function

' Funcao auxiliar para ler configuracoes
Private Function LerConfiguracao(sessao As String, chave As String) As String
    Dim buffer As String
    Dim resultado As Long
    
    buffer = String(1024, vbNullChar)
    resultado = MDFE_ConfigLerValor(sessao, chave, buffer, Len(buffer))
    
    If resultado > 0 Then
        LerConfiguracao = Left(buffer, resultado)
    Else
        LerConfiguracao = "(não configurado)"
    End If
End Function

' Funcao auxiliar para obter ultimo retorno
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

' Funcao auxiliar para extrair valores de XML
Private Function ExtrairEntreXML(xml As String, tag As String) As String
    Dim posIni As Long
    Dim posFim As Long
    
    posIni = InStr(xml, "<" & tag & ">")
    If posIni > 0 Then
        posIni = posIni + Len(tag) + 2
        posFim = InStr(posIni, xml, "</" & tag & ">")
        If posFim > posIni Then
            ExtrairEntreXML = Mid(xml, posIni, posFim - posIni)
        End If
    End If
End Function

' Teste simples para chamada rapida
Public Sub TesteRapido()
    Call TestarInstalacaoCompleta
End Sub

' Funcao para configurar certificado (para uso manual)
Public Function ConfigurarCertificado(caminhoArquivo As String, senha As String) As Boolean
    On Error GoTo ErroConfig
    
    Dim resultado As Long
    
    resultado = MDFE_Inicializar("C:\Projetos\MDFe\NFE\ACBrLibMDFe.ini", "")
    If resultado <> 0 Then
        ConfigurarCertificado = False
        Exit Function
    End If
    
    resultado = MDFE_ConfigGravarValor("Certificados", "Arquivo", caminhoArquivo)
    If resultado <> 0 Then
        Call MDFE_Finalizar
        ConfigurarCertificado = False
        Exit Function
    End If
    
    resultado = MDFE_ConfigGravarValor("Certificados", "Senha", senha)
    ConfigurarCertificado = (resultado = 0)
    
    Call MDFE_Finalizar
    Exit Function
    
ErroConfig:
    Call MDFE_Finalizar
    ConfigurarCertificado = False
End Function

' ============================================================================
' EXEMPLO DE USO NO IMMEDIATE WINDOW (VB6)
' ============================================================================
'
' Para executar os testes, digite no Immediate Window do VB6:
'
' TesteRapido
' 
' Ou:
' 
' ?TestarInstalacaoCompleta()
'
' Para configurar certificado:
'
' ?ConfigurarCertificado("C:\certificado.pfx", "senha123")
'
' ============================================================================