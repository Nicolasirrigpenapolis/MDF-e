' ConfiguradorACBr.bas
' Modulo para configurar ACBrLibMDFe baseado nas configuracoes existentes do NFE.INI
' Integração com o sistema atual

Option Explicit

' ============================================================================
' FUNCOES PARA INTEGRAR CONFIGURACOES DO NFE.INI COM ACBrLibMDFe
' ============================================================================

Public Function ConfigurarACBrComBaseNoSistema() As Boolean
    On Error GoTo ErroConfig
    
    Dim arquivoACBr As String
    Dim arquivoNFE As String
    Dim numero As Integer
    
    ' Caminhos dos arquivos
    arquivoACBr = "C:\Projetos\MDFe\NFE\ACBrLibMDFe.ini"
    arquivoNFE = "C:\Projetos\MDFe\NFE\NFE.INI"
    
    Debug.Print "Configurando ACBrLibMDFe baseado no sistema atual..."
    
    ' Verificar se arquivos existem
    If Dir(arquivoNFE) = "" Then
        Debug.Print "ERRO: Arquivo NFE.INI não encontrado em: " & arquivoNFE
        ConfigurarACBrComBaseNoSistema = False
        Exit Function
    End If
    
    ' Ler configurações do NFE.INI
    Dim bancoDados As String
    Dim servidor As String
    Dim diretorioConsultas As String
    
    bancoDados = LerINI(arquivoNFE, "DESKTOP-CTAJU78 - Geral", "Nome de DADOSNFE", "DADOSNFE")
    servidor = LerINI(arquivoNFE, "DESKTOP-CTAJU78 - Geral", "Server de DADOSNFE", "DESKTOP-CTAJU78\SQLEXPRESS02")
    diretorioConsultas = LerINI(arquivoNFE, "DESKTOP-CTAJU78 - Geral", "Diretório das consultas", "C:\Projetos\MDFe\NFE\")
    
    Debug.Print "Configurações lidas do NFE.INI:"
    Debug.Print "  Banco: " & bancoDados
    Debug.Print "  Servidor: " & servidor
    Debug.Print "  Diretório: " & diretorioConsultas
    
    ' Criar arquivo ACBrLibMDFe.ini personalizado
    numero = FreeFile
    Open arquivoACBr For Output As numero
    
    ' Cabeçalho
    Print #numero, "# =============================================================================="
    Print #numero, "# CONFIGURACAO ACBrLibMDFe - GERADA AUTOMATICAMENTE"
    Print #numero, "# Baseado no arquivo NFE.INI existente"
    Print #numero, "# Data: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #numero, "# =============================================================================="
    Print #numero, ""
    
    ' Configurações principais
    Print #numero, "[Principal]"
    Print #numero, "LogLevel=2"
    Print #numero, "LogPath=" & diretorioConsultas & "Logs\"
    Print #numero, "ArquivoUnico=0"
    Print #numero, ""
    
    ' Configurações DFe
    Print #numero, "[DFe]"
    Print #numero, "UF=SP"  ' Ajuste conforme necessário
    Print #numero, "Ambiente=1"  ' Produção - mude para 2 se quiser homologação
    Print #numero, "Visualizar=0"
    Print #numero, "SalvarWS=1"
    Print #numero, "RetirarAcentos=1"
    Print #numero, "FormatoAlerta=clAsterisco"
    Print #numero, "PathSchemas="
    Print #numero, "VersaoDFe=3.00"
    Print #numero, ""
    
    ' WebService
    Print #numero, "[WebService]"
    Print #numero, "UF=SP"
    Print #numero, "Ambiente=1"
    Print #numero, "Visualizar=0"
    Print #numero, "SalvarEnvio=1"
    Print #numero, "SalvarResposta=1"
    Print #numero, "AjustaAguardaConsultaRet=1"
    Print #numero, "AguardarConsultaRet=1000"
    Print #numero, "Tentativas=5"
    Print #numero, "IntervaloTentativas=2000"
    Print #numero, "TimeOut=60000"
    Print #numero, "ProxyHost="
    Print #numero, "ProxyPort="
    Print #numero, "ProxyUser="
    Print #numero, "ProxyPass="
    Print #numero, ""
    
    ' Certificados - deixar vazio para configuração manual
    Print #numero, "[Certificados]"
    Print #numero, "Arquivo="
    Print #numero, "Senha="
    Print #numero, "NumeroSerie="
    Print #numero, "CacheLib=1"
    Print #numero, "CryptoLib=1"
    Print #numero, "HttpLib=1"
    Print #numero, "XmlSignLib=1"
    Print #numero, ""
    
    ' Arquivos
    Print #numero, "[Arquivos]"
    Print #numero, "PastaMensal=1"
    Print #numero, "AddLiteral=0"
    Print #numero, "EmissaoPathMDFe=1"
    Print #numero, "SalvarEvento=1"
    Print #numero, "SepararPorCNPJ=0"
    Print #numero, "PathMDFe=" & diretorioConsultas & "XML\"
    Print #numero, "PathEvento=" & diretorioConsultas & "XML\"
    Print #numero, ""
    
    ' DAMDFE
    Print #numero, "[DAMDFE]"
    Print #numero, "TipoDAMDFE=0"
    Print #numero, "PathPDF=" & diretorioConsultas & "PDF\"
    Print #numero, "PathLogo="
    Print #numero, "Visualizar=1"
    Print #numero, "ImprimirHoraSaida=0"
    Print #numero, "ImprimirHoraSaida_Hora=12:00:00"
    Print #numero, "TamanhoPapel=0"
    Print #numero, "Margem_Sup=8"
    Print #numero, "Margem_Inf=8"
    Print #numero, "Margem_Esq=6"
    Print #numero, "Margem_Dir=6"
    Print #numero, "FonteDAMDFE_Nome=Times New Roman"
    Print #numero, "FonteDAMDFE_Tamanho=9"
    Print #numero, ""
    
    ' Email
    Print #numero, "[Email]"
    Print #numero, "Nome="
    Print #numero, "Email="
    Print #numero, "Usuario="
    Print #numero, "Senha="
    Print #numero, "Servidor="
    Print #numero, "Porta=587"
    Print #numero, "SSL=1"
    Print #numero, "TLS=1"
    Print #numero, "Assunto=Manifesto Eletronico de Documentos Fiscais"
    Print #numero, "Mensagem=Segue em anexo o Manifesto Eletronico de Documentos Fiscais (MDF-e)."
    Print #numero, ""
    
    ' Responsável Técnico
    Print #numero, "[RespTec]"
    Print #numero, "CNPJ="
    Print #numero, "xContato=Sistema MDFe"
    Print #numero, "email="
    Print #numero, "fone="
    Print #numero, "idCSRT=01"
    Print #numero, "hashCSRT="
    Print #numero, ""
    
    ' Banco de Dados - baseado no NFE.INI
    Print #numero, "[BancodeDados]"
    Print #numero, "Nome=" & bancoDados
    Print #numero, "Tipo=8"
    Print #numero, "Server=" & servidor
    Print #numero, "TrustedConnection=1"
    Print #numero, "Provider=SQLOLEDB.1"
    Print #numero, "Driver={SQL Server Native Client 10.0}"
    Print #numero, ""
    
    ' Sistema - baseado no NFE.INI
    Print #numero, "[Sistema]"
    Print #numero, "DiretorioConsultas=" & diretorioConsultas
    Print #numero, "TempoRefresh=25"
    Print #numero, "WindowState=2"
    Print #numero, ""
    
    ' Comentários finais
    Print #numero, "# =============================================================================="
    Print #numero, "# IMPORTANTE:"
    Print #numero, "# 1. Configure o certificado digital na seção [Certificados]"
    Print #numero, "# 2. Ajuste a UF conforme seu estado nas seções [DFe] e [WebService]"
    Print #numero, "# 3. Para testes, mude Ambiente=2 nas seções [DFe] e [WebService]"
    Print #numero, "# 4. Verifique se as pastas existem: Logs, XML, PDF"
    Print #numero, "# =============================================================================="
    
    Close numero
    
    ' Criar pastas se não existirem
    Call CriarPastaSe(diretorioConsultas & "Logs")
    Call CriarPastaSe(diretorioConsultas & "XML")
    Call CriarPastaSe(diretorioConsultas & "PDF")
    
    Debug.Print "✓ Arquivo ACBrLibMDFe.ini criado com sucesso!"
    Debug.Print "✓ Pastas criadas: Logs, XML, PDF"
    Debug.Print ""
    Debug.Print "PRÓXIMOS PASSOS:"
    Debug.Print "1. Configure o certificado digital"
    Debug.Print "2. Ajuste a UF se necessário"
    Debug.Print "3. Teste em homologação primeiro (Ambiente=2)"
    
    ConfigurarACBrComBaseNoSistema = True
    Exit Function
    
ErroConfig:
    If numero > 0 Then Close numero
    Debug.Print "ERRO ao configurar ACBr: " & Err.Description
    ConfigurarACBrComBaseNoSistema = False
End Function

' Função para ler valores do arquivo INI
Private Function LerINI(arquivoINI As String, secao As String, chave As String, valorPadrao As String) As String
    Dim buffer As String
    Dim tamanho As Long
    
    buffer = String(255, vbNullChar)
    tamanho = GetPrivateProfileString(secao, chave, valorPadrao, buffer, Len(buffer), arquivoINI)
    
    If tamanho > 0 Then
        LerINI = Left(buffer, tamanho)
    Else
        LerINI = valorPadrao
    End If
End Function

' Função para criar pasta se não existir
Private Sub CriarPastaSe(caminho As String)
    On Error Resume Next
    If Dir(caminho, vbDirectory) = "" Then
        MkDir caminho
        Debug.Print "✓ Pasta criada: " & caminho
    End If
    On Error GoTo 0
End Sub

' Função para atualizar configurações específicas no ACBrLibMDFe.ini
Public Function AtualizarConfiguracaoACBr(secao As String, chave As String, valor As String) As Boolean
    On Error GoTo ErroAtualizar
    
    Dim arquivoACBr As String
    arquivoACBr = "C:\Projetos\MDFe\NFE\ACBrLibMDFe.ini"
    
    ' Usar API do Windows para escrever no INI
    Call WritePrivateProfileString(secao, chave, valor, arquivoACBr)
    
    Debug.Print "✓ Configuração atualizada: [" & secao & "] " & chave & "=" & valor
    AtualizarConfiguracaoACBr = True
    Exit Function
    
ErroAtualizar:
    Debug.Print "ERRO ao atualizar configuração: " & Err.Description
    AtualizarConfiguracaoACBr = False
End Function

' Função para configurar certificado A1
Public Function ConfigurarCertificadoA1(caminhoArquivo As String, senha As String) As Boolean
    ConfigurarCertificadoA1 = AtualizarConfiguracaoACBr("Certificados", "Arquivo", caminhoArquivo) And _
                              AtualizarConfiguracaoACBr("Certificados", "Senha", senha) And _
                              AtualizarConfiguracaoACBr("Certificados", "NumeroSerie", "")
End Function

' Função para configurar certificado A3
Public Function ConfigurarCertificadoA3(numeroSerie As String) As Boolean
    ConfigurarCertificadoA3 = AtualizarConfiguracaoACBr("Certificados", "Arquivo", "") And _
                              AtualizarConfiguracaoACBr("Certificados", "Senha", "") And _
                              AtualizarConfiguracaoACBr("Certificados", "NumeroSerie", numeroSerie)
End Function

' Função para alternar entre produção e homologação
Public Function AlterarAmbiente(ambiente As Integer) As Boolean
    Dim ambienteStr As String
    ambienteStr = CStr(ambiente)
    
    AlterarAmbiente = AtualizarConfiguracaoACBr("DFe", "Ambiente", ambienteStr) And _
                      AtualizarConfiguracaoACBr("WebService", "Ambiente", ambienteStr)
    
    If AlterarAmbiente Then
        Debug.Print "✓ Ambiente alterado para: " & IIf(ambiente = 1, "PRODUÇÃO", "HOMOLOGAÇÃO")
    End If
End Function

' Função para configurar UF
Public Function ConfigurarUF(uf As String) As Boolean
    ConfigurarUF = AtualizarConfiguracaoACBr("DFe", "UF", uf) And _
                   AtualizarConfiguracaoACBr("WebService", "UF", uf)
    
    If ConfigurarUF Then
        Debug.Print "✓ UF configurado para: " & uf
    End If
End Function

' Declarações da API do Windows para INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, _
     ByVal lpFileName As String) As Long

' ============================================================================
' EXEMPLOS DE USO NO IMMEDIATE WINDOW (VB6):
' 
' Para configurar automaticamente baseado no NFE.INI:
' ?ConfigurarACBrComBaseNoSistema()
' 
' Para configurar certificado A1:
' ?ConfigurarCertificadoA1("C:\certificado.pfx", "minhasenha")
' 
' Para configurar certificado A3:
' ?ConfigurarCertificadoA3("1234567890")
' 
' Para alterar para homologação:
' ?AlterarAmbiente(2)
' 
' Para alterar para produção:
' ?AlterarAmbiente(1)
' 
' Para configurar UF:
' ?ConfigurarUF("SP")
' ============================================================================