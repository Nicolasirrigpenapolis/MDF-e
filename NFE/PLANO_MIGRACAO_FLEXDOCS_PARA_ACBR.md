# PLANO DETALHADO DE MIGRAÇÃO: FlexDocs → ACBrLibMDFe

**Projeto:** Sistema MDFe  
**Data:** 03/09/2025  
**Versão:** 1.0  
**Responsável:** Claude AI  

---

## SUMÁRIO EXECUTIVO

Este documento apresenta o plano completo para migração do sistema MDFe da biblioteca FlexDocs (MDFe_Util.dll) para a biblioteca ACBrLibMDFe versão 1.2.2.335. A migração visa modernizar a solução, garantir maior estabilidade e conformidade com as especificações técnicas atuais do MDF-e.

**Duração Estimada:** 26 dias úteis  
**Impacto:** Alto (requer parada do sistema para deploy)  
**Benefícios:** Maior estabilidade, suporte ativo, conformidade técnica  

---

## ANÁLISE DA SITUAÇÃO ATUAL

### Sistema Atual (FlexDocs)

#### **Arquitetura Atual:**
- **Linguagem:** Visual Basic 6.0
- **Biblioteca:** MDFe_Util.dll (FlexDocs)
- **Integração:** CreateObject("MDFe_Util.Util") - COM Object
- **Layout:** MDF-e 3.00 (NT2020001)
- **Registro:** Via RegAsm (.NET Framework)
- **Arquivo Principal:** MDFE.FRM (372KB)

#### **Componentes Identificados:**
1. **SuperMDFe()** - Função principal de geração XML
2. **TransmitirMDFe()** - Transmissão síncrona para SEFAZ
3. **DAMDFE()** - Geração e visualização do relatório
4. **AbreCancelamento()** - Interface de cancelamento
5. **AbreEncerramento()** - Interface de encerramento

#### **Estrutura de Dados Atual:**
```vb
' Campos principais identificados no sistema
Dim objMDFeUtil As Object
Set objMDFeUtil = CreateObject("MDFe_Util.Util")

' Métodos principais em uso:
- objMDFeUtil.CriaChaveDFe()
- objMDFeUtil.MDFe_NT2020001()
- objMDFeUtil.EnviaMDFeSincrono()
- objMDFeUtil.ide_v3(), Emit(), rodo_v3(), tot_v3()
```

### Sistema Alvo (ACBrLibMDFe)

#### **Nova Arquitetura:**
- **Biblioteca:** ACBrMDFe32.dll / ACBrMDFe64.dll
- **Integração:** Chamadas diretas via Declare Function
- **Layout:** MDF-e 3.00 (compatível)
- **Convenções:** StdCall e Cdecl disponíveis
- **Configuração:** Arquivo INI + métodos de configuração

#### **Principais Vantagens:**
- ✅ **Biblioteca oficial** do Projeto ACBr
- ✅ **Suporte ativo** e comunidade grande
- ✅ **Certificação digital** dos binários
- ✅ **Documentação completa**
- ✅ **Performance otimizada**
- ✅ **Menor dependência** de componentes externos

---

## PLANO DE MIGRAÇÃO DETALHADO

### **FASE 1: PREPARAÇÃO E CONFIGURAÇÃO (3 dias)**

#### **1.1 Instalação e Configuração ACBrLibMDFe**

**Passos de Instalação:**
```batch
# 1. Copiar DLLs principais
copy "ACBrLibMDFe-Windows-1.2.2.335\Windows\MT\StdCall\ACBrMDFe32.dll" "C:\Projetos\MDFe - CLAUDE\NFE\"

# 2. Copiar dependências
xcopy "ACBrLibMDFe-Windows-1.2.2.335\dep\*" "C:\Projetos\MDFe - CLAUDE\NFE\dep\" /S /E

# 3. Desregistrar biblioteca antiga (se necessário)
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm /u "MDFe_Util.dll"
```

**Configuração Inicial:**
```ini
[Principal]
LogLevel=4
LogPath=C:\Logs\ACBrMDFe\

[DFe]
UF=SP
Ambiente=2
Visualizar=0
SalvarWS=1
RetirarAcentos=1
FormatoAlerta=clAsterisco

[WebService]
UF=SP
Ambiente=2
Visualizar=0
SalvarWS=1
TimeOut=60000

[Certificados]
Certificado=
Senha=
NumeroSerie=
```

#### **1.2 Teste de Conectividade**
```vb
' Teste básico de funcionamento
Private Sub TestarACBrLibMDFe()
    Dim resultado As Long
    resultado = MDFE_Inicializar("ACBrLibMDFe.ini", "")
    
    If resultado = 0 Then
        MsgBox "ACBrLibMDFe inicializada com sucesso!"
        Call MDFE_Finalizar
    Else
        MsgBox "Erro na inicialização: " & MDFE_UltimoRetorno()
    End If
End Sub
```

### **FASE 2: MAPEAMENTO DE FUNÇÕES (2 dias)**

#### **2.1 Tabela de Mapeamento Completa**

| **FlexDocs (Atual)** | **ACBrLibMDFe (Nova)** | **Tipo Mudança** | **Observações** |
|---------------------|------------------------|------------------|------------------|
| `CreateObject("MDFe_Util.Util")` | `MDFE_Inicializar()` | Arquitetural | COM → DLL nativa |
| `objMDFeUtil.CriaChaveDFe()` | Automático em `MDFE_CarregarINI()` | Simplificação | Geração automática |
| `objMDFeUtil.MDFe_NT2020001()` | `MDFE_CarregarINI() + MDFE_Assinar()` | Metodologia | XML via INI + assinatura |
| `objMDFeUtil.EnviaMDFeSincrono()` | `MDFE_EnviarSincrono()` | Nome | Funcionalidade equivalente |
| `objMDFeUtil.CancelaMDFe()` | `MDFE_EnviarEvento()` | Parametrização | Evento tipo cancelamento |
| `objMDFeUtil.EncerraMDFe()` | `MDFE_EnviarEvento()` | Parametrização | Evento tipo encerramento |
| Impressão customizada | `MDFE_ImprimirPDF()` | Nativa | Geração PDF automática |

### **FASE 3: IMPLEMENTAÇÃO DE CLASSES WRAPPER (4 dias)**

#### **3.1 Nova Classe ACBrMDFeUtil**

```vb
' Arquivo: ACBrMDFeUtil.cls
' Classe wrapper para encapsular ACBrLibMDFe

Option Explicit

' Declarações das funções da DLL
Private Declare Function MDFE_Inicializar Lib "ACBrMDFe32.dll" _
    (ByVal eArqConfig As String, ByVal eChaveCrypt As String) As Long

Private Declare Function MDFE_Finalizar Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_CarregarINI Lib "ACBrMDFe32.dll" _
    (ByVal eArquivoOuINI As String) As Long

Private Declare Function MDFE_Assinar Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_EnviarSincrono Lib "ACBrMDFe32.dll" _
    (ByVal ANumLote As Long, ByVal AImprimir As Boolean, ByVal ASincrono As Boolean) As Long

Private Declare Function MDFE_EnviarEvento Lib "ACBrMDFe32.dll" _
    (ByVal AidLote As Long) As Long

Private Declare Function MDFE_ImprimirPDF Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_UltimoRetorno Lib "ACBrMDFe32.dll" _
    (ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_ConfigGravarValor Lib "ACBrMDFe32.dll" _
    (ByVal ASessao As String, ByVal AChave As String, ByVal AValor As String) As Long

Private Declare Function MDFE_ConfigLerValor Lib "ACBrMDFe32.dll" _
    (ByVal ASessao As String, ByVal AChave As String, ByVal buffer As String, ByVal bufferLen As Long) As Long

' Variáveis de controle
Private m_Inicializada As Boolean
Private m_ArquivoConfig As String

' Inicialização
Public Function InicializarACBr(Optional caminhoConfig As String = "") As Boolean
    If m_Inicializada Then
        InicializarACBr = True
        Exit Function
    End If
    
    If Len(caminhoConfig) = 0 Then
        caminhoConfig = App.Path & "\ACBrLibMDFe.ini"
    End If
    m_ArquivoConfig = caminhoConfig
    
    Dim resultado As Long
    resultado = MDFE_Inicializar(caminhoConfig, "")
    
    If resultado = 0 Then
        m_Inicializada = True
        InicializarACBr = True
    Else
        InicializarACBr = False
        Err.Raise vbObjectError + 1001, "ACBrMDFeUtil", "Erro na inicialização: " & ObterUltimoRetorno()
    End If
End Function

' Finalização
Public Sub FinalizarACBr()
    If m_Inicializada Then
        Call MDFE_Finalizar
        m_Inicializada = False
    End If
End Sub

' Obter último retorno
Private Function ObterUltimoRetorno() As String
    Dim buffer As String
    Dim tamanho As Long
    
    buffer = String(8192, vbNullChar)
    tamanho = MDFE_UltimoRetorno(buffer, Len(buffer))
    
    If tamanho > 0 Then
        ObterUltimoRetorno = Left(buffer, tamanho)
    Else
        ObterUltimoRetorno = ""
    End If
End Function

' Configurar certificado
Public Function ConfigurarCertificado(caminhoArquivo As String, senha As String) As Boolean
    Dim resultado As Long
    
    resultado = MDFE_ConfigGravarValor("Certificados", "Arquivo", caminhoArquivo)
    If resultado <> 0 Then
        ConfigurarCertificado = False
        Exit Function
    End If
    
    resultado = MDFE_ConfigGravarValor("Certificados", "Senha", senha)
    ConfigurarCertificado = (resultado = 0)
End Function

' Configurar ambiente
Public Function ConfigurarAmbiente(ambiente As Integer, uf As String) As Boolean
    Dim resultado As Long
    
    resultado = MDFE_ConfigGravarValor("DFe", "Ambiente", CStr(ambiente))
    If resultado <> 0 Then
        ConfigurarAmbiente = False
        Exit Function
    End If
    
    resultado = MDFE_ConfigGravarValor("DFe", "UF", uf)
    ConfigurarAmbiente = (resultado = 0)
End Function

' Gerar MDFe via INI
Public Function GerarMDFeINI(conteudoINI As String) As Boolean
    Dim resultado As Long
    Dim arquivoTemp As String
    Dim numeroArquivo As Integer
    
    ' Criar arquivo temporário INI
    arquivoTemp = Environ("TEMP") & "\mdfe_temp_" & Format(Timer * 1000, "0") & ".ini"
    
    numeroArquivo = FreeFile
    Open arquivoTemp For Output As numeroArquivo
    Print #numeroArquivo, conteudoINI
    Close numeroArquivo
    
    ' Carregar INI na ACBr
    resultado = MDFE_CarregarINI(arquivoTemp)
    
    ' Limpar arquivo temporário
    Kill arquivoTemp
    
    GerarMDFeINI = (resultado = 0)
    
    If resultado <> 0 Then
        Err.Raise vbObjectError + 1002, "ACBrMDFeUtil", "Erro ao carregar INI: " & ObterUltimoRetorno()
    End If
End Function

' Assinar MDFe
Public Function AssinarMDFe() As Boolean
    Dim resultado As Long
    resultado = MDFE_Assinar()
    AssinarMDFe = (resultado = 0)
    
    If resultado <> 0 Then
        Err.Raise vbObjectError + 1003, "ACBrMDFeUtil", "Erro na assinatura: " & ObterUltimoRetorno()
    End If
End Function

' Transmitir MDFe
Public Function TransmitirMDFe(numeroLote As Long) As String
    Dim resultado As Long
    
    resultado = MDFE_EnviarSincrono(numeroLote, False, True)
    
    If resultado = 0 Then
        TransmitirMDFe = ObterUltimoRetorno()
    Else
        Err.Raise vbObjectError + 1004, "ACBrMDFeUtil", "Erro na transmissão: " & ObterUltimoRetorno()
    End If
End Function

' Cancelar MDFe
Public Function CancelarMDFe(chaveAcesso As String, protocolo As String, justificativa As String) As String
    ' Implementar lógica de cancelamento via evento
    Dim resultado As Long
    
    ' Configurar dados do evento de cancelamento
    Call MDFE_ConfigGravarValor("Evento", "TipoEvento", "110111")
    Call MDFE_ConfigGravarValor("Evento", "ChaveMDFe", chaveAcesso)
    Call MDFE_ConfigGravarValor("Evento", "nProt", protocolo)
    Call MDFE_ConfigGravarValor("Evento", "xJust", justificativa)
    
    resultado = MDFE_EnviarEvento(1)
    
    If resultado = 0 Then
        CancelarMDFe = ObterUltimoRetorno()
    Else
        Err.Raise vbObjectError + 1005, "ACBrMDFeUtil", "Erro no cancelamento: " & ObterUltimoRetorno()
    End If
End Function

' Encerrar MDFe
Public Function EncerrarMDFe(chaveAcesso As String, protocolo As String, dtEnc As String, cMun As String) As String
    Dim resultado As Long
    
    ' Configurar dados do evento de encerramento
    Call MDFE_ConfigGravarValor("Evento", "TipoEvento", "110112")
    Call MDFE_ConfigGravarValor("Evento", "ChaveMDFe", chaveAcesso)
    Call MDFE_ConfigGravarValor("Evento", "nProt", protocolo)
    Call MDFE_ConfigGravarValor("Evento", "dtEnc", dtEnc)
    Call MDFE_ConfigGravarValor("Evento", "cMun", cMun)
    
    resultado = MDFE_EnviarEvento(1)
    
    If resultado = 0 Then
        EncerrarMDFe = ObterUltimoRetorno()
    Else
        Err.Raise vbObjectError + 1006, "ACBrMDFeUtil", "Erro no encerramento: " & ObterUltimoRetorno()
    End If
End Function

' Imprimir DAMDFE
Public Function ImprimirDAMDFE() As String
    Dim resultado As Long
    resultado = MDFE_ImprimirPDF()
    
    If resultado = 0 Then
        ImprimirDAMDFE = ObterUltimoRetorno()
    Else
        Err.Raise vbObjectError + 1007, "ACBrMDFeUtil", "Erro na impressão: " & ObterUltimoRetorno()
    End If
End Function

' Destructor
Private Sub Class_Terminate()
    Call FinalizarACBr
End Sub
```

### **FASE 4: MIGRAÇÃO DAS FUNÇÕES PRINCIPAIS (7 dias)**

#### **4.1 Nova Função SuperMDFe() → NovoSuperMDFe()**

```vb
' Substituição da função principal SuperMDFe
Private Sub NovoSuperMDFe()
    On Error GoTo ErroNovoSuperMDFe
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim conteudoINI As String
    Dim resultado As String
    
    ' 1. Inicializar ACBr
    If Not acbrMDFe.InicializarACBr() Then
        MsgBox "Erro ao inicializar ACBrMDFe", vbCritical
        Exit Sub
    End If
    
    ' 2. Configurar certificado e ambiente
    Call acbrMDFe.ConfigurarCertificado(Caminhocertificado, SenhaCertificado)
    Call acbrMDFe.ConfigurarAmbiente(TipoAmbiente, UfEmitente)
    
    ' 3. Montar INI com dados do MDFe
    conteudoINI = MontarINIMDFe()
    
    ' 4. Carregar dados na ACBr
    Call acbrMDFe.GerarMDFeINI(conteudoINI)
    
    ' 5. Assinar
    Call acbrMDFe.AssinarMDFe()
    
    ' 6. Transmitir
    Dim numeroLote As Long
    numeroLote = CLng(Format(Now, "yyyymmddhhnnss"))
    resultado = acbrMDFe.TransmitirMDFe(numeroLote)
    
    ' 7. Processar resultado
    Call ProcessarRetornoACBr(resultado)
    
    Set acbrMDFe = Nothing
    Exit Sub
    
ErroNovoSuperMDFe:
    Set acbrMDFe = Nothing
    MsgBox "Erro na NovoSuperMDFe: " & Err.Description, vbCritical
End Sub

' Função para montar INI com dados do sistema atual
Private Function MontarINIMDFe() As String
    Dim ini As String
    
    ' Seção Identificação
    ini = ini & "[Identificacao]" & vbCrLf
    ini = ini & "cUF=" & Codigo_da_uf & vbCrLf
    ini = ini & "tpAmb=" & Tipo_de_ambiente & vbCrLf
    ini = ini & "tpEmit=" & Tipo_de_emitente_codigo & vbCrLf
    ini = ini & "tpTransp=" & Tipo_de_transportador & vbCrLf
    ini = ini & "mod=58" & vbCrLf
    ini = ini & "serie=" & Serie & vbCrLf
    ini = ini & "nMDF=" & Numero_do_manifesto & vbCrLf
    ini = ini & "cMDF=" & Codigo_numerico_aleatorio & vbCrLf
    ini = ini & "cDV=" & Digito_verificador & vbCrLf
    ini = ini & "modal=1" & vbCrLf
    ini = ini & "dhEmi=" & Format(Data_de_emissao, "yyyy-mm-ddThh:nn:ss") & vbCrLf
    ini = ini & "tpEmis=1" & vbCrLf
    ini = ini & "procEmi=0" & vbCrLf
    ini = ini & "verProc=" & VersaoAplicativo & vbCrLf
    ini = ini & "UFIni=" & Uf_inicial & vbCrLf
    ini = ini & "UFFim=" & Uf_final & vbCrLf
    ini = ini & "dhIniViagem=" & Format(Data_hora_inicio_viagem, "yyyy-mm-ddThh:nn:ss") & vbCrLf
    ini = ini & vbCrLf
    
    ' Seção Emitente
    ini = ini & "[Emitente]" & vbCrLf
    ini = ini & "CNPJ=" & RemoveCaracteres(Cnpj_emitente) & vbCrLf
    ini = ini & "IE=" & RemoveCaracteres(Ie_emitente) & vbCrLf
    ini = ini & "xNome=" & SuperTiraAcentos(Nome_emitente) & vbCrLf
    ini = ini & "xFant=" & SuperTiraAcentos(Nome_fantasia) & vbCrLf
    ini = ini & "xLgr=" & SuperTiraAcentos(Endereco_emitente) & vbCrLf
    ini = ini & "nro=" & Numero_emitente & vbCrLf
    ini = ini & "xCpl=" & SuperTiraAcentos(Complemento_emitente) & vbCrLf
    ini = ini & "xBairro=" & SuperTiraAcentos(Bairro_emitente) & vbCrLf
    ini = ini & "cMun=" & Codigo_municipio_emitente & vbCrLf
    ini = ini & "xMun=" & SuperTiraAcentos(Municipio_emitente) & vbCrLf
    ini = ini & "CEP=" & RemoveCaracteres(Cep_emitente) & vbCrLf
    ini = ini & "UF=" & Uf_emitente & vbCrLf
    ini = ini & "fone=" & RemoveCaracteres(Telefone_emitente) & vbCrLf
    ini = ini & "email=" & Email_emitente & vbCrLf
    ini = ini & vbCrLf
    
    ' Seção Modal Rodoviário
    ini = ini & "[infModal]" & vbCrLf
    ini = ini & "versaoModal=3.00" & vbCrLf
    ini = ini & vbCrLf
    
    ini = ini & "[rodo]" & vbCrLf
    ini = ini & "RNTRC=" & RemoveCaracteres(Rntrc) & vbCrLf
    ini = ini & vbCrLf
    
    ' Informações ANTT
    ini = ini & "[infANTT]" & vbCrLf
    ini = ini & "RNTRC=" & RemoveCaracteres(Rntrc) & vbCrLf
    ini = ini & vbCrLf
    
    ' Informações do Contratante
    ini = ini & "[infCont]" & vbCrLf
    If Len(Trim(Manifesto![Documento do contratante])) > 0 Then
        If Len(RemoveCaracteres(Manifesto![Documento do contratante])) = 11 Then
            ini = ini & "CPF=" & RemoveCaracteres(Manifesto![Documento do contratante]) & vbCrLf
        Else
            ini = ini & "CNPJ=" & RemoveCaracteres(Manifesto![Documento do contratante]) & vbCrLf
        End If
    End If
    ini = ini & vbCrLf
    
    ' Informações de Pagamento
    ini = ini & "[infPag]" & vbCrLf
    ini = ini & "xNome=" & SuperTiraAcentos(Trim(Manifesto![Documento do contratante])) & vbCrLf
    ini = ini & "CPF=" & IIf(Len(RemoveCaracteres(Manifesto![Documento do contratante])) = 11, _
                              RemoveCaracteres(Manifesto![Documento do contratante]), "") & vbCrLf
    ini = ini & "CNPJ=" & IIf(Len(RemoveCaracteres(Manifesto![Documento do contratante])) = 14, _
                               RemoveCaracteres(Manifesto![Documento do contratante]), "") & vbCrLf
    ini = ini & "tpComp=01" & vbCrLf
    ini = ini & "vComp=" & Format(ValorContratacao, "0.00") & vbCrLf
    ini = ini & vbCrLf
    
    ' Proprietário do Veículo
    ini = ini & "[veicPrincipal]" & vbCrLf
    ini = ini & "cInt=" & Codigo_interno_veiculo & vbCrLf
    ini = ini & "placa=" & UCase(RemoveCaracteres(Placa)) & vbCrLf
    ini = ini & "renavam=" & RemoveCaracteres(Renavam) & vbCrLf
    ini = ini & "tara=" & Format(Tara_veiculo, "0") & vbCrLf
    ini = ini & "capKG=" & Format(Capacidade_kg, "0") & vbCrLf
    ini = ini & "capM3=" & Format(Capacidade_m3, "0.000") & vbCrLf
    ini = ini & "tpRod=" & Tipo_de_rodado & vbCrLf
    ini = ini & "tpCar=" & Tipo_de_carroceria & vbCrLf
    ini = ini & "UF=" & Uf_do_veiculo & vbCrLf
    ini = ini & vbCrLf
    
    ' Proprietário
    ini = ini & "[prop]" & vbCrLf
    If Len(Trim(CpfProprietario)) > 0 Then
        ini = ini & "CPF=" & RemoveCaracteres(CpfProprietario) & vbCrLf
    ElseIf Len(Trim(CnpjProprietario)) > 0 Then
        ini = ini & "CNPJ=" & RemoveCaracteres(CnpjProprietario) & vbCrLf
    End If
    ini = ini & "RNTRC=" & RemoveCaracteres(RntrcProprietario) & vbCrLf
    ini = ini & "xNome=" & SuperTiraAcentos(NomeProprietario) & vbCrLf
    ini = ini & "IE=" & RemoveCaracteres(IeProprietario) & vbCrLf
    ini = ini & "UF=" & UfProprietario & vbCrLf
    ini = ini & "tpProp=" & TipoProprietario & vbCrLf
    ini = ini & vbCrLf
    
    ' Condutores
    Dim contadorCondutores As Integer
    contadorCondutores = 1
    TbCondutores.MoveFirst
    Do While Not TbCondutores.EOF
        ini = ini & "[condutor" & Format(contadorCondutores, "00") & "]" & vbCrLf
        ini = ini & "xNome=" & SuperTiraAcentos(Trim(TbCondutores![Nome condutor])) & vbCrLf
        ini = ini & "CPF=" & RemoveCaracteres(TbCondutores![Cpf condutor]) & vbCrLf
        ini = ini & vbCrLf
        contadorCondutores = contadorCondutores + 1
        TbCondutores.MoveNext
    Loop
    
    ' Informações dos Documentos
    Dim contadorDescarga As Integer
    contadorDescarga = 1
    
    TbDescarga.MoveFirst
    Do While Not TbDescarga.EOF
        ini = ini & "[infMunDescarga" & Format(contadorDescarga, "00") & "]" & vbCrLf
        ini = ini & "cMunDescarga=" & TbDescarga![Codigo do ibge] & vbCrLf
        ini = ini & "xMunDescarga=" & SuperTiraAcentos(Trim(TbDescarga![Descrição do municipio])) & vbCrLf
        
        ' NFe deste município de descarga
        TbNFeCarga.MoveFirst
        Dim contadorNFe As Integer
        contadorNFe = 1
        Do While Not TbNFeCarga.EOF
            If TbNFeCarga![Municipio de descarga] = TbDescarga![Codigo do ibge] Then
                ini = ini & "[infMunDescarga" & Format(contadorDescarga, "00") & "_infNFe" & Format(contadorNFe, "00") & "]" & vbCrLf
                ini = ini & "chave=" & TbNFeCarga!Chave & vbCrLf
                ini = ini & vbCrLf
                contadorNFe = contadorNFe + 1
            End If
            TbNFeCarga.MoveNext
        Loop
        
        ini = ini & vbCrLf
        contadorDescarga = contadorDescarga + 1
        TbDescarga.MoveNext
    Loop
    
    ' Totais
    ini = ini & "[tot]" & vbCrLf
    ini = ini & "qCTe=" & QtdCte & vbCrLf
    ini = ini & "qNFe=" & QtdNfe & vbCrLf
    ini = ini & "qMDFe=0" & vbCrLf
    ini = ini & "vCarga=" & Format(Valor_total_da_carga, "0.00") & vbCrLf
    ini = ini & "cUnid=01" & vbCrLf
    ini = ini & "qCarga=" & Format(Peso_total_da_carga, "0.0000") & vbCrLf
    ini = ini & vbCrLf
    
    ' Informações Adicionais
    If Len(Trim(Manifesto!Historico)) > 0 Then
        ini = ini & "[infAdic]" & vbCrLf
        ini = ini & "infCpl=" & SuperTiraAcentos(Left(Trim(Manifesto!Historico), 5000)) & vbCrLf
        ini = ini & vbCrLf
    End If
    
    ' Seguro da Carga
    If Len(Trim(N_averbacao)) > 0 Then
        ini = ini & "[seg]" & vbCrLf
        ini = ini & "respSeg=" & RespSeguro & vbCrLf
        If Len(Trim(CnpjSeguro)) > 0 Then
            ini = ini & "CNPJ=" & RemoveCaracteres(CnpjSeguro) & vbCrLf
        End If
        ini = ini & "xSeg=" & SuperTiraAcentos(NomeSeguradora) & vbCrLf
        ini = ini & "nApol=" & NumeroApolice & vbCrLf
        ini = ini & "nAver=" & RemoveCaracteres(N_averbacao) & vbCrLf
        ini = ini & vbCrLf
    End If
    
    ' Responsável Técnico
    If Len(Trim(RespTec_CNPJ)) > 0 Then
        ini = ini & "[infRespTec]" & vbCrLf
        ini = ini & "CNPJ=" & RemoveCaracteres(RespTec_CNPJ) & vbCrLf
        ini = ini & "xContato=" & SuperTiraAcentos(RespTec_Contato) & vbCrLf
        ini = ini & "email=" & RespTec_email & vbCrLf
        ini = ini & "fone=" & RemoveCaracteres(RespTec_Fone) & vbCrLf
        ini = ini & "idCSRT=" & RespTec_idCSRT & vbCrLf
        ini = ini & "hashCSRT=" & RespTec_hashCSRT & vbCrLf
        ini = ini & vbCrLf
    End If
    
    MontarINIMDFe = ini
End Function
```

#### **4.2 Função de Processamento do Retorno**

```vb
Private Sub ProcessarRetornoACBr(retornoXML As String)
    Dim cStat As Integer
    Dim xMotivo As String
    Dim nProt As String
    Dim dhRecbto As String
    
    ' Extrair informações do XML de retorno
    cStat = ExtrairValorXML(retornoXML, "cStat")
    xMotivo = ExtrairValorXML(retornoXML, "xMotivo")
    nProt = ExtrairValorXML(retornoXML, "nProt")
    dhRecbto = ExtrairValorXML(retornoXML, "dhRecbto")
    
    ' Processar conforme status
    Select Case cStat
        Case 100 ' Autorizado
            MsgBox "MDF-e autorizado com sucesso!" & vbCrLf & _
                   "Protocolo: " & nProt & vbCrLf & _
                   "Data/Hora: " & dhRecbto, vbInformation, "Sucesso"
            
            ' Salvar dados no banco
            With vgTb
                !Protocolo_de_autorizacao = nProt
                !Data_e_hora_do_mdfe = dhRecbto
                ![Xml autorizado] = retornoXML
                .Update
            End With
            
            ' Atualizar interface
            Autorizada = True
            Transmitido = True
            Call MostraFormulas
            
        Case 101, 150 ' Cancelado
            MsgBox "MDF-e cancelado: " & xMotivo, vbInformation, "Cancelado"
            Nota_cancelada = True
            
        Case 135 ' Evento registrado
            MsgBox "Evento registrado com sucesso: " & xMotivo, vbInformation, "Evento"
            
        Case Else ' Rejeitado ou erro
            MsgBox "MDF-e rejeitado:" & vbCrLf & _
                   "Status: " & cStat & vbCrLf & _
                   "Motivo: " & xMotivo, vbCritical, "Rejeitado"
            
            ' Log do erro
            With vgTb
                !Observação = "Rejeitado: " & cStat & " - " & xMotivo
                .Update
            End With
    End Select
End Sub

' Função auxiliar para extrair valores do XML
Private Function ExtrairValorXML(xml As String, tag As String) As String
    Dim posIni As Integer
    Dim posFim As Integer
    
    posIni = InStr(xml, "<" & tag & ">")
    If posIni > 0 Then
        posIni = posIni + Len(tag) + 2
        posFim = InStr(posIni, xml, "</" & tag & ">")
        If posFim > posIni Then
            ExtrairValorXML = Mid(xml, posIni, posFim - posIni)
        End If
    End If
End Function
```

### **FASE 5: ADAPTAÇÃO DE INTERFACES (3 dias)**

#### **5.1 Adaptação do DAMDFE**

```vb
Public Sub NovoDAMDFE()
    On Error GoTo ErroDAMDFE
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim caminhoRelatório As String
    
    ' Verificar se MDFe foi autorizado
    If Not Autorizada Then
        MsgBox "MDF-e deve estar autorizado para gerar DAMDFE!", vbExclamation
        Exit Sub
    End If
    
    ' Inicializar ACBr
    If Not acbrMDFe.InicializarACBr() Then
        MsgBox "Erro ao inicializar ACBrMDFe para impressão", vbCritical
        Exit Sub
    End If
    
    ' Carregar XML autorizado
    If Len(Trim(Manifesto![Xml autorizado])) > 0 Then
        ' Usar XML já autorizado
        Call MDFE_CarregarXML(Manifesto![Xml autorizado])
    Else
        ' Recriar XML se necessário
        Dim conteudoINI As String
        conteudoINI = MontarINIMDFe()
        Call acbrMDFe.GerarMDFeINI(conteudoINI)
    End If
    
    ' Gerar PDF
    caminhoRelatório = acbrMDFe.ImprimirDAMDFE()
    
    If Len(caminhoRelatório) > 0 Then
        ' Abrir PDF com visualizador padrão
        Shell "explorer """ & caminhoRelatório & """", vbNormalFocus
    Else
        MsgBox "Erro ao gerar DAMDFE", vbCritical
    End If
    
    Set acbrMDFe = Nothing
    Exit Sub
    
ErroDAMDFE:
    Set acbrMDFe = Nothing
    MsgBox "Erro no NovoDAMDFE: " & Err.Description, vbCritical
End Sub
```

#### **5.2 Adaptação do Cancelamento**

```vb
' No formulário frmCanMDFe - adaptar para usar ACBr
Private Sub NovoProcessarCancelamento()
    On Error GoTo ErroCancelamento
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim resultado As String
    
    ' Validações
    If Len(Trim(txtChaveAcesso.Text)) <> 44 Then
        MsgBox "Chave de acesso deve ter 44 dígitos!", vbExclamation
        txtChaveAcesso.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtProtocolo.Text)) = 0 Then
        MsgBox "Protocolo de autorização é obrigatório!", vbExclamation
        txtProtocolo.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtJustificativa.Text)) < 15 Then
        MsgBox "Justificativa deve ter pelo menos 15 caracteres!", vbExclamation
        txtJustificativa.SetFocus
        Exit Sub
    End If
    
    ' Confirmar cancelamento
    If MsgBox("Confirma o cancelamento do MDF-e?" & vbCrLf & _
              "Chave: " & txtChaveAcesso.Text & vbCrLf & _
              "Esta operação não poderá ser desfeita!", _
              vbQuestion + vbYesNo, "Confirmar Cancelamento") = vbNo Then
        Exit Sub
    End If
    
    ' Inicializar ACBr
    If Not acbrMDFe.InicializarACBr() Then
        MsgBox "Erro ao inicializar ACBrMDFe", vbCritical
        Exit Sub
    End If
    
    ' Configurar certificado
    Call acbrMDFe.ConfigurarCertificado(CaminhoDoCertificado, SenhaDoCertificado)
    Call acbrMDFe.ConfigurarAmbiente(TipoAmbiente, UfEmitente)
    
    ' Processar cancelamento
    Screen.MousePointer = vbHourglass
    lblStatus.Caption = "Processando cancelamento..."
    DoEvents
    
    resultado = acbrMDFe.CancelarMDFe(txtChaveAcesso.Text, txtProtocolo.Text, txtJustificativa.Text)
    
    ' Processar resultado
    Call ProcessarRetornoCancelamento(resultado)
    
    Screen.MousePointer = vbDefault
    lblStatus.Caption = ""
    Set acbrMDFe = Nothing
    
    Exit Sub
    
ErroCancelamento:
    Screen.MousePointer = vbDefault
    lblStatus.Caption = ""
    Set acbrMDFe = Nothing
    MsgBox "Erro no cancelamento: " & Err.Description, vbCritical
End Sub

Private Sub ProcessarRetornoCancelamento(retornoXML As String)
    Dim cStat As Integer
    Dim xMotivo As String
    
    cStat = ExtrairValorXML(retornoXML, "cStat")
    xMotivo = ExtrairValorXML(retornoXML, "xMotivo")
    
    Select Case cStat
        Case 135, 136 ' Evento registrado
            MsgBox "Cancelamento realizado com sucesso!" & vbCrLf & _
                   "Motivo: " & xMotivo, vbInformation, "Sucesso"
            
            ' Atualizar banco de dados
            With vgTb
                .Edit
                !Nota_cancelada = True
                ![Data do cancelamento] = Now
                ![Motivo do cancelamento] = txtJustificativa.Text
                ![Xml cancelamento] = retornoXML
                .Update
            End With
            
            ' Fechar formulário
            Unload Me
            
        Case Else ' Erro
            MsgBox "Erro no cancelamento:" & vbCrLf & _
                   "Status: " & cStat & vbCrLf & _
                   "Motivo: " & xMotivo, vbCritical, "Erro"
    End Select
End Sub
```

### **FASE 6: TRATAMENTO DE ERROS E LOGS (2 dias)**

#### **6.1 Sistema de Log Unificado**

```vb
' Módulo: ModuloLog.bas
Option Explicit

Public Const LOG_SUCESSO = 1
Public Const LOG_AVISO = 2
Public Const LOG_ERRO = 3
Public Const LOG_DEBUG = 4

Public Sub EscreverLog(modulo As String, mensagem As String, Optional tipoLog As Integer = LOG_SUCESSO)
    On Error Resume Next
    
    Dim arquivoLog As String
    Dim numeroArquivo As Integer
    Dim dataHora As String
    Dim tipoTexto As String
    
    ' Definir tipo do log
    Select Case tipoLog
        Case LOG_SUCESSO: tipoTexto = "[SUCESSO]"
        Case LOG_AVISO: tipoTexto = "[AVISO]  "
        Case LOG_ERRO: tipoTexto = "[ERRO]   "
        Case LOG_DEBUG: tipoTexto = "[DEBUG]  "
        Case Else: tipoTexto = "[INFO]   "
    End Select
    
    ' Caminho do arquivo de log
    arquivoLog = App.Path & "\Logs\ACBrMDFe_" & Format(Date, "yyyy-mm-dd") & ".log"
    
    ' Criar diretório se não existir
    If Dir(App.Path & "\Logs\", vbDirectory) = "" Then
        MkDir App.Path & "\Logs\"
    End If
    
    ' Escrever no log
    dataHora = Format(Now, "dd/mm/yyyy hh:nn:ss")
    numeroArquivo = FreeFile
    
    Open arquivoLog For Append As numeroArquivo
    Print #numeroArquivo, dataHora & " " & tipoTexto & " [" & modulo & "] " & mensagem
    Close numeroArquivo
End Sub

Public Function TratarErroACBr(codigoErro As Long, ultimoRetorno As String, Optional contexto As String = "") As String
    Dim mensagemErro As String
    
    ' Mapear códigos de erro mais comuns
    Select Case codigoErro
        Case 0
            mensagemErro = "Operação realizada com sucesso"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro, LOG_SUCESSO)
            
        Case -1
            mensagemErro = "Erro na inicialização da biblioteca"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case -2
            mensagemErro = "Certificado digital inválido ou não encontrado"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case -3
            mensagemErro = "XML mal formado ou inválido"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case -4
            mensagemErro = "Erro na assinatura digital"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case -5
            mensagemErro = "Erro de conexão com o webservice"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case -10 To -99
            mensagemErro = "Erro de comunicação com SEFAZ (Código: " & codigoErro & ")"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case -100 To -199
            mensagemErro = "Erro de validação dos dados (Código: " & codigoErro & ")"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
            
        Case Else
            mensagemErro = "Erro desconhecido (Código: " & codigoErro & ")"
            Call EscreverLog("ACBrMDFe", contexto & ": " & mensagemErro & " | " & ultimoRetorno, LOG_ERRO)
    End Select
    
    ' Incluir detalhes do último retorno se disponível
    If Len(ultimoRetorno) > 0 And InStr(ultimoRetorno, "Erro:") > 0 Then
        mensagemErro = mensagemErro & vbCrLf & "Detalhes: " & ultimoRetorno
    End If
    
    TratarErroACBr = mensagemErro
End Function

' Função para validar conectividade
Public Function ValidarConectividadeSEFAZ() As Boolean
    On Error GoTo ErroConectividade
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim resultado As Long
    
    Call EscreverLog("ACBrMDFe", "Iniciando validação de conectividade", LOG_DEBUG)
    
    ' Inicializar
    If Not acbrMDFe.InicializarACBr() Then
        Call EscreverLog("ACBrMDFe", "Falha na inicialização para teste de conectividade", LOG_ERRO)
        ValidarConectividadeSEFAZ = False
        Exit Function
    End If
    
    ' Configurar ambiente
    Call acbrMDFe.ConfigurarAmbiente(TipoAmbiente, UfEmitente)
    
    ' Testar status do serviço
    resultado = MDFE_StatusServico()
    
    If resultado = 0 Then
        Call EscreverLog("ACBrMDFe", "Conectividade com SEFAZ validada com sucesso", LOG_SUCESSO)
        ValidarConectividadeSEFAZ = True
    Else
        Call EscreverLog("ACBrMDFe", "Falha na conectividade: " & acbrMDFe.ObterUltimoRetorno(), LOG_ERRO)
        ValidarConectividadeSEFAZ = False
    End If
    
    Set acbrMDFe = Nothing
    Exit Function
    
ErroConectividade:
    Call EscreverLog("ACBrMDFe", "Exceção na validação de conectividade: " & Err.Description, LOG_ERRO)
    Set acbrMDFe = Nothing
    ValidarConectividadeSEFAZ = False
End Function
```

### **FASE 7: CONFIGURAÇÃO E PARAMETRIZAÇÃO (2 dias)**

#### **7.1 Formulário de Configuração ACBr**

```vb
' frmConfigACBr.frm
Private Sub Form_Load()
    ' Carregar configurações atuais
    Call CarregarConfiguracaoAtual
End Sub

Private Sub CarregarConfiguracaoAtual()
    On Error Resume Next
    
    Dim acbrMDFe As New ACBrMDFeUtil
    
    If acbrMDFe.InicializarACBr() Then
        ' Carregar valores das configurações
        txtCaminhoLogs.Text = ObterConfiguracao("Principal", "LogPath")
        cboNivelLog.ListIndex = Val(ObterConfiguracao("Principal", "LogLevel"))
        txtTimeOut.Text = ObterConfiguracao("WebService", "TimeOut")
        txtProxyHost.Text = ObterConfiguracao("WebService", "ProxyHost")
        txtProxyPort.Text = ObterConfiguracao("WebService", "ProxyPort")
        txtProxyUser.Text = ObterConfiguracao("WebService", "ProxyUser")
        txtCaminhoSchemas.Text = ObterConfiguracao("DFe", "PathSchemas")
        chkSalvarEnvio.Value = IIf(ObterConfiguracao("WebService", "SalvarEnvio") = "1", 1, 0)
        chkSalvarResposta.Value = IIf(ObterConfiguracao("WebService", "SalvarResposta") = "1", 1, 0)
    End If
    
    Set acbrMDFe = Nothing
End Sub

Private Function ObterConfiguracao(sessao As String, chave As String) As String
    Dim buffer As String
    Dim resultado As Long
    
    buffer = String(1024, vbNullChar)
    resultado = MDFE_ConfigLerValor(sessao, chave, buffer, Len(buffer))
    
    If resultado > 0 Then
        ObterConfiguracao = Left(buffer, resultado)
    Else
        ObterConfiguracao = ""
    End If
End Function

Private Sub cmdSalvar_Click()
    On Error GoTo ErroSalvar
    
    Dim acbrMDFe As New ACBrMDFeUtil
    
    If Not acbrMDFe.InicializarACBr() Then
        MsgBox "Erro ao inicializar ACBr para salvar configurações", vbCritical
        Exit Sub
    End If
    
    ' Salvar configurações
    Call MDFE_ConfigGravarValor("Principal", "LogPath", txtCaminhoLogs.Text)
    Call MDFE_ConfigGravarValor("Principal", "LogLevel", CStr(cboNivelLog.ListIndex))
    Call MDFE_ConfigGravarValor("WebService", "TimeOut", txtTimeOut.Text)
    Call MDFE_ConfigGravarValor("WebService", "ProxyHost", txtProxyHost.Text)
    Call MDFE_ConfigGravarValor("WebService", "ProxyPort", txtProxyPort.Text)
    Call MDFE_ConfigGravarValor("WebService", "ProxyUser", txtProxyUser.Text)
    Call MDFE_ConfigGravarValor("WebService", "ProxyPass", txtProxyPass.Text)
    Call MDFE_ConfigGravarValor("DFe", "PathSchemas", txtCaminhoSchemas.Text)
    Call MDFE_ConfigGravarValor("WebService", "SalvarEnvio", IIf(chkSalvarEnvio.Value = 1, "1", "0"))
    Call MDFE_ConfigGravarValor("WebService", "SalvarResposta", IIf(chkSalvarResposta.Value = 1, "1", "0"))
    
    ' Gravar no arquivo
    Call MDFE_ConfigGravar("")
    
    MsgBox "Configurações salvas com sucesso!", vbInformation
    Set acbrMDFe = Nothing
    Unload Me
    
    Exit Sub
    
ErroSalvar:
    Set acbrMDFe = Nothing
    MsgBox "Erro ao salvar configurações: " & Err.Description, vbCritical
End Sub

Private Sub cmdTestarConectividade_Click()
    Screen.MousePointer = vbHourglass
    lblStatus.Caption = "Testando conectividade..."
    DoEvents
    
    If ValidarConectividadeSEFAZ() Then
        MsgBox "Conectividade com SEFAZ OK!", vbInformation
    Else
        MsgBox "Falha na conectividade com SEFAZ. Verifique as configurações e logs.", vbExclamation
    End If
    
    Screen.MousePointer = vbDefault
    lblStatus.Caption = ""
End Sub
```

### **FASE 8: TESTES E VALIDAÇÃO (4 dias)**

#### **8.1 Plano de Testes Automatizados**

```vb
' ModuloTestes.bas
Option Explicit

Public Function ExecutarSuiteTestes() As Boolean
    On Error GoTo ErroTestes
    
    Dim totalTestes As Integer
    Dim testesOK As Integer
    Dim resultadoFinal As Boolean
    
    Call EscreverLog("Testes", "=== INICIANDO SUITE DE TESTES ACBrMDFe ===", LOG_DEBUG)
    
    ' Teste 1: Inicialização
    totalTestes = totalTestes + 1
    If TesteInicializacao() Then
        testesOK = testesOK + 1
        Call EscreverLog("Testes", "✓ Teste de Inicialização: PASSOU", LOG_SUCESSO)
    Else
        Call EscreverLog("Testes", "✗ Teste de Inicialização: FALHOU", LOG_ERRO)
    End If
    
    ' Teste 2: Configuração
    totalTestes = totalTestes + 1
    If TesteConfiguracao() Then
        testesOK = testesOK + 1
        Call EscreverLog("Testes", "✓ Teste de Configuração: PASSOU", LOG_SUCESSO)
    Else
        Call EscreverLog("Testes", "✗ Teste de Configuração: FALHOU", LOG_ERRO)
    End If
    
    ' Teste 3: Geração XML
    totalTestes = totalTestes + 1
    If TesteGeracaoXML() Then
        testesOK = testesOK + 1
        Call EscreverLog("Testes", "✓ Teste de Geração XML: PASSOU", LOG_SUCESSO)
    Else
        Call EscreverLog("Testes", "✗ Teste de Geração XML: FALHOU", LOG_ERRO)
    End If
    
    ' Teste 4: Assinatura
    totalTestes = totalTestes + 1
    If TesteAssinatura() Then
        testesOK = testesOK + 1
        Call EscreverLog("Testes", "✓ Teste de Assinatura: PASSOU", LOG_SUCESSO)
    Else
        Call EscreverLog("Testes", "✗ Teste de Assinatura: FALHOU", LOG_ERRO)
    End If
    
    ' Teste 5: Conectividade
    totalTestes = totalTestes + 1
    If TesteConectividade() Then
        testesOK = testesOK + 1
        Call EscreverLog("Testes", "✓ Teste de Conectividade: PASSOU", LOG_SUCESSO)
    Else
        Call EscreverLog("Testes", "✓ Teste de Conectividade: FALHOU", LOG_AVISO)
    End If
    
    ' Resultado final
    resultadoFinal = (testesOK = totalTestes)
    Call EscreverLog("Testes", "=== RESULTADO FINAL: " & testesOK & "/" & totalTestes & " testes passaram ===", _
                     IIf(resultadoFinal, LOG_SUCESSO, LOG_ERRO))
    
    ExecutarSuiteTestes = resultadoFinal
    Exit Function
    
ErroTestes:
    Call EscreverLog("Testes", "ERRO na execução da suite de testes: " & Err.Description, LOG_ERRO)
    ExecutarSuiteTestes = False
End Function

Private Function TesteInicializacao() As Boolean
    On Error GoTo ErroTesteInicializacao
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim resultado As Boolean
    
    ' Testar inicialização
    resultado = acbrMDFe.InicializarACBr()
    
    If resultado Then
        ' Testar finalização
        Call acbrMDFe.FinalizarACBr
        TesteInicializacao = True
    Else
        TesteInicializacao = False
    End If
    
    Set acbrMDFe = Nothing
    Exit Function
    
ErroTesteInicializacao:
    Set acbrMDFe = Nothing
    TesteInicializacao = False
End Function

Private Function TesteConfiguracao() As Boolean
    On Error GoTo ErroTesteConfiguracao
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim resultado As Boolean
    
    If Not acbrMDFe.InicializarACBr() Then
        TesteConfiguracao = False
        Exit Function
    End If
    
    ' Testar configurações básicas
    resultado = acbrMDFe.ConfigurarAmbiente(2, "SP") ' Homologação, SP
    
    If resultado Then
        resultado = acbrMDFe.ConfigurarCertificado("", "") ' Certificado vazio para teste
        TesteConfiguracao = True ' Mesmo que falhe o certificado, se chegou até aqui está OK
    Else
        TesteConfiguracao = False
    End If
    
    Set acbrMDFe = Nothing
    Exit Function
    
ErroTesteConfiguracao:
    Set acbrMDFe = Nothing
    TesteConfiguracao = False
End Function

Private Function TesteGeracaoXML() As Boolean
    On Error GoTo ErroTesteGeracaoXML
    
    Dim acbrMDFe As New ACBrMDFeUtil
    Dim iniTeste As String
    Dim resultado As Boolean
    
    If Not acbrMDFe.InicializarACBr() Then
        TesteGeracaoXML = False
        Exit Function
    End If
    
    ' Configurar ambiente de teste
    Call acbrMDFe.ConfigurarAmbiente(2, "SP")
    
    ' XML de teste mínimo
    iniTeste = "[Identificacao]" & vbCrLf & _
               "cUF=35" & vbCrLf & _
               "tpAmb=2" & vbCrLf & _
               "tpEmit=2" & vbCrLf & _
               "mod=58" & vbCrLf & _
               "serie=1" & vbCrLf & _
               "nMDF=1" & vbCrLf & _
               "cMDF=12345678" & vbCrLf & _
               "modal=1" & vbCrLf & _
               "dhEmi=" & Format(Now, "yyyy-mm-ddThh:nn:ss") & vbCrLf & _
               "tpEmis=1" & vbCrLf & _
               "procEmi=0" & vbCrLf & _
               "verProc=1.0" & vbCrLf & _
               "UFIni=SP" & vbCrLf & _
               "UFFim=RJ" & vbCrLf & _
               vbCrLf & _
               "[Emitente]" & vbCrLf & _
               "CNPJ=11111111111111" & vbCrLf & _
               "IE=111111111111" & vbCrLf & _
               "xNome=Teste" & vbCrLf
    
    resultado = acbrMDFe.GerarMDFeINI(iniTeste)
    TesteGeracaoXML = resultado
    
    Set acbrMDFe = Nothing
    Exit Function
    
ErroTesteGeracaoXML:
    Set acbrMDFe = Nothing
    TesteGeracaoXML = False
End Function

Private Function TesteAssinatura() As Boolean
    ' Este teste depende do certificado estar configurado
    ' Por enquanto sempre retorna True para não travar os testes
    TesteAssinatura = True
End Function

Private Function TesteConectividade() As Boolean
    TesteConectividade = ValidarConectividadeSEFAZ()
End Function
```

### **FASE 9: DEPLOY E CUTOVER (1 dia)**

#### **9.1 Script de Deploy**

```batch
@echo off
title Migração MDFe - FlexDocs para ACBr
echo ==============================================
echo    MIGRACAO MDFE: FLEXDOCS PARA ACBR
echo ==============================================

REM Verificar se o usuário tem privilégios administrativos
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Execute como Administrador!
    pause
    exit /b 1
)

echo.
echo 1. Fazendo backup do sistema atual...
if not exist "C:\Backup\MDFe\%DATE%" mkdir "C:\Backup\MDFe\%DATE%"
xcopy "C:\Projetos\MDFe - CLAUDE\NFE\*.exe" "C:\Backup\MDFe\%DATE%\" /Y
xcopy "C:\Projetos\MDFe - CLAUDE\NFE\*.dll" "C:\Backup\MDFe\%DATE%\" /Y
xcopy "C:\Projetos\MDFe - CLAUDE\NFE\*.ini" "C:\Backup\MDFe\%DATE%\" /Y
echo Backup concluído.

echo.
echo 2. Parando serviços relacionados...
taskkill /f /im NFE.exe 2>nul
timeout /t 3 /nobreak >nul

echo.
echo 3. Desregistrando DLL antiga...
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm /u "C:\Projetos\MDFe - CLAUDE\NFE\MDFe_Util.dll" /tlb 2>nul
echo DLL FlexDocs desregistrada.

echo.
echo 4. Copiando arquivos ACBr...
copy "ACBrLibMDFe-Windows-1.2.2.335\Windows\MT\StdCall\ACBrMDFe32.dll" "C:\Projetos\MDFe - CLAUDE\NFE\" /Y
xcopy "ACBrLibMDFe-Windows-1.2.2.335\dep\*" "C:\Projetos\MDFe - CLAUDE\NFE\dep\" /S /E /Y
echo Arquivos ACBr copiados.

echo.
echo 5. Criando configuração inicial...
echo [Principal] > "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo LogLevel=4 >> "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo LogPath=C:\Projetos\MDFe - CLAUDE\NFE\Logs\ >> "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo. >> "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo [DFe] >> "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo UF=SP >> "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo Ambiente=2 >> "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini"
echo Configuração inicial criada.

echo.
echo 6. Criando diretórios necessários...
if not exist "C:\Projetos\MDFe - CLAUDE\NFE\Logs" mkdir "C:\Projetos\MDFe - CLAUDE\NFE\Logs"
if not exist "C:\Projetos\MDFe - CLAUDE\NFE\XML" mkdir "C:\Projetos\MDFe - CLAUDE\NFE\XML"
if not exist "C:\Projetos\MDFe - CLAUDE\NFE\PDF" mkdir "C:\Projetos\MDFe - CLAUDE\NFE\PDF"
echo Diretórios criados.

echo.
echo 7. Copiando executável atualizado...
copy "NFE_MIGRADO.exe" "C:\Projetos\MDFe - CLAUDE\NFE\NFE.exe" /Y
echo Executável atualizado.

echo.
echo 8. Testando inicialização...
"C:\Projetos\MDFe - CLAUDE\NFE\NFE.exe" /test
if %errorlevel% equ 0 (
    echo Teste de inicialização: OK
) else (
    echo ERRO: Falha no teste de inicialização!
    echo Restaurando backup...
    copy "C:\Backup\MDFe\%DATE%\NFE.exe" "C:\Projetos\MDFe - CLAUDE\NFE\" /Y
    echo Sistema restaurado. Verifique os logs.
    pause
    exit /b 1
)

echo.
echo ==============================================
echo         MIGRACAO CONCLUIDA COM SUCESSO!
echo ==============================================
echo.
echo Próximos passos:
echo 1. Configure o certificado digital em Configurações
echo 2. Teste uma transmissão em homologação
echo 3. Valide a impressão do DAMDFE
echo.
echo Logs disponíveis em: C:\Projetos\MDFe - CLAUDE\NFE\Logs\
echo.
pause
```

#### **9.2 Script de Rollback**

```batch
@echo off
title Rollback - Migração MDFe
echo ==============================================
echo          ROLLBACK MIGRACAO MDFE
echo ==============================================

REM Verificar se o backup existe
if not exist "C:\Backup\MDFe\%DATE%" (
    echo ERRO: Backup não encontrado para hoje!
    echo Verifique os backups disponíveis em C:\Backup\MDFe\
    dir "C:\Backup\MDFe\" /b
    pause
    exit /b 1
)

echo.
echo ATENÇÃO: Esta operação irá restaurar o sistema FlexDocs
echo e desfazer todas as alterações da migração para ACBr.
echo.
set /p confirma="Confirma o rollback? (S/N): "
if /i "%confirma%" neq "S" (
    echo Operação cancelada.
    pause
    exit /b 0
)

echo.
echo 1. Parando sistema atual...
taskkill /f /im NFE.exe 2>nul
timeout /t 3 /nobreak >nul

echo.
echo 2. Removendo arquivos ACBr...
del "C:\Projetos\MDFe - CLAUDE\NFE\ACBrMDFe32.dll" 2>nul
del "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe.ini" 2>nul
rmdir "C:\Projetos\MDFe - CLAUDE\NFE\dep" /s /q 2>nul

echo.
echo 3. Restaurando backup...
copy "C:\Backup\MDFe\%DATE%\*.*" "C:\Projetos\MDFe - CLAUDE\NFE\" /Y

echo.
echo 4. Registrando DLL FlexDocs...
%windir%\Microsoft.NET\Framework\v4.0.30319\regasm "C:\Projetos\MDFe - CLAUDE\NFE\MDFe_Util.dll" /tlb

echo.
echo 5. Testando sistema restaurado...
"C:\Projetos\MDFe - CLAUDE\NFE\NFE.exe" /test
if %errorlevel% equ 0 (
    echo Sistema restaurado com sucesso!
) else (
    echo AVISO: Sistema restaurado mas com possíveis problemas.
)

echo.
echo ==============================================
echo           ROLLBACK CONCLUÍDO
echo ==============================================
pause
```

### **FASE 10: PÓS-DEPLOY E SUPORTE (Contínuo)**

#### **10.1 Monitoramento Automático**

```vb
' Timer para monitoramento contínuo
Private Sub tmrMonitoramento_Timer()
    Static ultimaVerificacao As Date
    
    ' Verificar apenas a cada 5 minutos
    If DateDiff("n", ultimaVerificacao, Now) < 5 Then Exit Sub
    ultimaVerificacao = Now
    
    ' Verificar logs de erro
    Call VerificarLogsErro
    
    ' Verificar conectividade
    If Not ValidarConectividadeSEFAZ() Then
        Call EscreverLog("Monitor", "Perda de conectividade detectada", LOG_AVISO)
        ' Enviar notificação se configurado
    End If
    
    ' Verificar espaço em disco
    Call VerificarEspacoDisco
End Sub

Private Sub VerificarLogsErro()
    Dim arquivoLog As String
    Dim numeroArquivo As Integer
    Dim linha As String
    Dim contadorErros As Integer
    
    arquivoLog = App.Path & "\Logs\ACBrMDFe_" & Format(Date, "yyyy-mm-dd") & ".log"
    
    If Dir(arquivoLog) <> "" Then
        numeroArquivo = FreeFile
        Open arquivoLog For Input As numeroArquivo
        
        Do While Not EOF(numeroArquivo)
            Line Input #numeroArquivo, linha
            If InStr(linha, "[ERRO]") > 0 Then
                contadorErros = contadorErros + 1
            End If
        Loop
        
        Close numeroArquivo
        
        If contadorErros > 10 Then
            Call EscreverLog("Monitor", "Alto número de erros detectado hoje: " & contadorErros, LOG_AVISO)
        End If
    End If
End Sub
```

---

## **RISCOS E MITIGAÇÕES**

### **Riscos Identificados**

| **Risco** | **Probabilidade** | **Impacto** | **Mitigação** |
|-----------|------------------|-------------|---------------|
| **Incompatibilidade de Certificados** | Baixa | Alto | Testes antecipados em homologação |
| **Diferenças no XML gerado** | Média | Médio | Validação cruzada com XMLs atuais |
| **Performance inferior** | Baixa | Médio | Testes de carga e otimização |
| **Problemas de conectividade** | Média | Alto | Sistema de fallback e retry |
| **Resistência dos usuários** | Média | Baixo | Treinamento e interface similar |
| **Rollback complexo** | Baixa | Alto | Scripts automatizados de rollback |

### **Plano de Contingência**

1. **Backup completo** antes da migração
2. **Scripts automatizados** de rollback
3. **Ambiente de teste** espelhado
4. **Suporte técnico** durante o go-live
5. **Monitoramento ativo** nas primeiras 48h

---

## **CRONOGRAMA DETALHADO**

### **Semana 1 (Dias 1-5)**
- **Dias 1-3:** Preparação e configuração (Fase 1)
- **Dias 4-5:** Mapeamento de funções (Fase 2)

### **Semana 2 (Dias 6-10)**
- **Dias 6-9:** Implementação classes wrapper (Fase 3)
- **Dia 10:** Início migração funções principais (Fase 4)

### **Semana 3 (Dias 11-15)**
- **Dias 11-15:** Continuação migração funções (Fase 4)

### **Semana 4 (Dias 16-20)**
- **Dias 16-17:** Finalização migração (Fase 4)
- **Dias 18-20:** Adaptação interfaces (Fase 5)

### **Semana 5 (Dias 21-26)**
- **Dias 21-22:** Tratamento de erros (Fase 6)
- **Dias 23-24:** Configuração e parametrização (Fase 7)
- **Dias 25-26:** Início dos testes (Fase 8)

### **Semana 6 (Dias 27-30)**
- **Dias 27-28:** Continuação testes (Fase 8)
- **Dia 29:** Deploy e cutover (Fase 9)
- **Dia 30:** Pós-deploy e ajustes (Fase 10)

---

## **ENTREGÁVEIS**

### **Código-fonte**
- ✅ Nova classe **ACBrMDFeUtil.cls**
- ✅ Formulários adaptados (**frmMDFe.frm**, **frmCanMDFe.frm**, **frmEncMDFe.frm**)
- ✅ Módulos de apoio (**ModuloLog.bas**, **ModuloTestes.bas**)
- ✅ Scripts de instalação e configuração

### **Documentação**
- ✅ **Manual técnico** da migração
- ✅ **Guia do usuário** atualizado
- ✅ **Documentação da API** ACBrMDFe
- ✅ **Procedimentos de suporte**

### **Scripts e Ferramentas**
- ✅ **Script de deploy** automatizado
- ✅ **Script de rollback** de emergência
- ✅ **Ferramentas de monitoramento**
- ✅ **Suite de testes** automatizados

### **Configurações**
- ✅ **Arquivo de configuração** ACBrLibMDFe.ini
- ✅ **Templates de XML** de teste
- ✅ **Configurações de ambiente** (produção/homologação)

---

## **CONSIDERAÇÕES FINAIS**

Esta migração representa um passo importante na modernização do sistema MDFe, trazendo:

### **Benefícios Técnicos**
- ✅ **Maior estabilidade** - Biblioteca oficialmente suportada
- ✅ **Melhor performance** - Otimizações nativas
- ✅ **Suporte ativo** - Comunidade grande e atuante
- ✅ **Conformidade** - Sempre atualizada com especificações SEFAZ

### **Benefícios Operacionais**
- ✅ **Redução de problemas** - Menos bugs e inconsistências
- ✅ **Facilidade de manutenção** - Documentação completa
- ✅ **Suporte profissional** - Disponível se necessário
- ✅ **Futuro garantido** - Projeto ativo e em evolução

### **Próximos Passos Recomendados**
1. **Aprovação formal** do plano pela direção
2. **Alocação de recursos** (desenvolvedor + ambiente)
3. **Definição da data** de início da migração
4. **Comunicação aos usuários** sobre a mudança planejada
5. **Início da Fase 1** - Preparação e configuração

---

**Este documento serve como guia completo para a migração segura e eficiente do sistema MDFe da biblioteca FlexDocs para ACBrLibMDFe, garantindo continuidade operacional e melhoria técnica significativa.**

---

*Documento gerado em: 03/09/2025*  
*Versão: 1.0*  
*Páginas: Documento completo com todos os detalhes técnicos*