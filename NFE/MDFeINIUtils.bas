Attribute VB_Name = "MDFeINIUtils"
Option Explicit

' Módulo para geração de arquivos INI para ACBrMDFe
' Substitui as chamadas FlexDocs por geração estruturada

Public Sub CriarMDFeINI(caminhoINI As String, _
                       ide_cUF As String, ide_tpAmb As String, ide_tpEmit As String, _
                       ide_tpTransp As String, ide_mod As String, ide_serie As String, _
                       ide_nMDF As String, ide_cMDF As String, ide_cDV As String, _
                       ide_Modal As String, ide_dhEmi As String, ide_tpEmis As String, _
                       ide_procEmi As String, ide_verProc As String, ide_UFIni As String, _
                       ide_UFFim As String, ide_dhIniViagem As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    ' Criar arquivo INI
    Open caminhoINI For Output As #arq
    
    ' Seção infMDFe
    Print #arq, "[infMDFe]"
    Print #arq, "Id="
    Print #arq, "versao=3.00"
    Print #arq, ""
    
    ' Seção ide
    Print #arq, "[ide]"
    Print #arq, "cUF=" & ide_cUF
    Print #arq, "tpAmb=" & ide_tpAmb
    Print #arq, "tpEmit=" & ide_tpEmit
    Print #arq, "tpTransp=" & ide_tpTransp
    Print #arq, "mod=" & ide_mod
    Print #arq, "serie=" & ide_serie
    Print #arq, "nMDF=" & ide_nMDF
    Print #arq, "cMDF=" & ide_cMDF
    Print #arq, "cDV=" & ide_cDV
    Print #arq, "modal=" & ide_Modal
    Print #arq, "dhEmi=" & ide_dhEmi
    Print #arq, "tpEmis=" & ide_tpEmis
    Print #arq, "procEmi=" & ide_procEmi
    Print #arq, "verProc=" & ide_verProc
    Print #arq, "UFIni=" & ide_UFIni
    Print #arq, "UFFim=" & ide_UFFim
    If ide_dhIniViagem <> "" Then
        Print #arq, "dhIniViagem=" & ide_dhIniViagem
    End If
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarEmitente(caminhoINI As String, _
                            emit_CNPJ As String, emit_IE As String, emit_xNome As String, _
                            emit_xFant As String, emit_xLgr As String, emit_nro As String, _
                            emit_xCpl As String, emit_xBairro As String, emit_cMun As String, _
                            emit_xMun As String, emit_CEP As String, emit_UF As String, _
                            emit_fone As String, emit_email As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[emit]"
    Print #arq, "CNPJ=" & emit_CNPJ
    Print #arq, "IE=" & emit_IE
    Print #arq, "xNome=" & emit_xNome
    If emit_xFant <> "" Then Print #arq, "xFant=" & emit_xFant
    Print #arq, "xLgr=" & emit_xLgr
    Print #arq, "nro=" & emit_nro
    If emit_xCpl <> "" Then Print #arq, "xCpl=" & emit_xCpl
    Print #arq, "xBairro=" & emit_xBairro
    Print #arq, "cMun=" & emit_cMun
    Print #arq, "xMun=" & emit_xMun
    Print #arq, "CEP=" & emit_CEP
    Print #arq, "UF=" & emit_UF
    If emit_fone <> "" Then Print #arq, "fone=" & emit_fone
    If emit_email <> "" Then Print #arq, "email=" & emit_email
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarModalRodoviario(caminhoINI As String, _
                                   versaoModal As String, rntrc As String, _
                                   ciot As String, contratante_cpf As String, _
                                   contratante_cnpj As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[rodo]"
    Print #arq, "versaoModal=" & versaoModal
    Print #arq, ""
    
    Print #arq, "[infANTT]"
    Print #arq, "RNTRC=" & rntrc
    If ciot <> "" Then Print #arq, "CIOT=" & ciot
    Print #arq, ""
    
    If contratante_cpf <> "" Or contratante_cnpj <> "" Then
        Print #arq, "[infCont01]"
        If contratante_cpf <> "" Then Print #arq, "CPF=" & contratante_cpf
        If contratante_cnpj <> "" Then Print #arq, "CNPJ=" & contratante_cnpj
        Print #arq, ""
    End If
    
    Close #arq
End Sub

Public Sub AdicionarVeiculoPrincipal(caminhoINI As String, _
                                    veic_cInt As String, veic_placa As String, _
                                    veic_renavam As String, veic_tara As String, _
                                    veic_capKG As String, veic_capM3 As String, _
                                    veic_tpRod As String, veic_tpCar As String, _
                                    veic_UF As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[veicPrincipal]"
    If veic_cInt <> "" Then Print #arq, "cInt=" & veic_cInt
    Print #arq, "placa=" & veic_placa
    If veic_renavam <> "" Then Print #arq, "RENAVAM=" & veic_renavam
    Print #arq, "tara=" & veic_tara
    Print #arq, "capKG=" & veic_capKG
    If veic_capM3 <> "" Then Print #arq, "capM3=" & veic_capM3
    Print #arq, "tpRod=" & veic_tpRod
    Print #arq, "tpCar=" & veic_tpCar
    Print #arq, "UF=" & veic_UF
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarCondutor(caminhoINI As String, indiceCondutor As Integer, _
                            condutor_xNome As String, condutor_CPF As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[condutor" & Format(indiceCondutor, "00") & "]"
    Print #arq, "xNome=" & condutor_xNome
    Print #arq, "CPF=" & condutor_CPF
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarProprietario(caminhoINI As String, _
                                prop_CPF As String, prop_CNPJ As String, _
                                prop_RNTRC As String, prop_xNome As String, _
                                prop_IE As String, prop_UF As String, _
                                prop_tpProp As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[prop]"
    If prop_CPF <> "" Then Print #arq, "CPF=" & prop_CPF
    If prop_CNPJ <> "" Then Print #arq, "CNPJ=" & prop_CNPJ
    Print #arq, "RNTRC=" & prop_RNTRC
    Print #arq, "xNome=" & prop_xNome
    If prop_IE <> "" Then Print #arq, "IE=" & prop_IE
    Print #arq, "UF=" & prop_UF
    Print #arq, "tpProp=" & prop_tpProp
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarMunicipioCarregamento(caminhoINI As String, indice As Integer, _
                                         cMunCarrega As String, xMunCarrega As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[infMunCarrega" & Format(indice, "00") & "]"
    Print #arq, "cMunCarrega=" & cMunCarrega
    Print #arq, "xMunCarrega=" & xMunCarrega
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarPercurso(caminhoINI As String, indice As Integer, ufPer As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[infPercurso" & Format(indice, "00") & "]"
    Print #arq, "UFPer=" & ufPer
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarMunicipioDescarregamento(caminhoINI As String, indice As Integer, _
                                           cMunDescarga As String, xMunDescarga As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[infMunDescarga" & Format(indice, "00") & "]"
    Print #arq, "cMunDescarga=" & cMunDescarga
    Print #arq, "xMunDescarga=" & xMunDescarga
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarNFe(caminhoINI As String, indiceMun As Integer, indiceNFe As Integer, _
                       chNFe As String, segCodBarra As String, indReentrega As String, _
                       infUnidTransp_tpUnidTransp As String, infUnidTransp_idUnidTransp As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[infMunDescarga" & Format(indiceMun, "00") & "_infNFe" & Format(indiceNFe, "00") & "]"
    Print #arq, "chNFe=" & chNFe
    If segCodBarra <> "" Then Print #arq, "segCodBarra=" & segCodBarra
    If indReentrega <> "" Then Print #arq, "indReentrega=" & indReentrega
    If infUnidTransp_tpUnidTransp <> "" Then Print #arq, "tpUnidTransp=" & infUnidTransp_tpUnidTransp
    If infUnidTransp_idUnidTransp <> "" Then Print #arq, "idUnidTransp=" & infUnidTransp_idUnidTransp
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarCTe(caminhoINI As String, indiceMun As Integer, indiceCTe As Integer, _
                       chCTe As String, segCodBarra As String, indReentrega As String, _
                       infUnidTransp_tpUnidTransp As String, infUnidTransp_idUnidTransp As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[infMunDescarga" & Format(indiceMun, "00") & "_infCTe" & Format(indiceCTe, "00") & "]"
    Print #arq, "chCTe=" & chCTe
    If segCodBarra <> "" Then Print #arq, "segCodBarra=" & segCodBarra
    If indReentrega <> "" Then Print #arq, "indReentrega=" & indReentrega
    If infUnidTransp_tpUnidTransp <> "" Then Print #arq, "tpUnidTransp=" & infUnidTransp_tpUnidTransp
    If infUnidTransp_idUnidTransp <> "" Then Print #arq, "idUnidTransp=" & infUnidTransp_idUnidTransp
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarTotalizadores(caminhoINI As String, _
                                 qNFe As String, qCTe As String, qMDFe As String, _
                                 vCarga As String, cUnid As String, qCarga As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[tot]"
    If qNFe <> "" Then Print #arq, "qNFe=" & qNFe
    If qCTe <> "" Then Print #arq, "qCTe=" & qCTe
    If qMDFe <> "" Then Print #arq, "qMDFe=" & qMDFe
    Print #arq, "vCarga=" & vCarga
    Print #arq, "cUnid=" & cUnid
    Print #arq, "qCarga=" & qCarga
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarInformacoesAdicionais(caminhoINI As String, _
                                         infAdFisco As String, infCpl As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[infAdic]"
    If infAdFisco <> "" Then Print #arq, "infAdFisco=" & infAdFisco
    If infCpl <> "" Then Print #arq, "infCpl=" & infCpl
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarSeguro(caminhoINI As String, indice As Integer, _
                          respSeg As String, xSeg As String, CNPJ As String, _
                          nApolice As String, nAver As String, _
                          tpCTe As String, docCTe As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[seg" & Format(indice, "00") & "]"
    Print #arq, "respSeg=" & respSeg
    If xSeg <> "" Then Print #arq, "xSeg=" & xSeg
    If CNPJ <> "" Then Print #arq, "CNPJ=" & CNPJ
    If nApolice <> "" Then Print #arq, "nApolice=" & nApolice
    If nAver <> "" Then Print #arq, "nAver=" & nAver
    If tpCTe <> "" Then Print #arq, "tpCTe=" & tpCTe
    If docCTe <> "" Then Print #arq, "docCTe=" & docCTe
    Print #arq, ""
    
    Close #arq
End Sub

Public Sub AdicionarLacre(caminhoINI As String, indice As Integer, nLacre As String)
    
    Dim arq As Integer
    arq = FreeFile
    
    Open caminhoINI For Append As #arq
    
    Print #arq, "[lacre" & Format(indice, "00") & "]"
    Print #arq, "nLacre=" & nLacre
    Print #arq, ""
    
    Close #arq
End Sub

' Função auxiliar para remover caracteres especiais
Public Function RemoveCaracteres(texto As String, Optional manterNumeros As Boolean = False) As String
    Dim resultado As String
    Dim i As Integer
    Dim char As String
    
    For i = 1 To Len(texto)
        char = Mid(texto, i, 1)
        If IsNumeric(char) Or (Not manterNumeros And IsAlpha(char)) Then
            resultado = resultado & char
        End If
    Next i
    
    RemoveCaracteres = resultado
End Function

' Função auxiliar para verificar se é letra
Private Function IsAlpha(char As String) As Boolean
    Dim asciiValue As Integer
    asciiValue = Asc(UCase(char))
    IsAlpha = (asciiValue >= 65 And asciiValue <= 90)
End Function

' Função para gerar chave de acesso usando ACBr
Public Function GerarChaveMDFe(m_ACBrMDFe As ACBrMDFe, _
                               cUF As Long, cNumerico As Long, modelo As Long, _
                               serie As Long, numero As Long, tpEmis As Long, _
                               dataEmissao As Date, cnpjCpf As String) As String
    
    GerarChaveMDFe = m_ACBrMDFe.GerarChave(cUF, cNumerico, modelo, serie, numero, tpEmis, dataEmissao, cnpjCpf)
End Function
