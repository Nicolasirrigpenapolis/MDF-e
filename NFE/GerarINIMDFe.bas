Attribute VB_Name = "GerarINIMDFe"
Option Explicit

' Função principal que gera arquivo INI completo do MDFe substituindo FlexDocs
Public Function GerarINIMDFeCompleto(caminhoINI As String) As Boolean
    On Error GoTo ErroGerarINI
    
    Dim arq As Integer
    arq = FreeFile
    
    ' Criar/sobrescrever arquivo INI
    Open caminhoINI For Output As #arq
    
    ' ========== SEÇÃO IDE ==========
    Print #arq, "[ide]"
    Print #arq, "cUF=" & ide_cUF
    Print #arq, "tpAmb=" & IIf(Parametros_da_NFe!ambiente = 0, "1", "2")
    Print #arq, "tpEmit=" & ide_tpEmit
    Print #arq, "tpTransp=" & IIf(Ide_tpTransp = "", "1", Ide_tpTransp)
    Print #arq, "mod=" & ide_mod
    Print #arq, "serie=" & ide_serie
    Print #arq, "nMDF=" & Format(ide_nMDF, "00000")
    Print #arq, "cMDF=" & ide_cMDF
    Print #arq, "cDV=" & ide_cDV
    Print #arq, "modal=" & ide_Modal
    Print #arq, "dhEmi=" & ide_dhEmi
    Print #arq, "tpEmis=" & ide_tpEmis
    Print #arq, "procEmi=" & ide_procEmi
    Print #arq, "verProc=" & ide_verProc
    Print #arq, "UFIni=" & ide_UFIni
    Print #arq, "UFFim=" & ide_UFFim
    If Ide_dhIniViagem <> "" Then Print #arq, "dhIniViagem=" & Ide_dhIniViagem
    Print #arq, ""
    
    ' ========== MUNICÍPIOS DE CARREGAMENTO ==========
    If Not TbCarrega Is Nothing And TbCarrega.RecordCount > 0 Then
        Dim indiceMunCarrega As Integer
        indiceMunCarrega = 1
        TbCarrega.MoveFirst
        Do While Not TbCarrega.EOF
            Print #arq, "[CARR" & Format(indiceMunCarrega, "000") & "]"
            Print #arq, "cMunCarrega=" & TbCarrega![Codigo do IBGE]
            Print #arq, "xMunCarrega=" & RemoveCaracteres(SuperTiraAcentos(Trim(TbCarrega![Descrição Do Municipio])), True)
            Print #arq, ""
            
            indiceMunCarrega = indiceMunCarrega + 1
            TbCarrega.MoveNext
        Loop
        TbCarrega.MoveFirst
    End If
    
    ' ========== PERCURSO ==========
    If Not TbPercurso Is Nothing And TbPercurso.RecordCount > 0 Then
        Dim indicePercurso As Integer
        indicePercurso = 1
        TbPercurso.MoveFirst
        Do While Not TbPercurso.EOF
            Print #arq, "[perc" & Format(indicePercurso, "000") & "]"
            Print #arq, "UFPer=" & TbPercurso!Uf
            Print #arq, ""
            
            indicePercurso = indicePercurso + 1
            TbPercurso.MoveNext
        Loop
        TbPercurso.MoveFirst
    End If
    
    ' ========== EMITENTE ==========
    If Not TbEmitente Is Nothing Then
        Print #arq, "[emit]"
        If Parametros_da_nfe!ambiente = 0 Then
            Print #arq, "CNPJCPF=" & RemoveCaracteres(TbEmitente!Cnpj)
            Print #arq, "IE=" & RemoveCaracteres(TbEmitente!Ie)
            Print #arq, "xNome=" & RemoveCaracteres(TbEmitente![Razão Social])
        Else
            Print #arq, "CNPJCPF=99999999999999"
            Print #arq, "IE=999999999999" 
            Print #arq, "xNome=TESTE HOMOLOGACAO"
        End If
        Print #arq, "xFant=" & IIf(IsNull(TbEmitente![Nome Fantasia]), "", RemoveCaracteres(TbEmitente![Nome Fantasia]))
        Print #arq, "xLgr=" & RemoveCaracteres(TbEmitente!Logradouro)
        Print #arq, "nro=" & TbEmitente!Nro
        Print #arq, "xCpl=" & IIf(IsNull(TbEmitente!Complemento), "", RemoveCaracteres(TbEmitente!Complemento))
        Print #arq, "xBairro=" & RemoveCaracteres(TbEmitente!Bairro)
        Print #arq, "cMun=" & TbEmitente![Codigo do IBGE]
        Print #arq, "xMun=" & RemoveCaracteres(TbEmitente!Municipio)
        Print #arq, "CEP=" & RemoveCaracteres(TbEmitente!Cep)
        Print #arq, "UF=" & TbEmitente!Uf
        Print #arq, "fone=" & RemoveCaracteres(Substitui(TbEmitente!Fone, " ", "", SO_UM))
        Print #arq, "email=" & IIf(IsNull(TbEmitente!Email), "", TbEmitente!Email)
        Print #arq, ""
    End If
    
    ' ========== MODAL RODOVIÁRIO ==========
    Print #arq, "[Rodo]"
    Print #arq, "codAgPorto="
    Print #arq, ""
    
    Print #arq, "[infANTT]"
    Print #arq, "RNTRC=" & Manifesto!Rntrc
    Print #arq, ""
    
    ' Contratante
    If Manifesto![Tipo de contratante] = 0 Then ' Física
        Print #arq, "[infCont001]"
        Print #arq, "CPF=" & RemoveCaracteres(Manifesto![Documento Do contratante])
        Print #arq, ""
    Else ' Jurídica
        Print #arq, "[infCont001]"
        Print #arq, "CNPJ=" & RemoveCaracteres(Manifesto![Documento Do contratante])
        Print #arq, ""
    End If
    
    ' ========== VEÍCULO PRINCIPAL ==========
    Print #arq, "[veicPrincipal]"
    Print #arq, "cInt=001"
    Print #arq, "placa=" & RemoveCaracteres(Manifesto!Placa)
    Print #arq, "RENAVAM="
    Print #arq, "tara=" & RemoveCaracteres(Manifesto!Tara)
    Print #arq, "capKG=" & RemoveCaracteres(Manifesto![Capacidade kg])
    Print #arq, "capM3="
    
    ' Tipo de rodado
    If Manifesto![Tipo de rodado] = "Truck" Then
        Print #arq, "tpRod=01"
    ElseIf Manifesto![Tipo de rodado] = "Toco" Then
        Print #arq, "tpRod=02"
    ElseIf Manifesto![Tipo de rodado] = "Cavalo Mecânico" Then
        Print #arq, "tpRod=03"
    ElseIf Manifesto![Tipo de rodado] = "VAN" Then
        Print #arq, "tpRod=04"
    ElseIf Manifesto![Tipo de rodado] = "Utilitário" Then
        Print #arq, "tpRod=05"
    ElseIf Manifesto![Tipo de rodado] = "Outros" Then
        Print #arq, "tpRod=06"
    Else
        Print #arq, "tpRod=06"
    End If
    
    ' Tipo de carroceria
    Print #arq, "tpCar=" & Format(Manifesto![Tipo de carroceria], "00")
    Print #arq, "UF=" & Manifesto![Uf Do veiculo]
    Print #arq, ""
    
    ' ========== CONDUTORES ==========
    If Not TbCondutores Is Nothing And TbCondutores.RecordCount > 0 Then
        Dim indiceCondutor As Integer
        indiceCondutor = 1
        TbCondutores.MoveFirst
        Do While Not TbCondutores.EOF
            Print #arq, "[condutor" & Format(indiceCondutor, "00") & "]"
            Print #arq, "xNome=" & RemoveCaracteres(TbCondutores![Nome Condutor])
            Print #arq, "CPF=" & RemoveCaracteres(TbCondutores![Cpf Condutor])
            Print #arq, ""
            
            indiceCondutor = indiceCondutor + 1
            TbCondutores.MoveNext
        Loop
        TbCondutores.MoveFirst
    End If
    
    ' ========== PROPRIETÁRIO (se não for do emitente) ==========
    If Not Vazio(Manifesto!Proprietario) And Trim(Manifesto!Proprietario) <> Trim(TbEmitente![Razão Social]) Then
        Print #arq, "[prop]"
        If Manifesto![Tipo de proprietario] = 1 Then ' CPF
            Print #arq, "CPF=" & RemoveCaracteres(Manifesto![Cpf Proprietario])
        Else ' CNPJ
            Print #arq, "CNPJ=" & RemoveCaracteres(Manifesto![Cnpj Proprietario])
        End If
        Print #arq, "RNTRC=" & RemoveCaracteres(Manifesto![Rntrc proprietario])
        Print #arq, "xNome=" & RemoveCaracteres(Manifesto![Nome Proprietario])
        If Not Vazio(Manifesto![Ie Proprietario]) Then
            Print #arq, "IE=" & RemoveCaracteres(Manifesto![Ie Proprietario])
        End If
        Print #arq, "UF=" & Manifesto![Uf proprietario]
        Print #arq, "tpProp=" & IIf(Manifesto![Tipo de proprietario] = 1, "1", "2")
        Print #arq, ""
    End If
    
    ' ========== DOCUMENTOS FISCAIS ==========
    GerarDocumentosFiscais arq
    
    ' ========== TOTALIZADORES ==========
    GerarTotalizadores arq
    
    ' ========== SEGURO ==========
    GerarSeguro arq
    
    ' ========== INFORMAÇÕES ADICIONAIS ==========
    If Not Vazio(Manifesto!Observação) Then
        Print #arq, "[infAdic]"
        Print #arq, "infCpl=" & Replace(Manifesto!Observação, vbCrLf, " ")
        Print #arq, ""
    End If
    
    Close #arq
    
    GerarINIMDFeCompleto = True
    Exit Function
    
ErroGerarINI:
    If arq <> 0 Then Close #arq
    GerarINIMDFeCompleto = False
    MsgBox "Erro ao gerar INI: " & Err.Description, vbCritical
End Function

Private Sub GerarDocumentosFiscais(arq As Integer)
    ' Gerar seções de documentos por município de descarga
    If Not TbDescarga Is Nothing And TbDescarga.RecordCount > 0 Then
        Dim indiceMunDescarga As Integer
        indiceMunDescarga = 1
        
        TbDescarga.MoveFirst
        Do While Not TbDescarga.EOF
            ' Município de descarga
            Print #arq, "[infMunDescarga" & Format(indiceMunDescarga, "00") & "]"
            Print #arq, "cMunDescarga=" & TbDescarga![Codigo do IBGE]
            Print #arq, "xMunDescarga=" & RemoveCaracteres(SuperTiraAcentos(Trim(TbDescarga![Descrição Do municipio])), True)
            Print #arq, ""
            
            ' NFes para este município
            Dim TbNFeCarga As Object
            Set TbNFeCarga = vgDb.OpenRecordSet("SELECT [Chave da NFe] As Chave From [NFe da Carga] Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " and [Descrição do Municipio] = '" & TbDescarga![Descrição Do municipio] & "'")
            
            If Not TbNFeCarga Is Nothing And TbNFeCarga.RecordCount > 0 Then
                Dim indiceNFe As Integer
                indiceNFe = 1
                TbNFeCarga.MoveFirst
                Do While Not TbNFeCarga.EOF
                    Print #arq, "[infMunDescarga" & Format(indiceMunDescarga, "00") & "_infNFe" & Format(indiceNFe, "00") & "]"
                    Print #arq, "chNFe=" & TbNFeCarga!Chave
                    Print #arq, ""
                    
                    indiceNFe = indiceNFe + 1
                    TbNFeCarga.MoveNext
                Loop
            End If
            
            ' CTes para este município  
            Dim TbCTeCarga As Object
            Set TbCTeCarga = vgDb.OpenRecordSet("SELECT [Chave do CTe] As Chave From [CTe da Carga] Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " and [Descrição do Municipio] = '" & TbDescarga![Descrição Do municipio] & "'")
            
            If Not TbCTeCarga Is Nothing And TbCTeCarga.RecordCount > 0 Then
                Dim indiceCTe As Integer
                indiceCTe = 1
                TbCTeCarga.MoveFirst
                Do While Not TbCTeCarga.EOF
                    Print #arq, "[infMunDescarga" & Format(indiceMunDescarga, "00") & "_infCTe" & Format(indiceCTe, "00") & "]"
                    Print #arq, "chCTe=" & TbCTeCarga!Chave
                    Print #arq, ""
                    
                    indiceCTe = indiceCTe + 1
                    TbCTeCarga.MoveNext
                Loop
            End If
            
            indiceMunDescarga = indiceMunDescarga + 1
            TbDescarga.MoveNext
        Loop
        TbDescarga.MoveFirst
    End If
End Sub

Private Sub GerarTotalizadores(arq As Integer)
    Print #arq, "[tot]"
    
    ' Contar NFes e CTes
    Dim qNFe As Long, qCTe As Long
    Dim vCarga As Double, qCarga As Double
    
    ' Contar documentos
    If Not TbNFeCarga Is Nothing Then
        qNFe = TbNFeCarga.RecordCount
    End If
    If Not TbCTeCarga Is Nothing Then
        qCTe = TbCTeCarga.RecordCount  
    End If
    
    ' Valores da carga (podem ser calculados dos documentos ou informados)
    vCarga = 0 ' Valor total da carga
    qCarga = 1000 ' Quantidade total da carga em kg
    
    If qNFe > 0 Then Print #arq, "qNFe=" & qNFe
    If qCTe > 0 Then Print #arq, "qCTe=" & qCTe
    Print #arq, "vCarga=" & Format(vCarga, "0.00")
    Print #arq, "cUnid=01" ' Kg
    Print #arq, "qCarga=" & Format(qCarga, "0.0000")
    Print #arq, ""
End Sub

Private Sub GerarSeguro(arq As Integer)
    ' Informações de seguro se disponíveis
    If Not Vazio(Manifesto![Responsavel do seguro]) Then
        Print #arq, "[seg01]"
        Print #arq, "respSeg=" & Manifesto![Responsavel do seguro]
        If Not Vazio(Manifesto![Nome da seguradora]) Then
            Print #arq, "xSeg=" & RemoveCaracteres(Manifesto![Nome da seguradora])
        End If
        If Not Vazio(Manifesto![Cnpj da seguradora]) Then
            Print #arq, "CNPJ=" & RemoveCaracteres(Manifesto![Cnpj da seguradora])
        End If
        If Not Vazio(Manifesto![N da apolice]) Then
            Print #arq, "nApolice=" & Manifesto![N da apolice]
        End If
        If Not Vazio(Manifesto![N averbação]) Then
            Print #arq, "nAver=" & Manifesto![N averbação]
        End If
        Print #arq, ""
    End If
End Sub
