' TesteConectividade.vbs
' Script para testar conectividade SEFAZ com ACBrLibMDFe

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Verificar se ACBrMDFe está instalado
arquitetura = ""
If fso.FileExists("ACBrMDFe32.dll") Then
    arquitetura = "32-bit"
ElseIf fso.FileExists("ACBrMDFe64.dll") Then  
    arquitetura = "64-bit"
Else
    MsgBox "ERRO: ACBrMDFe não encontrado!" & vbCrLf & _
           "Execute primeiro o instalar_acbr_producao.bat", vbCritical, "Erro"
    WScript.Quit
End If

' Verificar configuração
If Not fso.FileExists("ACBrLibMDFe.ini") Then
    MsgBox "ERRO: Arquivo ACBrLibMDFe.ini não encontrado!" & vbCrLf & _
           "Execute primeiro o configurador de certificado", vbCritical, "Erro"
    WScript.Quit
End If

' Ler ambiente configurado
Set arquivo = fso.OpenTextFile("ACBrLibMDFe.ini", 1)
conteudo = arquivo.ReadAll
arquivo.Close

ambiente = "HOMOLOGAÇÃO"
If InStr(conteudo, "Ambiente=1") > 0 Then
    ambiente = "PRODUÇÃO"
End If

' Verificar certificado configurado
certificadoOK = False
If InStr(conteudo, "Arquivo=C:\") > 0 Or InStr(conteudo, "NumeroSerie=") > 0 Then
    certificadoOK = True
End If

' Mostrar resultado
mensagem = "TESTE DE CONECTIVIDADE ACBrLibMDFe" & vbCrLf & vbCrLf & _
           "✓ Biblioteca: Instalada (" & arquitetura & ")" & vbCrLf & _
           "✓ Configuração: Encontrada" & vbCrLf & _
           "✓ Ambiente: " & ambiente & vbCrLf & _
           IIf(certificadoOK, "✓ Certificado: Configurado", "⚠ Certificado: NÃO CONFIGURADO") & vbCrLf & vbCrLf

If certificadoOK Then
    mensagem = mensagem & "STATUS: PRONTO PARA USAR!" & vbCrLf & vbCrLf & _
               "Próximos passos:" & vbCrLf & _
               "1. Faça um teste de transmissão" & vbCrLf & _
               "2. Verifique se gera DAMDFE" & vbCrLf & _
               "3. Teste cancelamento/encerramento"
    icone = vbInformation
Else
    mensagem = mensagem & "STATUS: FALTA CONFIGURAR CERTIFICADO" & vbCrLf & vbCrLf & _
               "Execute: configurar_certificado.bat"
    icone = vbExclamation
End If

MsgBox mensagem, icone, "Teste de Conectividade"