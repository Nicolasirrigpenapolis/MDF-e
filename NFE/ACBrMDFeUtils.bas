Attribute VB_Name = "ACBrMDFeUtils"
Option Explicit

Public Function CreateMDFe(Optional ByVal eArqConfig As String = "", _
                          Optional ByVal eChaveCrypt As String = "") As Object
    Dim mdfe As Object
    
    ' Tentar criar instancia usando diferentes abordagens
    On Error Resume Next
    Set mdfe = CreateObject("ACBrMDFe.ACBrMDFe")
    If Err.Number <> 0 Then
        Err.Clear
        Set mdfe = New ACBrMDFe
    End If
    On Error GoTo 0
    
    ' NAO INICIALIZAR AUTOMATICAMENTE - Deixar para ser feito manualmente
    ' para evitar erro "Bad DLL calling convention" no Form_Load
    ' Se quiser inicializar: mdfe.InicializarLib eArqConfig, eChaveCrypt
    
    Set CreateMDFe = mdfe
End Function