# 📋 MIGRAÇÃO MDFe FlexDocs → ACBrLib - Documentação Completa

## 📑 ÍNDICE
1. [Status Atual](#status-atual)
2. [Etapa 1 - Teste SuperMDFe](#etapa-1)
3. [Próximas Etapas](#proximas-etapas)
4. [Estrutura do Banco](#estrutura-banco)
5. [Arquivos ACBr](#arquivos-acbr)
6. [Código - Antes/Depois](#codigo)
7. [Testes](#testes)
8. [Chamadas FlexDocs](#chamadas-flexdocs)
9. [Regras Críticas](#regras-criticas)
10. [Informações Técnicas](#info-tecnicas)

---

## 🚧 STATUS ATUAL {#status-atual}

### 📊 STATUS ATUAL (Janeiro 2025):
- ✅ **ETAPA 1 CONCLUÍDA COM SUCESSO!** 
- ✅ **ACBr inicializado, INI gerado, XML criado (1157 chars), dados salvos no banco**
- ✅ **Problema "Bad DLL calling convention" RESOLVIDO**
- 📋 **PRÓXIMO PASSO**: Etapa 2 - Transmissão do MDFe

---

## 🎯 ETAPA 1 - TESTE SuperMDFe() {#etapa-1}
**Status**: ✅ **CONCLUÍDA COM SUCESSO!**

**O que foi implementado:**
- ✅ Função `SuperMDFe_Teste()` criada em mdfe 2.txt (linha ~4580)
- ✅ Chamada redirecionada: linha 4745 de `SuperMDFe` → `SuperMDFe_Teste`
- ✅ DLL correta identificada e instalada (ST/StdCall)
- ✅ Debug extensivo adicionado para troubleshooting

**🔧 SOLUÇÃO PARA "Bad DLL calling convention":**

### ❌ PROBLEMA IDENTIFICADO:
Estávamos usando a **DLL ERRADA**:
- ❌ **ACBrMDFe32.dll MT/StdCall** (Multi-threaded) → Causava erro "Bad DLL calling convention"
- ❌ **ACBrMDFe32.dll MT/Cdecl** (Multi-threaded, Cdecl) → Incompatível com VB6

### ✅ SOLUÇÃO APLICADA:
**Usar a DLL CORRETA**: `ACBrMDFe32.dll ST/StdCall` (Single-threaded, StdCall)

**Localização da DLL correta:**
```
C:\Projetos\MDFe\NFE\ACBrLibMDFe-Windows-1.2.2.335\Windows\ST\StdCall\ACBrMDFe32.dll
```

**Comando usado para correção:**
```bash
cp "C:\Projetos\MDFe\NFE\ACBrLibMDFe-Windows-1.2.2.335\Windows\ST\StdCall\ACBrMDFe32.dll" "C:\Projetos\MDFe\NFE\ACBrMDFe32.dll"
```

### 📋 TIPOS DE DLL DISPONÍVEIS:
```
ACBrLibMDFe-Windows-1.2.2.335\Windows\
├── MT\                    (Multi-threaded - NÃO funcionou)
│   ├── Cdecl\            ❌ Incompatível com VB6
│   └── StdCall\          ❌ Causou "Bad DLL calling convention"
└── ST\                    (Single-threaded - FUNCIONOU!)
    ├── Cdecl\            ⚠️ Não testado
    └── StdCall\          ✅ SOLUÇÃO QUE FUNCIONOU!
```

**Teste implementado e funcionando:**
1. ✅ Inicializa ACBr com DLL ST/StdCall
2. ✅ Gera arquivo INI básico  
3. ✅ Carrega INI no ACBr
4. ✅ Gera XML (1157 caracteres)
5. ✅ Salva dados no banco + arquivo em \Temp\

**Resultado obtido:**
- ✅ Mensagem: "TESTE ETAPA 1 CONCLUÍDO!"
- ✅ XML gerado com sucesso (1157 chars)
- ✅ Dados salvos no banco corretamente
- ✅ Arquivo XML salvo em \Temp\ para verificação

---

## 📋 PRÓXIMAS ETAPAS {#proximas-etapas}
1. **ETAPA 2**: Migrar `TransmitirMDFe()` após sucesso da Etapa 1
2. **ETAPA 3**: Migrar `RetornoMDFe()` após sucesso da Etapa 2
3. **ETAPA 4**: Substituir funções originais pelas migradas
4. **ETAPA 5**: Testes integrados e homologação

---

## 🗄️ ESTRUTURA DO BANCO DE DADOS {#estrutura-banco}

### 1. Manifesto (Tabela Principal)
```sql
SELECT [Sequencia do manifesto], [Numero do manifesto], [Data de emissão], 
       Uf, [Tipo de emitente], [Uf de descarregamento], Observação, Rntrc, 
       [Tipo de carroceria], [Uf do veiculo], [Tipo de rodado], Placa, Tara, 
       Renavam, [Capacidade kg], [Codigo do emitente], Transmitido, 
       [Nota cancelada], Autorizada, Historico, Proprietario, [Cpf Proprietario], 
       [Cnpj Proprietario], [Rntrc proprietario], [Nome Proprietario], 
       [Ie Proprietario], [Uf proprietario], [Tipo de proprietario], 
       [Tipo documento], [Numero do recibo], [Chave de acesso], XmlAssinado, 
       [Protocolo de autorização], [Data e hora do mdfe], XmlAutorizado, 
       Encerrado, [Responsavel do seguro], [Tipo do responsavel], 
       [Documento do responsavel], [Nome da seguradora], [Cnpj da seguradora], 
       [N da apolice], [N averbação], [Tipo de contratante], 
       [Documento do contratante], [Produto Predominante], 
       [Latitude de Carregamento], [Longitude de Carregamento], 
       [Latitude de Descrregamento], [Longitude de Descarregamento], 
       [CEP Carregamento], [CEP Descarregamento]
FROM Manifesto
```

### 2. Emitentes MDFe
```sql
SELECT [Sequencia do emitente], Cnpj, Ie, [Razão social], [Nome fantasia], 
       Logradouro, Nro, Complemento, Bairro, [Codigo do ibge], Municipio, 
       Cep, Uf, Fone, Email, Inativo, [Certificado digital], [Chave flexdocs]
FROM [Emitentes mdfe]
```

---

## 📁 ARQUIVOS ACBr ORGANIZADOS {#arquivos-acbr}

### ✅ ARQUIVOS NECESSÁRIOS (no NFE.vbp):
- **ACBrMDFe.cls** - Classe principal do ACBr
- **ACBrMDFeUtils.bas** - Utilitários para criar instâncias  
- **ACBrComum.bas** - Funções de conversão UTF-8
- **MDFeINIUtils.bas** - Utilitários para criar INI
- **GerarINIMDFe.bas** - Funções auxiliares para INI

### ❌ ARQUIVOS REMOVIDOS (duplicatas):
- ~~ACBrComum_UTF8.bas~~ (duplicata)
- ~~MDFeINIUtils_UTF8.bas~~ (duplicata)  
- ~~GerarINIMDFe_UTF8.bas~~ (duplicata)

---

## 🔧 CÓDIGO - ANTES/DEPOIS {#codigo}

### ❌ ANTES (FlexDocs - ATUAL):
```vb
Set objMDFEUtil = CreateObject("MDFe_Util.Util")
resultado = objMDFEUtil.infMunCarrega(codigo, municipio)
Consolidacao = objMDFEUtil.MDFe_NT2020001(...)
```

### ✅ DEPOIS (ACBr - IMPLEMENTADO):
```vb
Set m_ACBrMDFe = CreateMDFe("", "")
m_ACBrMDFe.InicializarLib App.Path & "\ACBrLibMDFe.ini", ""
Call CriarMDFeINI(caminhoINI, ...)
m_ACBrMDFe.CarregarINI caminhoINI
m_ACBrMDFe.Assinar
xmlGerado = m_ACBrMDFe.ObterXml(0)
```

---

## 🧪 TESTES A REALIZAR {#testes}

### ETAPA 1 - SuperMDFe():
**Como testar:**
1. Execute o sistema MDFe
2. Crie/edite um manifesto
3. Clique "Gerar MDFe"
4. Verifique mensagens nas observações
5. Confirme arquivos em \Temp\

**Se der erro:**
- Verificar se ACBrMDFE32.dll existe
- Conferir permissões da pasta \Temp\
- Validar configuração ACBrLibMDFe.ini

### FUNCIONALIDADES PENDENTES:
1. ⚠️ **SuperMDFe()** - EM TESTE (geração XML)
2. 📅 **TransmitirMDFe()** - AGUARDANDO ETAPA 2
3. 📅 **RetornoMDFe()** - AGUARDANDO ETAPA 3
4. 📅 **CancelaMDFe()** - AGUARDANDO ETAPA 4
5. 📅 **EncerraMDFe()** - AGUARDANDO ETAPA 5

---

## 📋 CHAMADAS FLEXDOCS PARA MIGRAÇÃO {#chamadas-flexdocs}

### SuperMDFe() (ETAPA 1):
- `objMDFeUtil.infMunCarrega()` → Arquivo INI
- `objMDFeUtil.CriaChaveDFe()` → m_ACBrMDFe.GerarChave()
- `objMDFeUtil.MDFe_NT2020001()` → CarregarINI + Assinar + ObterXml

### TransmitirMDFe() (ETAPA 2):
- `objMDFeUtil.EnviaMDFe()` → m_ACBrMDFe.Enviar()

### RetornoMDFe() (ETAPA 3):
- `objMDFeUtil.BuscaMDFe()` → m_ACBrMDFe.ConsultarRecibo()

---

## 🚨 REGRAS CRÍTICAS DA MIGRAÇÃO {#regras-criticas}

### ❌ NUNCA MODIFICAR ESTRUTURA DO BANCO:
- **JAMAIS** alterar nomes de campos
- **JAMAIS** alterar nomes de tabelas  
- **JAMAIS** adicionar/remover campos
- **SEMPRE** usar os campos EXATOS conforme tabelas.txt
- **Migração DEVE** ser apenas no código, não no banco

### 📋 ESTRUTURA COMPLETA DO BANCO (OFICIAL)

**1. Condutores:**
```sql
SELECT [Sequencia do manifesto], [Nome condutor], [Cpf condutor]
FROM Codutores
```

**2. CTe da Carga:**
```sql
SELECT [Sequencia descarrega cte], [Sequencia do manifesto], [Chave do cte], 
       [Descrição do municipio], [Valor total], [Peso bruto]
FROM [Cte da carga]
```

**3. NFe da Carga:**
```sql
SELECT [Sequencia descarrega], [Sequencia do manifesto], [Chave da nfe], 
       [Descrição do municipio], [Valor total], [Peso bruto]
FROM [Nfe da carga]
```

**4. Municípios:**
```sql
SELECT [Sequencia do municipio], [Descrição do municipio], Uf, [Codigo do ibge], 
       Cep, Inativo, Latitude, Longitude
FROM Municipios
```

**5. Manifesto (TABELA PRINCIPAL):**
```sql
SELECT [Sequencia do manifesto], [Numero do manifesto], [Data de emissão], Uf, 
       [Tipo de emitente], [Uf de descarregamento], Observação, Rntrc, 
       [Tipo de carroceria], [Uf do veiculo], [Tipo de rodado], Placa, Tara, 
       Renavam, [Capacidade kg], [Codigo do emitente], Transmitido, 
       [Nota cancelada], Autorizada, Historico, Proprietario, [Cpf Proprietario], 
       [Cnpj Proprietario], [Rntrc proprietario], [Nome Proprietario], 
       [Ie Proprietario], [Uf proprietario], [Tipo de proprietario], 
       [Tipo documento], [Numero do recibo], [Chave de acesso], XmlAssinado, 
       [Protocolo de autorização], [Data e hora do mdfe], XmlAutorizado, 
       Encerrado, [Responsavel do seguro], [Tipo do responsavel], 
       [Documento do responsavel], [Nome da seguradora], [Cnpj da seguradora], 
       [N da apolice], [N averbação], [Tipo de contratante], 
       [Documento do contratante], [Produto Predominante], 
       [Latitude de Carregamento], [Longitude de Carregamento], 
       [Latitude de Descrregamento], [Longitude de Descarregamento], 
       [CEP Carregamento], [CEP Descarregamento]
FROM Manifesto
```

**6. Local de Carregamento:**
```sql
SELECT [Sequencia do manifesto], Sequencia, Uf, [Descrição do municipio], 
       [Codigo do ibge]
FROM [Local de carregamento]
```

**7. Emitentes MDFe:**
```sql
SELECT [Sequencia do emitente], Cnpj, Ie, [Razão social], [Nome fantasia], 
       Logradouro, Nro, Complemento, Bairro, [Codigo do ibge], Municipio, 
       Cep, Uf, Fone, Email, Inativo, [Certificado digital], [Chave flexdocs]
FROM [Emitentes mdfe]
```

**8. Descarregamento NFe:**
```sql
SELECT [Sequencia descarrega], [Sequencia do manifesto], [Codigo do ibge], 
       [Descrição do municipio]
FROM [Descarregamento nfe]
```

**9. Descarregamento CTe:**
```sql
SELECT [Sequencia descarregamento cte], [Sequencia do manifesto], 
       [Codigo do ibge], [Descrição do municipio]
FROM [Descarregamento cte]
```

**10. UF de Percurso:**
```sql
SELECT [Sequencia do manifesto], Sequencia, Uf
FROM [Uf de percurso]
```

## ⚠️ IMPORTANTE - SITUAÇÃO ATUAL:
- **Sistema ainda usa FlexDocs** em produção
- **Teste ACBr implementado** mas não ativado por padrão  
- **Migração incremental** para evitar quebrar funcionamento
- **Backup automático** - função original preservada
- **ESTRUTURA DO BANCO**: Intocável e imutável

---

## 🔧 INFORMAÇÕES TÉCNICAS PARA CONTINUIDADE {#info-tecnicas}

### 📍 LOCALIZAÇÃO DAS MODIFICAÇÕES:
- **Arquivo principal**: `mdfe 2.txt`
- **Função de teste**: `SuperMDFe_Teste()` - linha 8617
- **Chamada redirecionada**: linha 4745 (`SuperMDFe` → `SuperMDFe_Teste`)
- **Projeto VB6**: `NFE.vbp` (atualizado com módulos corretos)

### 🛠️ DEPENDÊNCIAS VERIFICADAS:
- ✅ **ACBrMDFE32.dll** - Biblioteca principal (deve estar no PATH)
- ✅ **ACBrLibMDFe.ini** - Configuração ACBr (RS, Homologação)
- ✅ **ACBrMDFe.cls** - Classe wrapper VB6
- ✅ **ACBrMDFeUtils.bas** - Funções CreateMDFe()
- ✅ **MDFeINIUtils.bas** - Funções CriarMDFeINI(), AdicionarEmitente(), etc.

### 🔄 COMO REVERTER SE NECESSÁRIO:
```vb
' Para voltar ao FlexDocs temporariamente, alterar linha 4745:
' DE: SuperMDFe_Teste ' ETAPA 1: Testando migração ACBr
' PARA: SuperMDFe
```

### 📋 TROUBLESHOOTING:
**Se der erro "Object not found" ou "DLL not found":**
1. Verificar se `ACBrMDFE32.dll` está na pasta do sistema
2. Registrar DLL: `regsvr32 ACBrMDFE32.dll`
3. Verificar permissões da pasta `\Temp\`

**Se der erro "Bad DLL calling convention":**
1. Problema na inicialização ACBr
2. Verificar se `ACBrLibMDFe.ini` existe
3. Tentar executar como Administrador

**Se XML não for gerado:**
1. Verificar dados obrigatórios (CNPJ, UF, etc.)
2. Conferir se funções INI estão sendo chamadas
3. Validar arquivo INI criado em `\Temp\`

### 📊 LOGS E DEPURAÇÃO:
- **Observações do Manifesto**: Contém progresso detalhado
- **Arquivos temporários**: `\Temp\MDFe_Teste_*.ini` e `\Temp\MDFe_Teste_XML_*.xml`
- **Mensagem final**: "TESTE ETAPA 1 CONCLUÍDO!" (se sucesso)

### 🎯 CRITÉRIOS DE SUCESSO ETAPA 1:
1. ✅ Sistema inicializa sem erro
2. ✅ Mensagem "TESTE ETAPA 1" aparece nas observações  
3. ✅ Arquivo INI é criado em \Temp\
4. ✅ XML é gerado com tamanho > 100 caracteres
5. ✅ Dados são salvos no banco (Chave de acesso + XmlAssinado)
6. ✅ Mensagem final de sucesso é exibida

---

## 📋 RESUMO EXECUTIVO

### 🎯 SITUAÇÃO ATUAL:
**ETAPA 1 IMPLEMENTADA - AGUARDANDO TESTE**

### ⚡ AÇÃO NECESSÁRIA:
1. **EXECUTAR O SISTEMA MDFe**
2. **CRIAR/GERAR UM MANIFESTO** 
3. **VERIFICAR SE APARECE "TESTE ETAPA 1"**

### 🔄 PRÓXIMO PASSO:
- **SE SUCESSO**: Avançar para Etapa 2 (TransmitirMDFe)
- **SE ERRO**: Analisar logs e corrigir

### 📞 PARA CONTINUIDADE:
**Comando**: "Olhe o claude.md" - Esta documentação contém TUDO necessário para continuar a migração.

---
*Documentação atualizada em: Janeiro 2025*
*Status: Etapa 1 pronta para teste*