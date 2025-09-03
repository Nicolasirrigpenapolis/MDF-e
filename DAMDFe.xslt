<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html" indent="yes"/>

  <!-- formata CNPJ/CPF -->
  <xsl:template name="formatDoc">
    <xsl:param name="doc"/>
    <xsl:choose>
      <xsl:when test="string-length($doc)=14">
        <xsl:value-of select="concat(
          substring($doc,1,2),'.',
          substring($doc,3,3),'.',
          substring($doc,6,3),'/', 
          substring($doc,9,4),'-',
          substring($doc,13,2)
        )"/>
      </xsl:when>
      <xsl:when test="string-length($doc)=11">
        <xsl:value-of select="concat(
          substring($doc,1,3),'.',
          substring($doc,4,3),'.',
          substring($doc,7,3),'-',
          substring($doc,10,2)
        )"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$doc"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- formata valores numéricos -->
  <xsl:template name="formatValor">
    <xsl:param name="v"/>
    <xsl:value-of select="format-number($v,'###,##0.00')"/>
  </xsl:template>

  <!-- *** TEMPLATES DE MAPEAMENTO DE CÓDIGOS *** -->
  <!-- tpEmit -->
  <xsl:template name="getTipoEmitente">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='0'">Remetente</xsl:when>
      <xsl:when test="$codigo='1'">Expedidor</xsl:when>
      <xsl:when test="$codigo='2'">Recebedor</xsl:when>
      <xsl:when test="$codigo='3'">Destinatário</xsl:when>
      <xsl:when test="$codigo='4'">Expedidor/Remetente</xsl:when>
      <xsl:when test="$codigo='5'">Recebedor/Destinatário</xsl:when>
      <xsl:when test="$codigo='6'">Emitente</xsl:when>
      <xsl:otherwise>Desconhecido</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- modal -->
  <xsl:template name="getModal">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='1'">Rodoviário</xsl:when>
      <xsl:when test="$codigo='2'">Aéreo</xsl:when>
      <xsl:when test="$codigo='3'">Ferroviário</xsl:when>
      <xsl:when test="$codigo='4'">Aquaviário</xsl:when>
      <xsl:when test="$codigo='5'">Dutoviário</xsl:when>
      <xsl:when test="$codigo='6'">Multimodal</xsl:when>
      <xsl:otherwise>Desconhecido</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- tpRod -->
  <xsl:template name="getTipoRodado">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='00'">Não especificado</xsl:when>
      <xsl:when test="$codigo='01'">Simples</xsl:when>
      <xsl:when test="$codigo='02'">Duplo</xsl:when>
      <xsl:when test="$codigo='03'">Triplo</xsl:when>
      <xsl:when test="$codigo='04'">Roda-Simples</xsl:when>
      <xsl:when test="$codigo='05'">Outros</xsl:when>
      <xsl:otherwise>Desconhecido</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- tpCar -->
  <xsl:template name="getTipoCarroceria">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='01'">Baú</xsl:when>
      <xsl:when test="$codigo='02'">Graneleiro</xsl:when>
      <xsl:when test="$codigo='03'">Tanque</xsl:when>
      <xsl:when test="$codigo='04'">Reboque</xsl:when>
      <xsl:when test="$codigo='05'">Container</xsl:when>
      <xsl:when test="$codigo='06'">Outros</xsl:when>
      <xsl:otherwise>Desconhecido</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- tpEmis -->
  <xsl:template name="getTipoEmissao">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='1'">Emissão Normal</xsl:when>
      <xsl:when test="$codigo='2'">Contingência FS-IA</xsl:when>
      <xsl:when test="$codigo='3'">Contingência DPEC</xsl:when>
      <xsl:when test="$codigo='4'">Contingência FS-DA</xsl:when>
      <xsl:when test="$codigo='5'">Contingência SVC-RS</xsl:when>
      <xsl:when test="$codigo='6'">Contingência SVC-SP</xsl:when>
      <xsl:otherwise>Desconhecido</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- respSeg -->
  <xsl:template name="getTipoResponsavelSeguro">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='0'">Não Segurado</xsl:when>
      <xsl:when test="$codigo='1'">Emitente</xsl:when>
      <xsl:when test="$codigo='2'">Tomador</xsl:when>
      <xsl:when test="$codigo='3'">Transportador</xsl:when>
      <xsl:otherwise>Desconhecido</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- tpCarga -->
  <xsl:template name="getTipoCarga">
    <xsl:param name="codigo"/>
    <xsl:choose>
      <xsl:when test="$codigo='01'">Granel Sólido</xsl:when>
      <xsl:when test="$codigo='02'">Granel Líquido</xsl:when>
      <xsl:when test="$codigo='03'">Granel Gasoso</xsl:when>
      <xsl:when test="$codigo='04'">Solta/Unitizada</xsl:when>
      <xsl:when test="$codigo='05'">Conteinerizada</xsl:when>
      <xsl:when test="$codigo='06'">Neogranel</xsl:when>
      <xsl:when test="$codigo='07'">Emergência</xsl:when>
      <xsl:when test="$codigo='09'">Carga Viva</xsl:when>
      <xsl:otherwise>Outros</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- *** FIM TEMPLATES DE CÓDIGOS *** -->

  <xsl:template match="/">
    <html>
      <head>
        <meta charset="UTF-8"/>
        <style>
          body { font-family:Arial,sans-serif; margin:20px }
          h1 { text-align:center; margin-bottom:20px }
          h2 { margin-top:30px; border-bottom:1px solid #ccc; padding-bottom:5px }
          table { width:100%; border-collapse:collapse; margin-top:10px }
          th,td { border:1px solid #666; padding:8px; text-align:left }
          th { background:#eee; width:30% }
        </style>
      </head>
      <body>
        <h1>DAMDFE – Dados Completos</h1>

       
	   <!-- GRUPO 1: Identificação do Manifesto -->
<h2>1. Identificação do Manifesto</h2>
<table>
  <tr>
    <th>1 – Número do Manifesto</th>
    <td>
      <xsl:value-of select="//*[local-name()='nMDF']"/>
    </td>
  </tr>
  <tr>
    <th>2 – Data de Emissão</th>
    <td>
      <xsl:value-of select="//*[local-name()='dhEmi']"/>
    </td>
  </tr>
  <tr>
    <th>3 – UF Inicial</th>
    <td>
      <xsl:value-of select="//*[local-name()='UFIni']"/>
    </td>
  </tr>
  <tr>
    <th>4 – Tipo de Emitente</th>
    <td>
      <xsl:call-template name="getTipoEmitente">
        <xsl:with-param name="codigo"
          select="//*[local-name()='tpEmit']"/>
      </xsl:call-template>
    </td>
  </tr>
  <tr>
    <th>5 – UF de Descarregamento</th>
    <td>
      <xsl:value-of select="//*[local-name()='UFFim']"/>
    </td>
  </tr>
  <!-- UF de Percurso: se existir infPercurso, lista todos UFPer; caso contrário, mostra UFIni e UFFim -->
  <tr>
    <th>— UF de Percurso</th>
    <td>
      <xsl:choose>
        <xsl:when test="//*[local-name()='infPercurso']">
          <xsl:for-each select="//*[local-name()='infPercurso']">
            <xsl:value-of select="*[local-name()='UFPer']"/>
            <xsl:if test="position() != last()">, </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="//*[local-name()='UFIni']"/><xsl:text>, </xsl:text>
          <xsl:value-of select="//*[local-name()='UFFim']"/>
        </xsl:otherwise>
      </xsl:choose>
    </td>
  </tr>
  <tr>
    <th>6 – Observação (ID)</th>
    <td>
      <xsl:value-of select="//*[local-name()='infMDFe']/@Id"/>
    </td>
  </tr>
  <tr>
    <th>7 – RNTRC</th>
    <td>
      <xsl:value-of select="//*[local-name()='RNTRC']"/>
    </td>
  </tr>
</table>

	   
	   
	   

        <!-- GRUPO 2: Veículo -->
        <h2>2. Dados do Veículo</h2>
        <table>
          <tr><th>8 – Tipo de Carroceria</th>
              <td>
                <xsl:call-template name="getTipoCarroceria">
                  <xsl:with-param name="codigo"
                    select="//*[local-name()='tpCar']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>9 – UF do Veículo</th>
              <td><xsl:value-of select="//*[local-name()='veicTracao']/*[local-name()='UF']"/></td></tr>
          <tr><th>10 – Tipo de Rodado</th>
              <td>
                <xsl:call-template name="getTipoRodado">
                  <xsl:with-param name="codigo"
                    select="//*[local-name()='tpRod']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>11 – Placa</th>
              <td><xsl:value-of select="//*[local-name()='placa']"/></td></tr>
          <tr><th>12 – Tara</th>
              <td><xsl:value-of select="//*[local-name()='tara']"/></td></tr>
          <tr><th>13 – Renavam</th>
              <td><xsl:value-of select="//*[local-name()='veicTracao']/*[local-name()='renavam']"/></td></tr>
          <tr><th>14 – Capacidade (kg)</th>
              <td><xsl:value-of select="//*[local-name()='qCarga']"/></td></tr>
        </table>

        <!-- GRUPO 3: Emitente e Autorização -->
        <h2>3. Emitente e Autorização</h2>
        <table>
          <tr><th>15 – CNPJ do Emitente</th>
              <td>
                <xsl:call-template name="formatDoc">
                  <xsl:with-param name="doc"
                    select="//*[local-name()='emit']/*[local-name()='CNPJ']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>16 – Versão do Processo</th>
              <td><xsl:value-of select="//*[local-name()='verProc']"/></td></tr>
          <tr><th>17 – Status da Emissão</th>
              <td><xsl:value-of select="//*[local-name()='cStat']"/></td></tr>
          <tr><th>18 – Motivo</th>
              <td><xsl:value-of select="//*[local-name()='xMotivo']"/></td></tr>
        </table>

        <!-- GRUPO 4: Contratante e Proprietário -->
        <h2>4. Contratante e Proprietário</h2>
        <table>
          <tr><th>19 – Documento do Contratante</th>
              <td>
                <xsl:call-template name="formatDoc">
                  <xsl:with-param name="doc"
                    select="//*[local-name()='infContratante']/*[local-name()='CNPJ' or local-name()='CPF']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>20 – CPF do Condutor</th>
              <td><xsl:value-of select="//*[local-name()='condutor']/*[local-name()='CPF']"/></td></tr>
          <tr><th>21 – RNTRC do Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='infANTT']/*[local-name()='RNTRC']"/></td></tr>
          <tr><th>22 – Nome do Proprietário</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infContratante']/*[local-name()='xNome']">
                    <xsl:value-of select="//*[local-name()='infContratante']/*[local-name()='xNome']"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="//*[local-name()='emit']/*[local-name()='xNome']"/>
                  </xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>23 – IE do Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='IE']"/></td></tr>
          <tr><th>24 – UF do Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='emit']/*[local-name()='enderEmit']/*[local-name()='UF']"/></td></tr>
        </table>

        <!-- GRUPO 5: Protocolos e Chave -->
        <h2>5. Protocolos e Chave</h2>
        <table>
          <tr><th>25 – Tipo de Documento</th>
              <td>
                <xsl:call-template name="getTipoEmissao">
                  <xsl:with-param name="codigo"
                    select="//*[local-name()='tpEmis']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>26 – Número do Recibo</th>
              <td><xsl:value-of select="//*[local-name()='protMDFe']/*[local-name()='nProt']"/></td></tr>
          <tr><th>27 – Chave de Acesso</th>
              <td>
                <xsl:value-of select="substring(//*[local-name()='infMDFe']/@Id,5)"/>
              </td></tr>
          <tr><th>28 – Data e Hora da Autorização</th>
              <td><xsl:value-of select="//*[local-name()='protMDFe']/*[local-name()='dhRecbto']"/></td></tr>
        </table>

        <!-- GRUPO 6: Seguro -->
        <h2>6. Seguro</h2>
        <table>
          <tr><th>29 – Responsável pelo Seguro</th>
              <td>
                <xsl:call-template name="getTipoResponsavelSeguro">
                  <xsl:with-param name="codigo"
                    select="//*[local-name()='infResp']/*[local-name()='respSeg']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>30 – Documento do Responsável</th>
              <td>
                <xsl:call-template name="formatDoc">
                  <xsl:with-param name="doc"
                    select="//*[local-name()='infSeg']/*[local-name()='CNPJ' or local-name()='CPF']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>31 – Nome da Seguradora</th>
              <td><xsl:value-of select="//*[local-name()='infSeg']/*[local-name()='xSeg']"/></td></tr>
          <tr><th>32 – CNPJ da Seguradora</th>
              <td>
                <xsl:call-template name="formatDoc">
                  <xsl:with-param name="doc"
                    select="//*[local-name()='infSeg']/*[local-name()='CNPJ']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>33 – Nº da Apólice</th>
              <td><xsl:value-of select="//*[local-name()='nApol']"/></td></tr>
          <tr><th>34 – Nº da Averbação</th>
              <td><xsl:value-of select="//*[local-name()='nAver']"/></td></tr>
        </table>

        <!-- GRUPO 7: Localização -->
        <h2>7. Localização</h2>
        <table>
          <tr><th>35 – CEP de Carregamento</th>
              <td><xsl:value-of select="//*[local-name()='infLocalCarrega']/*[local-name()='CEP']"/></td></tr>
          <tr><th>36 – Latitude de Carregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalCarrega']/*[local-name()='latitude']">
                    <xsl:value-of select="//*[local-name()='infLocalCarrega']/*[local-name()='latitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>37 – Longitude de Carregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalCarrega']/*[local-name()='longitude']">
                    <xsl:value-of select="//*[local-name()='infLocalCarrega']/*[local-name()='longitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>38 – CEP de Descarregamento</th>
              <td><xsl:value-of select="//*[local-name()='infLocalDescarrega']/*[local-name()='CEP']"/></td></tr>
          <tr><th>39 – Latitude de Descarregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalDescarrega']/*[local-name()='latitude']">
                    <xsl:value-of select="//*[local-name()='infLocalDescarrega']/*[local-name()='latitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>40 – Longitude de Descarregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalDescarrega']/*[local-name()='longitude']">
                    <xsl:value-of select="//*[local-name()='infLocalDescarrega']/*[local-name()='longitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
        </table>

        <!-- GRUPO 8: Dados Adicionais -->
        <h2>8. Dados Adicionais</h2>
        <table>
          <tr><th>41 – Sequência do Manifesto</th>
              <td><xsl:value-of select="//*[local-name()='cMDF']"/></td></tr>
          <tr><th>42 – Código do Emitente</th>
              <td><xsl:value-of select="//*[local-name()='emit']/*[local-name()='CNPJ']"/></td></tr>
          <tr><th>43 – Transmitido</th>
              <td><xsl:value-of select="//*[local-name()='ide']/*[local-name()='tpAmb']"/></td></tr>
          <tr><th>44 – Nota Cancelada</th>
              <td>Não</td></tr>
          <tr><th>45 – Autorizada</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='cStat']='100'">Sim</xsl:when>
                  <xsl:otherwise>Não</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>46 – Histórico</th>
              <td><xsl:value-of select="//*[local-name()='xMotivo']"/></td></tr>
          <tr><th>47 – Proprietário</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infContratante']/*[local-name()='xNome']">
                    <xsl:value-of select="//*[local-name()='infContratante']/*[local-name()='xNome']"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="//*[local-name()='emit']/*[local-name()='xNome']"/>
                  </xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>48 – CPF Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='condutor']/*[local-name()='CPF']"/></td></tr>
          <tr><th>49 – CNPJ Proprietário</th>
              <td>
                <xsl:call-template name="formatDoc">
                  <xsl:with-param name="doc"
                    select="//*[local-name()='infContratante']/*[local-name()='CNPJ']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>50 – RNTRC Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='infANTT']/*[local-name()='RNTRC']"/></td></tr>
          <tr><th>51 – IE Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='IE']"/></td></tr>
          <tr><th>52 – UF Proprietário</th>
              <td><xsl:value-of select="//*[local-name()='emit']/*[local-name()='enderEmit']/*[local-name()='UF']"/></td></tr>
          <tr><th>53 – Tipo de Proprietário</th>
              <td>
                <xsl:call-template name="getTipoEmitente">
                  <xsl:with-param name="codigo"
                    select="//*[local-name()='tpEmit']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>54 – XML Assinado</th>
              <td><xsl:copy-of select="/*/*[local-name()='MDFe']"/></td></tr>
          <tr><th>55 – Protocolo de Autorização</th>
              <td><xsl:value-of select="//*[local-name()='protMDFe']/*[local-name()='nProt']"/></td></tr>
          <tr><th>56 – Data e Hora do MDFe</th>
              <td><xsl:value-of select="//*[local-name()='protMDFe']/*[local-name()='dhRecbto']"/></td></tr>
          <tr><th>57 – XML Autorizado</th>
              <td><xsl:copy-of select="/*/*[local-name()='protMDFe']"/></td></tr>
          <tr><th>58 – Encerrado</th>
              <td>Não</td></tr>
          <tr><th>59 – Tipo de Contratante</th>
              <td>
                <xsl:call-template name="getTipoEmitente">
                  <xsl:with-param name="codigo"
                    select="//*[local-name()='tpEmit']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>60 – Documento do Contratante</th>
              <td>
                <xsl:call-template name="formatDoc">
                  <xsl:with-param name="doc"
                    select="//*[local-name()='infContratante']/*[local-name()='CNPJ' or local-name()='CPF']"/>
                </xsl:call-template>
              </td></tr>
          <tr><th>61 – Produto Predominante</th>
              <td><xsl:value-of select="//*[local-name()='prodPred']/*[local-name()='xProd']"/></td></tr>
          <tr><th>62 – Latitude de Carregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalCarrega']/*[local-name()='latitude']">
                    <xsl:value-of select="//*[local-name()='infLocalCarrega']/*[local-name()='latitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>63 – Longitude de Carregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalCarrega']/*[local-name()='longitude']">
                    <xsl:value-of select="//*[local-name()='infLocalCarrega']/*[local-name()='longitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>64 – Latitude de Descarregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalDescarrega']/*[local-name()='latitude']">
                    <xsl:value-of select="//*[local-name()='infLocalDescarrega']/*[local-name()='latitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>65 – Longitude de Descarregamento</th>
              <td>
                <xsl:choose>
                  <xsl:when test="//*[local-name()='infLocalDescarrega']/*[local-name()='longitude']">
                    <xsl:value-of select="//*[local-name()='infLocalDescarrega']/*[local-name()='longitude']"/>
                  </xsl:when>
                  <xsl:otherwise>Não informado</xsl:otherwise>
                </xsl:choose>
              </td></tr>
          <tr><th>66 – CEP de Carregamento</th>
              <td><xsl:value-of select="//*[local-name()='infLocalCarrega']/*[local-name()='CEP']"/></td></tr>
          <tr><th>67 – CEP de Descarregamento</th>
              <td><xsl:value-of select="//*[local-name()='infLocalDescarrega']/*[local-name()='CEP']"/></td></tr>
        </table>

      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>
