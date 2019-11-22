#Include "TOPCONN.CH"
#Include "FWMVCDef.ch"
#Include "COLORS.CH"
#Include "FONT.CH"
#Include "ApWizard.ch"
#Include "FILEIO.CH"
#include "TOTVS.CH"

#DEFINE BR Chr(10)
#DEFINE _CRLF Chr(13) + Chr(10)

//------------------------------------------------------------------------------------------
/* {Protheus.doc} zMATCCusto
Monitor de Custos

@author    Ronaldo Tapia
@version   12.1.17
@since     10/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
USER Function zMATCCusto()

	// Variáveis utilizadas na exportação do arquivo texto
	Private aDadosAmb := {}
	Private aFontes   := {}
	Private aVlCampos := {}
	Private aProc	  := {}
	Private aPE		  := {}
	Private aParam	  := {}
	Private oDlg	  := Nil

	// Variaveis utilizadas para analise de custo
	Private oFolder2  := Nil
	Private cProduto  := Space(15)
	Private cDescr    := Space(40)
	Private dUltFech  := dDataI := dDataF := CTOD("  /  /    ")
	Private cLocal    := Space(02)
	Private dDtProces := dDataBase
	Private aKarLocal := {}
	Private aKarLote  := {}
	Private oListCusto:= Nil
	Private oLBCabec  := Nil
	Private oLBDetalhe:= Nil
	Private aArrayCV8 := {{'','','','','','',''}}
	Private aArrayDET := {{'','','','','','',''}}
	Private cCrlLot   := ""
	Private cCrlEnd   := ""
	Private c1UM      := ""
	Private c2UM      := ""
	Private cZona     := ""
	Private cPict     := PesqPict('SB2','B2_QATU')
	Private cPict2UM  := PesqPict('SB2','B2_QTSEGUM')
	Private cPictD5   := PesqPict('SD5','D5_QUANT')
	Private cPictDB   := PesqPict('SDB','DB_QUANT')
	Private cPictB1   := PesqPict('SB1','B1_CONV')
	Private cPictLoc  := PesqPict('SB2','B2_LOCAL')
	Private nSldKarLocal := nSldKarLote := nSldKarEnde := 0
	Private aSx3Box      := RetSx3Box(Posicione('SX3',2,'BE_STATUS','X3CBox()'),,,1)
	Private _lReturn     := .F.
	Private lNumLote 	 := SuperGetMV('MV_LOTEUNI', .F., .F.)
	Private lCusfilA 	 := .F.
	Private lCusfilF 	 := .F.
	Private lCusfilE 	 := .F.
	Private oQtSB9
	Private nQtSB9
	Private oVlSB9
	Private nVLSB9
	Private oVlOpSB9
	Private nVlOpSB9
	Private oQtSD1
	Private nQtSD1
	Private oVlSD1
	Private nVlSD1
	Private oQtSD2
	Private nQtSD2
	Private oVlSD2
	Private nVlSD2
	Private nQtPSD3
	Private nVlPSD3
	Private nVlOpPSD3
	Private oQtPSD3
	Private oVlPSD3
	Private oVlOpPSD3
	Private oQtRSD3
	Private nQtRSD3
	Private oVlRSD3
	Private nVlRSD3
	Private oVlOpRSD3
	Private nVlOpRSD3
	Private oQtDSD3
	Private nQtDSD3
	Private oVlDSD3
	Private nVlDSD3
	Private oVlOpDSD3
	Private nVlOpDSD3
	Private oFecQtSB2
	Private nFecQtSB2
	Private oFecVlSB2
	Private nFecVlSB2
	Private oFecVlOpSB2
	Private nFecVlOpSB2
	Private oFecMovQt
	Private nFecMovQt
	Private oFecMovVl
	Private nFecMovVl
	Private oFecMovOp
	Private nFecMovOp
	Private oSay
	Private oBtErro2 := Nil
	Private oBtErro  := Nil

	// Contabilização
	Private oVlContab
	Private nVlContab
	Private oVlCT2
	Private nVlCT2

	// Variaveis para analise de divergencias de produtos
	Private cProdAnDe  := Space(20)
	Private cProdAnAte := Space(20)
	Private cLocAnDe   := Space(02)
	Private cLocAnAte  := Space(02)
	Private dIniAn     := CTOD("  /  /    ")
	Private dFimAn	   := CTOD("  /  /    ")
	Private oFather    := Nil
	Private dFechAn	   := CTOD("  /  /    ")
	Private aProdAn    := {{'','','',''}}
	Private oTempTable := Nil
	Private lLogDec	   := .F.

	//Variáveis utilizadas para atribuição das propriedades do CSS
	Private cCSSTFol  := ""
	Private cCSSTCBrw := ""
	Private cCSSTBut  := ""

	//Outras
	Private lKardex1  := .T. // Ordena por NUMSEQ
	Private lKardex2  := .F. // Ordena por SEQCALC

	// Monta interface para exibição dos dados
	MATCDiag(aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCDiag
Monta interface para exibição dos dados

@Param
aDadosAmb - array com dados do Ambiente (Server)
aFontes   - array com os procipais fontes de custos
aVlCampos - array com os campos para validação de decimal
aProc     - array de procedures
aPE		  - array de pontos de entrada
aParam    - array de parametros

@author    Ronaldo Tapia
@version   12.1.17
@since     10/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCDiag(aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam)

	Local aSizeMain	:= FWGetDialogSize( oMainWnd )
	Local aTFolder  := {}
	Local oFolder1  := Nil
	Local oLayer1	:= Nil
	Local oLayer2	:= Nil
	Local oLayer3	:= Nil
	Local oLayer4	:= Nil
	Local oLayer5	:= Nil
	Local oScr1		:= Nil
	Local oScr2		:= Nil
	Local oScr3		:= Nil
	Local oScr4		:= Nil

	DEFINE MSDIALOG oDlg TITLE "Monitor de Custos" FROM aSizeMain[1], aSizeMain[2] TO aSizeMain[3], aSizeMain[4] COLORS 0, 16777215 PIXEL

	// Cria os objetos tipo FWLayer
	oLayer1   := FWLayer():New()
	oLayer2   := FWLayer():New()
	oLayer3   := FWLayer():New()
	oLayer4   := FWLayer():New()
	oLayer5   := FWLayer():New()

	// Cria Folder para separar telas
	aTFolder := { "Ambiente","Validações","Parâmetros","Análise","Recálculo"}
	oFolder1 := TFolder():New( 1,2,aTFolder,,oDlg,,,,.T.,,aSizeMain[4]/2-5, aSizeMain[3]/2 )
	oFolder1:lVisible
	oFolder1:show()
	oFolder1:bSetOption := {|nAtu| MATCValFo(nAtu) } // Validação na mudança de aba

	// Função para setar as propriedades do CSS
	MATCSCSS()

	// Instancia objeto CSS
	oFolder1:SetCss(cCSSTFol)

	// Layer 1 - aDialogs 1
	oLayer1:Init( oFolder1:aDialogs[1], .T. )
	oLayer1:AddLine( "LINE01", 60 )
	oLayer1:AddCollumn( "BOX01", 33,, "LINE01" )
	oLayer1:AddWindow( "BOX01", "PANEL01", "Diagnóstico do Ambiente", 100, .F.,,, "LINE01" ) //Diagnóstico do Sistema
	FPanel01( oLayer1:GetWinPanel( "BOX01", "PANEL01", "LINE01" ),aDadosAmb ) //Construção do painel - Diagnóstico do Sistema

	oLayer1:AddCollumn( "BOX02", 34,, "LINE01" )
	oLayer1:AddWindow( "BOX02", "PANEL02", "Principais Fontes", 100, .F.,,, "LINE01" ) //Principais Fontes
	FPanel01_1( oLayer1:GetWinPanel( "BOX02", "PANEL02", "LINE01" ),aFontes ) //Construção do painel - Principais Fontes

	oLayer1:AddLine( "LINE02", 40 )
	oLayer1:AddCollumn( "BOX04", 33,, "LINE02" )
	oLayer1:AddWindow( "BOX04", "PANEL04", "Funções", 100, .F.,,, "LINE02" ) //Funções
	FPanel01_2( oLayer1:GetWinPanel( "BOX04", "PANEL04", "LINE02" )) //Construção do painel - Funções

	oLayer1:AddCollumn( "BOX03", 34,, "LINE02" )
	oLayer1:AddWindow( "BOX03", "PANEL03", "Procedures", 100, .F.,,, "LINE02" ) //Procedures
	FPanel01_3( oLayer1:GetWinPanel( "BOX03", "PANEL03", "LINE02" ),aProc ) //Construção do painel - Procedures

	oLayer1:AddCollumn( "BOX05", 33,, "LINE01" )
	oLayer1:AddWindow( "BOX05", "PANEL05", "Pontos de Entrada", 100, .F.,,, "LINE01" ) //Pontos de Entrada
	/*Cria um objeto tipo TScrollBox dentro do Painel 1 / Box05*/
	oScr1 := TScrollBox():New(oLayer1:GetWinPanel( "BOX05", "PANEL05", "LINE01" ),001,001,aSizeMain[3]/2-149,aSizeMain[4]/2-466,.T.,.T.,.F.)
	FPanel01_4( oLayer1:GetWinPanel( "BOX05", "PANEL05", "LINE01" ),oScr1,aPE ) //Construção do painel - Pontos de Entrada

	oLayer1:AddCollumn( "BOX06", 33,, "LINE02" )
	oLayer1:AddWindow( "BOX06", "PANEL06", "Links Importantes", 100, .F.,,, "LINE02" ) //"Links Importantes"
	FPanel01_5( oLayer1:GetWinPanel( "BOX06", "PANEL06", "LINE02" ),aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam ) //Construção do painel - "Links Importantes"

	// Layer 2 - aDialogs 2
	oLayer2:Init( oFolder1:aDialogs[2], .T. )
	oLayer2:AddLine( "LINE01", 100 )
	oLayer2:AddCollumn( "BOX01", 100,, "LINE01" )
	oLayer2:AddWindow( "BOX01", "PANEL01", "Validação de Campos", 100, .F.,,, "LINE01" ) //Validação de Campos

	// Cria um objeto tipo TScrollBox dentro do Painel 2
	oScr2 := TScrollBox():New(oLayer2:GetWinPanel( "BOX01", "PANEL01", "LINE01" ),003,003,220,530,.F.,.T.,.F.)
	FPanel02( oLayer2:GetWinPanel( "BOX01", "PANEL01", "LINE01" ),oScr2,aVlCampos ) //Construção do painel - "Validação de Campos"

	// Layer 3 - aDialogs 3
	oLayer3:Init( oFolder1:aDialogs[3], .T. )
	oLayer3:AddLine( "LINE01", 100 )
	oLayer3:AddCollumn( "BOX01", 100,, "LINE01" )
	oLayer3:AddWindow( "BOX01", "PANEL01", "Parâmetros)", 100, .F.,,, "LINE01" ) //"Parâmetros"

	// Cria um objeto tipo TScrollBox dentro do Painel 4
	oScr3 := TScrollBox():New(oLayer3:GetWinPanel( "BOX01", "PANEL01", "LINE01" ),003,003,aSizeMain[3]/2-80,aSizeMain[4]/2-30,.T.,.T.,.F.)
	FPanel03( oLayer3:GetWinPanel( "BOX01", "PANEL01", "LINE01" ),oScr3 ,aParam) //Construção do painel - "Parâmetros"

	// Layer 4 - aDialogs 4
	oLayer4:Init( oFolder1:aDialogs[4], .T. )
	oLayer4:AddLine( "LINE01", 100 )
	oLayer4:AddCollumn( "BOX01", 100,, "LINE01" )
	oLayer4:AddWindow( "BOX01", "PANEL01", "Análise de Custos", 100, .F.,,, "LINE01" ) //Análise de Custos
	FPanel04( oLayer4:GetWinPanel( "BOX01", "PANEL01", "LINE01" ), ) //Construção do painel - Análise

	// Layer 5 - aDialogs 5
	oLayer5:Init( oFolder1:aDialogs[5], .T. )
	oLayer5:AddLine( "LINE01", 23 )
	oLayer5:AddCollumn( "BOX01", 100,, "LINE01" )
	oLayer5:AddWindow( "BOX01", "PANEL01", "Recálculo", 100, .F.,,, "LINE01" ) //Recálculo
	FPanel05( oLayer5:GetWinPanel( "BOX01", "PANEL01", "LINE01" ), ) //Construção do painel - Recálculo

	oLayer5:AddLine( "LINE02", 77 )
	oLayer5:AddCollumn( "BOX02", 100,, "LINE02" )
	oLayer5:AddWindow( "BOX02", "PANEL02", "Logs de Execução", 100, .F.,,, "LINE02" )  //Recalculo
	FPanel05_1( oLayer5:GetWinPanel( "BOX02", "PANEL02", "LINE02" ),aFontes ) //Construção do painel - Logs de Execução

	ACTIVATE MSDIALOG oDlg CENTERED

Return


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel01
Painel1 - Folder1 - Carrega Dados do Ambiente

@Param
oPanel    - objeto da tela
aDadosAmb - array com dados do server

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------

Static Function FPanel01(oPanel,aDadosAmb)

	Local nLin	 := 005
	Local aDados1, aDados2, aDados3, aDados4, aDados5, aDados6, aDados7, aDados8, aDados9, aDados10 := {}
	Local oValor :=	TFont():New( "Courier new",, -14,, .F. )
	Local oValorB :=	TFont():New( "Courier new",, -14,, .T. )

	// Dados do Server
	Local cVBuild 	:= GetBuild() 		// Traz informações do Application Server
	Local cVClient 	:= GetBuild( .T. )  // Traz informações do SmartClient
	Local cVDBAcc	:= TCGetBuild() 	// Traz informações do DBAccess
	Local cData		:=	"" 				// Valida data de Lib
	Local cRpoRelease := GetRpoRelease()
	Local lTopMemo	:= GetSrvProfString("TOPMEMOMEGA","") == "1" // Valida seção TopMemoMega
	Local lSrvUnix  := IsSrvUnix()
	Local cAmbiente := Capital(GetEnvServer())
	Local cLocFiles := GetSrvProfString("LocalFiles","Ctree")
	Local cTopDaBase:= TCGetDB()

	Default aDadosAmb  := {} // Array para exportar dados em txt

	TSay():New( nLin, 007, { || "Dados do Server"}        , oPanel,,oValorB,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 15

	// Build Server
	TSay():New( nLin, 007, { || "Build Server: " + cVBuild }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"Build Server",cVBuild})

	// Build Client
	TSay():New( nLin, 007, { || "SmartClient: " + cVClient }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"SmartClient",cVClient})

	// DbAccess
	TSay():New( nLin, 007, { || "Build DbAccess: " + cVDBAcc }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"Build DbAccess",cVDBAcc})

	// Valida a LIB
	If FindFunction( "__FWLibVersion" ) //Disponével somente a partir da label 20151008
		cData := __FWLibVersion()
	Else
		//Verifica data de um fonte da LIB que foi alterado na release 12.1.005
		cData := DToS( GetAPOInfo( "PROTHEUSFUNCTIONMVC.PRX" )[4] )
	EndIf
	TSay():New( nLin, 007, { || "LIB: " + cData }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"LIB",cData})

	// Release do RPO
	TSay():New( nLin, 007, { || "Release do RPO: " + cRpoRelease }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"Release do RPO",cRpoRelease})

	// Ambiente
	TSay():New( nLin, 007, { || "Ambiente: " + cAmbiente }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"Ambiente",cAmbiente})

	// Local Files
	TSay():New( nLin, 007, { || "Local Files: " + cLocFiles }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"Local Files",cLocFiles})

	// DataBase
	TSay():New( nLin, 007, { || "TOP DataBase: " + cTopDaBase }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aAdd(aDadosAmb, {"TOP DataBase",cTopDaBase})

	// Valida a chave TopMemoMega no .INI
	If lTopMemo
		TSay():New( nLin, 007, { || "TopMemoMega: Habilitado" }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
		nLin := nLin + 10
		aAdd(aDadosAmb, {"TopMemoMega", "Habilitado"})
	Else
		TSay():New( nLin, 007, { || "TopMemoMega: Desabilitado" }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
		nLin := nLin + 10
		aAdd(aDadosAmb, {"TopMemoMega","Desabilitado"})
	EndIf

	If ( lSrvUnix )
		TSay():New( nLin, 007, { || "Servidor rodando em ambiente: Linux" }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
		nLin := nLin + 10
		aAdd(aDadosAmb, {"Servidor", "Rodando em ambiente Linux"})
	Else
		TSay():New( nLin, 007, { || "Servidor rodando em ambiente: Windows" }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
		nLin := nLin + 10
		aAdd(aDadosAmb, {"Servidor", "Rodando em ambiente Windows"})
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel01_1
FPanel01_1 - Folder1 - Carrega Dados do Ambiente

@Param
oPanel    - objeto da tela
aFontes   - array com principais fontes de custos

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel01_1(oPanel,aFontes)

	// Data dos Fontes
	Local aDados1, aDados2, aDados3, aDados4, aDados5, aDados6, aDados7, aDados8, aDados9, aDados10, aDados11, aDados12 := {}
	Local oBtFontes	:=	Nil
	Local nLin 		:= 005
	Local oValor 	:=	TFont():New( "Courier new",, -14,, .F. )
	Local oValorB   :=	TFont():New( "Courier new",, -14,, .T. )

	Default aFontes  := {} // Array para exportar dados em txt

	TSay():New( nLin, 007, { || "Principais Fontes - Custos"}        , oPanel,,oValorB,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 15

	// Pego os dados do fonte com a função GetAPOInfo
	aDados1 := GetAPOInfo("sigacus.prw")
	aAdd(aFontes, aDados1[1]+ "  - " + cValToChar(aDados1[4])+ " - " + aDados1[5] )
	TSay():New( nLin, 007, { || aDados1[1]+ "  - " + cValToChar(aDados1[4])+ " - " + aDados1[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados2 := GetAPOInfo("sigacusa.prx")
	aAdd(aFontes, aDados2[1]+ " - " + cValToChar(aDados2[4])+ " - " + aDados2[5] )
	TSay():New( nLin, 007, { || aDados2[1]+ " - " + cValToChar(aDados2[4])+ " - " + aDados2[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados3 := GetAPOInfo("sigacusb.prx")
	aAdd(aFontes, aDados3[1]+ " - " + cValToChar(aDados3[4])+ " - " + aDados3[5] )
	TSay():New( nLin, 007, { || aDados3[1]+ " - " + cValToChar(aDados3[4])+ " - " + aDados3[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados4 := GetAPOInfo("MATA280.PRX")
	aAdd(aFontes, aDados4[1]+ "  - " + cValToChar(aDados4[4])+ " - " + aDados4[5] )
	TSay():New( nLin, 007, { || aDados4[1]+ "  - " + cValToChar(aDados4[4])+ " - " + aDados4[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados5 := GetAPOInfo("MATA300.PRX")
	aAdd(aFontes, aDados5[1]+ "  - " + cValToChar(aDados5[4])+ " - " + aDados5[5] )
	TSay():New( nLin, 007, { || aDados5[1]+ "  - " + cValToChar(aDados5[4])+ " - " + aDados5[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados6 := GetAPOInfo("MATA330.PRX")
	aAdd(aFontes, aDados6[1]+ "  - " + cValToChar(aDados6[4])+ " - " + aDados6[5] )
	TSay():New( nLin, 007, { || aDados6[1]+ "  - " + cValToChar(aDados6[4])+ " - " + aDados6[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados7 := GetAPOInfo("mata331.prx")
	aAdd(aFontes, aDados7[1]+ "  - " + cValToChar(aDados7[4])+ " - " + aDados7[5] )
	TSay():New( nLin, 007, { || aDados7[1]+ "  - " + cValToChar(aDados7[4])+ " - " + aDados7[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados8 := GetAPOInfo("matr900.prx")
	aAdd(aFontes, aDados8[1]+ "  - " + cValToChar(aDados8[4])+ " - " + aDados8[5] )
	TSay():New( nLin, 007, { || aDados8[1]+ "  - " + cValToChar(aDados8[4])+ " - " + aDados8[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados9 := GetAPOInfo("matr910.prx")
	aAdd(aFontes, aDados9[1]+ "  - " + cValToChar(aDados9[4])+ " - " + aDados9[5] )
	TSay():New( nLin, 007, { || aDados9[1]+ "  - " + cValToChar(aDados9[4])+ " - " + aDados9[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados10 := GetAPOInfo("cfgx051.prw")
	aAdd(aFontes, aDados10[1]+ "  - " + cValToChar(aDados10[4])+ " - " + aDados10[5] )
	TSay():New( nLin, 007, { || aDados10[1]+ "  - " + cValToChar(aDados10[4])+ " - " + aDados10[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
	nLin := nLin + 10
	aDados11 := GetAPOInfo("mata333.prx")
	If Len(aDados11) > 0
		aAdd(aFontes, aDados11[1]+ " - " + cValToChar(aDados11[4])+ " - " + aDados11[5] )
		TSay():New( nLin, 007, { || aDados11[1]+ "  - " + cValToChar(aDados11[4])+ " - " + aDados11[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
		nLin := nLin + 10
	EndIf
	aDados12 := GetAPOInfo("m333jctb.prx")
	If Len(aDados12) > 0
		aAdd(aFontes, aDados12[1]+ " - " + cValToChar(aDados12[4])+ " - " + aDados12[5] )
		TSay():New( nLin, 007, { || aDados12[1]+ " - " + cValToChar(aDados12[4])+ " - " + aDados12[5] }        , oPanel,,oValor,,,, .T.,    ,, 200, 010 )
		nLin := nLin + 10
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel01_2
FPanel01_2 - Folder1 - Funções

@Param
oPanel - objeto da tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel01_2(oPanel)

	// Data dos Fontes
	Local oBtFontes,oBtHist,oBtExpor :=	Nil
	Local oValorN 	:=	TFont():New( "Courier new",, -14,, .T. )

	TSay():New( 005, 007, { || "Data dos Fontes" }, oPanel,,oValorN,,,, .T.,CLR_BLACK,, 300, 010 ) //"Data dos Fontes"
	TSay():New( 032, 007, { || "Histórico de Atualizações" }, oPanel,,oValorN,,,, .T.,CLR_BLACK,, 500, 010 ) //"Histórico de Atualizações"
	TSay():New( 059, 007, { || "Exportação de Dados"},oPanel,,oValorN,,,, .T.,CLR_BLACK,, 500, 010 )

	// Botão para chamar a tela de diagnóstico de fontes
	oBtFontes := TButton():New( 015, 007, "Verificar Fontes",oPanel,{||MATCFontes()}, 60,10,,,.F.,.T.,.F.,,.F.,,,.F. ) //"Fontes"
	oBtHist   := TButton():New( 042, 007, "Histórico RPO",oPanel,{||MATCHAtu()}, 60,10,,,.F.,.T.,.F.,,.F.,,,.F. ) //"Hist. RPO"
	oBtExpor  := TButton():New( 069, 007, "Gerar Arquivo (.zip)",oPanel,{||MACTWizFon(Nil,3,aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam)}, 60,10,,,.F.,.T.,.F.,,.F.,,,.F. ) //"Gerar Arquivo"

	// Instancia objeto CSS
	oBtFontes:SetCss(cCSSTBut)
	oBtHist:SetCss(cCSSTBut)
	oBtExpor:SetCss(cCSSTBut)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel01_3
FPanel01_3 - Folder1 - Valida Procedures

@Param
oPanel - objeto da tela
aProc  - array de procedures

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel01_3(oPanel,aProc)

	Local oValorN := TFont():New( "Courier new",, -14,, .T. )
	Local nLin    := 005

	Local cSPMAT004 := GetSPName("MAT004","19")
	Local cSPMAT007 := GetSPName("MAT007","19")
	Local cSPMAT009 := GetSPName("MAT009","19")
	Local cSPMAT016 := GetSPName("MAT016","19")
	Local cSPMAT052 := GetSPName("MAT052","19")
	Local cSPMAT054 := GetSPName("MAT054","19")

	Local lExProc := ExistProc(cSPMAT004) .And. ;
	ExistProc(cSPMAT007) .And. ;
	ExistProc(cSPMAT009) .And. ;
	ExistProc(cSPMAT016) .And. ;
	ExistProc(cSPMAT052) .And. ;
	ExistProc(cSPMAT054)

	// Verifica a utilizacao de alguma procedures de custo em partes
	Local lExistProcCP := ExistProc("M330INB2CP") .Or. ;
	ExistProc("M330INC2CP") .Or. ;
	ExistProc("MA280INB9CP") .Or. ;
	ExistProc("MA280INC2CP") .Or. ;
	ExistProc("MA330CP")

	Default aProc := {}

	If lExProc
		TSay():New( nLin, 007, { || "Procedures de Recálculo de Custo Médio Instaladas!" }        , oPanel,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aProc, "Procedures de Recálculo de Custo Médio Instaladas!")
	Else
		TSay():New( nLin, 007, { || "Procedure de Recálculo de Custo Médio não Instaladas. Verifique!" }        , oPanel,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
		aAdd(aProc, "Procedure de Recálculo de Custo Médio não Instalada. Verifique!")
	EndIf
	nLin := nLin + 15

	If lExistProcCP
		TSay():New( nLin, 007, { || "Procedures de Custo em Partes Instaladas!" }        , oPanel,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aProc, "Procedures de Custo em Partes Instaladas!" )
	Else
		TSay():New( nLin, 007, { || "Procedure de Custo em Partes não Instaladas" }        , oPanel,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
		aAdd(aProc, "Procedure de Custo em Partes não Instalada!")
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel01_4
FPanel01_4 - Folder1 - Valida existencia de pontos de entrada

@Param
oPanel - objeto da tela
oScr1  - objeto do scroll da tela
aPE    - array de pontos de entrada

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel01_4(oPanel,oScr1,aPE)

	Local oValorN   :=	TFont():New( "Courier new",, -14,, .T. )
	Local nLin		:= 004

	// ************* SIGACUSA *************
	Local lMTAB2D1   := ExistBlock("MTAB2D1")
	Local lMTAATUD1  := ExistBlock("MTAATUD1")
	Local lMTAB2D1R  := ExistBlock("MTAB2D1R")
	Local lMTAB2D2   := ExistBlock("MTAB2D2")
	Local lMTAATUD2  := ExistBlock("MTAATUD2")
	Local lMTAB2D2R  := ExistBlock("MTAB2D2R")
	Local lMTAB2D3   := ExistBlock("MTAB2D3")
	Local lMTAATUD3  := ExistBlock("MTAATUD3")
	Local lMTAB2D3R  := ExistBlock("MTAB2D3R")
	Local lSigaCus	 := .T.

	// ************* MATA280 *************
	Local lA280OK    := ExistBlock("A280OK")
	Local lMA280CON  := ExistBlock("MA280CON")
	Local lA280SB9   := ExistBlock("A280SB9")
	Local lMA280FIM  := ExistBlock("MA280FIM")
	Local lMA280FI   := ExistBlock("MA280FI")
	Local lMA280BAT  := ExistBlock("MA280BAT")
	Local lA280SBJ   := ExistBlock("A280SBJ")
	Local lA280SBK   := ExistBlock("A280SBK")
	Local lMata280   := .T.

	// ************* MATA300 *************
	Local lM300SB8   := ExistBlock("M300SB8")
	Local lM300SBF   := ExistBlock("M300SBF")
	Local lMA300OK   := ExistBlock("MA300OK")
	Local lMata300   := .T.

	// ************* MATA330 *************
	Local lMA330CP   := ExistBlock("MA330CP")
	Local lM330TMP2  := ExistBlock("M330TMP2")
	Local lMA330FIL  := ExistBlock("MA330FIL")
	Local lMA330TRB  := ExistBlock("MA330TRB")
	Local lM330TMP1  := ExistBlock("M330TMP1")
	Local lMA330FIM  := ExistBlock("MA330FIM")
	Local lMA330MOD  := ExistBlock("MA330MOD")
	Local lA330GRUP  := ExistBlock("A330GRUP")
	Local lA330QTMO  := ExistBlock("A330QTMO")
	Local lM330FCC   := ExistBlock("M330FCC")
	Local lMA330SEQ  := ExistBlock("MA330SEQ")
	Local lMA330D1   := ExistBlock("MA330D1")
	Local lMA330D3   := ExistBlock("MA330D3")
	Local lMA330D2   := ExistBlock("MA330D2")
	Local lM330CD2   := ExistBlock("M330CD2")
	Local lM330CD1   := ExistBlock("M330CD1")
	Local lM330CD3   := ExistBlock("M330CD3")
	Local lMA330P3   := ExistBlock("MA330P3")
	Local lA330CDEV  := ExistBlock("A330CDEV")
	Local lMA330AL   := ExistBlock("MA330AL")
	Local lM330CMU   := ExistBlock("M330CMU")
	Local lMA330OK   := ExistBlock("MA330OK")
	Local lMA330PRC  := ExistBlock("MA330PRC")
	Local lMA330TRF  := ExistBlock("MA330TRF")
	Local lMA330PGI  := ExistBlock("MA330PGI")
	Local lMA330PGF  := ExistBlock("MA330PGF")
	Local lMA330IND  := ExistBlock("MA330IND")
	Local lMata330	 := .T.

	// ************* MATA331 *************
	Local lM331DATA  := ExistBlock("M331DATA")
	Local lMA331FIM  := ExistBlock("MA331FIM")
	Local lMA331OK   := ExistBlock("MA331OK")
	Local lMata331   := .T.

	Default aPE		 := {}

	// SIGACUSA
	TSay():New( nLin, 007, { || "Fonte: SIGACUS" }        , oScr1,,oValorN,,,, .T.,     ,, 300, 010 )
	nLin := nLin + 10

	If !(lMTAB2D1) .And. !(lMTAATUD1) .And. !(lMTAB2D1R) .And. !(lMTAB2D2) .And. !(lMTAATUD2) .And. !(lMTAB2D2R) .And. !(lMTAB2D3) .And. !(lMTAATUD3) .And. !(lMTAB2D3R)
		TSay():New( nLin, 007, { || "Nenhum ponto de entrada encontrada!" }        , oScr1,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aPE,{"SIGACUS","Nenhum ponto de entrada encontrada!"})
		nLin := nLin + 10
		lSigaCus := .F.
	EndIf
	If lSigaCus
		If lMTAB2D1
			TSay():New( nLin, 007, { || "Ponto de entrada MTAB2D1 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAB2D1 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAATUD1
			TSay():New( nLin, 007, { || "Ponto de entrada MTAATUD1 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAATUD1 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAB2D1R
			TSay():New( nLin, 007, { || "Ponto de entrada MTAB2D1R encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAB2D1R encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAB2D2
			TSay():New( nLin, 007, { || "Ponto de entrada MTAB2D2 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAB2D2 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAATUD2
			TSay():New( nLin, 007, { || "Ponto de entrada MTAATUD2 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAATUD2 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAB2D2R
			TSay():New( nLin, 007, { || "Ponto de entrada MTAB2D2R encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAB2D2R encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAB2D3
			TSay():New( nLin, 007, { || "Ponto de entrada MTAB2D3 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAB2D3 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAATUD3
			TSay():New( nLin, 007, { || "Ponto de entrada MTAATUD3 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAATUD3 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMTAB2D3R
			TSay():New( nLin, 007, { || "Ponto de entrada MTAB2D3R encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"SIGACUS","Ponto de entrada MTAB2D3R encontrado."})
			nLin := nLin + 10
		EndIf
	EndIf

	nLin := nLin + 5

	// MATA280
	TSay():New( nLin, 007, { || "Fonte: MATA280" }        , oScr1,,oValorN,,,, .T.,     ,, 300, 010 )
	nLin := nLin + 10

	If !(lA280OK) .And. !(lMA280CON) .And. !(lA280SB9) .And. !(lMA280FIM) .And. !(lMA280FI) .And. !(lMA280BAT) .And. !(lA280SBJ) .And. !(lA280SBK)
		TSay():New( nLin, 007, { || "Nenhum ponto de entrada encontrada!" }        , oScr1,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aPE,{"MATA280","Nenhum ponto de entrada encontrada!"})
		nLin := nLin + 10
		lMata280   := .F.
	EndIf
	If lMata280
		If lA280OK
			TSay():New( nLin, 007, { || "Ponto de entrada A280OK encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada A280OK encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA280CON
			TSay():New( nLin, 007, { || "Ponto de entrada MA280CON encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada MA280CON encontrado."})
			nLin := nLin + 10
		EndIf
		If lA280SB9
			TSay():New( nLin, 007, { || "Ponto de entrada A280SB9 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada A280SB9 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA280FIM
			TSay():New( nLin, 007, { || "Ponto de entrada MA280FIM encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada MA280FIM encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA280FI
			TSay():New( nLin, 007, { || "Ponto de entrada MA280FI encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada MA280FI encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA280BAT
			TSay():New( nLin, 007, { || "Ponto de entrada MA280BAT encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada MA280BAT encontrado."})
			nLin := nLin + 10
		EndIf
		If lA280SBJ
			TSay():New( nLin, 007, { || "Ponto de entrada A280SBJ encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada A280SBJ encontrado."})
			nLin := nLin + 10
		EndIf
		If lA280SBK
			TSay():New( nLin, 007, { || "Ponto de entrada A280SBK encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA280","Ponto de entrada A280SBK encontrado."})
			nLin := nLin + 10
		EndIf
	EndIf

	nLin := nLin + 5

	// MATA300
	TSay():New( nLin, 007, { || "Fonte: MATA300" }        , oScr1,,oValorN,,,, .T.,     ,, 300, 010 )
	nLin := nLin + 10

	If !(lM300SB8) .And. !(lM300SBF) .And. !(lMA300OK)
		TSay():New( nLin, 007, { || "Nenhum ponto de entrada encontrada!" }        , oScr1,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aPE,{"MATA300","Nenhum ponto de entrada encontrada!"})
		nLin := nLin + 10
		lMata300   := .F.
	EndIf
	If lMata300
		If lM300SB8
			TSay():New( nLin, 007, { || "Ponto de entrada M300SB8 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA300","Ponto de entrada M300SB8 encontrado."})
			nLin := nLin + 10
		EndIf
		If lM300SBF
			TSay():New( nLin, 007, { || "Ponto de entrada M300SBF encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA300","Ponto de entrada M300SBF encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA300OK
			TSay():New( nLin, 007, { || "Ponto de entrada MA300OK encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA300","Ponto de entrada MA300OK encontrado."})
			nLin := nLin + 10
		EndIf
	EndIf

	nLin := nLin + 5

	// MATA330
	TSay():New( nLin, 007, { || "Fonte: MATA330" }        , oScr1,,oValorN,,,, .T.,     ,, 300, 010 )
	nLin := nLin + 10

	If !(lMA330CP) .And. !(lM330TMP2) .And. !(lMA330FIL) .And. !(lMA330TRB) .And. !(lM330TMP1) .And. !(lMA330FIM) .And. !(lMA330MOD) .And. !(lA330GRUP) .And.;
	!(lA330QTMO) .And. !(lM330FCC) .And. !(lMA330SEQ) .And. !(lMA330D1) .And. !(lMA330D3) .And. !(lMA330D2) .And. !(lM330CD2) .And. !(lM330CD1) .And.;
	!(lM330CD3) .And. !(lMA330P3) .And. !(lA330CDEV) .And. !(lMA330AL) .And. !(lMA330OK) .And. !(lMA330PRC) .And. !(lMA330TRF) .And. !(lMA330PGI) .And.;
	!(lMA330PGF) .And. !(lMA330IND)
		TSay():New( nLin, 007, { || "Nenhum ponto de entrada encontrada!" }        , oScr1,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aPE,{"MATA330","Nenhum ponto de entrada encontrada!"})
		nLin := nLin + 10
		lMata330	 := .F.
	EndIf
	If lMata330
		If lMA330CP
			TSay():New( nLin, 007, { || "Ponto de entrada MA330CP encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330CP encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330TMP2
			TSay():New( nLin, 007, { || "Ponto de entrada M330TMP2 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330TMP2 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330FIL
			TSay():New( nLin, 007, { || "Ponto de entrada MA330FIL encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330FIL encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330TRB
			TSay():New( nLin, 007, { || "Ponto de entrada MA330TRB encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330TRB encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330TMP1
			TSay():New( nLin, 007, { || "Ponto de entrada M330TMP1 encontrado."}        , oScr1,,oValorN,,,, .T.,CLR_BLUE,, 500, 010 )
			nLin := nLin + 10
			TSay():New( nLin, 007, { || "Compilado pelo Monitor de Custos!" }           , oScr1,,oValorN,,,, .T.,CLR_BLUE,, 500, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330TMP1 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330FIM
			TSay():New( nLin, 007, { || "Ponto de entrada MA330FIM encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330FIM encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330MOD
			TSay():New( nLin, 007, { || "Ponto de entrada MA330MOD encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330MOD encontrado."})
			nLin := nLin + 10
		EndIf
		If lA330GRUP
			TSay():New( nLin, 007, { || "Ponto de entrada A330GRUP encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada A330GRUP encontrado."})
			nLin := nLin + 10
		EndIf
		If lA330QTMO
			TSay():New( nLin, 007, { || "Ponto de entrada A330QTMO encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada A330QTMO encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330FCC
			TSay():New( nLin, 007, { || "Ponto de entrada M330FCC encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330FCC encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330SEQ
			TSay():New( nLin, 007, { || "Ponto de entrada MA330SEQ encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330SEQ encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330D1
			TSay():New( nLin, 007, { || "Ponto de entrada MA330D1 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330D1 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330D3
			TSay():New( nLin, 007, { || "Ponto de entrada MA330D3 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330D3 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330D2
			TSay():New( nLin, 007, { || "Ponto de entrada MA330D2 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330D2 encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330CD2
			TSay():New( nLin, 007, { || "Ponto de entrada M330CD2 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330CD2 encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330CD1
			TSay():New( nLin, 007, { || "Ponto de entrada M330CD1 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330CD1 encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330CD3
			TSay():New( nLin, 007, { || "Ponto de entrada M330CD3 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330CD3 encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330P3
			TSay():New( nLin, 007, { || "Ponto de entrada MA330P3 encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330P3 encontrado."})
			nLin := nLin + 10
		EndIf
		If lA330CDEV
			TSay():New( nLin, 007, { || "Ponto de entrada A330CDEV encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada A330CDEV encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330AL
			TSay():New( nLin, 007, { || "Ponto de entrada MA330AL encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330AL encontrado."})
			nLin := nLin + 10
		EndIf
		If lM330CMU
			TSay():New( nLin, 007, { || "Ponto de entrada M330CMU encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada M330CMU encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330OK
			TSay():New( nLin, 007, { || "Ponto de entrada MA330OK encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330OK encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330PRC
			TSay():New( nLin, 007, { || "Ponto de entrada MA330PRC encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330PRC encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330TRF
			TSay():New( nLin, 007, { || "Ponto de entrada MA330TRF encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330TRF encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330PGI
			TSay():New( nLin, 007, { || "Ponto de entrada MA330PGI encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330PGI encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330PGF
			TSay():New( nLin, 007, { || "Ponto de entrada MA330PGF encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330PGF encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA330IND
			TSay():New( nLin, 007, { || "Ponto de entrada MA330IND encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA330","Ponto de entrada MA330IND encontrado."})
			nLin := nLin + 10
		EndIf
	EndIf

	nLin := nLin + 5

	// MATA331
	TSay():New( nLin, 007, { || "Fonte: MATA331" }        , oScr1,,oValorN,,,, .T.,     ,, 300, 010 )
	nLin := nLin + 10

	If !(lM331DATA) .And. !(lMA331FIM) .And. !(lMA331OK)
		TSay():New( nLin, 007, { || "Nenhum ponto de entrada encontrada!" }        , oScr1,,oValorN,,,, .T.,CLR_GREEN,, 300, 010 )
		aAdd(aPE,{"MATA331","Nenhum ponto de entrada encontrada!"})
		nLin := nLin + 10
		lMata331   := .F.
	EndIf
	If lMata331
		If lM331DATA
			TSay():New( nLin, 007, { || "Ponto de entrada M331DATA encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA331","Ponto de entrada M331DATA encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA331FIM
			TSay():New( nLin, 007, { || "Ponto de entrada MA331FIM encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA331","Ponto de entrada MA331FIM encontrado."})
			nLin := nLin + 10
		EndIf
		If lMA331OK
			TSay():New( nLin, 007, { || "Ponto de entrada MA331OK encontrado." }        , oScr1,,oValorN,,,, .T.,CLR_HRED,, 300, 010 )
			aAdd(aPE,{"MATA331","Ponto de entrada MA331OK encontrado."})
			nLin := nLin + 10
		EndIf
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel01_5
FPanel01_5 - Folder5 - Links Importantes

@Param
oPanel - objeto da tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel01_5(oPanel)

	Local oBtlinks  := Nil
	Local oBtExpor	:= Nil

	// Links do TDN
	@005,007 SAY oBtlinks PROMPT "<u>"+ 'Dicas de Uso - Monitor de Custos' +"</u>" SIZE 300,010 OF oPanel HTML PIXEL
	oBtlinks:bLClicked := {|| ShellExecute("open","http://tdn.totvs.com.br/display/PROT/Monitor+de+Custos","","",1) }

	@017,007 SAY oBtlinks PROMPT "<u>"+ 'Página centralizadora de Custo Médio, FIFO e Recálculo' +"</u>" SIZE 300,010 OF oPanel HTML PIXEL
	oBtlinks:bLClicked := {|| ShellExecute("open","http://tdn.totvs.com/pages/viewpage.action?pageId=340361781","","",1) }

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel02
Painel02 - Folder2 - Valida decimais dos campos

@Param
oPanel    - objeto da tela
oScr2     - objeto de scroll da tela
aVlCampos - array com campos para validaçao de decimais

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel02(oPanel,oScr2,aVlCampos)

	Local oValor  := TFont():New( "Courier new",, -14,, .F. )
	Local oValorN := TFont():New( "Courier new",, -14,, .T. )
	Local aArea		:= GetArea()
	Local nTamDec	:= 0
	Local nTamDecOld:= 0
	Local aCampos	:={{"SB9","B9_VINI"},{"SB2","B2_VFIM"},{"SD1","D1_CUSTO"},{"SD2","D2_CUSTO"},{"SD3","D3_CUSTO"},{"SC2","C2_VINI"},{"SC2","C2_VFIM"}}
	Local nz,nx,cCampo
	Local cMoeda330C:= SuperGetMv('MV_MOEDACM',.F.,"2345")
	Local cNomeTab	:= ""
	Local cNCampo	:= ""
	Local nLin      := 005

	Default aVlCampos := {}
	Private aLogDec   := {}

	// Posiciono no SX2 para buscar o nome da tabela
	dbSelectArea("SX2")
	SX2->(dbSetOrder(1))

	// Posiciono no SX3 para buscar o nome do campo
	dbSelectArea("SX3")
	SX3->(dbSetOrder(2))

	For nz:=1 to 5
		// Verifica se moeda devera ser considerada
		If nz # 1 .And. !(Str(nz,1,0) $ cMoeda330C)
			Loop
		EndIf
		For nx:=1 to Len(aCampos)
			cCampo:=aCampos[nx,2]+Strzero(nz,1,0)
			// Campo que foge a regra -> D1_CUSTO
			If cCampo == "D1_CUSTO1"
				cCampo:="D1_CUSTO"
			EndIf
			nTamDec:=TamSX3(cCampo)[2]
			// Busco o nome da tabela
			If SX2->(DBSeek(aCampos[nx,1]))
				cNomeTab := SX2->X2_NOME
			EndIf
			// Busco o nome do camp
			If SX3->(DBSeek(cCampo))
				cNCampo := SX3->X3_TITULO
			EndIf
			AADD(aLogDec,{aCampos[nx,1],cNomeTab,cCampo,cNCampo,cValtoChar(nTamDec)})
			If nx > 1 .And. nTamDec # nTamDecOld
				lLogDec:=.T.
			EndIf
			nTamDecOld:=nTamDec
		Next nx
	Next nz
	RestArea(aArea)

	If lLogDec
		//Informa que possui campos com decimais divergentes
		TSay():New( nLin, 007, { || "ATENÇÃO: Existem campos com decimais divergentes, poderao ocorrer diferenças de arredondamento!" }        , oScr2,,oValorN,,,, .T.,CLR_HRED,, 500, 010 )
		nLin := nLin + 10
		@ nLin,007 SAY oBtErro PROMPT "<u>" + "Clique aqui e selecione a aba 'Precisão de Calculo' para mais informações sobre casas decimais." + "</u>" SIZE 400,009 OF oScr2 HTML PIXEL
		oBtErro:bLClicked := {|| ShellExecute("open","http://tdn.totvs.com/x/NYJJF","","",1) }

		// Mostra na tela os campos caso haja diferenças decimais
		MATCMBrow(oPanel,aLogDec,oScr2,.T.,.F.,002,190)
	Else
		TSay():New( nLin, 007, { || "Campos validados corretamente." }        , oScr2,,oValorN,,,, .T.,CLR_GREEN,, 550, 010 )
	EndIf

	// Copia os dados para o array aVlCampos para exportação de dados
	aVlCampos := aClone(aLogDec)
	aAdd(aVlCampos, lLogDec)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCMBrow
Função Genérica para mostrar dados no browser

@Param
oPanel  - objeto da tela
aCols   - array com dados do browser
oScr    - objeto scroll da tela posicionada
lValid  - indica se exibe o browser para validação de campos
lParam  - indica se exibe o browser para validação de parametros
nRow    - posição incial para criação do browser
nHeight - altura do objeto do browser
lCusto  - indica se exibe o browser para validação do custo

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCMBrow(oPanel,aCols,oScr,lValid,lParam,nRow,nHeight,lCusto)

	Local Nx 		:= 0
	Local blDblClick:= ""
	Local bLine		:= ""
	Local aColSizes	:= {}
	Local aHeader   := {}
	Local oListBox  := Nil

	Default oPanel  := Nil
	Default aCols	:= {}
	Default oScr	:= Nil
	Default lValid	:= .F.
	Default lParam	:= .F.
	Default nRow	:= 001
	Default nHeight	:= 190

	If Empty(bLine)
		nTamCol := Len(aCols[01])
		bLine 	:= "{|| {"
		For Nx := 1 To nTamCol
			bLine += "aCols[oListBox:nAt]["+StrZero(Nx,3)+"]"
			If Nx < nTamCol
				bLine += ","
			EndIf
		Next
		bLine += "} }"
	EndIf

	// Define qual cabeçalho e monta browser para exibir os dados
	If lValid
		aColSizes := {40,130,50,80,20}
		aHeader   := {'Tabela','Descrição','Campo','Descrição','Decimais'}
		oListBox := TCBrowse():New(nRow,001,488,nHeight,,aHeader,,oScr,'Campos',,,,,,,,,,,,,,,,,,.F.)
	ElseIf lParam
		aColSizes := {60,100}
		aHeader   := {'Parâmetro','Conteúdo'}
		oListBox := TCBrowse():New(nRow,001,300,nHeight,,aHeader,,oScr,'Campos',,,,,,,,,,,,,,,,,,.F.)
	EndIf

	// Define os dados do browser
	oListBox:SetArray( aCols )
	oListBox:bLine := &bLine
	If !Empty( aColSizes )
		oListBox:aColSizes := aColSizes
	EndIf
	oListBox:SetFocus()

	// Abre link do TDN
	If lParam
		oListBox:bLDblClick := { || MsgRun( "Abrindo página WEB... Aguarde", "Monitor de Custos", {|| ShellExecute("open","http://tdn.totvs.com/x/NYJJF","","",1) } ) }
	EndIf

	// Instancia objeto CSS
	oListBox:SetCss(cCSSTCBrw)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel03
Painel03 - Folder3 - Valida parâmetros do sistema

@Param
oPanel - objeto da tela
oScr4  - objeto scroll da tela
aParam - array com parametros

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel03(oPanel,oScr4,aParam)

	Local oValor    := TFont():New( "Courier new",, -14,, .F. )
	Local oValorN   := TFont():New( "Courier new",, -14,, .T. )
	Local nLin	    := 005
	Local aPar1     := {}
	Local aPar2     := {}
	Local aPar3     := {}
	Local aPar4     := {}
	Local Nx 		:= 0
	Local oListBox	:= Nil
	Local lAllClient:= .F.
	Local blDblClick:= ""
	Local bLine		:= ""
	Local oBtlinks  := Nil

	Default aParam  := {}

	// Adiciono o conteúdo dos parametros
	aAdd(aPar1,{"MV_CUSFIL",cValtoChar(SuperGetMv("MV_CUSFIL",.F.,"A"))})
	aAdd(aPar1,{"MV_CUSZERO",cValtoChar(SuperGetMv("MV_CUSZERO",.F.,"N"))})
	aAdd(aPar1,{"MV_CUSREP",cValtoChar(SuperGetMv("MV_CUSREP",.F.,.F.))})
	aAdd(aPar1,{"MV_SEQ300",cValtoChar(SuperGetMv("MV_SEQ300",.F.,.F.))})
	aAdd(aPar1,{"MV_SEQ500",cValtoChar(SuperGetMv("MV_SEQ500",.F.,.T.))})
	aAdd(aPar1,{"MV_CUSLIFO",cValtoChar(SuperGetMv("MV_CUSLIFO",.F.,.F.))})
	aAdd(aPar1,{"MV_CUSFIFO",cValtoChar(SuperGetMv("MV_CUSFIFO",.F.,.F.))})
	aAdd(aPar1,{"MV_PROCQE6",cValtoChar(SuperGetMv("MV_PROCQE6",.F.,.F.))})
	aAdd(aPar1,{"MV_M330CON",cValtoChar(SuperGetMv("MV_M330CON",.F.,.F.))})
	aAdd(aPar1,{"MV_GERIMPV",cValtoChar(SuperGetMv("MV_GERIMPV",.F.,"N"))})
	aAdd(aPar1,{"MV_NGMNTES",cValtoChar(SuperGetMv("MV_NGMNTES",.F.,"N"))})
	aAdd(aPar1,{"MV_M330TCF",cValtoChar(SuperGetMv("MV_M330TCF",.F.,"RE5/DE5/RE6/DE6"))})
	aAdd(aPar1,{"MV_AGCUSTO",cValtoChar(SuperGetMv("MV_AGCUSTO",.F.,.F.))})
	aAdd(aPar1,{"MV_M330TRF",cValtoChar(SuperGetMv("MV_M330TRF",.F.,.F.))})
	aAdd(aPar1,{"MV_330ATCM",cValtoChar(SuperGetMv("MV_330ATCM",.F.,.F.))})
	aAdd(aPar1,{"MV_PRODPR0",cValtoChar(SuperGetMv("MV_PRODPR0",.F.,1))})
	aAdd(aPar1,{"MV_CUSTDEV",cValtoChar(SuperGetMv("MV_CUSTDEV",.F.,.F.))})
	aAdd(aPar1,{"MV_DOCSEQ",cValtoChar(SuperGetMv("MV_DOCSEQ",.F.,""))})

	aAdd(aPar2,{"MV_ULMES",cValtoChar(SuperGetMv("MV_ULMES",.F.,STOD("")))})
	aAdd(aPar2,{"MV_CUSTEXC",cValtoChar(SuperGetMv("MV_CUSTEXC",.F.,"N"))})
	aAdd(aPar2,{"MV_DBLQMOV",cValtoChar(SuperGetMv("MV_DBLQMOV",.F.,STOD("")))})
	aAdd(aPar2,{"MV_AJUSNFC",cValtoChar(SuperGetMv("MV_AJUSNFC",.F.,.F.))})
	aAdd(aPar2,{"MV_NGMNTPC",cValtoChar(SuperGetMv("MV_NGMNTPC",.F.,"N"))})
	aAdd(aPar2,{"MV_NEGESTR",cValtoChar(SuperGetMv("MV_NEGESTR",.F.,.F.))})
	aAdd(aPar2,{"MV_CUSMED",cValtoChar(SuperGetMv("MV_CUSMED",.F.,"M"))})

	aAdd(aPar3,{"MV_MUDATRT",cValtoChar(SuperGetMv("MV_MUDATRT",.F.,.F.))})
	aAdd(aPar3,{"MV_A330GRV",cValtoChar(SuperGetMv("MV_A330GRV",.F.,.T.))})
	aAdd(aPar3,{"MV_A330190",cValtoChar(SuperGetMv("MV_A330190",.F.,"S"))})
	aAdd(aPar3,{"MV_PROCCV3",cValtoChar(SuperGetMv("MV_PROCCV3",.F.,.T.))})
	aAdd(aPar3,{"MV_A330DRV",cValtoChar(SuperGetMv("MV_A330DRV",.F.,"DBFCDX"))})
	aAdd(aPar3,{"MV_THRSEQ",cValtoChar(SuperGetMv("MV_THRSEQ",.F.,.F.))})
	aAdd(aPar3,{"MV_M330JCM",cValtoChar(SuperGetMv("MV_M330JCM",.F.,""))})
	aAdd(aPar3,{"MV_I330FSM",cValtoChar(SuperGetMv("MV_I330FSM",.F.,.F.))})
	aAdd(aPar3,{"MV_MOEDACM",cValtoChar(SuperGetMv("MV_MOEDACM",.F.,"2345"))})

	aAdd(aPar4,{"MV_CONTERC",cValtoChar(SuperGetMv("MV_CONTERC",.F.,.F.))})
	aAdd(aPar4,{"MV_LOCALIZA",cValtoChar(SuperGetMv("MV_LOCALIZA",.F.,"N"))})
	aAdd(aPar4,{"MV_RASTRO",cValtoChar(SuperGetMv("MV_RASTRO",.F.,"N"))})
	aAdd(aPar4,{"MV_LOCPROC",cValtoChar(SuperGetMv("MV_LOCPROC",.F.,"99"))})
	aAdd(aPar4,{"MV_NIVALT",cValtoChar(SuperGetMv("MV_NIVALT",.F.,"N"))})
	aAdd(aPar4,{"MV_CQ",cValtoChar(SuperGetMv("MV_CQ",.F.,"98"))})
	aAdd(aPar4,{"MV_PCOINTE",cValtoChar(SuperGetMv("MV_PCOINTE",.F.,"2"))})
	aAdd(aPar4,{"MV_PRODMNT",cValtoChar(SuperGetMv("MV_PRODMNT",.F.,"MANUTENCAO"))})
	aAdd(aPar4,{"MV_DEPTRAN",cValtoChar(SuperGetMv("MV_DEPTRAN",.F.,"95"))})
	aAdd(aPar4,{"MV_NGMNTCM",cValtoChar(SuperGetMv("MV_NGMNTCM",.F.,"N"))})

	// Clona dados para exportação do arquivo texto
	aAdd(aParam,aPar1)
	aAdd(aParam,aPar2)
	aAdd(aParam,aPar3)
	aAdd(aParam,aPar4)

	TSay():New( nLin, 007, { || "Parâmetros utilizados no cálculo do custo do produto - (Duplo clique na linha para mais informações)" }        , oScr4,,oValorN,,,, .T.,    ,, 600, 010 )
	// Mostra os parametros (aPar1) na tela
	MATCMBrow(oPanel,aPar1,oScr4,.F.,.T.,002,175)
	nLin := nLin + 215

	TSay():New( nLin, 007, { || "Parâmetros utilizados no processo de fechamento de estoque - (Duplo clique na linha para mais informações)" }        , oScr4,,oValorN,,,, .T.,    ,, 600, 010 )
	// Mostra os parametros (aPar2) na tela
	MATCMBrow(oPanel,aPar2,oScr4,.F.,.T.,017,82)
	nLin := nLin + 120

	TSay():New( nLin, 007, { || "Parâmetros para ganho de performance - (Duplo clique na linha para mais informações)" }        , oScr4,,oValorN,,,, .T.,    ,, 600, 010 )
	// Mostra os parametros (aPar3) na tela
	MATCMBrow(oPanel,aPar3,oScr4,.F.,.T.,026,99)
	nLin := nLin + 140

	TSay():New( nLin, 007, { || "Outros parâmetros importantes - (Duplo clique na linha para mais informações)" }        , oScr4,,oValorN,,,, .T.,    ,, 600, 010 )
	// Mostra os parametros (aPar4) na tela
	MATCMBrow(oPanel,aPar4,oScr4,.F.,.T.,036,107)
	nLin := nLin + 150

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel05
Painel05 - Folder5 - KARDEX e Recalculo

@Param
oPanel    - objeto da tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel05(oPanel)

	Local oBtRecalc := Nil
	Local oValor	:=	TFont():New( "Courier new",, -14,, .T. )
	Local oValorN	:=	TFont():New( "Courier new",, -14,, .F. )

	Default oPanel := Nil

	TSay():New( 012, 015, { || "Rodar Recálculo" }, oPanel,,oValor,,,, .T.,CLR_BLACK,, 500, 010 ) //"Rodar Recálculo"
	TSay():New( 012, 077, { || " - O arquivo temporario será salvo e exportado na opção 'Gerar Arquivo (.zip)'." }, oPanel,,oValorN,,,, .T.,CLR_BLACK,, 500, 010 )

	oBtRecalc := TButton():New( 022, 015, "Recálculo do Custo Médio",oPanel,{||MATA330()}, 80,10,,,.F.,.T.,.F.,,.F.,,,.F. ) //"Recálculo do Custo Médio"

	// Instancia objeto CSS
	oBtRecalc:SetCss(cCSSTBut)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel05_1
FPanel05_1 - Folder5 - Recalculo

@Param
oPanel    - objeto da tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel05_1(oPanel)

	//Local aArrayCV8 := {{'','','','','','',''}}
	Local cAliasCV8 := ''
	Local cQuery    := ''
	Local cWRotina  := 'MATA330'
	Local oValor	:=	TFont():New( "Courier new",, -14,, .F. )

	Default oPanel := Nil

	// Cria a TwBrowse do Log de Execucoes (Cabecalho)
	TSay():New( 005, 008, { || "Selecione o ID" }, oPanel,,oValor,,,, .T.,CLR_BLACK,, 500, 010 ) //"Selecione o ID"
	aHeaderCab   := {'ID Processo','Rotina','Usuário','Dt. Inicio','Hr. Inicio','Dt. Fim','Hr. Fim'}
	oLBCabec := TCBrowse():New(001,001,640,060,,aHeaderCab,,oPanel,'Campos',,,,,,,,,,,,,,,,,.T.,.T.)
	oLBCabec:SetArray(aArrayCV8)
	oLBCabec:bChange := { || MACTLDetail(oPanel,aArrayCV8[oLbCabec:nAt,1],oLBDetalhe) }
	oLBCabec:bLine := { || { aArrayCV8[oLBCabec:nAT][1], aArrayCV8[oLBCabec:nAT][2], aArrayCV8[oLBCabec:nAT][3], aArrayCV8[oLBCabec:nAT][4], aArrayCV8[oLBCabec:nAT][5], aArrayCV8[oLBCabec:nAT][6], aArrayCV8[oLBCabec:nAT][7]} }
	oLBCabec:Refresh()

	// Cria a TwBrowse do Log de Execucoes (Detalhes)
	TSay():New( 076, 008, { || "Detalhes do Processamento" }, oPanel,,oValor,,,, .T.,CLR_BLACK,, 500, 010 ) //"Detalhes do Processamento"
	aHeaderDet   := {'ID Processo','Filial','Dt. Inicio','Hr. Inicio','Mensagem','Sub. Processo','Info'}
	oLBDetalhe := TCBrowse():New(006,001,640,104,,aHeaderDet,,oPanel,'Campos',,,,,,,,,,,,,,,,,.T.,.T.)
	oLBDetalhe:SetArray(aArrayDET)
	oLBDetalhe:bLDblClick := { || MATCLMemo(oPanel,oLBDetalhe:nAt) }
	oLBDetalhe:bLine := { || { aArrayDET[oLBDetalhe:nAT][1],aArrayDET[oLBDetalhe:nAT][2],aArrayDET[oLBDetalhe:nAT][3],aArrayDET[oLBDetalhe:nAT][4],aArrayDET[oLBDetalhe:nAT][5],aArrayDET[oLBDetalhe:nAT][6],aArrayDET[oLBDetalhe:nAT][7] } }
	oLBDetalhe:Refresh()

	// Instancia objeto CSS
	oLBCabec:SetCss(cCSSTCBrw)
	oLBDetalhe:SetCss(cCSSTCBrw)

	cAliasCV8 := GetNextAlias()
	cQuery := "SELECT"

	cQuery += " CV8_FILIAL,"
	cQuery += " CV8_IDMOV,"
	cQuery += " CV8_PROC,"
	cQuery += " CV8_USER,"
	cQuery += " ( SELECT IsNull(Min(CV8_DATA),'') "
	cQuery +=     " FROM "+RetSqlName("CV8")+" "
	cQuery += " WHERE  CV8_IDMOV = CV8.CV8_IDMOV "
	cQuery +=    " AND CV8_PROC = '" + Padr(cWRotina,TamSX3("CV8_PROC")[1]) + "' "
	cQuery +=     " AND CV8_INFO = '1' "
	cQuery +=     " AND CV8_SBPROC = '' "
	cQuery +=     " AND D_E_L_E_T_ = ' ' ) DATAINI, "

	cQuery += " ( SELECT IsNull(Min(CV8_HORA),'') "
	cQuery +=     " FROM "+RetSqlName("CV8")+" "
	cQuery += " WHERE  CV8_IDMOV = CV8.CV8_IDMOV "
	cQuery +=    " AND CV8_PROC = '" + Padr(cWRotina,TamSX3("CV8_PROC")[1]) + "' "
	cQuery +=     " AND CV8_INFO = '1' "
	cQuery +=     " AND CV8_SBPROC = '' "
	cQuery +=     " AND D_E_L_E_T_ = ' ' ) HORAINI, "

	cQuery += " ( SELECT IsNull(MAX(CV8_DATA),'') "
	cQuery +=     " FROM "+RetSqlName("CV8")+" "
	cQuery += " WHERE  CV8_IDMOV = CV8.CV8_IDMOV "
	cQuery +=    " AND CV8_PROC = '" + Padr(cWRotina,TamSX3("CV8_PROC")[1]) + "' "
	cQuery +=     " AND CV8_INFO = '2' "
	cQuery +=     " AND CV8_SBPROC = '' "
	cQuery +=     " AND D_E_L_E_T_ = ' ' ) DATAFIM, "

	cQuery += " ( SELECT IsNull(MAX(CV8_HORA),'') "
	cQuery +=     " FROM "+RetSqlName("CV8")+" "
	cQuery += " WHERE  CV8_IDMOV = CV8.CV8_IDMOV "
	cQuery +=    " AND CV8_PROC = '" + Padr(cWRotina,TamSX3("CV8_PROC")[1]) + "' "
	cQuery +=     " AND CV8_INFO = '2' "
	cQuery +=     " AND CV8_SBPROC = '' "
	cQuery +=     " AND D_E_L_E_T_ = ' ' ) HORAFIM

	cQuery +=  " FROM "+RetSqlName("CV8")+" CV8 "
	cQuery += " WHERE CV8_PROC = '" + Padr(cWRotina,TamSX3("CV8_PROC")[1]) + "' "
	cQuery +=   " AND D_E_L_E_T_ = ' ' "

	cQuery +=   " GROUP BY  CV8_FILIAL,CV8_IDMOV,CV8_PROC,CV8_USER"
	cQuery +=   " ORDER BY  CV8_FILIAL,CV8_IDMOV,CV8_PROC,CV8_USER"
	cQuery := ChangeQuery(cQuery, .F.)

	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasCV8,.T.,.T.)

	TcSetField(cAliasCV8,"DATAINI","D",TamSX3("CV8_DATA")[1],TamSX3("CV8_DATA")[2])
	TcSetField(cAliasCV8,"DATAFIM","D",TamSX3("CV8_DATA")[1],TamSX3("CV8_DATA")[2])
	// Carrega as informacoes para o Log de processamento
	Do While !Eof()
		aAdd(aArrayCV8,{(cAliasCV8)->CV8_IDMOV,(cAliasCV8)->CV8_PROC, (cAliasCV8)->CV8_USER, (cAliasCV8)->DATAINI, (cAliasCV8)->HORAINI,(cAliasCV8)->DATAFIM, (cAliasCV8)->HORAFIM })
		(cAliasCV8)->(dbSkip())
	EndDo

	// Remove a linha em branco
	If Len(aArrayCV8) >= 2
		If Empty(aArrayCV8[1,1])
			Adel(aArrayCV8,1)
			ASize(aArrayCV8,Len(aArrayCV8)-1)
		EndIf
	EndIf

	// Atualiza a TWBrowse com os dados da rotina selecionada
	oLBCabec:SetArray(aArrayCV8)
	oLBCabec:bLine := { || {aArrayCV8[oLBCabec:nAT][1],aArrayCV8[oLBCabec:nAT][2],aArrayCV8[oLBCabec:nAT][3],aArrayCV8[oLBCabec:nAT][4],aArrayCV8[oLBCabec:nAT][5],aArrayCV8[oLBCabec:nAT][6],aArrayCV8[oLBCabec:nAT][7] } }
	oLBCabec:Refresh()

	//Atualiza a tabela de Logs pela primeira vez
	MACTLDetail(oPanel,aArrayCV8[oLbCabec:nAt,1],oLBDetalhe)

	If Select(cAliasCV8) > 0
		(cAliasCV8)->(dbCloseArea())
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCFontes
Carrega array com todos os fontes do repositório

@aParam lGeraTXT - Se chamada através da rotina de exportação gera txt sem montar tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018
@protected

@Return aFontes
[1] - Nome do Fonte
[2] - Data de compilação
[3] - Hora de compilação
*/
//------------------------------------------------------------------------------------------
Static Function MATCFontes(lGeraTXT)

	Local aGetFontes := {}
	Local aRetFontes := {}
	Local aFontes	 := {}
	Local nx
	Local cBarra    := If(issrvunix(), "/", "\")
	Local cPath	    := GetSrvProfString("StartPath", "") + If( Right( GetSrvProfString("StartPath",""), 1 ) == cBarra, "", cBarra )
	Local cFontes	:= "Fontes"

	Default lGeraTXT := .F.

	// Carrega todos os fontes do rpo
	aGetFontes := GetSrcArray("*")

	For nx:= 1 to len(aGetFontes)
		aRetFontes := GetApoInfo(aGetFontes[nx])
		aAdd(aFontes,{aRetFontes[1],aRetFontes[4],aRetFontes[5]})
	Next nx

	If !lGeraTXT
		MATCLisInf( { "Fonte", "Data", "Hora" }, aFontes, { 60, 50, 20 }, "Analise de Ambiente",.F. )
	Else
		FWMsgRun(,{|| CursorWait(),MACTExpFon(Nil,cFontes,cPath,aFontes,.T.)},,"Aguarde! Gerando arquivo..." )
	EndIf

Return (aFontes)


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCLisInf
Lista todos os fontes compilados no RPO

@Param
lSrcFonte - Pesquisar Fonte ou Tabela
aHeader   - aHeader do Grid
aCols     - aCols do Grid
aColSizes - Tamanho do Grid
cTitulo   - Titulo da Tela
lApliPath - Se .T. identifica que a tela foi chamada para exibição do histórico de atualizações

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018
@protected

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static function MATCLisInf( aHeader, aCols, aColSizes, cTitulo, lApliPath )

	Local Nx 		:= 0
	Local oListBox	:= nil
	Local oArea		:= Nil
	Local oValorN	:= TFont():New("Arial",08,14,,.T.,,,,.T.)
	Local oList		:= Nil
	Local aCoord	:= {}
	Local aWindow	:= {}
	Local aColumn	:= {}
	Local oButt1, oButt2, oButt3 := Nil
	Local cFonte	:= Space(20)

	Default aColSizes	:= {}
	Default cTitulo		:= "Analise de Ambiente"
	Default lApliPath	:= .F.

	//Ajusta ao objeto pai da tela
	aCoord 	:= {000,000,400,800}
	oList	:= oFather
	aWindow := {020,073}

	oArea := FWLayer():New()
	oFather := tDialog():New(aCoord[1],aCoord[2],aCoord[3],aCoord[4],cTitulo,,,,,CLR_BLACK,CLR_WHITE,,,.T.)
	oArea:Init(oFather,.F., .F. )

	oArea:AddLine("L01",100,.T.)

	oArea:AddCollumn("L01C01",99,.F.,"L01") //dados
	oArea:AddWindow("L01C01","TEXT","Funções",aWindow[01],.F.,.F.,/*bAction*/,"L01",/*bGotFocus*/)
	oText	:= oArea:GetWinPanel("L01C01","TEXT","L01")

	If lApliPath // Mostra tela com dados do histórico de atualizações
		TSay():New( 001, 002, { || "Histórico de Atualizações"}, oText,,oValorN,,,, .T.,CLR_BLACK,, 350, 020 )
		TSay():New( 010, 002, { || "Total de patchs aplicados: " + cValtoChar(Len(aCols)) }, oText,,oValorN,,,, .T.,CLR_GREEN,, 350, 020 )
	Else // Mostra tela com diagnóstico dos fontes
		TSay():New(005,002,{||'Pesquisar Fonte:'},oText,,,,,,.T.,,,200,20)
		TGet():New(003,050,{|u| if( PCount() > 0, cFonte := u, cFonte ) },oText,150,009,"@!",,,,,,,.T.,,,,.T.,,,.F.,,"","cFonte",,,,.T.,.T.)
		oButt3 := tButton():New(003,180,"Pesquisar",oText,{||MACTFilFon(oListBox,cFonte)}, 45,11,,,.F.,.T.,.F.,,.F.,,,.F. )
	EndIf

	oArea:AddWindow("L01C01","LIST","Fontes",aWindow[02],.F.,.F.,/*bAction*/,"L01",/*bGotFocus*/)
	oList	:= oArea:GetWinPanel("L01C01","LIST","L01")

	// Exportar Dados
	If lApliPath
		oButt2 := tButton():New(003,243,"&Exportar",oText,{|| MACTWizFon(oListBox,2) },45,11,,,.F.,.T.,.F.,,.F.,,,.F. )
	Else
		oButt2 := tButton():New(003,240,"&Exportar",oText,{|| MACTWizFon(oListBox,1)},45,11,,,.F.,.T.,.F.,,.F.,,,.F. )
	EndIf

	// Sair
	oButt1 := tButton():New(003,301,"&Sair",oText,{|| oFather:End() },45,11,,,.F.,.T.,.F.,,.F.,,,.F. )

	nTamCol := Len(aCols[01])
	bLine 	:= "{|| {"
	For Nx := 1 To nTamCol
		bLine += "aCols[oListBox:nAt]["+StrZero(Nx,3)+"]"
		If Nx < nTamCol
			bLine += ","
		EndIf
	Next
	bLine += "} }"

	oListBox := TCBrowse():New(0,0,386,130,,aHeader,,oList,'Fonte')
	oListBox:SetArray( aCols )
	oListBox:bLine := &bLine

	If !Empty( aColSizes )
		oListBox:aColSizes := aColSizes
	EndIf
	oListBox:SetFocus()

	// Instancia objeto CSS
	oListBox:SetCss(cCSSTCBrw)
	oButt1:SetCss(cCSSTBut)
	oButt2:SetCss(cCSSTBut)

	If oButt3 <> Nil
		oButt3:SetCss(cCSSTBut)
	EndIf

	oFather:Activate(,,,.T.,/*valid*/,,/*On Init*/)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MACTFilFon
Pesquisa um fonte na tela

@Param
oListBox - objeto da tela
cFonte	 - fonte para pesquisa

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MACTFilFon(oListBox,cFonte)

	Local nPos  	 := 0
	Local lRet  	 := .F.
	Local ni		 := 1
	Local lPosPesq	 := .F.
	Default oListBox := Nil
	Default cFonte	 := ""

	If RAT(".", cFonte) > 0
		cFonte := Alltrim(LEFT(cFonte, RAT(".", cFonte) - 1))
	Else
		cFonte := AllTrim(cFonte)
	EndIf

	// Faz um scan no objeto para encontrar a posição e posicionar no browser
	If Valtype(cFonte) = "C" .And. !Empty(cFonte)
		nPos := aScan( oListBox:aArray, {|x| x[1] == cFonte+".PRW" } )

		If nPos > 0
			oListBox:GoPosition(nPos)
			oListBox:Refresh()
			lRet  := .T.
		Else
			nPos := aScan( oListBox:aArray, {|x| x[1] == cFonte+".PRX" } )
			If nPos > 0
				oListBox:GoPosition(nPos)
				oListBox:Refresh()
				lRet  := .T.
			EndIf
		EndIf

		// Pesquisa fonte parcial
		If !lRet
			For ni := 1 to Len(oListBox:aArray)
				lPosPesq := cFonte $ oListBox:aArray[ni][1]
				If lPosPesq
					Exit
				EndIf
			Next ni
			If lPosPesq
				oListBox:GoPosition(ni)
				oListBox:Refresh()
			EndIf
		EndIf
	EndIf

	If nPos == 0 .And. !lPosPesq
		MsgAlert("Não foi possível encontrar o fonte " + cFonte + " na pesquisa.")
	EndIf

Return

//------------------------------------------------------------------------------------------
/* {Protheus.doc} MACTWizFon
Monta um wizard para exportar os dados

@Param
oListBox  - Objeto da area de trabalho da tela
nOpc	  - 1(Data Fontes) ;2(Hist. Atualizações); 3(Exporta ZIP); 4(Produtos Divergentes)
aDadosAmb - array com dados do ambiente
aFontes   - array com principais fontes
aVlCampos - array com campos para validação de decimais
aProc     - array de procedures
aPE       - array de pontos de entrada
aParam    - array de parametros

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MACTWizFon(oListBox,nOpc,aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam)

	Local oWizard 		:= Nil
	Local aPWiz 		:= {}
	Local aRetWiz		:= {}
	Local nOpcRot		:= 0
	Local aArea			:= GetArea()

	Default oListBox    := Nil
	Default nOpc		:= 0

	//Wizard
	aAdd(aPWiz,{ 1,"Nome do arquivo",Space(50),"","","","",50,	.T.}) //"Nome do arquivo"
	aAdd(aPWiz,{ 6,"Diretorio de Gravacao",Space(100),"","","",100,.T.,"", "",GETF_LOCALHARD + GETF_RETDIRECTORY }) //"Diretorio de Gravacao"
	aAdd(aRetWiz,Space(50))
	aAdd(aRetWiz,Space(50))
	aAdd(aRetWiz,.F.)

	DEFINE WIZARD oWizard TITLE "Exportar Dados"; // "Exportar Dados"
	HEADER "Monitor de Custos" ;  // "Monitor de Custos"
	MESSAGE "Parâmetros Iniciais..." 	 ; //"Parâmetros Iniciais..."
	TEXT "Essa rotina irá exportar os dados do monitor de custos."; // "Essa rotina irá exportar os dados do monitor de custos."
	NEXT {||.T.} ;
	FINISH {|| .F. } ;
	PANEL

	//Wizard
	CREATE PANEL oWizard ;
	HEADER "Dados do Arquivo"; //"Dados do Arquivo"
	MESSAGE "";
	BACK {|| .T. } ;
	NEXT {|| .T. } ;
	FINISH {|| nOpcRot := 1 , .T. } ;
	PANEL

	ParamBox(aPWiz,"Parâmetros...",@aRetWiz,,,,,,oWizard:GetPanel(2)) //"Parâmetros..."

	ACTIVATE WIZARD oWizard CENTERED

	If nOpcRot == 1
		Do Case
			Case nOpc == 1
			FWMsgRun(,{|| CursorWait(),MACTExpFon(oListBox,aRetWiz[1],aRetWiz[2])},,"Aguarde! Gerando arquivo..." )
			Case nOpc == 2
			FWMsgRun(,{|| CursorWait(),MATCExAtu(oListBox,aRetWiz[1],aRetWiz[2])},,"Aguarde! Gerando arquivo..." )
			Case nOpc == 3
			FWMsgRun(,{|| CursorWait(),MATCGerZip(aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam,aRetWiz[1],aRetWiz[2])},,"Aguarde! Gerando arquivo..." )
			Case nOpc == 4
			FWMsgRun(,{|| CursorWait(),MATCExprd(aRetWiz[1],aRetWiz[2])},,"Aguarde! Gerando arquivo..." )
		EndCase
	EndIf

	RestArea(aArea)

Return()

//------------------------------------------------------------------------------------------
/* {Protheus.doc} MACTExpFon
Exporta a data dos fontes para um arquivo texto

@Param
oListBox - Objeto da area de trabalho da tela
cNomeArq - Nome do arquivo
cCaminho - Caminho do arquivo
aDados   - Array com os dados para apresentação
lExpFon  - Define se a chamada foi realizado pela rotina de exportação

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018
@protected

@Return (Nil)
*/
//------------------------------------------------------------------------------------------
Static Function MACTExpFon(oListBox,cNomeArq,cCaminho,aDados,lExpFon)

	Local cTexto	:= ""
	Local nHandle	:= 0
	Local cFileTxt	:= ""
	Local lRetorno	:= .T.
	Local cFile		:= ""
	Local nx

	Default oListBox := Nil
	Default cNomeArq := ""
	Default cCaminho := ""
	Default aDados:= {}
	Default lExpFon	 := .F.

	cCaminho := AllTrim(cCaminho)

	If Empty(cNomeArq) .And. Empty(cCaminho)
		Aviso("GeraArquivo", "Parametros em Branco, verifique!", {"Ok"}) // "GeraArquivo" ## "Parametros em Branco, verifique!" ## "Ok"
		lRetorno := .F.
	EndIf

	If Len(aDados) < 0
		lRetorno := .F.
	EndIf

	If lRetorno
		cTexto := "----------------------------------------------------------------------------------------------------------------------"
		cTexto += _CRLF
		cTexto += "                         				Diagnóstico de Fontes" // "                         Diagnóstico de Fonte
		cTexto += _CRLF
		cTexto += "----------------------------------------------------------------------------------------------------------------------"
		cTexto += _CRLF + _CRLF

		cTexto += "Fonte                   Data          Hora"
		cTexto += _CRLF
		For nx := 1 To Len(aDados)
			cTexto += PADR((aDados[nx][1]),20) + "    " + cValtoChar(aDados[nx][2]) + "    " + (aDados[nx][3])
			cTexto += _CRLF
		Next nx

		cFile    := Alltrim(cNomeArq) + ".txt"
		cFileTxt := cCaminho + cFile

		// Cria arquivo texto
		nHandle := MsFCreate(cFileTxt)

		If nHandle < 0
			Aviso("Geração de Arquivo Texto","Não foi possível criar o arquivo: " + cFile + "." + " Erro: " + IIf(cValToChar( FError() ) == "13", "Sem permissão de acesso ao diretório, verifique!",cValToChar( FError() )),{"Ok"},3) // "Geração de Arquivo Texto" ## "Não foi possível criar o arquivo: " ## "Erro: " ## "Ok"
		Else
			WrtStrTxt(nHandle,cTexto)
			If !lExpFon
				Aviso("Geração de Arquivo Texto","Arquivo: " + cFile + " " + "gerado com sucesso no Diretório: " + Alltrim(cCaminho) + ".",{"Ok"},3) // "Geração de Arquivo Texto" ## "Arquivo: " ## "gerado com sucesso no Diretório: " ## "Ok"
			EndIf
		EndIf

		FClose(nHandle)
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCHAtu
Retorna o histórico de atualizações aplicadas no RPO

@aParam lGeraHist - Se chamada através da rotina de exportação gera txt sem montar tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018
@protected

@Return aDados - Array com dados dos pacotes aplicados no RPO
*/
//------------------------------------------------------------------------------------------
Static Function MATCHAtu(lGeraHist)

	Local aData  := GetRpoLog()
	Local aDados := {}
	Local nx
	Local cBarra     := If(issrvunix(), "/", "\")
	Local cPath	     := GetSrvProfString("StartPath", "") + If( Right( GetSrvProfString("StartPath",""), 1 ) == cBarra, "", cBarra )
	Local cExpAtAmb := "Atualiza_Ambiente"

	Default lGeraHist := .F.

	If len(aData) > 2
		For nx := 3 To Len(aData)
			aAdd(aDados,{cValtoChar(aData[nx][1]),cValtoChar(aData[nx][2]),cValtoChar(aData[nx][4]),cValtoChar(aData[nx][6]),aData[nx]})
		Next nx
	EndIf

	// Monta tela para exibir os dados dos fontes
	If !lGeraHist
		If Len(aDados) > 0
			MATCLisInf( {'Nome do patch','Data de geração do patch','Data de aplicação do patch','Número de programas'}, aDados, {75,75,75,75}, "Analise de Dados - Histórico de Atualizações", .T.)
		Else
			MsgInfo("Não foram aplicadas patchs neste ambiente!")
		EndIf
	Else
		FWMsgRun(,{|| CursorWait(),MATCExAtu(Nil,cExpAtAmb,cPath,aDados,lGeraHist)},,"Aguarde! Gerando arquivo..." )
	EndIf

Return (aDados)


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCExAtu
Exporta o histórico de atualizações

@Param
oListBox  - Objeto da area de trabalho da tela
cNomeArq  - Nome do arquivo para gravação
cCaminho  - Diretório para gravação
aInfPatc  - Array com os dados
lExpPatc  - Define se a rotina foi chamada pela rotina de exportação ZIP

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018
@protected

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCExAtu(oListBox,cNomeArq,cCaminho,aInfPatc,lExpPatc)

	Local cTexto	:= ""
	Local nHandle	:= 0
	Local cFileTxt	:= ""
	Local lRetorno	:= .T.
	Local cFile		:= ""
	Local nx
	Local ny

	Default oListBox := Nil
	Default cNomeArq := ""
	Default cCaminho := ""
	Default aInfPatc := {}
	Default lExpPatc := .F.

	cCaminho := AllTrim(cCaminho)

	If oListBox <> Nil
		aInfPatc := oListBox:aArray
	EndIf

	If Empty(cNomeArq) .And. Empty(cCaminho)
		Aviso("GeraArquivo", "Parametros em Branco, verifique!", {"Ok"}) // "GeraArquivo" ## "Parametros em Branco, verifique!" ## "Ok"
		lRetorno := .F.
	EndIf

	If lRetorno
		cTexto := "----------------------------------------------------------------------------------------------------------------------"
		cTexto += _CRLF
		cTexto += "                         				RPO - Histórico de Atualizações" // "                         RPO - Histórico de Atualizações"
		cTexto += _CRLF
		cTexto += "----------------------------------------------------------------------------------------------------------------------"
		cTexto += _CRLF + _CRLF

		cTexto += "Quantidade de paths aplicados: " + cValtoChar(Len(aInfPatc))
		cTexto += _CRLF + _CRLF
		For nx := 1 To Len(aInfPatc)
			cTexto += "Nome do patch                 Data de geração do patch         Data de aplicação do patch          Número de programas"
			cTexto += _CRLF
			cTexto += aInfPatc[nx][1] + "                  " + aInfPatc[nx][2] + "                       " + aInfPatc[nx][3] + "                          " + aInfPatc[nx][4]
			cTexto += _CRLF + _CRLF
			cTexto += "Fonte                          Data"
			cTexto += _CRLF
			For ny := 7 to Len(aInfPatc[nx][5])
				If ValType(aInfPatc[nx][5][ny]) == "A"
					cTexto += PADR(aInfPatc[nx][5][ny][1],20) + "             " + cValtoChar(aInfPatc[nx][5][ny][2])
					cTexto += _CRLF
				EndIf
			Next ny
			cTexto += _CRLF + _CRLF
		Next nx

		cFile    := Alltrim(cNomeArq) + ".txt"
		cFileTxt := cCaminho + cFile

		// Cria arquivo texto
		nHandle := MsFCreate(cFileTxt)

		If nHandle < 0
			Aviso("Geração de Arquivo Texto","Não foi possível criar o arquivo: " + cFile + "." + " Erro: " + IIf(cValToChar( FError() ) == "13", "Sem permissão de acesso ao diretório, verifique!",cValToChar( FError() )),{"Ok"},3) // "Geração de Arquivo Texto" ## "Não foi possível criar o arquivo: " ## "Erro: " ## "Ok"
		Else
			WrtStrTxt(nHandle,cTexto)
			If !lExpPatc
				Aviso("Geração de Arquivo Texto","Arquivo: " + cFile + " " + "gerado com sucesso no Diretório: " + Alltrim(cCaminho) + ".",{"Ok"},3) // "Geração de Arquivo Texto" ## "Arquivo: " ## "gerado com sucesso no Diretório: " ## "Ok"
			EndIf
		EndIf

		FClose(nHandle)
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCGerZip
MATCGerZip() - FPanel01_5 - Folder1 - Exportar Dados
Exporta consulta para arquivo texto

@Param
aDadosAmb - array com dados do ambiente
aFontes   - array com principais fontes
aVlCampos - array com campos para validação de decimais
aProc     - array de procedures
aPE       - array de pontos de entrada
aParam    - array de parametros
cNome	  - nome para geração do arquivo
cCaminho  - diretório para exportação

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCGerZip(aDadosAmb,aFontes,aVlCampos,aProc,aPE,aParam,cNome,cCaminho)

	Local cTexto	:= ""
	Local nHandle	:= 0
	Local cFile		:= ""
	Local cFileTxt	:= ""
	Local lRetorno	:= .T.
	Local nx
	Local ny
	Local cBarra    := If(issrvunix(), "/", "\")
	Local cPath	    := GetSrvProfString("StartPath", "") + If( Right( GetSrvProfString("StartPath",""), 1 ) == cBarra, "", cBarra )
	Local cFileTRX	:= "bkarq_trx"
	Local cFileTRT	:= "bkarq_trb"
	Local cEnvSrv     := GetEnvServer()
	Local cLocalFiles := Upper(GetPvProfString(cEnvSrv, "LocalFiles", "ADS", GetADV97()))
	Local cExtArq	  := ""
	Local cDirFile	:= Alltrim(cNome)+".ZIP"
	Local cNomeArq	:= "Monitor_Custos.txt"
	Local cFontes	:= "Fontes.txt"
	Local cAtual	:= "Historico_RPO.txt"
	Local aArquivos := {}
	Local nPos		:= 0
	Local aHistAtu	:= {}
	Local bOk

	Default aDadosAmb := {}
	Default aFontes   := {}
	Default aVlCampos := {}
	Default aProc	  := {}
	Default aPE		  := {}
	Default aParam    := {}
	Default cNome    := ""
	Default cCaminho := ""

	// Verifica tipo de arquivo utilizado
	If cLocalFiles == "ADS"
		cExtArq := ".dbf"
	ElseIf cLocalFiles == "CTREE"
		cExtArq := ".dtc"
	EndIf

	cTexto := "----------------------------------------------------------------------------------------------------------------------"
	cTexto += _CRLF
	cTexto += "                         MONTIOR DE CUSTOS - Diagnóstico detalhado para análise de Custos" // "                         MONTIOR DE CUSTOS - Diagnóstico detalhada para análise de Custos"
	cTexto += _CRLF
	cTexto += "----------------------------------------------------------------------------------------------------------------------"
	cTexto += _CRLF + _CRLF

	// Dados do Ambiente
	If Len(aDadosAmb) > 0
		cTexto += "Dados do Ambiente"
		cTexto += _CRLF
	EndIf

	For nx := 1 To Len(aDadosAmb)
		If Len(aDadosAmb[nx]) > 0
			cTexto += PADR(aDadosAmb[nx][1],15) + " - " +  aDadosAmb[nx][2]
			cTexto += _CRLF
		EndIf
	Next nx

	cTexto += _CRLF + _CRLF

	// Data dos Fontes
	If Len(aFontes) > 0
		cTexto += "Data dos Fontes"
		cTexto += _CRLF
	EndIf

	For nx := 1 To Len(aFontes)
		If Len(aFontes[nx]) > 0
			cTexto += aFontes[nx]
			cTexto += _CRLF
		EndIf
	Next nx

	cTexto += _CRLF + _CRLF

	// Validação dos decimais dos campos
	If Len(aVlCampos) > 0
		cTexto += "Validação de Campos"
		cTexto += _CRLF
	EndIf

	// Pego a última posição do array aVLCampos
	nPos := Len(aVlCampos)
	If ValType(aVlCampos[nPos]) == "L" .And. !(aVlCampos[nPos])
		cTexto += "Campos validados corretamente."
	Else
		cTexto += "ATENÇÃO: Existem campos com decimais divergentes, poderao ocorrer diferenças de arredondamento!"
		cTexto += _CRLF + _CRLF
		For nx := 1 To Len(aVlCampos)
			If nx <> nPos .And. (Len(aVlCampos[nx]) > 0)
				cTexto += aVlCampos[nx][1] + "        " + PADR(aVlCampos[nx][2],30) + "        " + PADR(cValtoChar(aVlCampos[nx][3]),9) + "        " + PADR(cValtoChar(aVlCampos[nx][4]),15) + "        " + cValtoChar(aVlCampos[nx][5])
				cTexto += _CRLF
			EndIf
		Next nx
	EndIf

	cTexto += _CRLF + _CRLF

	// Procedures
	If Len(aProc) > 0
		cTexto += "Procedures"
		cTexto += _CRLF
	EndIf

	For nx := 1 To Len(aProc)
		If Len(aProc[nx]) > 0
			cTexto += aProc[nx]
			cTexto += _CRLF
		EndIf
	Next nx

	cTexto += _CRLF + _CRLF

	// Pontos de Entrada
	If Len(aPE) > 0
		cTexto += "Pontos de Entrada"
		cTexto += _CRLF
	EndIf

	For nx := 1 To Len(aPE)
		If Len(aPE[nx]) > 0
			cTexto += aPE[nx][1] + " - " + aPE[nx][2]
			cTexto += _CRLF
		EndIf
	Next nx

	cTexto += _CRLF + _CRLF

	// Parametros
	If Len(aParam) > 0
		cTexto += "Parâmetros"
		cTexto += _CRLF
	EndIf

	For nx := 1 To 4 // Tamanho do array aParam
		For ny := 1 To Len(aParam[nx])
			cTexto += PADR(aParam[nx][ny][1],12) + " = " + aParam[nx][ny][2]
			cTexto += _CRLF
		Next ny
	Next nx

	cTexto += _CRLF + _CRLF

	//Carrega todos os fontes do RPO
	MATCFontes(.T.)
	aAdd(aArquivos,cFontes)

	//Carrega as atualizações aplicadas no RPO
	aHistAtu := MATCHAtu(.T.)
	If Len(aHistAtu) > 0
		aAdd(aArquivos,cAtual)
	Else
		cTexto += "Histórico de Atualizações"
		cTexto += _CRLF
		cTexto += "Não foram aplicadas patchs neste ambiente!"
	EndIf

	// Nome e local do arqruivo
	cFileTxt := cPath + cNomeArq

	nHandle := MsFCreate(cFileTxt)

	If nHandle < 0
		Aviso("Geração de Arquivo Texto","Não foi possível criar o arquivo: " + cNomeArq + "." + " Erro: " + IIf(cValToChar( FError() ) == "13", "Sem permissão de acesso ao diretório, verifique!",cValToChar( FError() )),{"Ok"},3) // "Geração de Arquivo Texto" ## "Não foi possível criar o arquivo: " ## "Erro: " ## "Ok"
	Else
		WrtStrTxt(nHandle,cTexto)
	EndIf

	FClose(nHandle)

	// Arquivos temporários do recalculo
	cFileTRX := cFileTRX+cExtArq
	cFileTRT := cFileTRT+cExtArq
	If File(cFileTRX) .And. File(cFileTRT)
		aAdd(aArquivos,cFileTRX)
		aAdd(aArquivos,cFileTRT)
	EndIf

	// Adiciono os arquivos em Array para gerar .zip
	aAdd(aArquivos,cNomeArq)

	// Gera arquivo ZIP
	nRet := FZip(cDirFile,aArquivos,cPath)

	// Move o arquivo para o diretório especificado pelo usuário
	bOk := CpyS2T( cPath+cDirFile, cCaminho,.F. )

	If nRet!=0 .Or. !bOk
		MsgStop("Não foi possível criar o arquivo zip. Verifique permissões de usuário no diretório escolhido!")
	Else
		Aviso("Geração de Arquivo ZIP","Arquivo: " + cDirFile + " " + "gerado com sucesso no diretório: " + cCaminho + ".",{"Ok"},3) // "Geração de Arquivo Texto" ## "Arquivo: " ## "gerado com sucesso no Diretório: " ## "Ok"
	EndIf

	// Apaga arquivos
	For nx := 1 to Len(aArquivos)
		If File(aArquivos[nx])
			FERASE(aArquivos[nx])
		EndIf
	Next nx
	If File(cDirFile)
		FERASE(cDirFile)
	EndIf

Return()

//------------------------------------------------------------------------------------------
/* {Protheus.doc} FPanel04
Painel04 - Folder4 - Analise

@Param
oPanel    - objeto da tela

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function FPanel04(oPanel)

	Local aArea        := GetArea()
	Local oNoMarked    := LoaDBitmap( GetResources(), "LBNO" )
	Local oMarked      := LoaDBitmap( GetResources(), "LBOK" )
	Local oBtProces, oBtProcKa, oBtSair, oBtPrDiv := Nil
	Local cCusfil	   := cValtoChar(SuperGetMv("MV_CUSFIL",.F.,""))
	Local aTFolder2    := {}

	Default oPanel     := Nil

	AAdd(aKarLocal ,{CTOD("  /  /    "),Space(06),Space(02),Space(14),Space(06),Space(06),Space(06),Space(06),Space(06),Space(06),Space(06),Space(06),Space(06)})

	DEFINE FONT oBold   NAME "Arial" SIZE 0, -12 BOLD
	DEFINE FONT oBold1  NAME "Arial" SIZE 0, -12
	DEFINE FONT oBold2  NAME "Arial" SIZE 0, -40 BOLD

	DBSelectArea("SB1")

	aTFolder2 := {"Fechamento","Kardex"}
	oFolder2 := TFolder():New(060,006,aTFolder2,,oPanel,,,,.T.,,648, 190)
	oFolder2:lVisible
	oFolder2:show()

	@ 014,010 SAY "Produto"                                 SIZE 040,10 PIXEL OF oPanel FONT oBold
	@ 010,040 MSGET oVar   VAR cProduto Picture "@!"        SIZE 060,10 PIXEL OF oPanel F3 "SB1" VALID(MATCDescPrd())
	@ 010,105 MSGET oVar   VAR cDescr   Picture "@!"        SIZE 156,10 PIXEL OF oPanel When .F.
	@ 032,010 SAY "Local"                                   SIZE 040,10 PIXEL OF oPanel FONT oBold
	@ 028,040 MSGET oVar   VAR cLocal Picture cPictLoc     	SIZE 030,10 PIXEL OF oPanel VALID(GetDatas(1))
	@ 049,010 SAY "Limite"                                  SIZE 070,10 PIXEL OF oPanel FONT oBold
	@ 045,040 MSGET oVar   VAR dDtProces Picture "99/99/9999" SIZE 040,10 PIXEL OF oPanel VALID .T.

	@ 032,095 SAY "1 UM"                                    SIZE 070,10 PIXEL OF oPanel FONT oBold
	@ 028,114 MSGET oVar   VAR c1UM      Picture "@!"       SIZE 015,10 PIXEL OF oPanel When .F.
	@ 032,150 SAY "2 UM"                                    SIZE 070,10 PIXEL OF oPanel FONT oBold
	@ 028,190 MSGET oVar   VAR c2UM      Picture "@!"       SIZE 015,10 PIXEL OF oPanel When .F.

	@ 014,270 SAY "Ultimo Fechamento"                       SIZE 070,10 PIXEL OF oPanel FONT oBold
	@ 010,330 MSGET oVar   VAR dUltFech  Picture "99/99/9999" SIZE 040,10 PIXEL OF oPanel When .F.
	@ 029,270 SAY "Data Inicial"                            SIZE 070,10 PIXEL OF oPanel FONT oBold
	@ 025,330 MSGET oVar   VAR dDataI    Picture "99/99/9999" SIZE 040,10 PIXEL OF oPanel When .F.
	@ 044,270 SAY "Data Final"                              SIZE 070,10 PIXEL OF oPanel FONT oBold
	@ 040,330 MSGET oVar   VAR dDataF    Picture "99/99/9999" SIZE 040,10 PIXEL OF oPanel VALID .T.

	@ 049,095 SAY "Lote"                                    SIZE 040,10 PIXEL OF oPanel FONT oBold
	@ 045,114 MSGET oVar   VAR cCrlLot   Picture "@!"       SIZE 025,10 PIXEL OF oPanel When .F.
	@ 049,150 SAY "Localização"                             SIZE 040,10 PIXEL OF oPanel FONT oBold
	@ 045,190 MSGET oVar   VAR cCrlEnd   Picture "@!"       SIZE 025,10 PIXEL OF oPanel When .F.
	@ 049,220 SAY "Zona"                                    SIZE 040,10 PIXEL OF oPanel FONT oBold
	@ 045,240 MSGET oVar   VAR cZona     Picture "@!"       SIZE 021,10 PIXEL OF oPanel When .F.

	/*
	CheckBox MV_CUSFIL - Parametro utilizado para verificar se o   |
	|sistema utiliza custo unificado por:                          |
	|      F = Custo Unificado por Filial                          |
	|      E = Custo Unificado por Empresa                         |
	|      A = Custo Unificado por Armazem                         |
	*/

	@ 014,380 SAY "MV_CUSFIL" SIZE 070,10 PIXEL OF oPanel FONT oBold
	If cCusfil == "A"
		lCusfilA := .T.
		TSay():New( 014, 420, { || "__________" }        , oPanel,,oBold,,,, .T.,CLR_HGREEN,, 100, 010 )
	ElseIf cCusfil == "F"
		lCusfilF := .T.
		TSay():New( 026, 420, { || "_______" }        , oPanel,,oBold,,,, .T.,CLR_HGREEN,, 100, 010 )
	ElseIf cCusfil == "E"
		lCusfilE := .T.
		TSay():New( 038, 420, { || "__________" }        , oPanel,,oBold,,,, .T.,CLR_HGREEN,, 100, 010 )
	EndIf
	TCheckBox():New( 013,420,"Amarzem",bSETGET(lCusfilA),oPanel,150,009,,,,,,,,.T.,,,)
	TCheckBox():New( 025,420,"Filial" ,bSETGET(lCusfilF),oPanel,150,009,,,,,,,,.T.,,,)
	TCheckBox():New( 037,420,"Empresa",bSETGET(lCusfilE),oPanel,150,009,,,,,,,,.T.,,,)

	@ 014,465  SAY "Kardex" SIZE 070,10 PIXEL OF oPanel FONT oBold
	TCheckBox():New( 013,494,"Digitação",bSETGET(lKardex1),oPanel,150,009,,,,,,,,.T.,,,)
	TCheckBox():New( 025,494,"Sequência" ,bSETGET(lKardex2),oPanel,150,009,,,,,,,,.T.,,,)

	oBtProces := TButton():New( 013, 545, "Processar",oPanel,{||MATCProces(oFolder2)}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )
	oBtSair   := TButton():New( 013, 600, "Sair",oPanel,{||MACTFecha()}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )

	TSay():New( 037, 545, { || "Consultar Produtos Divergentes" }        , oPanel,,oBold,,,, .T.,CLR_HBLUE,, 100, 010 )
	oBtPrDiv := TButton():New( 047, 545, "Consultar",oPanel,{||MATCPrDiv()}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )

	/*Saldo Inicial*/
	@ 007,007 TO 86, 91 PROMPT "Saldo Inicial - SB9" PIXEL OF oFolder2:aDialogs[1]
	@ 022,012 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 018,032 MSGET oQtSB9  VAR nQtSB9 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 040,012 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 036,032 MSGET oVlSB9  VAR nVLSB9 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 057,012 SAY "Fechamento por OP"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold1
	@ 071,012 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 067,032 MSGET oVlOpSB9  VAR nVlOpSB9 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	/*Saldo Entradas*/
	@ 007,99 TO 86, 183 PROMPT "Entradas - SD1" PIXEL OF oFolder2:aDialogs[1]
	@ 022,104 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 018,124 MSGET oQtSD1  VAR nQtSD1 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 040,104 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 036,124 MSGET oVlSD1  VAR nVlSD1 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	/*Saldo Saídas*/
	@ 007,191 TO 86, 275 PROMPT "Saídas - SD2"  PIXEL OF oFolder2:aDialogs[1]
	@ 022,196 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 018,216 MSGET oQtSD2  VAR nQtSD2 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 040,196 SAY "Custo"                        SIZE 25,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 036,216 MSGET oVlSD2  VAR nVlSD2 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	/*Saldo Movimentações Internas*/
	@ 007,283 TO 86, 559 PROMPT "Movimentações Internas - SD3" PIXEL OF oFolder2:aDialogs[1]

	@ 015,291 TO 84, 373 PROMPT "Saldo Produção" PIXEL OF oFolder2:aDialogs[1]
	@ 024,296 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 022,316 MSGET oQtPSD3  VAR nQtPSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 042,296 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 038,316 MSGET oVlPSD3  VAR nVlPSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 057,296 SAY "Fechamento por OP"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold1
	@ 071,296 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 067,316 MSGET oVlOpPSD3  VAR nVlOpPSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	@ 015,380 TO 84, 462 PROMPT "Saldo Requisição" PIXEL OF oFolder2:aDialogs[1]
	@ 024,385 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 022,405 MSGET oQtRSD3  VAR nQtRSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 042,385 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 038,405 MSGET oVlRSD3  VAR nVlRSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 057,385 SAY "Fechamento por OP"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold1
	@ 071,385 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 067,405 MSGET oVlOpRSD3 VAR nVlOpRSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	@ 015,469 TO 84, 551 PROMPT "Saldo Devolução" PIXEL OF oFolder2:aDialogs[1]
	@ 024,473 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 022,493 MSGET oQtDSD3  VAR nQtDSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 042,473 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 038,493 MSGET oVlDSD3  VAR nVlDSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 057,473 SAY "Fechamento por OP"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold1
	@ 071,473 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 067,493 MSGET oVlOpDSD3  VAR nVlOpDSD3 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	/*Contabilização*/
	/*
	@ 007,566 TO 86, 637 PROMPT "Contabilização" PIXEL OF oFolder2:aDialogs[1]
	@ 020,570 SAY "Valor dos Movimentos"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 030,574 MSGET oVlContab  VAR nVlContab Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 050,570 SAY "Valor Contabilizado"           SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 060,574 MSGET oVlCT2  VAR nVlCT2 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	*/

	/*Saldo Fechamento*/
	@ 090,007 TO 165, 91 PROMPT "Fechamento: SB2" PIXEL OF oFolder2:aDialogs[1]
	@ 105,012 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 101,032 MSGET oFecQtSB2  VAR nFecQtSB2 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 123,012 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 119,032 MSGET oFecVlSB2  VAR nFecVlSB2 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 140,012 SAY "Fechamento por OP"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold1
	@ 154,012 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 150,032 MSGET oFecVlOpSB2  VAR nFecVlOpSB2 Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	/*Saldo Fechamento x Movimento*/
	@ 090,099 TO 165, 183 PROMPT "Fechamento: Movimento" PIXEL OF oFolder2:aDialogs[1]
	@ 105,104 SAY "Quant"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 101,124 MSGET oFecMovQt  VAR nFecMovQt Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 123,104 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 119,124 MSGET oFecMovVl  VAR nFecMovVl Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.
	@ 140,104 SAY "Fechamento por OP"            SIZE 080,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold1
	@ 154,104 SAY "Custo"                        SIZE 025,10 PIXEL OF oFolder2:aDialogs[1] FONT oBold
	@ 150,124 MSGET oFecMovOp  VAR nFecMovOp Picture cPict SIZE 055,10 PIXEL OF oFolder2:aDialogs[1] When .F.

	@ 090,191 TO 165, 367 PROMPT "Status do Fechamento" 	PIXEL OF oFolder2:aDialogs[1]

	@ 090,374 TO 165, 637 PROMPT "Divergências"             PIXEL OF oFolder2:aDialogs[1]

	// FOLDER 02
	// Kardex - Almoxarifado
	@ 003,010 SAY "Saldo Inicial : "+Trans(nQtSB9,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 003,090 SAY "Custo Inicial : "+Trans(nVLSB9,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold

	@ 162,010 SAY "T O T A I S : " SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 162,161 SAY Trans(nQtSD1 + nQtPSD3 + nQtDSD3,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 162,233 SAY Trans(nVlSD1 + nVlPSD3 + nVlDSD3,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 162,336 SAY Trans(nQtSD2 + nQtRSD3,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 162,404 SAY Trans(nVlSD2 + nVlRSD3,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 162,500 SAY Trans(nFecMovQt,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold
	@ 162,563 SAY Trans(nFecMovVl,cPict) SIZE 200,50 PIXEL OF oFolder2:aDialogs[2] FONT oBold

	oBtProcKa   := TButton():New( 003, 570, "Kardex - MATR900",oFolder2:aDialogs[2],{||MATR900()}, 60,08,,,.F.,.T.,.F.,,.F.,,,.F. )

	// Mostra browser com os valores de Custo
	MATCBrwCus(oFolder2:aDialogs[2],aKarLocal,001,145)

	// Instancia objeto CSS
	oBtProces:SetCss(cCSSTBut)
	oBtSair:SetCss(cCSSTBut)
	oBtProcKa:SetCss(cCSSTBut)
	oBtPrDiv:SetCss(cCSSTBut)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCDescPrd
Painel04 - Folder4 - MATCDescPrd

@Param

@author    Ronaldo Tapia
@version   12.1.17
@since     13/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCDescPrd()

	Local aArea    := GetArea()

	DBSelectArea("SB1")
	DBSetOrder(1)
	IF !Empty(cProduto) .AND. !DBSeek(xFilial("SB1")+cProduto)
		MsgStop("Produto não cadastrado, verifique!")
		Return (.F.)
	Endif

	cDescr   := SB1->B1_DESC
	cCrlLot  := If(SB1->B1_RASTRO ="L","Lote",If(SB1->B1_RASTRO ="S","Sub-Lote","Não"))
	cCrlEnd  := If(SB1->B1_LOCALIZ="S","Sim","Não")
	cLocal   := IF(Empty(cLocal),SB1->B1_LOCPAD,cLocal)
	c1UM     := SB1->B1_UM
	c2UM     := SB1->B1_SEGUM

	DBSelectArea("SB5")
	DBSetOrder(1)
	DBSeek(xFilial("SB5")+cProduto)
	cZona := SB5->B5_CODZON

	RestArea(aArea)

Return(.T.)


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCProces
Painel04 - Folder4 - MATCProces
Processa as informações de custo

@Param
oFolder2 - objeto tFolder da página

@author    Ronaldo Tapia
@version   12.1.17
@since     13/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCProces(oFolder2)

	Local cQuery   := ""
	Local cQtdMov  := ""
	Local aArea    := GetArea()
	Local Qrykdex, QrySC2, QrySB9, QrySB2 := ""
	Local aTamSD3  := TamSx3('D3_QUANT')
	Local lRet     := .T.
	Local lFound   := .F.
	Local nX	   := 0
	Local nCont    := 0
	Local aCusOp   := {}
	Local aAreaSM0  := {}
	// Zera variáveis da tela
	nQtSB9      := 0
	nVLSB9      := 0
	nVlOpSB9	:= 0
	nQtSD1		:= 0
	nVlSD1		:= 0
	nQtSD2		:= 0
	nVlSD2		:= 0
	nQtPSD3		:= 0
	nVlPSD3		:= 0
	nVlOpPSD3	:= 0
	nQtRSD3		:= 0
	nVlRSD3		:= 0
	nQtOpRSSD3	:= 0
	nVlOpRSD3	:= 0
	nQtDSD3		:= 0
	nVlDSD3		:= 0
	nVlOpDSD3	:= 0
	nFecQtSB2	:= 0
	nFecVlSB2	:= 0
	nFecVlOpSB2	:= 0
	nFecMovQt	:= 0
	nFecMovVl	:= 0
	nFecMovOp	:= 0
	nQuantArm   := 0

	//Contabilização
	nVlContab   := 0
	nVlCT2		:= 0


	DEFINE FONT oBold2  NAME "Arial" SIZE 0, -40 BOLD

	// Limpa objeto das mensagens
	If oSay <> Nil .And. !Empty(oSay:Ctitle)
		oSay:Ctitle := ""
		oSay:Refresh()
	EndIf
	If oBtErro <> Nil .And. !Empty(oBtErro:Ctitle)
		oBtErro:Ctitle := ""
		oBtErro:Refresh()
	EndIf

	// Zera valores do Kardex
	aSize(aKarLocal, 0)

	// Validações parâmetro MV_CUSFIL
	If lRet .And. (!lCusfilA .And. !lCusfilF .And. !lCusfilE)
		MsgStop("Parâmetro MV_CUSFIl não definido nas opções do Monitor. Verifique!")
		lRet := .F.
	EndIf

	If lRet .And. ((lCusfilA .And. lCusfilF) .Or. (lCusfilA .And. lCusfilE) .Or. (lCusfilF .And. lCusfilE))
		MsgStop("Só pode ser definido um tipo de custo na escolha do parâmetro MV_CUSLFIL. Verifique!")
		lRet := .F.
	EndIf

	If lRet .And. (!lKardex1 .And. !lKardex2)
		MsgStop("Defina um tipo de ordenação para o Kardex. Verifique!")
		lRet := .F.
	EndIf

	If lRet .And. (lKardex1 .And. lKardex2)
		MsgStop("Só pode ser definido um tipo de ordenação para o Kardex. Verifique!")
		lRet := .F.
	EndIf

	// Validações de Data
	If lRet .And. (Empty(dDataI) .OR. Empty(dDataF))
		MsgStop("Datas de Processamento não informadas. Verifique!")
		lRet := .F.
	EndIf

	If lRet .And. Empty(cProduto)
		MsgStop("Produto não informado. Verifique!")
		lRet := .F.
	EndIf

	If lRet
		// Valida o MV_CUSFIL selecionado
		If lCusfilA
			If cValtoChar(SuperGetMv("MV_CUSFIL",.F.,"A")) <> "A"
				MsgAlert("O parâmetro MV_CUSFIL definido na análise é diferente do configurado no dicionário, poderão ocorrer divergências nos valores de fechamento!")
			EndIf
		ElseIf lCusfilF
			If cValtoChar(SuperGetMv("MV_CUSFIL",.F.,"A")) <> "F"
				MsgAlert("O parâmetro MV_CUSFIL definido na análise é diferente do configurado no dicionário, poderão ocorrer divergências nos valores de fechamento!")
			EndIf
		ElseIf lCusfilE
			If cValtoChar(SuperGetMv("MV_CUSFIL",.F.,"A")) <> "E"
				MsgAlert("O parâmetro MV_CUSFIL definido na análise é diferente do configurado no dicionário, poderão ocorrer divergências nos valores de fechamento!")
			EndIf
		EndIf

		DBSelectArea("SB1")
		DBSetOrder(1)
		DBSeek(xFilial("SB1")+cProduto)

		ProcRegua(15)

		IF substr(cLocal,1,2) == "**"
			cLocAnDe  := "  "
			cLocAnAte := "ZZ"
		ELSE
			cLocAnDe  := cLocal
			cLocAnAte := cLocal
		ENDIF
		cProdAnDe  := cProduto
		cProdAnAte := cProduto

		// Saldo Inicial
		GetDatas() // Carrega a data do último fechamento
		IncProc("Processando Saldo Inicial")
		IF substr(cLocal,1,2) <> "**"
			cQuery := " SELECT SUM(B9_QINI) B9_QINI, SUM(B9_VINI1) B9_VINI1, B9_LOCAL AS ARMAZEM FROM "+RETSQLNAME("SB9")
		Else
			cQuery := " SELECT SUM(B9_QINI) B9_QINI, SUM(B9_VINI1) B9_VINI1 FROM "+RETSQLNAME("SB9")
		ENDIF
		cQuery += " WHERE B9_COD = '"+cProduto+"' "
		If lCusfilA // Custo por Armazem
			cQuery += " AND B9_DATA = '"+DTOS(dUltFech)+"' AND B9_FILIAL='"+xFilial("SB9")+"' "
			IF substr(cLocal,1,2) <> "**"
				cQuery += " AND B9_LOCAL = '"+cLocal+"' "
			ENDIF
		ElseIf lCusfilF // Custo por Filial
			cQuery += " AND B9_DATA = '"+DTOS(dUltFech)+"' AND B9_FILIAL = '"+xFilial("SB9")+"'"
		ElseIf lCusfilE // Custo por Empresa
			cQuery += " AND B9_DATA = '"+DTOS(dUltFech)+"'"
		EndIf
		cQuery += " AND "+RETSQLNAME("SB9")+".D_E_L_E_T_ = ' ' GROUP BY B9_LOCAL"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QrySB9",.F.,.T.)
		WHILE QrySB9->(!Eof())
			nQtSB9 += QrySB9->B9_QINI
			nVLSB9 += QrySB9->B9_VINI1
			IF substr(cLocal,1,2) <> "**" .AND. QrySB9->ARMAZEM == cLocal
				nQuantArm := QrySB9->B9_QINI
			ENDIF
			DBSKIP()
		ENDDO
		QrySB9->(DBCloseArea())

		/*--------------------------------------------------------------*/
		/*Variáveis para controle do valor de custo após as movimentações*/
		/*--------------------------------------------------------------*/
		nFecMovQt:= nQtSB9
		nFecMovVl:= nVLSB9

		IncProc("Processando Movimentos")
		// Entradas
		cQuery := "SELECT D1_FILIAL as Filial,"
		cQuery += " D1_EMISSAO as Emissao,"
		cQuery += " D1_CF as CF,"
		cQuery += " D1_LOCAL as Armazem,"
		cQuery += " D1_NUMSEQ as NUMSEQ,"
		cQuery += " D1_SEQCALC as SEQCALC,"
		cQuery += " D1_QUANT as QUANT,"
		cQuery += " D1_CUSTO as CUSTO"
		cQuery += " FROM "+RetSqlName("SD1")+" SD1,"+RetSqlName("SF4")+" SF4 "
		cQuery += " WHERE D1_COD = '"+cProduto+"'
		cQuery += " AND D1_ORIGLAN <> 'LF' AND D1_EMISSAO >= '"+DTOS(dDataI)+"' AND D1_EMISSAO <= '"+DTOS(dDataF)+"'"
		cQuery += " AND D1_TES=F4_CODIGO AND F4_ESTOQUE = 'S' AND F4_FILIAL = '"+xFilial("SF4")+"'"
		//Complemento do Where da tabela SD1 (Tratamento Filial)
		If lCusfilE
			If FWModeAccess("SD1") == 'E' .AND. FWModeAccess("SF4") == 'E'
				cQuery += " AND D1_FILIAL = F4_FILIAL"
			EndIf
		ElseIf lCusfilF
				cQuery += " AND D1_FILIAL = '"+xFilial("SD1")+"' "

		ElseIf lCusfilA
			IF SUBSTR(cLocal,1,2) <> "**"
				cQuery += " AND D1_LOCAL = '"+cLocal+"' "
			ENDIF
			cQuery += " AND D1_FILIAL = '"+xFilial("SD1")+"' "
		EndIf
		cQuery += " AND SD1.D_E_L_E_T_ = ' '"
		cQuery += " AND SF4.D_E_L_E_T_ = ' '"

		// Saídas
		cQuery += " UNION ALL"
		cQuery += " SELECT D2_FILIAL as Filial,"
		cQuery += " D2_EMISSAO as Emissao,"
		cQuery += " D2_CF as CF,"
		cQuery += " D2_LOCAL as Armazem,"
		cQuery += " D2_NUMSEQ as NUMSEQ,"
		cQuery += " D2_SEQCALC as SEQCALC,"
		cQuery += " D2_QUANT as QUANT,"
		cQuery += " D2_CUSTO1 as CUSTO"
		cQuery += " FROM "+RetSqlName("SD2")+" SD2,"+RetSqlName("SF4")+" SF4 "
		cQuery += " WHERE D2_COD = '"+cProduto+"'
		cQuery += " AND D2_ORIGLAN <> 'LF' AND D2_EMISSAO >= '"+DTOS(dDataI)+"' AND D2_EMISSAO <= '"+DTOS(dDataF)+"'"
		cQuery += " AND D2_TES = F4_CODIGO AND F4_ESTOQUE = 'S' AND F4_FILIAL = '"+xFilial("SF4")+"'"
		//Complemento do Where da tabela SD2 (Tratamento Filial)
		If lCusfilE
			If FWModeAccess("SD2") == 'E' .AND. FWModeAccess("SF4") == 'E'
				cQuery += " AND D2_FILIAL = F4_FILIAL"
			EndIf
		ElseIf lCusfilF
			cQuery += " AND D2_FILIAL = '"+xFilial("SD2")+"' "
		ElseIf lCusfilA
			IF SUBSTR(cLocal,1,2) <> "**"
				cQuery += " AND D2_LOCAL = '"+cLocal+"' "
			ENDIF
			cQuery += " AND D2_FILIAL = '"+xFilial("SD2")+"' "
		EndIf
		cQuery += " AND SD2.D_E_L_E_T_ = ' '"
		cQuery += " AND SF4.D_E_L_E_T_ = ' '"

		// Movimentação Interna
		cQuery += " UNION ALL"
		cQuery += " SELECT D3_FILIAL as Filial,"
		cQuery += " D3_EMISSAO as Emissao,"
		cQuery += " D3_CF as CF,"
		cQuery += " D3_LOCAL as Armazem,"
		cQuery += " D3_NUMSEQ as NUMSEQ,"
		cQuery += " D3_SEQCALC as SEQCALC,"
		cQuery += " D3_QUANT as QUANT,"
		cQuery += " D3_CUSTO1 as CUSTO"
		cQuery += " FROM "+RetSqlName("SD3")+" SD3 "
		cQuery += " WHERE D3_COD = '"+cProduto+"'"
		cQuery += " AND D3_ESTORNO = ' ' AND D3_EMISSAO >= '"+DTOS(dDataI)+"' AND D3_EMISSAO <= '"+DTOS(dDataF)+"'"
		//Complemento do Where da tabela SD3 (Tratamento Filial)
		If lCusfilF
			cQuery += " AND D3_FILIAL = '"+xFilial("SD3")+"' "
		ElseIf lCusfilA
			IF SUBSTR(cLocal,1,2) <> "**"
				cQuery += " AND D3_LOCAL = '"+cLocal+"' "
			ENDIF
			cQuery += " AND D3_FILIAL = '"+xFilial("SD3")+"' "
		EndIf
		cQuery += " AND SD3.D_E_L_E_T_ = ' '"

		If lKardex1  // Ordena por NUMSEQ
			cQuery += " ORDER BY NUMSEQ"
		ElseIf lKardex2  // Ordena por SEQCAL
			cQuery += " ORDER BY SEQCALC"
		EndIf

		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"Qrykdex",.F.,.T.)

		While Qrykdex->(!Eof())
			// Verifico qual CFOP para incluir no array aKarLocal que será carregado no browser
			If SUBSTRING(Qrykdex->CF,1,2) = 'PR' .Or. SUBSTRING(Qrykdex->CF,1,2) = 'DE' .Or. SUBSTRING(Qrykdex->CF,1,1) < '5'

				nFecMovQt += Qrykdex->QUANT // Quantidade Fechamento x Movimento
				nFecMovVl += Round(Qrykdex->CUSTO,TAMSX3("B2_CM1")[2]) // Valor Fechamento x Movimento
				AAdd(aKarLocal,{STOD(Qrykdex->Emissao), Qrykdex->CF,Qrykdex->Armazem, Qrykdex->NUMSEQ, Qrykdex->SEQCALC, Qrykdex->QUANT, Round(Qrykdex->CUSTO,aTamSD3[2]),Round((Qrykdex->CUSTO/Qrykdex->QUANT),aTamSD3[2]),SPACE(10),SPACE(10),SPACE(10),nFecMovQt,nFecMovVl})

				//Grava variáveis para apresentar os valores em tela
				If SUBSTRING(Qrykdex->CF,1,2) = 'PR'
					nQtPSD3 := nQtPSD3 + Qrykdex->QUANT
					nVlPSD3 := nVlPSD3 + Round(Qrykdex->CUSTO,aTamSD3[2])
				ElseIf SUBSTRING(Qrykdex->CF,1,2) = 'DE'
					nQtDSD3 := nQtDSD3 + Qrykdex->QUANT
					nVlDSD3 := nVlDSD3 + Round(Qrykdex->CUSTO,aTamSD3[2])
				ElseIf SUBSTRING(Qrykdex->CF,1,1) <= '5'
					nQtSD1  := nQtSD1 + Qrykdex->QUANT
					nVlSD1  := nVlSD1 + Round(Qrykdex->CUSTO,aTamSD3[2])
				EndIf
				// Pega a quanitdade do aramzém para prossessar o custo do  armazém mesmo que o custo seja por filial
				IF SUBSTR(cLocal,1,2) <> "**" .AND. lCusfilF .AND. Qrykdex->ARMAZEM == cLocal
					nQuantArm += Qrykdex->QUANT
				ENDIF

			ElseIf SUBSTRING(Qrykdex->CF,1,2) = 'RE' .Or. SUBSTRING(Qrykdex->CF,1,1) >= '5'
					nFecMovQt -= Qrykdex->QUANT // Quantidade Fechamento x Movimento
					nFecMovVl -= Round(Qrykdex->CUSTO,TAMSX3("B2_VFIM1")[2]) // Valor Fechamento x Movimento
				AAdd(aKarLocal,{STOD(Qrykdex->Emissao), Qrykdex->CF,Qrykdex->Armazem, Qrykdex->NUMSEQ, Qrykdex->SEQCALC,SPACE(10),SPACE(10),SPACE(10),Qrykdex->QUANT,Round(Qrykdex->CUSTO,aTamSD3[2]),Round((Qrykdex->CUSTO/Qrykdex->QUANT),aTamSD3[2]),nFecMovQt,nFecMovVl})

				//Grava variáveis para apresentar os valores em tela
				If SUBSTRING(Qrykdex->CF,1,2) = 'RE'
					nQtRSD3 := nQtRSD3 + Qrykdex->QUANT
					nVlRSD3 := nVlRSD3 + Round(Qrykdex->CUSTO,aTamSD3[2])
				ElseIf SUBSTRING(Qrykdex->CF,1,1) >= '5'
					nQtSD2  := nQtSD2 + Qrykdex->QUANT
					nVlSD2  := nVlSD2 + Round(Qrykdex->CUSTO,aTamSD3[2])
				EndIf
				// Pega a quanitdade do aramzém para prossessar o custo do  armazém mesmo que o custo seja por filial
				IF SUBSTR(cLocal,1,2) <> "**" .AND. lCusfilF .AND. Qrykdex->ARMAZEM == cLocal
					nQuantArm -= Qrykdex->QUANT
				ENDIF
			EndIf
			DBSkip()

		ENDDO

		Qrykdex->(DBCloseArea())

		IF SUBSTR(cLocal,1,2) <> "**" .AND. lCusfilF
			nFecMovVl := (nFecMovVl / nFecMovQt) * nQuantArm
			nFecMovQt := nQuantArm//cQtdMov
		ENDIF

		// Chama a função CalcCusOP para buscar o saldo das OPs
		aCusOp := CalcCusOP(cProduto,cProduto,cLocal, cLocal, dDataI,dDataF)

    	// LAÇO DE REPETIÇÃO PARA MONTAR O ARRAY DO CABEÇALHO
		FOR nCont := 1 TO LEN(aCusOp)
			nVlOpSB9    += aCusOp[nCont][2]
			nFecVlOpSB2 += aCusOp[nCont][3]
			nFecMovOp   += aCusOp[nCont][4]

		NEXT nCont

		// Saldo Fechamento x Movimento
		IncProc("Processando Saldo Fechamento x Movimento")
		DBSelectArea("SB2")
		DBSetOrder(1)

		cQuery := "SELECT B2_QFIM AS QFIM, B2_VFIM1 AS VFIM FROM "+RetSqlName("SB2")+" "

		cQuery += "WHERE B2_COD = '"+cProduto+"' AND D_E_L_E_T_ = ' ' "

		IF SUBSTR(cLocal,1,2) <> "**"
			cQuery += "AND B2_LOCAL = '"+cLocal+"' "
			IF !lCusfilE
				// custo por armazém ou filial
				cQuery += "AND B2_FILIAL = '"+xFilial("SB2")+"'"

			ENDIF
		ELSE
			IF !lCusfilE
				// custo por armazém ou filial
				cQuery += "AND B2_FILIAL = '"+xFilial("SB2")+"'"
			ENDIF
		ENDIF

		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QrySB2",.F.,.T.)

		WHILE QrySB2->(!Eof())
			nFecQtSB2 += QrySB2->QFIM
			nFecVlSB2 += QrySB2->VFIM
			lFound := .T.
			DBSKIP()
		ENDDO
		QrySB2->(DBCloseArea())
		IF !lFound
			MsgStop("Não foram encontrados registros para os parâmetros informados. Verifique!")
			Return .F.
		ENDIF
		// Processa as validações do fechamento
		MATCValFec()

		// Apresenta mensagem com resultado do fechamento
		If (ROUND(nFecMovQt,TAMSX3("B2_QATU")[2]) <> ROUND(nFecQtSB2,TAMSX3("B2_QATU")[2])) .Or. (ROUND(nFecMovVl,TAMSX3("B2_VATU1")[2]) <> ROUND(nFecVlSB2,TAMSX3("B2_VATU1")[2])) .or. (ROUND(nFecVlOpSB2,TAMSX3("B2_VATU1")[2]) <> ROUND(nFecMovOp,TAMSX3("B2_VATU1")[2]))
			@ 115,201 SAY oSay VAR "DIVERGÊNCIA" SIZE 200,50 PIXEL OF oFolder2:aDialogs[1] FONT oBold2 COLOR CLR_HRED
		Else
			@ 115,201 SAY oSay VAR "OK" SIZE 200,50 PIXEL OF oFolder2:aDialogs[1] FONT oBold2 COLOR CLR_GREEN
		EndIf

		// Atualiza objetos da tela
		oQtSB9:Refresh()
		oVlSB9:Refresh()
		oVlOpSB9:Refresh()
		oQtSD1:Refresh()
		oVlSD1:Refresh()
		oQtSD2:Refresh()
		oVlSD2:Refresh()
		oQtPSD3:Refresh()
		oVlPSD3:Refresh()
		oVlOpPSD3:Refresh()
		oQtRSD3:Refresh()
		oVlRSD3:Refresh()
		oVlOpRSD3:Refresh()
		oQtDSD3:Refresh()
		oVlDSD3:Refresh()
		oVlOpDSD3:Refresh()
		oFecQtSB2:Refresh()
		oFecVlSB2:Refresh()
		oFecVlOpSB2:Refresh()
		oFecMovQt:Refresh()
		oFecMovVl:Refresh()
		oFecMovOp:Refresh()
		oListCusto:Refresh()
		//oVlContab:Refresh()
		//oVlCT2:Refresh()

	ENDIF

Return


//------------------------------------------------------------------------------------------
/* {Protheus.doc} GetDatas
Painel04 - Folder4 - GetDatas

@Param nOpc - Opção para processamento do Saldo Inicial

@author    Ronaldo Tapia
@version   12.1.17
@since     13/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function GetDatas(nOpc)

	Local cQuery   := ""
	Local aArea    := GetArea()

	Default nOpc   := 1

	dDataI := CTOD("  /  /    ")

	cQuery := "SELECT MAX(B9_DATA) AS B9_DATA FROM "+RETSQLNAME('SB9')+" WHERE B9_FILIAL='"+xFilial("SB9")+"' AND D_E_L_E_T_ <> '*'
	If nOpc == 1
		IF substr(cLocal,1,2) <> "**"
			cQuery += " AND B9_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' "

		ENDIF
		cQuery += " AND B9_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"' AND B9_DATA < '"+DTOS(dDtProces)+"'"

	EndIf
	cQuery := ChangeQuery(cQuery)
	DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QrySB9",.F.,.T.)
	TCSETFIELD( "QrySB9","B9_DATA","D")
	dUltFech := QrySB9->B9_DATA

	If nOpc == 2
		dFechAn := QrySB9->B9_DATA
		dIniAn  := (dUltFech + 1)
	EndIf

	// Fixo Data Inicial
	If !Empty(dUltFech)
		dDataI := (dUltFech + 1)
	Else
		// Se não encontrar a data do último fechamento pega do MV_ULMES, validação para produtos que tiveram saldo inicial com campo B9_DATA em branco
		dDataI := SuperGetMV("MV_ULMES",.F.,"19961231",XFILIAl('SB9')) + 1
		//MsgAlert("Não foi realizada a virada de saldo para o produto " + cProduto + ". Verifique!")
	EndIf

	QrySB9->(DBCloseArea())

	RestArea(aArea)

Return(.T.)


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCValFec
Painel04 - Folder4 - MATCValFec
Processa as validações do fechamento
@Param

@author    Ronaldo Tapia
@version   12.1.17
@since     13/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCValFec()

	Local aArea    := GetArea()
	Local cQuery   := ""
	Local QryFSC2  := ""
	Local lValid   := .T.
	Local nLin     := 105

	If oBtErro2 <> Nil .And. !Empty(oBtErro2:Ctitle)
		oBtErro2:Ctitle := ""
		oBtErro2:Refresh()
	EndIf

	DEFINE FONT oBold   NAME "Arial" SIZE 0, -12 BOLD

	DBSelectArea("SB2")
	DBSetOrder(1)
	If DBSeek(xFilial("SB2")+cProduto+cLocal)
		//Valida valores negativos
		If (B2_VFIM1 < 0) .Or. (B2_QFIM < 0)
			@ nLin,384 SAY oBtErro2 VAR "Fechamento com valores negativos! Verifique o campo Vlr. Final (B2_VFIM1)." SIZE 300,50 PIXEL OF oFolder2:aDialogs[1] FONT oBold COLOR CLR_HRED
			nLin := nLin + 15
			lValid := .F.
		EndIf

		//Valida custo sem quantidade
		If (B2_VFIM1 > 0) .And. (B2_QFIM <= 0)
			@ nLin,384 SAY oBtErro2 VAR "Fechamento sem quantidade! Verifique o campo Qtd. Fim Mes (B2_QFIM)." SIZE 300,50 PIXEL OF oFolder2:aDialogs[1] FONT oBold COLOR CLR_HRED
			nLin := nLin + 15
			lValid := .F.
		EndIf
	EndIf

	//Valida saldo em processo com OP encerrada
	cQuery := " SELECT C2_TPOP, C2_DATRF, C2_QUJE, C2_QUANT, C2_VFIM1, C2_APRFIM1 FROM "+RETSQLNAME("SC2")
	cQuery += " WHERE C2_FILIAL='"+xFilial("SC2")+"' AND "+RETSQLNAME("SC2")+".D_E_L_E_T_ <> '*' AND C2_PRODUTO = '"+cProduto+"'"
	IF substr(cLocal,1,2) <> "**"
				cQuery += " AND C2_LOCAL = '"+cLocal+"' "
			ENDIF
	cQuery += " AND C2_EMISSAO >= '"+DTOS(dDataI)+"' AND C2_EMISSAO <= '"+DTOS(dDataF)+"'"
	cQuery := ChangeQuery(cQuery)
	DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QryFSC2",.F.,.T.)

	While QryFSC2->(!Eof())
		If QryFSC2->C2_TPOP == "F" .And. !Empty(QryFSC2->C2_DATRF) .And. QryFSC2->(C2_QUJE >= C2_QUANT) // OP Encerrada totalmente
			If QryFSC2->C2_VFIM1 > 0 //.Or. QryFSC2->C2_APRFIM1 > 0
				@ nLin,384 SAY oBtErro2 VAR "OP Encerrada com saldo em processo em aberto, verifique!" SIZE 300,50 PIXEL OF oFolder2:aDialogs[1] FONT oBold COLOR CLR_HRED
				nLin := nLin + 15
				lValid := .F.
			EndIf
		EndIf
		DBSkip()
	Enddo
	QryFSC2->(DBCloseArea())

	If !lValid
		@ nLin,384 SAY oBtErro PROMPT "<u>" + "Clique aqui para mais informações de como realizar a correção para o fechamento de custos." + "</u>" SIZE 250,010 OF oFolder2:aDialogs[1] HTML PIXEL
		oBtErro:bLClicked := {|| ShellExecute("open","http://tdn.totvs.com/x/NYJJF","","",1) }
	EndIf

	RestArea(aArea)

Return

//------------------------------------------------------------------------------------------
/* {Protheus.doc} MACTFecha
Fecha tela do Monitor de Custos

@Param

@author    Ronaldo Tapia
@version   12.1.17
@since     19/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MACTFecha()

	oDlg:End()

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCBrwCus
Função para mostrar dados do custo no browser

@Param
oPanel  - objeto da tela
aCols   - array com dados do browser
nRow    - posição incial para criação do browser
nHeight - altura do objeto do browser
lCusto  - indica se exibe o browser para validação do custo

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCBrwCus(oPanel,aCols,nRow,nHeight)

	Local Nx 		:= 0
	Local blDblClick:= ""
	Local bLine		:= ""
	Local aColSizes	:= {}
	Local aHeader   := {}
	Local oListBox  := Nil

	Default oPanel  := Nil
	Default aCols	:= {}
	Default nRow	:= 001
	Default nHeight	:= 190

	If Empty(bLine)
		nTamCol := Len(aCols[01])
		bLine 	:= "{|| {"
		For Nx := 1 To nTamCol
			bLine += "aCols[oListCusto:nAt]["+StrZero(Nx,3)+"]"
			If Nx < nTamCol
				bLine += ","
			EndIf
		Next
		bLine += "} }"
	EndIf

	aColSizes := {10,20,20,57,32,70,55,50,65,50,50,60,50}
	If lCusfilA
		aHeader   := {'Data','CF','Local','NUMSEQ','SEQCAL','Entradas - Quantidade','Entradas - Custo','CM Movimento','Saídas - Quantidade','Saídas - Custo','CM Armazem','Saldo - Quantidade','Saldo - Custo'}
	ElseIf lCusfilF
		aHeader   := {'Data','CF','Local','NUMSEQ','SEQCAL','Entradas - Quantidade','Entradas - Custo','CM Movimento','Saídas - Quantidade','Saídas - Custo','CM Filial','Saldo - Quantidade','Saldo - Custo'}
	ElseIf lCusfilE
		aHeader   := {'Data','CF','Local','NUMSEQ','SEQCAL','Entradas - Quantidade','Entradas - Custo','CM Movimento','Saídas - Quantidade','Saídas - Custo','CM Empresa','Saldo - Quantidade','Saldo - Custo'}
	EndIf
	oListCusto := TCBrowse():New(nRow,001,630,nHeight,,aHeader,,oPanel,'Campos',,,,,,,,,,,,,,,,,.T.,.T.)

	// Define os dados do browser
	oListCusto:SetArray( aCols )
	oListCusto:bLine := &bLine
	If !Empty( aColSizes )
		oListCusto:aColSizes := aColSizes
	EndIf
	oListCusto:SetFocus()

	// Instancia objeto CSS
	oListCusto:SetCss(cCSSTCBrw)

	oListCusto:Refresh()

Return()

//------------------------------------------------------------------------------------------
/* {Protheus.doc} MACTLDetail
Função para atualizar o Log de Processamento ao selecionar o ID

@Param
oPanel  - Objeto da tela
cID     - ID do processo
oLBDetalhe - Detalhes do Log

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MACTLDetail(oPanel,cID,oLBDetalhe)
	Local cAliasCV8 := ''
	Local cQuery    := ''
	Local aAreaAnt  := GetArea()
	Local cQRotina  := Iif(Type('cWRotina')=="C",cWRotina,'')

	// Carrega o Log de processamento
	cAliasCV8 := GetNextAlias()
	cQuery := "SELECT CV8_IDMOV, CV8_FILIAL, CV8_DATA, CV8_HORA, CV8_MSG, CV8_INFO, CV8_SBPROC, R_E_C_N_O_ RECNOCV8 "
	cQuery +=  " FROM "+RetSqlName("CV8")+" CV8 "
	cQuery += " WHERE CV8_IDMOV = '" + cID + "' "
	cQuery +=   " AND CV8_PROC = 'MATA330' "
	cQuery +=   " AND D_E_L_E_T_ = ' ' "
	cQuery +=   " ORDER BY  CV8_IDMOV,R_E_C_N_O_,CV8_PROC,CV8_USER,CV8_DATA,CV8_HORA "
	cQuery := ChangeQuery(cQuery)
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasCV8,.T.,.T.)
	aArrayDET   := {{'','','','','','',''}}
	aEval(CV8->(dbStruct()), {|x| If(x[2] <> "C", TcSetField(cAliasCV8,x[1],x[2],x[3],x[4]),Nil)})
	Do While !Eof()
		aAdd(aArrayDet,{(cAliasCV8)->CV8_IDMOV,(cAliasCV8)->CV8_FILIAL,(cAliasCV8)->CV8_DATA,(cAliasCV8)->CV8_HORA,(cAliasCV8)->CV8_MSG,(cAliasCV8)->CV8_SBPROC,(cAliasCV8)->CV8_INFO,STRZERO((cAliasCV8)->RECNOCV8,15)})
		(cAliasCV8)->(dbSkip())
	EndDo

	// Remove a linha em branco
	If Len(aArrayDET) >= 2
		If Empty(aArrayDET[1,1])
			Adel(aArrayDET,1)
			ASize(aArrayDET,Len(aArrayDET)-1)
		EndIf
	EndIf

	// Montagem do detalhe do Log
	If !Empty(cID) .And. Len(aArrayDET) > 0
		oLBDetalhe:SetArray(aArrayDET)
		oLBDetalhe:bLine := { || {aArrayDET[oLBDetalhe:nAT][1],aArrayDET[oLBDetalhe:nAT][2],aArrayDET[oLBDetalhe:nAT][3],aArrayDET[oLBDetalhe:nAT][4],aArrayDET[oLBDetalhe:nAT][5],aArrayDET[oLBDetalhe:nAT][6],aArrayDET[oLBDetalhe:nAT][7]} }
		oLBDetalhe:aColSizes := {TamSX3("CV8_IDMOV")[1],TamSX3("CV8_FILIAL")[1]+20,8,10,1,TamSX3("CV8_MSG")[1],TamSX3("CV8_SBPROC")[1],15}
		oLBDetalhe:Refresh()
	EndIf

	// Encerra o alias temporario
	If Select(cAliasCV8) > 0
		(cAliasCV8)->(dbCloseArea())
	EndIf

	RestArea(aAreaAnt)

Return


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCLMemo
Função para leitura do campo MEMO (Tabela CV8)

@Param
oPanel  - Objeto da tela
nLinha  - Posição da pesquisa

@author    Ronaldo Tapia
@version   12.1.17
@since     11/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCLMemo(oPanel, nLinha)
	Local oDlgMemo	:= Nil
	Local oBotOk	:= Nil
	Local nrecnoCV8 := 0
	Local oMemo, mMemo

	// Carrega os dados do campo MEMO (CV8_DET)
	If Type('aArrayDet')=="A" .And. Len(aArrayDet)>=nLinha
		nRecnoCV8 := Val(aArrayDet[nLinha,8])
		dbSelectArea("CV8")
		dbGoto(nRecnoCV8)
		mMemo := CV8->CV8_DET
	EndIf

	DEFINE DIALOG oDlgMemo FROM  005,005 TO 300,370 OF oPanel TITLE "Detalhes do Processamento" PIXEL	//"Detalhes do Processamento"

	//Campo memo
	@ 00,00 GET oMemo VAR mMemo OF oDlgMemo MEMO SIZE 185,148 FONT oDlgMemo:oFont

	oMemo:bRClicked := {||AllwaysTrue()}

	ACTIVATE DIALOG oDlgMemo CENTER

Return

//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCSCSS
Função para setar as propriedades dos métodos CSS

@Param

@author    Ronaldo Tapia
@version   12.1.17
@since     26/09/2018

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCSCSS()

	// Atribui propriedados do CSS ao TFolder
	cCSSTFol +=	BR+"QTabBar::tab";
	+BR+"{";
	+BR+"  color: #FFFFFF; /*cor da fonte*/";
	+BR+"  font-size: 14px; /*tamanho da fonte*/";
	+BR+"  background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,";
	+BR+"                     stop: 0 #465A91, stop: 1 #75A6D6); /*Cor de fundo*/";
	+BR+"  border: 1px solid #465A91; /*cor da borda*/";
	+BR+"  border-radius: 10px; /*arredondamento borda*/";
	+BR+"  min-width: 150px; /*largura minima botao*/";
	+BR+"  min-height: 14px; /*altura minima botao*/";
	+BR+"  padding: 2px; /*margin para arredondamento da borda*/";
	+BR+"}";
	+BR+"/* Acoes quando a aba for selecionada */";
	+BR+"QTabBar::tab:selected{";
	+BR+"  background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,";
	+BR+"                     stop: 0 #75A6D6, stop: 1 #75A6D6); /*Cor de fundo*/";
	+BR+"  border-bottom: none; /*cor da borda*/";
	+BR+"}";
	+BR+"/* Acoes quando a aba nao estiver selecionada */";
	+BR+"QTabBar::tab:!selected{";
	+BR+"  margin-top: 2px; /*margem superior*/";
	+BR+"}"

	// Atribui propriedados do CSS ao TCBrowse
	cCSSTCBrw += BR+"/* Componentes que herdam da TCBrowse:";
	+BR+"   BrGetDDb, MsBrGetDBase,MsSelBr, TGridContainer, TSBrowse, TWBrowse */";
	+BR+"QTableWidget {";
	+BR+"  gridline-color: #a9acad; /*Cor da grade*/";
	+BR+"  color: #000000; /*Cor da fonte*/";
	+BR+"  font-size: 11px; /*Tamanho da fonte*/";
	+BR+"  background: #FFFFFF; /*Cor do fundo da Grid*/";
	+BR+"  alternate-background-color: #e5e8ea; /*Cor do zebrado*/";
	+BR+"  selection-background-color: qlineargradient(x1: 0, y1: 0, x2: 0.8, y2: 0.8,";
	+BR+"                              stop: 0 #75A6D6, stop: 1 #75A6D6); /*Cor da linha selecionada*/";
	+BR+"}";
	+BR+"/* Acoes quando a celula for selecionada, aqui mudo a cor de fundo */";
	+BR+"QTableView:item:selected:focus {";
	+BR+"  background-color: #75A6D6} /*Cor da celula selecionada*/"

	// Atribui propriedados do CSS ao TButton
	cCSSTBut +=	BR+"QPushButton {";
	+BR+"  color: #FFFFFF; /*Cor da fonte*/";
	+BR+"  border: 2px solid #465A91; /*Cor da borda*/";
	+BR+"  border-radius: 6px; /*Arrerondamento da borda*/";
	+BR+"  background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,";
	+BR+"                                    stop: 0 #465A91, stop: 1 #75A6D6); /*Cor de fundo*/";
	+BR+"  min-width: 80px; /*Largura minima*/";
	+BR+"}";
	+BR+"/* Acoes quando pressionado botao, aqui mudo a cor de fundo */";
	+BR+"QPushButton:pressed {";
	+BR+"  background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,";
	+BR+"                                    stop: 0 #75A6D6, stop: 1 #75A6D6);";
	+BR+"}"

Return


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCPrDiv
Browser para listar os produtos com divergencia entre Fechamento: SB2 x Fechamento: Movimento

@Param
aHeader   - aHeader do Grid
aCols     - aCols do Grid

@author    Ronaldo Tapia
@version   12.1.17
@since     27/09/2018
@protected

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCPrDiv()

	Local Nx 		:= 0
	Local oListBox	:= nil
	Local oArea		:= Nil
	Local oValorN	:= TFont():New("Arial",08,14,,.T.,,,,.T.)
	Local oList		:= Nil
	Local oFather	:= Nil
	Local aCoord	:= {}
	Local aWindow	:= {}
	Local aColumn	:= {}
	Local oBtProces, oBtSair, oButExp := Nil
	Local aColSizes := {}
	Local bLine     := ""


	DEFINE FONT oBold   NAME "Arial" SIZE 0, -12 BOLD

	//Ajusta ao objeto pai da tela
	aCoord 	:= {000,000,500,800}
	oList	:= oFather
	//ListBox(aProdAn)

	nTamCol := Len(aProdAn[01])
	bLine 	:= "{|| {"
	For Nx := 1 To nTamCol
		bLine += "aProdAn[oListBox:nAt]["+StrZero(Nx,3)+"]"
		If Nx < nTamCol
			bLine += ","
		EndIf
	Next
	bLine += "} }"


	// Cria o objeto tipo tDialog
	DEFINE MSDIALOG oFather TITLE "Produtos com Divergência no Fechamento" FROM aCoord[1],aCoord[2] TO aCoord[3],aCoord[4] COLORS 0, 16777215 PIXEL

	// Cria o objeto tipo FWLayer
	oArea := FWLayer():New()
	oArea:Init(oFather,.F.)

	// Layer 1 - Analise de Divergências
	oArea:Init(oFather,.F., .F. )
	oArea:AddLine( "LINE01", 100 )
	oArea:AddCollumn( "BOX01", 99,, "LINE01" )
	oArea:AddWindow( "BOX01", "PANEL01", "Analise de Produtos", 95, .F.,,, "LINE01" ) //Analise de Produtos
	oList := oArea:GetWinPanel("BOX01","PANEL01","LINE01")

	@ 006,005 SAY "Do Produto"                       SIZE 040,10 PIXEL OF oList FONT oBold
	@ 002,050 MSGET oVar VAR cProdAnDe Picture "@!"  SIZE 060,10 PIXEL OF oList F3 "SB1" VALID .T.

	@ 026,005 SAY "Até o Produto"                    SIZE 040,10 PIXEL OF oList FONT oBold
	@ 022,050 MSGET oVar VAR cProdAnAte Picture "@!" SIZE 060,10 PIXEL OF oList F3 "SB1" VALID .T.

	@ 046,005 SAY "Do Local"                         SIZE 040,10 PIXEL OF oList FONT oBold
	@ 042,050 MSGET oVar VAR cLocAnDe Picture "@!"   SIZE 010,10 PIXEL OF oList VALID(GetDatas(2))

	@ 064,005 SAY "Até o Local"                      SIZE 040,10 PIXEL OF oList FONT oBold
	@ 060,050 MSGET oVar VAR cLocAnAte Picture "@!"  SIZE 010,10 PIXEL OF oList VALID(GetDatas(2))//VALID .T.

	@ 006,120 SAY "Último Fechamento"                SIZE 070,10 PIXEL OF oList FONT oBold
	@ 002,185 MSGET oVar VAR dFechAn  Picture "99/99/99" SIZE 040,10 PIXEL OF oList When .F.

	@ 026,120 SAY "Data Inicial"                     SIZE 070,10 PIXEL OF oList FONT oBold
	@ 022,185 MSGET oVar VAR dIniAn Picture "99/99/99" SIZE 040,10 PIXEL OF oList VALID .T.

	@ 046,120 SAY "Data Final"                       SIZE 070,10 PIXEL OF oList FONT oBold
	@ 042,185 MSGET oVar VAR dFIMAn Picture "99/99/99" SIZE 040,10 PIXEL OF oList VALID .T.

	oBtProces := TButton():New( 008, 260, "Processar",oList,{||MATCProcD(oListBox)}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )
	oBtSair   := TButton():New( 008, 320, "Sair",oList,{||oFather:End()}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )
	oButExp   := TButton():New( 028, 260, "&Exportar",oList,{|| MACTWizFon(oList,4) },40,12,,,.F.,.T.,.F.,,.F.,,,.F. )

	aHeader   := {'Filial','Produto','Armazem','Status'}
	aColSizes := {40,50,40,80}
	oListBox := TCBrowse():New(6,0,388,139,,aHeader,,oList,'Fonte')
	oListBox:SetArray(aProdAn)
	oListBox:bLine := &bLine
	oListBox:Refresh()

	If !Empty( aColSizes )
		oListBox:aColSizes := aColSizes
	EndIf
	oListBox:SetFocus()

	// Instancia objeto CSS
	oListBox:SetCss(cCSSTCBrw)
	oBtProces:SetCss(cCSSTBut)
	oBtSair:SetCss(cCSSTBut)
	oButExp:SetCss(cCSSTBut)

	//oListBox:DrawSelect()

	oFather:Activate(,,,.T.,/*valid*/,,/*On Init*/)

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCPCalc
Verifica se a divergência no valor de fechamento x movimento dos produtos informados

@Param
aHeader   - aHeader do Grid
aCols     - aCols do Grid

@author    Ronaldo Tapia
@version   12.1.17
@since     01/010/2018
@protected

@Return

*/
//------------------------------------------------------------------------------------------
Static Function MATCProcD(oListBox)

	Processa({||MATCPCalc(oListBox)})

Return

Static Function MATCPCalc(oListBox)

	Local cQuery      := ""
	Local cData       := "s"
	Local bLine       := ""
	Local QryTSD1An, QryTSD2An, QryTSD3An, QryTSD3AnP, QrySB9An, QrySB2An, QryTempDB, QrySB2FIL := ""
	Local aFields     := {}
	Local aSaldIni    := {}
	Local lRet		  := .T.
	Local nFecVlOpSB2 := 0
	LOcal nFecMovOp   := 0
	Local nValFin     := 0
	//Local oSaldIni    As object
	//Local oData       As object

	Default oListBox := Nil

	// Array com os campos utilizados na view da consulta de divergência de produtos
	Aadd(aFields, {"FILIAL","C",TamSX3("B9_FILIAL")[1],0})
	Aadd(aFields, {"COD","C",TamSX3("B9_COD")[1],0})
	Aadd(aFields, {"ARMAZEM","C",TamSX3("B9_LOCAL")[1],0})
	Aadd(aFields, {"CUSTO","N",TamSX3("B9_VINI1")[1],TamSX3("B9_VINI1")[2]})
	Aadd(aFields, {"QUANT","N",TamSX3("B9_QINI")[1],TamSX3("B9_QINI")[2]})

	// Validações
	If Empty(dIniAn) .OR. Empty(dFIMAn)
		MsgStop("Datas de Processamento não informadas. Verifique!")
		lRet := .F.
	EndIf

	If lRet .And. Empty(cProdAnAte)
		MsgStop("Produto não informado. Verifique!")
		lRet := .F.
	EndIf

	If lRet
		// Cria tabela temporario no banco de dados
		oTempTable := FWTemporaryTable():New( "MATCTmpAn" )
		oTemptable:SetFields( aFields )
		oTempTable:AddIndex("1", {"FILIAL","COD","ARMAZEM"} )
		oTempTable:Create()

		// Zero array de produtos divergentes para não ficar duplicado em um novo processo.
		aProdAn    := {{'','','',''}}
		GetDatas() //Grava a última data de fechamento na variável dUltFech

		// Saldo Inicial
		IncProc("Processando Saldo Inicial")

		cQuery := " SELECT B9_FILIAL, B9_COD, B9_LOCAL FROM "+RETSQLNAME("SB9")
		cQuery += " WHERE B9_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"'"
		If lCusfilA // Custo por Armazem
			cQuery += "  AND B9_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' AND B9_FILIAL = '"+xFILIAL("SB9")+"' "
		ElseIf lCusfilF // Custo por Filial
			cQuery += " AND B9_FILIAL = '"+xFILIAL("SB9")+"' "
		ElseIf lCusfilE // Custo por Empresa
			cQuery += " AND B9_DATA = '"+DTOS(dUltFech)+"'"
		EndIf
		cQuery += " AND "+RETSQLNAME("SB9")+".D_E_L_E_T_ = ' ' GROUP BY B9_FILIAL, B9_COD, B9_LOCAL"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QrySB9An",.F.,.T.)

		While QrySB9An->(!Eof())

			IF !IsProdMOD(QrySB9An->B9_COD) // Não processar produtos MOD / GGF
				aSaldIni := CalcEst(QrySB9An->B9_COD,QrySB9An->B9_LOCAL,dIniAn,xFILIAL("SB9"),.F.,.F.)
				RecLock( "MATCTmpAn", .T. )
				Replace MATCTmpAn->FILIAL	With	QrySB9An->B9_FILIAL
				Replace MATCTmpAn->COD    	With	QrySB9An->B9_COD
				Replace MATCTmpAn->ARMAZEM  With	QrySB9An->B9_LOCAL
				Replace MATCTmpAn->CUSTO	With	aSaldIni[2]
				Replace MATCTmpAn->QUANT	With	aSaldIni[1]
				//Replace MATCTmpAn->QUANT	With	QrySB9An->B9_DATA
				msUnlock()
			END
			QrySB9An->(DBSkip())
		Enddo
		QrySB9An->(DBCloseArea())

		// Saldo Entradas
		IncProc("Processando Saldo Entradas")
		cQuery := " SELECT D1_FILIAL, D1_COD, D1_LOCAL, SUM(D1_CUSTO) AS SD1_CUSTO, SUM(D1_QUANT) AS SD1_QUANT FROM "+RETSQLNAME("SD1")+", "+RETSQLNAME("SF4")
		cQuery += " WHERE D1_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"' "
		cQuery += " AND D1_ORIGLAN <> 'LF' AND D1_DTDIGIT >= '"+DTOS(dIniAn)+"' AND D1_DTDIGIT <= '"+DTOS(dFIMAn)+"'"
		cQuery += " AND D1_FILIAL = F4_FILIAL "
		cQuery += " AND D1_TES=F4_CODIGO AND F4_ESTOQUE = 'S' "
		//Complemento do Where da tabela SD1 (Tratamento Filial)
		If lCusfilE
			If FWModeAccess("SD1") == 'E' .AND. FWModeAccess("SF4") == 'E'
				cQuery += " AND D1_FILIAL = F4_FILIAL"
			EndIf
		ElseIf lCusfilF
			cQuery += " AND D1_FILIAL='"+xFilial("SD1")+"' "
		ElseIf lCusfilA
			cQuery += " AND D1_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' "
		EndIf
		cQuery += " AND "+RETSQLNAME("SD1")+".D_E_L_E_T_ = ' '"
		cQuery += " AND "+RETSQLNAME("SF4")+".D_E_L_E_T_ = ' '"
		cQuery += " GROUP BY D1_FILIAL,D1_COD,D1_LOCAL"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QryTSD1An",.F.,.T.)
		While QryTSD1An->(!Eof())
			RecLock( "MATCTmpAn", .T. )
			Replace MATCTmpAn->FILIAL	With	QryTSD1An->D1_FILIAL
			Replace MATCTmpAn->COD    	With	QryTSD1An->D1_COD
			Replace MATCTmpAn->ARMAZEM  With	QryTSD1An->D1_LOCAL
			Replace MATCTmpAn->CUSTO	With	QryTSD1An->SD1_CUSTO
			Replace MATCTmpAn->QUANT	With	QryTSD1An->SD1_QUANT
			msUnlock()
			QryTSD1An->(DBSkip())
		Enddo
		QryTSD1An->(DBCloseArea())

		// Saldo Movimentações Internas - Requisição
		IncProc("Processando Saldo Movimentações Internas - Requisição")
		cQuery := " SELECT D3_FILIAL, D3_COD, D3_LOCAL, SUM(D3_CUSTO1) AS SD3_CUSTO_R, SUM(D3_QUANT) AS SD3_QUANT_R FROM "+RETSQLNAME("SD3")
		cQuery += " WHERE D3_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"' "
		cQuery += " AND D3_ESTORNO=' ' AND D3_EMISSAO >= '"+DTOS(dIniAn)+"' AND D3_EMISSAO <= '"+DTOS(dFIMAn)+"'"
		//Complemento do Where da tabela SD3 (Tratamento Filial)
		If lCusfilF
			cQuery += " AND D3_FILIAL='"+xFilial("SD3")+"' "
		ElseIf lCusfilA
			cQuery += " AND D3_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' "
		EndIf
		cQuery += " AND D3_CF IN('RE0','RE1','RE2','RE3','RE4','RE5','RE6','RE7')"
		cQuery += " AND "+RETSQLNAME("SD3")+".D_E_L_E_T_ = ' '"
		cQuery += "GROUP BY D3_FILIAL,D3_COD,D3_LOCAL"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QryTSD3AnP",.F.,.T.)

		While QryTSD3AnP->(!Eof())
			IF !IsProdMOD(QryTSD3AnP->D3_COD) // Não processar produtos MOD / GGF
				RecLock( "MATCTmpAn", .T. )
				Replace MATCTmpAn->FILIAL	With	QryTSD3AnP->D3_FILIAL
				Replace MATCTmpAn->COD    	With	QryTSD3AnP->D3_COD
				Replace MATCTmpAn->ARMAZEM  With	QryTSD3AnP->D3_LOCAL
				Replace MATCTmpAn->CUSTO	With	-(QryTSD3AnP->SD3_CUSTO_R)
				Replace MATCTmpAn->QUANT	With	-(QryTSD3AnP->SD3_QUANT_R)
				msUnlock()
			END
			QryTSD3AnP->(DBSkip())
		Enddo
		QryTSD3AnP->(DBCloseArea())

		// Saldo Movimentações Internas - Producao e Devolução
		IncProc("Processando Saldo Movimentações Internas - Producao")
		cQuery := " SELECT D3_FILIAL, D3_COD, D3_LOCAL, SUM(D3_CUSTO1) AS SD3_CUSTO, SUM(D3_QUANT) AS SD3_QUANT FROM "+RETSQLNAME("SD3")
		cQuery += " WHERE D3_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"' "
		cQuery += " AND D3_ESTORNO=' ' AND D3_EMISSAO >= '"+DTOS(dIniAn)+"' AND D3_EMISSAO <= '"+DTOS(dFIMAn)+"'"
		//Complemento do Where da tabela SD3 (Tratamento Filial)
		If lCusfilF
			cQuery += " AND D3_FILIAL='"+xFilial("SD3")+"' "
		ElseIf lCusfilA
			cQuery += " AND D3_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' "
		EndIf
		cQuery += " AND D3_CF IN('PR0','PR1','DE0','DE1','DE2','DE3','DE4','DE5','DE6','DE7') "
		cQuery += " AND "+RETSQLNAME("SD3")+".D_E_L_E_T_ = ' '"
		cQuery += "GROUP BY D3_FILIAL,D3_COD,D3_LOCAL"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QryTSD3An",.F.,.T.)

		While QryTSD3An->(!Eof())
			IF !IsProdMOD(QryTSD3An->D3_COD) // Não processar produtos MOD / GGF
				RecLock( "MATCTmpAn", .T. )
				Replace MATCTmpAn->FILIAL	With	QryTSD3An->D3_FILIAL
				Replace MATCTmpAn->COD    	With	QryTSD3An->D3_COD
				Replace MATCTmpAn->ARMAZEM  With	QryTSD3An->D3_LOCAL
				Replace MATCTmpAn->CUSTO	With	QryTSD3An->SD3_CUSTO
				Replace MATCTmpAn->QUANT	With	QryTSD3An->SD3_QUANT
				msUnlock()
			END
			QryTSD3An->(DBSkip())
		Enddo
		QryTSD3An->(DBCloseArea())

		// Saldo Saídas
		IncProc("Processando Saldo Saídas")
		cQuery := " SELECT D2_FILIAL, D2_COD, D2_LOCAL, SUM(D2_CUSTO1) AS SD2_CUSTO, SUM(D2_QUANT) AS SD2_QUANT FROM "+RETSQLNAME("SD2")+", "+RETSQLNAME("SF4")
		cQuery += " WHERE D2_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"' "
		cQuery += " AND D2_ORIGLAN <> 'LF' AND D2_EMISSAO >= '"+DTOS(dIniAn)+"' AND D2_EMISSAO <= '"+DTOS(dFIMAn)+"'"
		cQuery += " AND D2_TES=F4_CODIGO AND F4_ESTOQUE = 'S' "
		//Complemento do Where da tabela SD2 (Tratamento Filial)
		If lCusfilE
			If FWModeAccess("SD2") == 'E' .AND. FWModeAccess("SF4") == 'E'
				cQuery += " AND D2_FILIAL = F4_FILIAL"
			EndIf
		ElseIf lCusfilF
			cQuery += " AND D2_FILIAL='"+xFilial("SD2")+"' "
		ElseIf lCusfilA
			cQuery += " AND D2_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' "
		EndIf
		cQuery += " AND "+RETSQLNAME("SD2")+".D_E_L_E_T_ = ' '"
		cQuery += " AND "+RETSQLNAME("SF4")+".D_E_L_E_T_ = ' '"
		cQuery += " GROUP BY D2_FILIAL,D2_COD,D2_LOCAL"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QryTSD2An",.F.,.T.)
		While QryTSD2An->(!Eof())

			RecLock( "MATCTmpAn", .T. )
			Replace MATCTmpAn->FILIAL	With	QryTSD2An->D2_FILIAL
			Replace MATCTmpAn->COD    	With	QryTSD2An->D2_COD
			Replace MATCTmpAn->ARMAZEM  With	QryTSD2An->D2_LOCAL
			Replace MATCTmpAn->CUSTO	With	-(QryTSD2An->SD2_CUSTO)
			Replace MATCTmpAn->QUANT	With	-(QryTSD2An->SD2_QUANT)
			msUnlock()

			QryTSD2An->(DBSkip())
		Enddo
		QryTSD2An->(DBCloseArea())

		// Saldo Fechamento x Movimento
		IncProc("Processando Saldo Fechamento")
		cQuery := " SELECT B2_FILIAL, B2_COD, B2_LOCAL, B2_VFIM1, B2_QFIM FROM "+RETSQLNAME("SB2")
		cQuery += " WHERE B2_COD BETWEEN '"+cProdAnDe+"' AND '"+cProdAnAte+"' "
		If lCusfilA // Custo por Armazem
			cQuery += "  AND B2_LOCAL BETWEEN '"+cLocAnDe+"' AND '"+cLocAnAte+"' AND B2_FILIAL = '"+xFilial("SB2")+"' "
		ElseIf lCusfilF  // Custo por Filial
			cQuery += " AND B2_FILIAL='"+xFilial("SB2")+"' "
		EndIf
		cQuery += " AND "+RETSQLNAME("SB2")+".D_E_L_E_T_ = ' '"
		cQuery := ChangeQuery(cQuery)
		DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QrySB2An",.F.,.T.)
		IncProc("Processando Saldo Fechamento x Movimento")
		While QrySB2An->(!Eof())

			// Faz um select na tabela temporaria para comparar os valores das movimentações com o valor do fechamento (SB2)
			cQuery := " SELECT TMP.FILIAL, TMP.COD, TMP.ARMAZEM, SUM(TMP.CUSTO) AS CUSTO, SUM(TMP.QUANT) AS QUANT FROM "+oTempTable:GetRealName()+" TMP"
			cQuery += " WHERE TMP.FILIAL = '"+QrySB2An->B2_FILIAL+"' AND TMP.COD = '"+QrySB2An->B2_COD+"' AND TMP.ARMAZEM = '"+QrySB2An->B2_LOCAL+"' "
			cQuery += " AND TMP.D_E_L_E_T_ = ' '"
			cQuery += " GROUP BY TMP.FILIAL, TMP.COD, TMP.ARMAZEM"
			cQuery := ChangeQuery(cQuery)
			DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QryTempDB",.F.,.T.)
			While QryTempDB->(!Eof())

				//Calcula o custo das OPs
				aCusOp := CalcCusOP(QryTempDB->COD, QryTempDB->COD, QryTempDB->ARMAZEM, QryTempDB->ARMAZEM, dDataI, dFIMAn)
				nFecVlOpSB2 := 0
				nFecMovOp   := 0
				aCusOp := {}
    			//For para montar o custo das OPs do produto
				FOR nCont := 1 TO LEN(aCusOp)
					nFecVlOpSB2 += aCusOp[nCont][3]
					nFecMovOp   += aCusOp[nCont][4]
				NEXT nCont

				//Mesmo o custo sendo por filial faz o cálculo do custo de apenas 1 armazém
				IF SuperGetMv("MV_CUSFIL",.F.,"A") == "F"
					cQuery := " SELECT TMP.FILIAL, TMP.COD, SUM(TMP.CUSTO) AS CUSTO, SUM(TMP.QUANT) AS QUANT FROM "+oTempTable:GetRealName()+" TMP"
					cQuery += " WHERE TMP.FILIAL = '"+QrySB2An->B2_FILIAL+"' AND TMP.COD = '"+QrySB2An->B2_COD+"'"
					cQuery += " AND TMP.D_E_L_E_T_ = ' '"
					cQuery += " GROUP BY TMP.FILIAL, TMP.COD"
					cQuery := ChangeQuery(cQuery)
					DBUseArea(.T.,"TOPCONN",TCGENQRY(,,cQuery),"QrySB2FIL",.F.,.T.)
					//cQuery += "AND TMP.ARMAZEM = '"+QrySB2FIL->B2_LOCAL+"'
					nValFin := ROUND((QrySB2FIL->CUSTO / QrySB2FIL->QUANT) * QryTempDB->QUANT, TAMSX3("B2_VATU1")[2])
					QrySB2FIL->(DBCloseArea())
				ELSE
					nValFin := Round(QryTempDB->CUSTO,TAMSX3("B2_VATU1")[2])
				END

				If (nValFin <> Round(QrySB2An->B2_VFIM1,TAMSX3("B2_VATU1")[2])) .Or. (QryTempDB->QUANT <> QrySB2An->B2_QFIM) .or. (Round(nFecVlOpSB2,TAMSX3("B2_QATU")[2]) <> Round(nFecMovOp,TAMSX3("B2_QATU")[2]))
					AAdd(aProdAn,{QrySB2An->B2_FILIAL,QrySB2An->B2_COD,QrySB2An->B2_LOCAL,"DIVERGENTE"})
				EndIf
				//oListBox:Refresh()
				QryTempDB->(DBSkip())
			Enddo
			QryTempDB->(DBCloseArea())
			QrySB2An->(DBSkip())
		Enddo
		QrySB2An->(DBCloseArea())

		// Remove Linha em Branco
		If Len(aProdAn) >= 2
			If Empty(aProdAn[1,1])
				Adel(aProdAn,1)
				ASize(aProdAn,Len(aProdAn)-1)
			EndIf
		EndIf

		// Teste Wagner
		If oListBox <> Nil
			MATCPrDiv()
			//ListBox(aProdAn)
			//oListBox:SetArray(aProdAn)
			//oListBox:bLine := &bLine
			//oListBox:Refresh()
			//oListBox:SetFocus()
		EndIf

		If Len(aProdAn) == 1 .And. Empty(aProdAn[1,1])
			FWAlertInfo("Não foram encontradas divergências para os parâmetros informados!")
		EndIf

		//Deleta tabela temporária criada no banco de dados
		oTempTable:Delete()
		oTempTable := Nil
	EndIf

	//************************************************************************************************************
	// ADICIONAR VALIDAÇÃO DA QUANTIDADE DAS MOVIMENTAÇÕES COM ARQUIVO TEMPORARIO
	//************************************************************************************************************

Return


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCExprd
Exporta o produtos divergentes: Custo Fechamento <> Custo Movimento

@Param
cNomeArq  - Nome do arquivo para gravação
cCaminho  - Diretório para gravação

@author    Ronaldo Tapia
@version   12.1.17
@since     02/10/2018
@protected

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCExprd(cNomeArq,cCaminho,oListBox,aInfPatc,lExpPatc)

	Local cTexto	:= ""
	Local nHandle	:= 0
	Local cFileTxt	:= ""
	Local lRetorno	:= .T.
	Local cFile		:= ""
	Local nx
	Local ny

	Default oListBox := Nil
	Default cNomeArq := ""
	Default cCaminho := ""
	Default aInfPatc := {}
	Default lExpPatc := .F.

	cCaminho := AllTrim(cCaminho)

	If oListBox <> Nil
		aInfPatc := oListBox:aArray
	EndIf

	If Empty(cNomeArq) .And. Empty(cCaminho)
		Aviso("GeraArquivo", "Parametros em Branco, verifique!", {"Ok"}) // "GeraArquivo" ## "Parametros em Branco, verifique!" ## "Ok"
		lRetorno := .F.
	EndIf

	If lRetorno
		cTexto := "----------------------------------------------------------------------------------------------------------------------"
		cTexto += _CRLF
		cTexto += "                         				Custo Fechamento <> Custo Movimento" // "                         Custo Fechamento <> Custo Movimento"
		cTexto += _CRLF
		cTexto += "----------------------------------------------------------------------------------------------------------------------"
		cTexto += _CRLF + _CRLF

		For nx := 1 To Len(aProdAn)
			cTexto += "Filial             Produto               Armazem               Status
			cTexto += _CRLF
			cTexto += aProdAn[nx][1] + "                 " + aProdAn[nx][2] + "       " + aProdAn[nx][3] + "                    " + aProdAn[nx][4]
			cTexto += _CRLF + _CRLF
		Next nx

		cFile    := Alltrim(cNomeArq) + ".txt"
		cFileTxt := cCaminho + cFile

		// Cria arquivo texto
		nHandle := MsFCreate(cFileTxt)

		If nHandle < 0
			Aviso("Geração de Arquivo Texto","Não foi possível criar o arquivo: " + cFile + "." + " Erro: " + IIf(cValToChar( FError() ) == "13", "Sem permissão de acesso ao diretório, verifique!",cValToChar( FError() )),{"Ok"},3) // "Geração de Arquivo Texto" ## "Não foi possível criar o arquivo: " ## "Erro: " ## "Ok"
		Else
			WrtStrTxt(nHandle,cTexto)
			Aviso("Geração de Arquivo Texto","Arquivo: " + cFile + " " + "gerado com sucesso no Diretório: " + Alltrim(cCaminho) + ".",{"Ok"},3) // "Geração de Arquivo Texto" ## "Arquivo: " ## "gerado com sucesso no Diretório: " ## "Ok"
		EndIf

		FClose(nHandle)
	EndIf

Return()


//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCValFo
Valida a mudança de aba

@Param
nTargetFolder - Folder selecionado

@author    Ronaldo Tapia
@version   12.1.17
@since     04/10/2018
@protected

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCValFo(nTargetFolder)

	Local oDlgValid := Nil
	Local oFont8   	:= Nil
	Local oButOK    := Nil

	Default nTargetFolder := 1

	// Validação da aba 2 - VALIDAÇÕES
	If nTargetFolder == 2 .And. lLogDec

		DEFINE FONT oFont8 NAME "Arial" BOLD
		DEFINE MSDIALOG oDlgValid FROM 264,182 TO 441,785 TITLE "Validação de Casas Decimais" OF oDlgValid PIXEL

		@ 007,007 SAY "ATENÇÃO: Existem campos com decimais divergentes, poderão ocorrer diferenças de arredondamento!"   OF oDlgValid PIXEL Size 295,010 FONT oFont8 COLOR CLR_HRED

		@ 017,007 SAY oBtECD PROMPT "<u>" + "Clique aqui e selecione a aba 'Precisão de Calculo' para mais informações sobre casas decimais" + "</u>" SIZE 400,010 OF oDlgValid HTML PIXEL
		oBtECD:bLClicked := {|| ShellExecute("open","http://tdn.totvs.com/x/NYJJF","","",1) }

		oButOK := tButton():New(060,240,"OK",oDlgValid,{|| oDlgValid:End() },036,12,,,.F.,.T.,.F.,,.F.,,,.F. )
		oButOK:SetCss(cCSSTBut)

		oDlgValid:Activate(,,,.T.,/*valid*/,,/*On Init*/)
	EndIf

Return

//------------------------------------------------------------------------------------------
/* {Protheus.doc} MATCValFoK
Valida os parametros para listar o Kardex em tela

@Param
nTargetFolder - Folder selecionado

@author    Ronaldo Tapia
@version   12.1.17
@since     04/10/2018
@protected

@Return ( Nil )
*/
//------------------------------------------------------------------------------------------
Static Function MATCValFoK()

	Local oDlgValid := Nil
	Local oFont8   	:= Nil
	Local oButOK    := Nil

	DEFINE FONT oFont8 NAME "Arial" BOLD
	DEFINE MSDIALOG oDlgValid FROM 264,182 TO 441,585 TITLE "Parâmetros Kardex" OF oDlgValid PIXEL

	@ 007,007 SAY "Tipo de ordenação: " OF oDlgValid PIXEL Size 295,010 FONT oFont8 COLOR CLR_BLACK
	TCheckBox():New( 007,70,"NUMSEQ",bSETGET(lKardex1),oDlgValid,150,009,,,,,,,,.T.,,,)
	TCheckBox():New( 017,70,"SEQCAL",bSETGET(lKardex2),oDlgValid,150,009,,,,,,,,.T.,,,)

	oButOK := tButton():New(060,140,"OK",oDlgValid,{|| MATCProces(oFolder2) },036,12,,,.F.,.T.,.F.,,.F.,,,.F. )
	oButOK:SetCss(cCSSTBut)

	oDlgValid:Activate(,,,.T.,,,)

Return


//------------------------------------------------------------------------------------------
/* {Protheus.doc} CalcCusOP

@ Calcula o custo das OPs com base nas requisições e produções

@author    Wagner Lima
@version   12.1.17
@since     29/07/2019
@protected

@Return (Retorna array com OP divergente)
*/
//------------------------------------------------------------------------------------------
Static Function CalcCusOP(cProdini, cProdFim, cLocAnDe, cLocAnAte, cDataIni, cDataFim)
Local nProdProp  	:= SuperGetMV("MV_PRODPR0",.F.,1)
//LOCAL lseq300    	:= SuperGetMv('MV_SEQ300',.F.,.F.)
Local cQuery     	:= ""
Local cAlias     	:= ""
Local cOpAnt     	:= ""
Local cOp           := ""
Local cParcTot      := "" //Apontamento parcial ou total
Local cDataAnt      := ""
Local nTotReq    	:= 0  //Total de requisições
Local nDevolucao    := 0  //Devoluções do produto requisistado contra a OP
Local nTotProd   	:= 0  //Total de produções
Local nMetAprop  	:= 2  // Metodo Apropriacao 1 = Sequencial / 2 = Mensal / 3 = Diaria
Local nSC2Ini	 	:= 0
Local nSC2Fin	 	:= 0
Local nSc2Quant  	:= 0
Local nPos          := 0
Local nCont      	:= 0
Local nContProd  	:= 0
Local nVlOpSB9   	:= 0
Local nCusOP	 	:= 0
Local nQuantApont   := 0
Local nReqOp        := 0
Local nDevOp        := 0
Local nTotApont     := 0
Local aStrut     	:= {}
Local aSd3Prod   	:= {}
Local aSd3Req    	:= {}
Local aArrSc2    	:= {}
Local aCusOp		:= {}
LOcal aDevProd      := {}
Local oTable     	As Object
Local oSd3Prod      As Object
Local oRet			As Object
DEFAULT cLocAnDe   	:= "**"

//Verifica o método de aprorpiação utilizado no recálculo do custo médio buscando da SX1
Pergunte("MTA330",.F.)
nMetAprop := mv_par14

//Função para geração da tabela temporária
cAlias := GeraTTP(cProdini, cProdFim, cDataFim, cDataIni,cLocAnDe,cLocAnAte)

While ((cAlias)->(EOF()) == .F.)

	nSC2Ini   := (cAlias)->C2_VINI
	nSC2Fin   := (cAlias)->C2_VFIM
	nSc2Quant := (cAlias)->C2_QUANT

		IF SUBSTR((cAlias)->D3_CF,1,2) $ "RE" // Gera o array com as requisições da OP

			AADD(aSd3Req, {(cAlias)->C2_OP,(cAlias)->D3_CUSTO, (cAlias)->D3_EMISSAO, (cAlias)->D3_NUMSEQ})
			nTotReq  +=  (cAlias)->D3_CUSTO

		ELSEIF SUBSTR((cAlias)->D3_CF,1,2) $ "PR" // Gera o aarray com os apontamentos de produção (PR0)

			AADD(aSd3Prod, {(cAlias)->C2_OP,(cAlias)->D3_CUSTO, (cAlias)->D3_QUANT, (cAlias)->D3_NUMSEQ, (cAlias)->D3_PARCTOT,(cAlias)->D3_EMISSAO})
			nTotProd    +=  (cAlias)->D3_CUSTO
			nQuantApont += (cAlias)->D3_QUANT
			cParcTot    := (cAlias)->D3_PARCTOT

		ELSEIF SUBSTR((cAlias)->D3_CF,1,3) $ "DE0-DE1" // Calcula as devoluções de produto requisitado contra a OP

			nDevolucao  +=  (cAlias)->D3_CUSTO
			AADD(aDevProd, {(cAlias)->C2_OP, (cAlias)->D3_CUSTO,(cAlias)->D3_EMISSAO, (cAlias)->D3_NUMSEQ})
		END
	cOpAnt    := (cAlias)->C2_OP
	cDataAnt  := (cAlias)->D3_EMISSAO
	(cAlias)->(DbSkip())

	IF cOpAnt <> (cAlias)->C2_OP
		AADD(aArrSc2, {cOpAnt, nSC2Ini, nTotReq + nSC2Ini - nDevolucao , nTotProd, nQuantApont, nSC2Fin, nSc2Quant, cParcTot, nDevolucao})
		nTotReq     := 0
		nTotProd    := 0
		nQuantApont := 0
		nDevolucao  := 0
	END
ENDDO

//Processa todas as OPs do produto Pai do array aArrSc2
FOR nCont := 1 TO LEN(aArrSc2)
	IF LEN(aSd3Prod) > 0

		nContProd := nCont
		nPos := 0
		nContProd := 1
		oSd3Prod := aToHM(aSd3Prod)

		IF !HMGet(oSd3Prod, aArrSc2[nCont][1], oRet)
			cOp     := aArrSc2[nCont][1]
			nSC2Ini := aArrSc2[nCont][2]
			nSC2Fin := aArrSc2[nCont][6]
			nCusOP  := aArrSc2[nCont][3] - aArrSc2[nCont][4]
		ELSE

			WHILE nPos < LEN(aSd3Prod)
				IF aArrSc2[nCont][1] == aSd3Prod[nContProd][1]
					IF  aArrSc2[nCont][8] == "T" //Calcula o custo para para o apontamento total da OP
						cOp     := aArrSc2[nCont][1]
						nSC2Ini := aArrSc2[nCont][2]
						nSC2Fin := aArrSc2[nCont][6]
						nCusOP  := aArrSc2[nCont][3] - aArrSc2[nCont][4]
					ELSE //Calcula o custo de apontamentos parciais

						DO Case
							CASE nMetAprop == 1 //Método de apropriação sequencial
								nReqOp := 1
								WHILE nReqOp <= LEN(aSd3Req) //Processo todas as requisições com data menos ou igual ao do apontamento da op
									IF aSd3Req[nReqOp][4] <= aSd3Prod[nContProd][4] .AND. aSd3Req[nReqOp][1] == aSd3Prod[nContProd][1]
										nTotReq += aSd3Req[nReqOp][2]
									END
									nReqOp ++
									LOOP
								END
								nDEvOp := 1
								WHILE nDEvOp <= LEN(aDevProd)
									IF aDevProd[nDEvOp][3] <= aSd3Prod[nContProd][6] .AND. aDevProd[nDEvOp][1] = aSd3Prod[nContProd][1]
										nTotReq -= 	aDevProd[nDEvOp][2]
									END
									nDEvOp ++
									LOOP
								END
								nCusOP  += (nTotReq - nCusOP)
								nTotApont +=  nTotReq
								nContProd ++
								nTotReq := 0

							CASE nMetAprop == 2 //Método de apropriação mensal

								cOp     := aArrSc2[nCont][1]
								nSC2Ini := aArrSc2[nCont][2]
								nSC2Fin := aArrSc2[nCont][6]
								IF nProdProp == 1 // MV_PRODPR0 com o conteúdo 1
									nCusOP  += aArrSc2[nCont][3]  / (aArrSc2[nCont][5] / aSd3Prod[nContProd][3])
								ELSEIF nProdProp == 3 // MV_PRODPR0 com o conteúdo 3
									nCusOP  += (aSd3Prod[nContProd][3]  / aArrSc2[nCont][7]) * aArrSc2[nCont][3]
								END
								nContProd ++

							CASE nMetAprop == 3 //Método de apropriação diário
								nReqOp := 1
								WHILE nReqOp <= LEN(aSd3Req) //Processo todas as requisições com data menos ou igual ao do apontamento da op
									IF aSd3Req[nReqOp][3] <= aSd3Prod[nContProd][6] .AND. aSd3Req[nReqOp][1] = aSd3Prod[nContProd][1]
										nTotReq += aSd3Req[nReqOp][2]
									END
									nReqOp ++
									LOOP
								END
								nDEvOp := 1
								WHILE nDEvOp <= LEN(aDevProd)
									IF aDevProd[nDEvOp][3] <= aSd3Prod[nContProd][6] .AND. aDevProd[nDEvOp][1] == aSd3Prod[nContProd][1]
										nTotReq -= 	aDevProd[nDEvOp][2]
									END
									nDEvOp ++
									LOOP
								END

								IF nProdProp == 1
									nCusOP  += nTotReq / (aArrSc2[nCont][5] / aSd3Prod[nContProd][3])
								ELSEIF nProdProp == 3
									nCusOP  += (aSd3Prod[nContProd][3]  / aArrSc2[nCont][7]) * nTotReq
								END
								nContProd ++
								nTotReq := 0
						ENDCASE

					END
					nPos++
				ELSE
					nContProd++
					nPos++
				END
				LOOP
				nTotApont := 0
			ENDDO
			IF nCusOP > 0
				nCusOP := aArrSc2[nCont][3] - nCusOP
			END
		END

		AADD(aCusOp, {cOp,  nSC2Ini, nSC2Fin, nCusOP})
		nCusOP := 0

	END

NEXT nCont

Return (aCusOp)

/*------ Função para gerar a tabela temporária com as movimentações da OP -----*/
/* ---------------------------Wagner Lima 16/09/2019-------------------------- */
/*Criei esta função para facilitar em futras modificações na estrutura da query*/

STATIC FUNCTION GeraTTP(cProdini, cProdFim, cDataFim, cDataIni,cLocAnDe,cLocAnAte)
Local cQuery     	:= ""
Local cAlias     	:= ""
Local aStrut        := {}
Local oTable     	As Object
Local oSd3Prod      As Object
DEFAULT cLocAnDe   	:= "**"

//Query para busca de requisições, devoluções e apontamentos de produção
cQuery := "SELECT SC2.C2_FILIAL As C2_FILIAL, SC2.C2_NUM + SC2.C2_ITEM + SC2.C2_SEQUEN AS C2_OP, SC2.C2_VINI1 AS C2_VINI, SC2.C2_VFIM1 AS C2_VFIM, "

cQuery += "SC2.C2_PRODUTO As C2_PROD, SC2.C2_QUANT AS C2_QUANT, SD3.D3_CF AS D3_CF, SD3.D3_CUSTO1 As D3_CUSTO, SD3.D3_EMISSAO As D3_EMISSAO, SD3.D3_NUMSEQ As D3_NUMSEQ, SD3.D3_QUANT As D3_QUANT, SD3.D3_PARCTOT AS D3_PARCTOT "

cQuery += "FROM "+RetSqlName("SC2")+" AS SC2 JOIN "+RetSqlName("SD3")+" AS SD3 ON SC2.C2_NUM + SC2.C2_ITEM + SC2.C2_SEQUEN  = SD3.D3_OP "

cQuery += "AND SC2.C2_FILIAL = '"+xFilial("SC2")+"' AND SC2.D_E_L_E_T_ <> '*' AND SD3.D3_FILIAL = '"+xFilial("SC2")+"' AND SD3.D_E_L_E_T_ <> '*' AND SD3.D3_ESTORNO = ' ' AND SD3.D3_EMISSAO BETWEEN '"+DTOS(cDataIni)+"' AND '"+DTOS(cDataFim)+"' "

cQuery += "AND SC2.C2_PRODUTO BETWEEN '"+cProdini+"' AND '"+cProdFim+"' AND SC2.C2_EMISSAO <= '"+DTOS(cDataFim)+"' AND (SC2.C2_DATRF = ' ' OR SC2.C2_DATRF >= '"+DTOS(cDataIni)+"') "

IF cLocAnDe != "**"
	cQuery += "AND SC2.C2_LOCAL BETWEEN  '"+cLocAnDe+"' AND '"+cLocAnAte+"' "
END

cQuery := ChangeQuery(cQuery)

//Pega o próximo alias temporário disponível
cAlias := GetNextAlias()
//oTable:GetRealName() RETONA A TABELA TEMPORÁRIA EM USO
oTable := FwTemporaryTable():New(cAlias)


//Cria a estrutura dos campos da tabela temporária
Aadd(aStrut, {"C2_FILIAL","C",TamSX3("C2_FILIAL")[1],0})
Aadd(aStrut, {"C2_VINI","N",TamSX3("B2_CM1")[1],8})
Aadd(aStrut, {"C2_VFIM","N",TamSX3("B2_CM1")[1],8})
Aadd(aStrut, {"C2_PROD","C",TamSX3("B1_COD")[1],0})
Aadd(aStrut, {"C2_OP","C",TamSX3("D3_OP")[1],0})
Aadd(aStrut, {"C2_QUANT","N",TamSX3("C2_QUANT")[1],2})
Aadd(aStrut, {"D3_CF","C",TamSX3("D3_CF")[1],0})
Aadd(aStrut, {"D3_CUSTO","N",TamSX3("B2_CM1")[1],8})
Aadd(aStrut, {"D3_EMISSAO","C",TamSX3("D3_EMISSAO")[1],0})
Aadd(aStrut, {"D3_NUMSEQ","C",TamSX3("D3_NUMSEQ")[1],0})
Aadd(aStrut, {"D3_QUANT","N",TamSX3("D3_QUANT")[1],2})
Aadd(aStrut, {"D3_PARCTOT","C",TamSX3("D3_PARCTOT")[1],0})
oTable:SetFields(aStrut)

//Cria  o índice temporário
oTable:AddIndex("1", {"C2_FILIAL","C2_PROD","C2_OP"} )
oTable:Create()

SQLToTrb(cQuery, aStrut, cAlias)
(cAlias)->(DbGotop())

RETURN (cAlias)

Static Function ListBox (aProdAn)

	nTamCol := Len(aProdAn[01])
	bLine 	:= "{|| {"
	For Nx := 1 To nTamCol
		bLine += "aProdAn[oListBox:nAt]["+StrZero(Nx,3)+"]"
		If Nx < nTamCol
			bLine += ","
		EndIf
	Next
	bLine += "} }"

Return
