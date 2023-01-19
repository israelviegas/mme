package mme;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import mme.util.Util;

public class AutomacaoMme {
	
	private static List<RelatorioCostTransactionExtract> listaRelatorioPlanilhaCostTransactionExtract =  null;
	private static List<RelatorioResourceTrend> listaRelatorioPlanilhaResourceTrend =  null;
	private static List<RelatorioMultiSegmentContractReport> listaRelatorioPlanilhaMultiSegmentContractReport =  null;
	private static String dataAtual = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss").format(new Date());
	private static String dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
	private static SimpleDateFormat formatoDataReferencia = new SimpleDateFormat("yyyy-MM-dd"); 
	private static int mesAtual = Integer.parseInt(new SimpleDateFormat("MM").format(new Date()));
	private static int mesAnterior = retornaMesAnterior(mesAtual);
	private static int anoAtual = Integer.parseInt(new SimpleDateFormat("yyyy").format(new Date()));
	private static int anoCorreto = retornaAnoCorreto(mesAnterior, anoAtual);
	private static java.sql.Date dataAtualBancoFinal = new java.sql.Date(new Date().getTime());
	private static Set<String> listaNomesDeContratosDistintos =  new HashSet<String>();
	private static ArrayList<RelatorioResourceTrend> listaMesAnoResourceTrend = null;
	private static ArrayList<RelatorioMultiSegmentContractReport> listaMesAnoMultiSegmentContractReport = null;
	private static String diretorioLogs = "";
	private static String subdiretorioRelatoriosBaixados = "";
	private static String mensagemResultadoRelatoriosCostTransactionExtract = "";
	private static String mensagemResultadoRelatoriosResourceTrend = "";
	private static String mensagemResultadoRelatoriosMultiSegmentContractReport = "";
	private static String anoReferenciaResourceTrend = "";
	private static String mesReferenciaResourceTrend = "";
	private static String anoReferenciaMultiSegmentContractReport = "";
	private static String mesReferenciaMultiSegmentContractReport = "";
	private static String mensagemFinalResultadoMme = "";
	private static int contadorErrosCalendario = 0;
	private static int contadorErrosArquivoInvalido = 0;
	private static int contadorParaAlterarNomeArquivos = 0;
	private static int executaAutomacaoMme = 0;
	
	
	public static void main(String[] args) throws InterruptedException, IOException, SQLException {
		
		WebDriver driver = null;

    		try {
    			
    		diretorioLogs = Util.getValor("caminho.diretorio.relatorios") + "/" + dataAtual;
    		String diretorioRelatorio = Util.getValor("caminho.download.relatorios") + "\\" + dataAtual;
    		subdiretorioRelatoriosBaixados = diretorioRelatorio + "\\" + "relatorios baixados " + dataAtual;
    		criaDiretorio(subdiretorioRelatoriosBaixados);
    		
			// Deleta os diret�rios que possu�rem data de cria��o anterior � data de 7 dias atr�s
			apagaDiretoriosDeRelatorios(Util.getValor("caminho.download.relatorios"));
			
			// As vezes o diret�rio que armazena dados tempor�rios do Chome simplesmente some, da� o Selenium d� pau na hora de chamar o browser
			// Com o m�todo abaixo, crio essa pasta se ela n�o existir
			criaDiretorioTemp();

    		executaAutomacaoMme(driver);
            
		} catch (Exception e) {
			gravarArquivo(diretorioLogs, "Erro Mme" + " " + dataAtual, ".txt", e.getMessage(), "Ocorreu um erro no Mme: ");
			inserirStatusExecucaoNoBanco("Mme", dataAtualPlanilhaFinal, "Erro");

		} finally {
			//mensagemErro("Houve um problema na extra��o dos pedidos no Mme\n");
			//fazerLogout(wait);
			if (driver != null) {
				driver.quit();
			}
			
			mataProcessosGoogle();
			mataProcessosFirefox();
		}
        
    }
	
    public static void executaAutomacaoMme(WebDriver driver) throws Exception{
    	
    	try {
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		mataProcessosFirefox();
    		
    		listaRelatorioPlanilhaCostTransactionExtract =  new ArrayList<RelatorioCostTransactionExtract>();
    		listaRelatorioPlanilhaResourceTrend =  new ArrayList<RelatorioResourceTrend>();
    		listaRelatorioPlanilhaMultiSegmentContractReport =  new ArrayList<RelatorioMultiSegmentContractReport>();
    		listaMesAnoResourceTrend = new ArrayList<RelatorioResourceTrend>();
    		listaMesAnoMultiSegmentContractReport = new ArrayList<RelatorioMultiSegmentContractReport>();
    		
    		System.out.println("In�cio: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		driver = getWebDriver();
    		Thread.sleep(5000);
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		WebDriverWait wait = new WebDriverWait(driver, 60);
    		
        	// Deleto arquivos que existirem no diret�rio relat�rio
        	apagaArquivosDiretorioDeRelatorios(Util.getValor("caminho.diretorio.relatorios"));
        	
        	// Deleto arquivos que existirem no sub-diret�rio de relat�rios baixados
        	apagaArquivosDiretorioDeRelatorios(subdiretorioRelatoriosBaixados);
        	
			// Contrato 9940191116 SW Factories
        	// Hist�rico Passado + Atuais
        	if (Util.getValor("contrato.9940191116.SW.Factories").equals("S")) {
        		contrato_9940191116_SW_Factories(driver, wait);
        	}
			
			// Contrato AM Faturamento
			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.AM.Faturamento").equals("S")) {
   				contrato_AM_Faturamento(driver, wait);
   			}
   			
			// Contrato Command Center
			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Command.Center").equals("S")) {
   				contrato_Command_Center(driver, wait);
   			}

   			// Contrato B2C SFA - Salesforce
			// Hist�rico Passado
   			if (Util.getValor("contrato.B2C.SFA.Salesforce").equals("S")) {
   				contrato_B2C_SFA_Salesforce(driver, wait);
   			}
			
			// Contrato Callidus
			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Callidus").equals("S")) {
   				contrato_Callidus(driver, wait);
   			}
			
			// Contrato GVT Proforma
			// Hist�rico Passado
   			if (Util.getValor("contrato.GVT.Proforma").equals("S")) {
   				contrato_GVT_Proforma(driver, wait);
   			}
			
			// Contrato Hybris - eCommerce
			// Hist�rico Passado
   			if (Util.getValor("contrato.Hybris.eCommerce").equals("S")) {
   				contrato_Hybris_eCommerce(driver, wait);
   			}

			// Contrato Digital Factory
			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Digital.Factory").equals("S")) {
   				contrato_Digital_Factory(driver, wait);
   			}
   			
			// Contrato AM Latam Brasil Fija Contrato Local
			// Hist�rico Passado
   			if (Util.getValor("contrato.AM.Latam.Brasil.Fija.Contrato.Local").equals("S")) {
   				contrato_AM_Latam_Brasil_Fija_Contrato_Local(driver, wait);
   			}
   			
			// Contrato Nova Fabrica Design
   			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Nova.Fabrica.Design").equals("S")) {
   				contrato_Nova_Fabrica_Design(driver, wait);
   			}
   			
			// Contrato Desligue Do Atis
   			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Desligue.Do.Atis").equals("S")) {
   				contrato_Desligue_Do_Atis(driver, wait);
   			}
  			
			// Contrato Portal_Terra
   			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Portal.Terra").equals("S")) {
   				contrato_Portal_Terra(driver, wait);
   			}
 			
			// Contrato Sustentacao VIVO GO
   			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.Sustentacao.VIVO.GO").equals("S")) {
   				contrato_Sustentacao_VIVO_GO(driver, wait);
   			}
   			
			// Contrato FiberCo Imp. Arquitet.BSS/OSS
			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.FiberCo.Imp.Arquitet.BSS.OSS").equals("S")) {
   				contrato_FiberCo_Imp_Arquitet_BSS_OSS(driver, wait);
   			}
   			
			// Contrato RPA - Blue Prism
			// Hist�rico Passado + Atuais
   			if (Util.getValor("contrato.RPA.Blue.Prism").equals("S")) {
   				contrato_RPA_Blue_Prism(driver, wait);
   			}
   			
			// Contrato Prote��o de dados
        	// Hist�rico Passado + Atuais
        	if (Util.getValor("contrato.Protecao.De.Dados").equals("S")) {
        		contrato_Protecao_De_Dados(driver, wait);
        	}
        	
			// Contrato B2B Transformation
        	// Hist�rico Passado + Atuais
        	if (Util.getValor("contrato.B2B.Transformation").equals("S")) {
        		contrato_B2B_Transformation(driver, wait);
        	}
        	
        	persisteRelatoriosCostTransactionExtractNoBanco();
   			
   			//persisteRelatoriosResourceTrend();
   			
   			persisteRelatoriosMultiSegmentContractReport();
   			
   			mensagemFinalResultadoMme = mensagemResultadoRelatoriosCostTransactionExtract + "\n" + mensagemResultadoRelatoriosResourceTrend + "\n" + mensagemResultadoRelatoriosMultiSegmentContractReport;
			
            gravarArquivo(diretorioLogs, "Resultado Mme" + " " + dataAtual, ".txt", "", mensagemFinalResultadoMme);
            
            inserirStatusExecucaoNoBanco("Mme", dataAtualPlanilhaFinal, "Sucesso");
            
            System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    	} catch (Exception e) {
			executaAutomacaoMme ++;
			// Executo at� 5 vezes se der erro no executaAutomacaoMme
			if (executaAutomacaoMme <= 5) {
				
				System.out.println("Deu erro no m�todo executaAutomacaoMme, tentativa de acerto: " + executaAutomacaoMme + "\n" + "Erro: " + e.getMessage());
				executaAutomacaoMme(driver);
			
			} else {
				throw new Exception("Ocorreu um erro no m�todo executaAutomacaoMme: " + e);
		    }

		}
    	
    }
    
    public static void persisteRelatoriosCostTransactionExtractNoBanco() throws Exception {
    	
    	deletaMesAtualEMesAnteriorEInsereRelatoriosCostTransactionExtractNoBanco();
    	
    }
    
    public static void deletaMesAtualEMesAnteriorEInsereRelatoriosCostTransactionExtractNoBanco() throws Exception {
    	
    	// Insere no banco a lista contendo todos os relat�rios baixados
    	if (listaRelatorioPlanilhaCostTransactionExtract != null && !listaRelatorioPlanilhaCostTransactionExtract.isEmpty()) {
    		
    		int contador = 0;
    		
    		if (listaNomesDeContratosDistintos != null && !listaNomesDeContratosDistintos.isEmpty()) {
    			
    			for (String nomeContrato : listaNomesDeContratosDistintos) {
    				
    				// Apagando o mes atual do contrato
    				MmeDao mmeDaoDeletarMesAtual = new MmeDao();
    				mmeDaoDeletarMesAtual.deletarRelatoriosCostTransactionExtract(mesAtual, anoAtual, nomeContrato);
    				System.out.println("Relatorio Cost Transaction Extract. Apaguei todos os dados do banco do m�s atual do contrato " + nomeContrato);
    				
    				// Apagando o mes anterior do contrato
    				MmeDao mmeDaoDeletarMesAnterior = new MmeDao();
    				mmeDaoDeletarMesAnterior.deletarRelatoriosCostTransactionExtract(mesAnterior, anoCorreto, nomeContrato);
    				System.out.println("Relatorio Cost Transaction Extract. Apaguei todos os dados do banco do m�s anterior do contrato " + nomeContrato);
    				
    			}
    			
    		}
    		
    		for (RelatorioCostTransactionExtract relatorioCostTransactionExtract : listaRelatorioPlanilhaCostTransactionExtract) {
    			
    			// Converto os valores de null para espa�o em branco para gravar branco no banco e n�o dar problema
    			// no relat�rio do sharepoint da Accenture
    			// Esse relat�rio de sharepoint da Accenture � gerado atrav�s do banco
    			Util.converteValorNullParaEspacoEmBrancoRelatorioCostTransactionExtract(relatorioCostTransactionExtract);
    			MmeDao mmeDaoInserirRelatorio = new MmeDao();
    			mmeDaoInserirRelatorio.inserirRelatoriosCostTransactionExtract(relatorioCostTransactionExtract);
    			contador++;
    			System.out.println("Relatorio Cost Transaction Extract. Inseri o relat�rio de n�mero: " + contador + " de um total de " + listaRelatorioPlanilhaCostTransactionExtract.size() + " nome do contrato: " + relatorioCostTransactionExtract.getContrato());
    			
    		}
    		
    		mensagemResultadoRelatoriosCostTransactionExtract = "Relat�rios Cost Transaction Extract do Mme gravados no banco com sucesso";
    		
    	} else {
    		
    		mensagemResultadoRelatoriosCostTransactionExtract = "N�o foram encontrados relat�rios Cost Transaction Extract do Mme para os contratos processados";
    		
    	}
    
    }
    
    public static void persisteRelatoriosResourceTrend() throws Exception {
    	
    	deletaMesAtualEInsereRelatoriosResourceTrendNoBanco();
    	
    }
    
    public static void deletaMesAtualEInsereRelatoriosResourceTrendNoBanco() throws Exception {
    	
    	// Insere no banco a lista contendo todos os relat�rios baixados
    	if (listaRelatorioPlanilhaResourceTrend != null && !listaRelatorioPlanilhaResourceTrend.isEmpty()) {
    		
    		int contador = 0;
    		
    		if (listaMesAnoResourceTrend != null && !listaMesAnoResourceTrend.isEmpty()) {
    			
    			for (RelatorioResourceTrend relatorioResourceTrend : listaMesAnoResourceTrend) {
    				
    				// Apagando o m�s atual do contrato
    				MmeDao mmeDaoDeletarMesAtual = new MmeDao();
    				mmeDaoDeletarMesAtual.deletarRelatoriosResourceTrend(Integer.parseInt(relatorioResourceTrend.getMes()), Integer.parseInt(relatorioResourceTrend.getAno()), relatorioResourceTrend.getContrato());
    				System.out.println("Relatorio Resource Trend. Apaguei todos os dados do banco do m�s atual do contrato " + relatorioResourceTrend.getContrato());
    				
    			}
    			
    		}
    		
    		for (RelatorioResourceTrend relatorioResourceTrend : listaRelatorioPlanilhaResourceTrend) {
    			
    			// Converto os valores de null para espa�o em branco para gravar branco no banco e n�o dar problema
    			// no relat�rio do sharepoint da Accenture
    			// Esse relat�rio de sharepoint da Accenture � gerado atrav�s do banco
    			Util.converteValorNullParaEspacoEmBrancoRelatorioResourceTrend(relatorioResourceTrend);
    			MmeDao mmeDaoInserirRelatorio = new MmeDao();
    			mmeDaoInserirRelatorio.inserirRelatoriosResourceTrend(relatorioResourceTrend);
    			contador++;
    			System.out.println("Relatorio Resource Trend. Inseri o relat�rio de n�mero: " + contador + " de um total de " + listaRelatorioPlanilhaResourceTrend.size() + " nome do contrato: " + relatorioResourceTrend.getContrato());
    			
    		}
    		
    		mensagemResultadoRelatoriosResourceTrend = "Relat�rios Resource Trend do Mme gravados no banco com sucesso";
    		
    	} else {
    		
    		mensagemResultadoRelatoriosResourceTrend = "N�o foram encontrados relat�rios Resource Trend do Mme para os contratos processados";
    		
    	}
    
    }
    
    public static void persisteRelatoriosMultiSegmentContractReport() throws Exception {
    	
    	deletaMesAtualEInsereRelatoriosMultiSegmentContractReportNoBanco();
    	
    }
    
    public static void deletaMesAtualEInsereRelatoriosMultiSegmentContractReportNoBanco() throws Exception {
    	
    	// Insere no banco a lista contendo todos os relat�rios baixados
    	if (listaRelatorioPlanilhaMultiSegmentContractReport != null && !listaRelatorioPlanilhaMultiSegmentContractReport.isEmpty()) {
    		
    		int contador = 0;
    		
    		if (listaMesAnoMultiSegmentContractReport != null && !listaMesAnoMultiSegmentContractReport.isEmpty()) {
    			
    			for (RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport : listaMesAnoMultiSegmentContractReport) {
    				// Apagando o m�s atual do contrato
    				MmeDao mmeDaoDeletarMesAtual = new MmeDao();
    				mmeDaoDeletarMesAtual.deletarRelatoriosMultiSegmentContractReport(Integer.parseInt(relatorioMultiSegmentContractReport.getMesMasterActive()), Integer.parseInt(relatorioMultiSegmentContractReport.getAnoMasterActive()), relatorioMultiSegmentContractReport.getContrato());
    				System.out.println("Multi-Segment Contract Report. Apaguei todos os dados do banco do m�s atual e futuro do contrato " + relatorioMultiSegmentContractReport.getContrato());
    				
    			}
    			
    		}
    		
    		for (RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport : listaRelatorioPlanilhaMultiSegmentContractReport) {
    			
    			// Converto os valores de null para espa�o em branco para gravar branco no banco e n�o dar problema
    			// no relat�rio do sharepoint da Accenture
    			// Esse relat�rio de sharepoint da Accenture � gerado atrav�s do banco
    			Util.converteValorNullParaEspacoEmBrancoRelatorioMultiSegmentContractReport(relatorioMultiSegmentContractReport);
    			MmeDao mmeDaoInserirRelatorio = new MmeDao();
    			mmeDaoInserirRelatorio.inserirRelatoriosMultiSegmentContractReport(relatorioMultiSegmentContractReport);
    			contador++;
    			System.out.println("Relatorio Multi-Segment Contract Report. Inseri o relat�rio de n�mero: " + contador + " de um total de " + listaRelatorioPlanilhaMultiSegmentContractReport.size() + " nome do contrato: " + relatorioMultiSegmentContractReport.getContrato());
    			
    		}
    		
    		mensagemResultadoRelatoriosMultiSegmentContractReport = "Relat�rios Multi-Segment Contract Report do Mme gravados no banco com sucesso";
    		
    	} else {
    		
    		mensagemResultadoRelatoriosMultiSegmentContractReport = "N�o foram encontrados relat�rios Multi-Segment Contract Report do Mme para os contratos processados";
    		
    	}
    
    }
    
    public static void fecharBarraInferiorComInformacoesDaAccenture(WebDriver driver) throws Exception {
    	
	    // Se aparecer a Hide, clico na aba para fechar
	    // Aguardo at� 5 segundos para a op��o aparecer
	    // Se ela n�o aparecer dar� erro, da� sigo adiante
    	try {
    		
    		WebDriverWait waitInformacoes = new WebDriverWait(driver, 5);
    		String textoHide = "Hide";
    		// Aguarda o surgimento da palavra Hide:
    		waitInformacoes.until(ExpectedConditions.elementToBeClickable(By.xpath("//div [text()='"+textoHide+"']"))).click();
		
    	} catch (Exception e) {
    		//System.out.println("Deu erro na hora de clicar na op��o Hide");
		}
    	
    }
    
    public static void rolagemParaBaixo(WebDriver driver) throws Exception {
    	
    	JavascriptExecutor jse = (JavascriptExecutor)driver;
    	jse.executeScript("scrollBy(0,250)", "");    	
    
    }
	
    public static void moverArquivosEntreDiretorios(String caminhoArquivoOrigem, String caminhoDiretorioDestino) throws Exception{
    	
    	boolean sucesso = true;
    	File arquivoOrigem = new File(caminhoArquivoOrigem);
        File diretorioDestino = new File(caminhoDiretorioDestino);
        if (arquivoOrigem.exists() && diretorioDestino.exists()) {
        	sucesso = arquivoOrigem.renameTo(new File(diretorioDestino, retornaNomeArquivoAlterado(arquivoOrigem.getName())));
        }
        
        if (!sucesso) {
        	throw new Exception("Ocorreu um erro no momento de mover o relat�rio " + caminhoArquivoOrigem + " para " + caminhoDiretorioDestino);
        }
        
    }
    
    public static String retornaNomeArquivoAlterado(String nomeArquivo) throws Exception{
    	
    	// Preciso alterar o nome dos arquivos antes de mov�-los porque temos nomes de arquivos iguais
    	String arquivo = nomeArquivo;
    	contadorParaAlterarNomeArquivos = contadorParaAlterarNomeArquivos + 1;
    	String[] partesNomeArquivo = arquivo.split("\\.");
    	
    	if (partesNomeArquivo != null && partesNomeArquivo.length > 0) {
    		String nome = partesNomeArquivo[0];
    		nome = nome + "_" + contadorParaAlterarNomeArquivos;
    		String extensao = partesNomeArquivo[1];
    		arquivo = nome + "." + extensao;
    	}
		
    	return arquivo;
        
    }
    
	@SuppressWarnings({ "resource" })
	public static void lerPlanilhaMultiSegmentContractReport(String planilha, String contrato) throws Exception {

		FileInputStream arquivo = null;
		
		try {
		   arquivo = new FileInputStream(new File(
				   planilha));
		
		   OPCPackage pkg = OPCPackage.open(new File(planilha));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetRelatorio = workbook.getSheet("RawData");
		   
		   // Uso o DataFormatter para deixar todos os campos como String, inclusive
		   // os que tem n�meros
		   DataFormatter formatter = new DataFormatter();
		   for (int i=0; i <= sheetRelatorio.getLastRowNum(); i++) {
		       Row row = sheetRelatorio.getRow(i);
		       
		       if (row != null) {
		    	   
		    	   RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport = new RelatorioMultiSegmentContractReport();
		    	   
		    	   if (row.getRowNum() == 0) {
		    		   continue;
		    	   }
		    	   
		           int indice = 0;
		           int contador = 0;
		           
		           // Level01
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel01(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level02
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel02(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level03
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel03(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level04
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel04(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level05
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel05(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level06
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel06(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level07
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel07(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level08
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel08(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level09
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel09(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level10
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel10(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level11
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel11(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level12
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setLevel12(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBSNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setWBSNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBSName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setWBSName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setCostCollectorNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setCostCollectorName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYear
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setFiscalYear(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalQuarter
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setFiscalQuarter(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Month
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setMonth(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Group
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setGroup(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Category
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setCategory(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Actual
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setActual(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Forecast
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setForecast(formatter.formatCellValue(row.getCell(indice)));
		           }

		           // ComparisonEAC
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMultiSegmentContractReport.setComparisonEAC(formatter.formatCellValue(row.getCell(indice)));
		           }

		           // Contrato
		           relatorioMultiSegmentContractReport.setContrato(contrato);

		           // Data Extra��o
		           relatorioMultiSegmentContractReport.setDataExtracao(dataAtualBancoFinal);
		           
		           // Data Refer�ncia
		           // Equivale ao campo Month do relat�rio, por�m vou salv�-lo no formato yyyy-MM-dd
		           String mesCampoMonth = "01";
		           String anoCampoMonth = "1999";
		           if (relatorioMultiSegmentContractReport.getMonth() != null && !relatorioMultiSegmentContractReport.getMonth().isEmpty()) {
		        	   
		        	   mesCampoMonth = retornaNumeroMes2(relatorioMultiSegmentContractReport.getMonth());

		        	   // Retirando os caracteres que n�o forem n�meros para trazer somente o ano
		        	   anoCampoMonth = relatorioMultiSegmentContractReport.getMonth().replaceAll("[^0-9]", "").trim();
		           
		           }
		           
		           Date dataReferencia = formatoDataReferencia.parse(anoCampoMonth + "-" + mesCampoMonth + "-" + "01");
		           java.sql.Date dataReferenciaFinal = new java.sql.Date(dataReferencia.getTime());
		           relatorioMultiSegmentContractReport.setDataReferencia(dataReferenciaFinal);
		           
		           if (relatorioMultiSegmentContractReport != null && !Util.relatorioMultiSegmentContractReportPossuiTodosCamposNulos(relatorioMultiSegmentContractReport)) {
		        	   listaRelatorioPlanilhaMultiSegmentContractReport.add(relatorioMultiSegmentContractReport);
		           }
		           
		       }
		       
		   }
		   
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de relat�rio n�o encontrado!");
			   throw new Exception("Arquivo Excel de relat�rio n�o encontrado!");
			
			} finally {
			
				if (arquivo != null) {
					arquivo.close();
				}
				
			}
		
			if (listaRelatorioPlanilhaMultiSegmentContractReport.size() == 0) {
			   //Pode ser que existam arquivos vazios, ent�o n�o posso lan�ar exce��o aqui
				//throw new Exception("Lista de projetos est� vazia");
			}
			
		}
    
	@SuppressWarnings({ "resource" })
	public static void lerPlanilhaResourceTrend(String planilha, String contrato, String ano, String mes) throws Exception {

		FileInputStream arquivo = null;
		
		try {
		   arquivo = new FileInputStream(new File(
				   planilha));
		
		   OPCPackage pkg = OPCPackage.open(new File(planilha));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetRelatorio = workbook.getSheet("Raw Data");
		   
		   // Uso o DataFormatter para deixar todos os campos como String, inclusive
		   // os que tem n�meros
		   DataFormatter formatter = new DataFormatter();
		   for (int i=0; i <= sheetRelatorio.getLastRowNum(); i++) {
		       Row row = sheetRelatorio.getRow(i);
		       
		       if (row != null) {
		    	   
		    	   RelatorioResourceTrend relatorioResourceTrend = new RelatorioResourceTrend();
		    	   
		    	   if (row.getRowNum() == 0) {
		    		   continue;
		    	   }
		    	   
		           int indice = 0;
		           int contador = 0;
		           
		           // ForecastVersion
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setForecastVersion(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // TypeCostCenterCareerTrack
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setTypeCostCenterCareerTrack(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Level
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setLevel(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMUID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setWMUID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMUName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setWMUName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMUOwner
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setWMUOwner(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setWBS(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollector
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setCostCollector(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setCostCollectorName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // RoleName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setRoleName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Name
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // PersonnelNumber
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setPersonnelNumber(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // HomeLocation
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setHomeLocation(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WorkLocation
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setWorkLocation(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Category
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setCategory(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Quantity
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setQuantity(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ActualForecast
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setActualForecast(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Date
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setDate(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // OrderBy
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setOrderBy(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // BillRateCardId
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setBillRateCardId(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // RateName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setRateName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // BillRateCardName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioResourceTrend.setBillRateCardName(formatter.formatCellValue(row.getCell(indice)));
		           }

		           // Contrato
		           relatorioResourceTrend.setContrato(contrato);

		           // Data Extra��o
		           relatorioResourceTrend.setDataExtracao(dataAtualBancoFinal);
		           
		           // Data Refer�ncia
		           Date dataReferencia = formatoDataReferencia.parse(ano + "-" + mes + "-" + "01");
		           java.sql.Date dataReferenciaFinal = new java.sql.Date(dataReferencia.getTime());
		           relatorioResourceTrend.setDataReferencia(dataReferenciaFinal);
		           
		           if (relatorioResourceTrend != null && !Util.relatorioResourceTrendPossuiTodosCamposNulos(relatorioResourceTrend)) {
		        	   listaRelatorioPlanilhaResourceTrend.add(relatorioResourceTrend);
		           }
		           
		       }
		       
		   }
		   
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de relat�rio n�o encontrado!");
			   throw new Exception("Arquivo Excel de relat�rio n�o encontrado!");
			
			} finally {
			
				if (arquivo != null) {
					arquivo.close();
				}
				
			}
		
			if (listaRelatorioPlanilhaResourceTrend.size() == 0) {
			   //Pode ser que existam arquivos vazios, ent�o n�o posso lan�ar exce��o aqui
				//throw new Exception("Lista de projetos est� vazia");
			}
			
		}
	
	@SuppressWarnings({ "resource" })
	public static void lerPlanilhaCostTransactionExtract(String planilha, String contrato, int mes, int ano, String periodo) throws Exception {

		FileInputStream arquivo = null;
		
		try {
		   arquivo = new FileInputStream(new File(
				   planilha));
		
		   OPCPackage pkg = OPCPackage.open(new File(planilha));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetRelatorio = workbook.getSheet("RawData");
		   
		   // Uso o DataFormatter para deixar todos os campos como String, inclusive
		   // os que tem n�meros
		   DataFormatter formatter = new DataFormatter();
		   Set<String> listaPeridosDeExtracaoDistintos = new HashSet<String>();
		   List<RelatorioCostTransactionExtract> listaRelatorioTemporaria = new ArrayList<RelatorioCostTransactionExtract>();
		   
		   for (int i=0; i <= sheetRelatorio.getLastRowNum(); i++) {
		       Row row = sheetRelatorio.getRow(i);
		       
		       if (row != null) {
		    	   
		    	   RelatorioCostTransactionExtract relatorioCostTransactionExtract = new RelatorioCostTransactionExtract();
		    	   
		    	   if (row.getRowNum() == 0 || row.getRowNum() == 1) {
		    		   continue;
		    	   }
		    	   
		           int indice = 0;
		           int contador = 0;
		           
		           // PostingPeriod
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setPostingPeriod(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYearPeriod
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setFiscalYearPeriod(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYearQuarter
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setFiscalYearQuarter(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYear
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setFiscalYear(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // PostingDate
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setPostingDate(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // DocumentDate
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setDocumentDate(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Category Group
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCategoryGroup(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Category
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCategory(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // DocumentType
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setDocumentType(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // DocumentNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setDocumentNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Description
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setDescription(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ReferenceNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setReferenceNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // AccountNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setAccountNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // AccountNbr Description
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setAccountNbrDescription(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Quantity Amount/Hours
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setQuantityAmountHours(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // GlobalAmount
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setGlobalAmount(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ObjectAmount
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setObjectAmount(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ObjectCurrency
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setObjectCurrency(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // TransactionalAmt
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setTransactionalAmt(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // TransactionalCurrency
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setTransactionalCurrency(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ReportingAmount
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setReportingAmount(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ReportingCurrency
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setReportingCurrency(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Enterprise ID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setEnterpriseID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Resource Name
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setResourceName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Personnel number
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setPersonnelNumber(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Current Worklocation
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCurrentWorklocation(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Current Home Office
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCurrentHomeOffice(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCenterID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCostCenterID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCenterName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCostCenterName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCostCollectorID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setCostCollectorName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWBS(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS Description
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWBSDescription(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS ProfitCenter
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWBSProfitCenter(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS ProfitCenterName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWBSProfitCenterName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Parent WBS
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setParentWBS(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ContractID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setContractID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS RaKey
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWBSRaKey(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel1
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel1(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel2
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel2(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel3
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel3(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel4
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel4(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel5
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel5(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel6
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel6(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel7
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel7(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel8
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel8(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel9
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel9(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel10
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioCostTransactionExtract.setWMULevel10(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Contrato
		           relatorioCostTransactionExtract.setContrato(contrato);

		           // Data Extra��o
		           relatorioCostTransactionExtract.setDataExtracao(dataAtualBancoFinal);
		           
		           // Data Refer�ncia
		           Date dataReferencia = formatoDataReferencia.parse(ano + "-" + mes + "-" + "01");
		           java.sql.Date dataReferenciaFinal = new java.sql.Date(dataReferencia.getTime());
		           relatorioCostTransactionExtract.setDataReferencia(dataReferenciaFinal);
		           
		           // Periodo
		           relatorioCostTransactionExtract.setPeriodo(periodo);
		           
		           // Armazeno os per�odos de extra��o distintos
		           //listaPeridosDeExtracaoDistintos.add(relatorioCostTransactionExtract.getPostingPeriod());

		           if (relatorioCostTransactionExtract != null && !Util.relatorioCostTransactionExtractPossuiTodosCamposNulos(relatorioCostTransactionExtract)) {
		        	   listaRelatorioPlanilhaCostTransactionExtract.add(relatorioCostTransactionExtract);
		           }
		           
		       }
		       
		   }
		   
		   // Existem alguns registros da planilha que possuem o campo PostingPeriod em branco
		   // Preencho esses campos com o valor dos outros PostingPeriod que possuem valor, por�m
		   // os salvo no campo Data_Referencia
		   //preencheListaRelatorioTemporariaPlanilhaCostTransactionExtractComDataReferencia(listaPeridosDeExtracaoDistintos, listaRelatorioTemporaria);
		   
		   // Por fim adiciono a listaRelatorioTemporaria na listaRelatorioPlanilhaCostTransactionExtract
		   //listaRelatorioPlanilhaCostTransactionExtract.addAll(listaRelatorioTemporaria);
		   
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de relat�rio n�o encontrado!");
			   throw new Exception("Arquivo Excel de relat�rio n�o encontrado!");
			
			} finally {
			
				if (arquivo != null) {
					arquivo.close();
				}
				
			}
		
			if (listaRelatorioPlanilhaCostTransactionExtract.size() == 0) {
			   //Pode ser que existam arquivos vazios, ent�o n�o posso lan�ar exce��o aqui
				//throw new Exception("Lista de projetos est� vazia");
			}
			
		}
	
	   public static void preencheListaRelatorioTemporariaPlanilhaCostTransactionExtractComDataReferencia(Set<String> listaPeridosDeExtracaoDistintos, List<RelatorioCostTransactionExtract> listaRelatorioTemporaria) throws Exception{
		   
		   String peridoDeExtracao  = "";
		   
		   if (listaPeridosDeExtracaoDistintos != null && !listaPeridosDeExtracaoDistintos.isEmpty()) {
			   
			   for (String dataDeExtracao : listaPeridosDeExtracaoDistintos) {
				   
				   if (dataDeExtracao != null && !dataDeExtracao.isEmpty()) {
					   
					   peridoDeExtracao = dataDeExtracao;
					   break;
					   
				   }
				
			   }
		   }   
			   
		   if (listaRelatorioTemporaria != null && !listaRelatorioTemporaria.isEmpty()) {
			   
			   for (RelatorioCostTransactionExtract relatorioMme : listaRelatorioTemporaria) {
				   
				   if (relatorioMme != null) {
					   
					   //relatorioMme.setDataReferencia(peridoDeExtracao);
					   
				   }
				   
			   }
			   
		   }
			   
	   }
	
	   public static boolean verificaArquivoValido(String item) throws Exception{
		   
			// O arquivo est� demorando para baixar, ent�o precio garantir que s� irei
			// process�-lo quando ele estiver �ntegro e com o nome completo baixado
		   
		   boolean isArquivoValido = true;
			
			int indiceArquivoValido = item.indexOf("Cost");
			while (indiceArquivoValido == -1) {
				isArquivoValido = false;
				Thread.sleep(3000);
				verificaArquivoValido(item);
			}
			isArquivoValido = true;
			
			return isArquivoValido;

	   }
	   
	   public static void listaEMovePlanilhasMultiSegmentContractReport(String caminhoDiretorioOrigem, String caminhoDiretorioDestino, String contrato) throws Exception{
		   
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorioOrigem);
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos arquivos
				String arquivos [] = diretorio.list();
				
				if (arquivos != null && arquivos.length > 0) {
					
					for (String item : arquivos){
						
						String caminhoArquivo = caminhoDiretorioOrigem + Util.getValor("separador.diretorio") + item;
						File arquivo = new File(caminhoArquivo);
						// Se existirem arquivos, os movo para a pasta de sa�da
						if (arquivo.exists() && arquivo.isFile()) {
							Thread.sleep(1000);
							// As vezes o arquivo est� demorando para baixar, ent�o precio garantir que s� irei
							// process�-lo quando ele estiver �ntegro e com o nome completo baixado
							// Ent�o fico tentando at� ele ter baixado completamente 
							int indiceArquivoValido = item.indexOf("Multi-Segment Contract Report");
							if (indiceArquivoValido == -1) {
								contadorErrosArquivoInvalido++;
								if (contadorErrosArquivoInvalido < 30) {
									System.out.println("Arquivo " + item + " ainda n�o est� pronto.  Tentativa de numero: " + contadorErrosArquivoInvalido);
									Thread.sleep(3000);
									listaEMovePlanilhasMultiSegmentContractReport(caminhoDiretorioOrigem, caminhoDiretorioDestino, contrato);
								} else {
									throw new Exception("Arquivo invalido: " + item );
								}
							}
							
							// L� os relat�rios baixados e armazena todas as linhas das planilhas em uma lista
				    		lerPlanilhaMultiSegmentContractReport(caminhoArquivo, contrato);
		            		//Move o relat�rio baixado do diret�rio relatorios para o diret�rio correto
		            		moverArquivosEntreDiretorios(caminhoArquivo, caminhoDiretorioDestino);
						}

					}
				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("N�o existe o diret�rio: " + caminhoDiretorioOrigem);
	        }
	        
	    }

	   
	   public static void listaEMovePlanilhasResourceTrend(String caminhoDiretorioOrigem, String caminhoDiretorioDestino, String contrato, String ano, String mes) throws Exception{
		   
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorioOrigem);
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos arquivos
				String arquivos [] = diretorio.list();
				
				if (arquivos != null && arquivos.length > 0) {
					
					for (String item : arquivos){
						
						String caminhoArquivo = caminhoDiretorioOrigem + Util.getValor("separador.diretorio") + item;
						File arquivo = new File(caminhoArquivo);
						// Se existirem arquivos, os movo para a pasta de sa�da
						if (arquivo.exists() && arquivo.isFile()) {
							Thread.sleep(1000);
							// As vezes o arquivo est� demorando para baixar, ent�o precio garantir que s� irei
							// process�-lo quando ele estiver �ntegro e com o nome completo baixado
							// Ent�o fico tentando at� ele ter baixado completamente 
							int indiceArquivoValido = item.indexOf("Resource Trend");
							if (indiceArquivoValido == -1) {
								contadorErrosArquivoInvalido++;
								if (contadorErrosArquivoInvalido < 30) {
									System.out.println("Arquivo " + item + " ainda n�o est� pronto.  Tentativa de numero: " + contadorErrosArquivoInvalido);
									Thread.sleep(3000);
									listaEMovePlanilhasResourceTrend(caminhoDiretorioOrigem, caminhoDiretorioDestino, contrato, ano, mes);
								} else {
									throw new Exception("Arquivo invalido: " + item );
								}
							}
							
							// L� os relat�rios baixados e armazena todas as linhas das planilhas em uma lista
				    		lerPlanilhaResourceTrend(caminhoArquivo, contrato, ano, mes);
		            		//Move o relat�rio baixado do diret�rio relatorios para o diret�rio correto
		            		moverArquivosEntreDiretorios(caminhoArquivo, caminhoDiretorioDestino);
						}

					}
				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("N�o existe o diret�rio: " + caminhoDiretorioOrigem);
	        }
	        
	    }

	   public static void listaEMovePlanilhasCostTransactionExtract(String caminhoDiretorioOrigem, String caminhoDiretorioDestino, String contrato, int mes, int ano, String periodo) throws Exception{
		   
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorioOrigem);
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos arquivos
				String arquivos [] = diretorio.list();
				
				if (arquivos != null && arquivos.length > 0) {
					
					for (String item : arquivos){
						
						String caminhoArquivo = caminhoDiretorioOrigem + Util.getValor("separador.diretorio") + item;
						File arquivo = new File(caminhoArquivo);
						// Se existirem arquivos, os movo para a pasta de sa�da
						if (arquivo.exists() && arquivo.isFile()) {
							Thread.sleep(1000);
							// As vezes o arquivo est� demorando para baixar, ent�o precio garantir que s� irei
							// process�-lo quando ele estiver �ntegro e com o nome completo baixado
							// Ent�o fico tentando at� ele ter baixado completamente 
							int indiceArquivoValido = item.indexOf("Cost");
							if (indiceArquivoValido == -1) {
								contadorErrosArquivoInvalido++;
								if (contadorErrosArquivoInvalido < 30) {
									System.out.println("Arquivo " + item + " ainda n�o est� pronto.  Tentativa de numero: " + contadorErrosArquivoInvalido);
									Thread.sleep(3000);
									listaEMovePlanilhasCostTransactionExtract(caminhoDiretorioOrigem, caminhoDiretorioDestino, contrato, mes, ano, periodo);
								} else {
									throw new Exception("Arquivo invalido: " + item );
								}
							}
							
							// L� os relat�rios baixados e armazena todas as linhas das planilhas em uma lista
				    		lerPlanilhaCostTransactionExtract(caminhoArquivo, contrato, mes, ano, periodo);
		            		//Move o relat�rio baixado do diret�rio relatorios para o diret�rio correto
		            		moverArquivosEntreDiretorios(caminhoArquivo, caminhoDiretorioDestino);
						}

					}
				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("N�o existe o diret�rio: " + caminhoDiretorioOrigem);
	        }
	        
	    }
	   
	  public static void apagaArquivosDiretorioDeRelatorios(String caminhoDiretorio) throws Exception{
	    	
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorio);
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos arquivos
				String arquivos [] = diretorio.list();
				
				if (arquivos != null && arquivos.length > 0) {
					
					for (String item : arquivos){
						
						File arquivo = new File(caminhoDiretorio + "/" + item);
						// Se existirem arquivos, os deleto
						if (arquivo.exists() && arquivo.isFile()) {
							arquivo.delete();
						}

					}
				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("Nao existe o diretorio: " + caminhoDiretorio);
	        }
	        
	    }
	  
	  public static void apagaDiretoriosDeRelatorios(String caminhoDiretorio) throws Exception{
	    	
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorio);
	        Date dataAtual = new Date();
	        Calendar cal = Calendar.getInstance();
	        cal.setTime(dataAtual);
	        cal.add(Calendar.DATE, -7);
	        Date dataAntes7Dias = cal.getTime();
	        
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos diret�rios
				String itens [] = diretorio.list();
				
				if (itens != null && itens.length > 0) {
					
					for (String item : itens){
						
						File pasta = new File(caminhoDiretorio + "/" + item);
						
						if (pasta.exists() && pasta.isDirectory()) {
							
							Long dataModificacaoPasta =  FileUtils.lastModified(pasta);
							Date dataModificacaoPasta2 = new Date(dataModificacaoPasta);
							
							// Se existirem diret�rios com a data anterior � data de 7 dias atr�s, os deleto
							if (dataModificacaoPasta2.before(dataAntes7Dias)) {
								FileUtils.deleteQuietly(pasta);
							}
							
						}

					}

				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("Nao existe o diretorio: " + caminhoDiretorio);
	        }
	        
	    }
	  
	   public static void mataProcessosFirefox() throws IOException, SQLException, InterruptedException{
		   
		   // Mata o Firefox
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.firefox"));
		   Thread.sleep(3000);
	   
	   }

	   public static void mataProcessosGoogle() throws IOException, SQLException, InterruptedException{
			  
		   // Mata o Google
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.google"));
		   Thread.sleep(3000);
		   
		   // Mata o chromedriver
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.chromedriver"));
		   Thread.sleep(3000);
	   
	   }
	   
	   public static void inserirStatusExecucaoNoBanco(String servico, String dataHora, String status) throws IOException, SQLException{
			  
		   HistoricoExecucaoDao historicoExecucaoDao = new HistoricoExecucaoDao();
		   historicoExecucaoDao.inserirStatusExecucao(servico, dataHora, status);
	   
	   }
	
	public static void abrirSiteMme(WebDriver driver, WebDriverWait wait) throws Exception {

		// Abrindo a URl do Mme
		// Quando o Dani executa o rob� na m�quina remota da Accenture, o bot�o de Select do pop-up de contratos fica escondido
		// Ent�o abro o Chrome em modo full-screen
		// Tamb�m tem que ser colocada essa op��o no m�todo getWebDriver que fica logo mais abaixo do c�digo
		driver.manage().window().maximize();
		driver.get(Util.getValor("url.mme"));
		Thread.sleep(10000);
		
	}
	
	public static void fecharPopUpMensagem(WebDriver driver, WebDriverWait wait) throws Exception {
		
		WebDriverWait waitFecharPopUpMensagem = new WebDriverWait(driver, 5);
		
		try {
			// Clicando no check box da mensagem Don't show this message again
			waitFecharPopUpMensagem.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"acn-modal\"]/div/div[2]/div[1]/label/span"))).click();
			Thread.sleep(1000);
		} catch (Exception e) {
			System.out.println("Provavelmente n�o foi encontrado pop-up para ser fechado");
		}
		
		try {
			// Fechando o pop-up
			waitFecharPopUpMensagem.until(ExpectedConditions.elementToBeClickable(By.className("acn-modal-del-icon"))).click();
			Thread.sleep(1000);
		} catch (Exception e) {
			System.out.println("Provavelmente n�o foi encontrado pop-up para ser fechado");
		}
		
	}

	
	public static void fecharMensagemBarraSuperior(WebDriver driver) throws Exception {

		WebDriverWait waitFecharMensagemBarraSuperior = new WebDriverWait(driver, 5);
		
		try {
			for (int i = 0; i < 2; i++) {
				// Tenta fechar mensagem na parte superior da tela
				// Estou executando 2 vezes porqu� j� vi aparecer duas mensagens
				waitFecharMensagemBarraSuperior.until(ExpectedConditions.elementToBeClickable(By.className("acn-alert-close-icon"))).click();
				Thread.sleep(1000);
			}
		} catch (Exception e) {
			System.out.println("Provavelmente a mensagem na parte superior da tela n�o apareceu");
		}
		
	}
	
	public static void abrirCaixaSelecaoContratos(WebDriver driver, WebDriverWait wait) throws Exception {
		
		fecharMensagemBarraSuperior(driver);
		
		fecharBarraInferiorComInformacoesDaAccenture(driver);
		
		WebDriverWait waitCancel = new WebDriverWait(driver, 5);
		
		try {

			// As vezes a caixa de selec�o de contratos j� aparece na tela
			// Ent�o tento cancelar a caixa
			// Se n�o conseguir cancelar � porque a caixa n�o apareceu, ent�o a abro mais abaixo 
			waitCancel.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Cancel')]"))).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			System.out.println("Provavelmente a caixa de sele��o de contratos n�o apareceu");
		}
		
        // Abrindo a caixa de sele��o de contratos
		try {
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='top-nav-breadcrumb-1']/div/div/button"))).click();
		} catch (Exception e) {
			wait.until(ExpectedConditions.elementToBeClickable(By.className("breadcrumbText"))).click();
		}
		
		Thread.sleep(2000);
	}
	
	public static void passosIniciaisCaixaSelecaoContratos(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Esperando renderizar a tela
		Thread.sleep(5000);
		
		// Limpando a pesquisa no campo de busca do Client Name
		String idCampoPesquisa = "//input[@id='ddlClient']";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idCampoPesquisa))).clear();
		Thread.sleep(3000);

		// Manda o texto TELEFONICA GROUP no campo de busca do Client Name
		String textoTelefonicaGroup = "TELEFONICA GROUP";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idCampoPesquisa))).sendKeys(textoTelefonicaGroup);
		Thread.sleep(3000);
		
		// Seleciona a op��o TELEFONICA GROUP
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='results']/mme-item/span"))).click();
		Thread.sleep(3000);

		// Limpando a pesquisa marcando e desmarcando o checkbox Retired
		for (int i = 0; i <= 1; i++) {
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//mat-checkbox[@id='retiredFilter']/label/span"))).click();
			Thread.sleep(4000);
		}
		
	}
	
	public static void passosIniciaisParaAbrirPesquisaDeContratos(WebDriver driver, WebDriverWait wait) throws Exception {

		abrirSiteMme(driver, wait);
		
		fecharPopUpMensagem(driver, wait);
		
		abrirCaixaSelecaoContratos(driver, wait);
		
		passosIniciaisCaixaSelecaoContratos(driver, wait);
		
	}
	
	public static void advancedReportingMenuTeste(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Acessando o Advanced Reporting
		driver.get(Util.getValor("url.advanced.reporting"));
		Thread.sleep(2000);
		
		fecharPopUpMensagem(driver, wait);
		
	}
	
	public static void advancedReportingMenu(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Clicando no link Update Forecast
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Update Forecast')]"))).click();
		Thread.sleep(2000);
		
		// Clicando na op��o Advanced Reporting do menu
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Advanced Reporting')]"))).click();
		Thread.sleep(2000);
		
		// Obtendo a url do link do relatório clássico
		String idRelatorioClassico = "//*[@id=\"reportingClassic\"]/a";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idRelatorioClassico)));
		WebElement linkRelatorioClassico = driver.findElement(By.xpath(idRelatorioClassico));
		String urlLinkRelatorioClassico = linkRelatorioClassico.getAttribute("href");
		urlLinkRelatorioClassico = urlLinkRelatorioClassico.replaceAll("https://mme.accenture.com//Cross", "https://mme.accenture.com/Cross");
		
		// Acessando o Advanced Reporting
		driver.get(urlLinkRelatorioClassico);
		Thread.sleep(5000);
				
		fecharPopUpMensagem(driver, wait);
		
		// Viegas
		// Quantidade de iframes
		//int size = driver.findElements(By.tagName("iframe")).size();
		
		// O relatório fica na parte de baixo da tela, que é um segundo frame
		// Troco para esse segundo frame para manipular os campos
		//WebDriver iframeMapa = driver.switchTo().frame(size - 1);
		//WebDriverWait waitSegundoIframe = new WebDriverWait(iframeMapa, 20);
		
		//List<String> windowHandles = new ArrayList<>(driver.getWindowHandles());
		//driver.switchTo().window(windowHandles.get(0));
		
		//((JavascriptExecutor) driver).executeScript("window.close()");
		
		// Clicando no link Update Forecast
		//waitSegundoIframe.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Update Forecast')]"))).click();
		
		// Acessando o Advanced Reporting
		//driver.get(Util.getValor("url.advanced.reporting"));
		//Thread.sleep(2000);
		
		//wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='top-nav-breadcrumb-1']/div/div/button"))).click();
		//Thread.sleep(2000);
		// Viegas
		
	}
	
	public static void passosFinaisPesquisaDeContratos(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Clicando no bot�o Select
		String textoSelect= "Select";
		Thread.sleep(5000);
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoSelect+"']"))).click();
		
		WebDriverWait waitCancel = new WebDriverWait(driver, 2);
		try {
			// O bot�o Select est� provavelmente desabilitado porqu� est� pesquisando no projeto j� selecionado anteriormente
			// Por�m mesmo desabilitado, o Selenium encontra o bot�o Select
			// Ent�o clico no bot�o Cancel para o modal sumir
			waitCancel.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Cancel')]"))).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			System.out.println("O bot�o Select est� provavelmente desabilitado porqu� est� pesquisando no projeto j� selecionado anteriormente. Ent�o clico no bot�o Cancel para o modal sumir");
		}
	
	}
	
	public static void contrato_9940191116_SW_Factories(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		//*[@id="tree"]/ez-node/li/ul/ez-node[1]/li/ul/ez-node[13]/li/div[1]/div
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		//*[@id="tree"]/ez-node/li/ul/ez-node[1]/li/ul/ez-node[10]/li/ul/ez-node[7]/li/div[1]/div
		Thread.sleep(3000);

		// Expandindo a op��o Global Village Telecom da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[22]/li/div/div"))).click();
		//*[@id="tree"]/ez-node/li/ul/ez-node[1]/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[22]/li/div[1]/div
		Thread.sleep(3000);

		// Expandindo a op��o GVT SW Factories da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[22]/li/ul/ez-node[2]/li/div/div"))).click();
		Thread.sleep(3000);

		// Clicando na op��o 9940191116 SW Factories da �rvores de op��es
		String id9940191116SwFactories = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[22]/li/ul/ez-node[2]/li/ul/ez-node/li/div[2]/span";
		String contrato = "9940191116 SW Factories";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(id9940191116SwFactories))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);
	}
	
	public static void contrato_Command_Center(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Command Center da �rvores de op��es
		String idAMFaturamento = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[9]/li/div/span";
		String contrato = "Command Center";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idAMFaturamento))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_AM_Faturamento(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o AM Faturamento da �rvores de op��es
		String idAMFaturamento = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[10]/li/div/span";
		String contrato = "AM Faturamento";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idAMFaturamento))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_B2C_SFA_Salesforce(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		// Clicando na op��o B2C SFA - Salesforce da �rvores de op��es
		String idB2C_SFA_Salesforce = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[13]/li/div[2]/span";
		String contrato = "B2C SFA - Salesforce";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idB2C_SFA_Salesforce))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_Callidus(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Callidus da �rvores de op��es
		String idCallidus = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[14]/li/div[2]/span";
		String contrato = "Callidus";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idCallidus))).click();
		Thread.sleep(5000);
		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_GVT_Proforma(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o Global Village Telecom da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[22]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o GVT Proforma da �rvores de op��es
		String idGVTProforma = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[7]/li/ul/ez-node[20]/li/ul/ez-node/li/div[2]/span";
		String contrato = "GVT Proforma";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idGVTProforma))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}

	public static void contrato_Hybris_eCommerce(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o Hybris - eCommerce da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o SCO - Ecommerce da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[7]/li/ul/ez-node[12]/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Hybris - eCommerce da �rvores de op��es
		String idHybris_eCommerce = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[7]/li/ul/ez-node[12]/li/ul/ez-node/li/ul/ez-node/li/div/span";
		String contrato = "Hybris - eCommerce";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idHybris_eCommerce))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_Digital_Factory(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		// Expandindo a op��o Telef�nica - XBD Digital Factory da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[32]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Digital Factory da �rvores de op��es
		String idDigitalFactory = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[32]/li/ul/ez-node/li/div[2]/span";
		String contrato = "Digital Factory";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idDigitalFactory))).click();
		Thread.sleep(5000);
		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_AM_Latam_Brasil_Fija_Contrato_Local(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o APOLLO LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o AM LATAM - Master Contract 9940089177 da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[8]/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o AM Latam Brasil Fija - Contrato Local da �rvores de op��es
		String idAmLatamBrasilFija = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[8]/li/ul/ez-node/li/ul/ez-node[4]/li/div[2]/span";
		String contrato = "AM Latam Brasil Fija - Contrato Local";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idAmLatamBrasilFija))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_Nova_Fabrica_Design(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Nova Fabrica Design da �rvores de op��es
		String idNovaFabricaDesign = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[24]/li/div[2]/span";
		String contrato = "Nova Fabrica Design";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idNovaFabricaDesign))).click();
		Thread.sleep(5000);
		
		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_Desligue_Do_Atis(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		// Clicando na op��o Desligue Do Atis da �rvores de op��es
		String idDesligueDoAtis = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[17]/li/div[2]/span";
		String contrato = "Desligue Do Atis";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idDesligueDoAtis))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);
	}
	
	public static void contrato_Portal_Terra(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Portal Terra da �rvores de op��es
		String idPortalTerra = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[25]/li/div[2]/span";
		String contrato = "Portal Terra";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idPortalTerra))).click();
		Thread.sleep(5000);
		passosFinaisPesquisaDeContratos(driver, wait);
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_Sustentacao_VIVO_GO(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		// Clicando na op��o Sustentacao VIVO GO da �rvores de op��es
		String idSustentacaoVIVOGO = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[31]/li/div/span";
		String contrato = "Sustentacao VIVO GO";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idSustentacaoVIVOGO))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);
		
	}
	
	public static void contrato_FiberCo_Imp_Arquitet_BSS_OSS(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o FiberCo Imp. Arquitet.BSS/OSS da �rvores de op��es
		String idFiberCo = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[19]/li/div[2]/span";
		String contrato = "FiberCo Imp. Arquitet.BSS/OSS";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idFiberCo))).click();
		Thread.sleep(5000);
		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_RPA_Blue_Prism(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		// Clicando na op��o RPA - Blue Prism da �rvores de op��es
		String idRpaBluePrism = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[27]/li/div[2]/span";
		String contrato = "RPA - Blue Prism";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idRpaBluePrism))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_Protecao_De_Dados(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		// Clicando na op��o Prote��o de dados da �rvores de op��es
		String idProtecaoDeDados = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[26]/li/div[2]/span";
		String contrato = "Prote��o de dados";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idProtecaoDeDados))).click();
		Thread.sleep(5000);
		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}
	
	public static void contrato_B2B_Transformation(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o B2B Transformation da �rvores de op��es
		String id_B2B_Transformation = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[11]/li/ul/ez-node[6]/li/ul/ez-node[12]/li/div[2]/span";
		String contrato = "B2B Transformation";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(id_B2B_Transformation))).click();
		Thread.sleep(5000);
		passosFinaisPesquisaDeContratos(driver, wait);
		
		listaNomesDeContratosDistintos.add(contrato);
		
		extracaoRelatorios(driver, wait, contrato);

	}

	public static void extracaoRelatorios(WebDriver driver, WebDriverWait wait, String contrato) throws Exception {
		
		advancedReportingMenu(driver, wait);
				
		if (Util.getValor("relatorio.cost.transaction.extract").equals("S")) {
			extracaoPlanilhaCostTransactionExtract(driver, wait, contrato);
		}

		if (Util.getValor("relatorio.resource.trend").equals("S")) {
			extracaoPlanilhaResourceTrend(driver, wait, contrato);
		}
		
		if (Util.getValor("relatorio.multi.segment.contract.report").equals("S")) {
			extracaoPlanilhaMultiSegmentContractReport(driver, wait, contrato);
		}
		
	}

	public static void extracaoPlanilhaMultiSegmentContractReport(WebDriver driver, WebDriverWait wait, String contrato) throws Exception {
		
		Thread.sleep(10000);
		
		// Abrindo o menu Selection
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_ImgBtnTree']"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Forecast Analysis 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Forecast Analysis')]"))).click();
		Thread.sleep(2000);	
		
		// Abrindo a op��o Multi-Segment Contract Report 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Multi-Segment Contract Report')]"))).click();
		Thread.sleep(5000);
			
		// Combo Baseline Version
		String idComboBaselineVersion = "//select[@id='ctl00_ctl00_c_dc_cntlForecastVersionAndChangeOrder_fvc_FV_ForecastVersion_dd']";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboBaselineVersion)));
		WebElement comboBaselineVersion = driver.findElement(By.xpath(idComboBaselineVersion));
		// Elementos do combo
		Select elementosComboBaselineVersion = new Select(comboBaselineVersion);
		String textoMasterActive = "Master Active";
		String mesAnoMasterActive = "";
		boolean encontrouMasterActive = false;
		for (WebElement elemento : elementosComboBaselineVersion.getOptions()) {
			if (elemento.getText().trim().contains(textoMasterActive)) {
				elementosComboBaselineVersion.selectByVisibleText(elemento.getText());
				mesAnoMasterActive = elemento.getText();
				encontrouMasterActive = true;
				break;
			}
		}
		
		if (encontrouMasterActive) {
			
			int mesMasterActive = retornaNumeroDoMes(mesAnoMasterActive);
			int anoMasterActive = retornaNumeroDoAno(mesMasterActive);
			extracaoPlanilhaMultiSegmentContractReportPorMesEAno(driver, wait, contrato, mesMasterActive, anoMasterActive);
			
		}

	}
	
	public static void extracaoPlanilhaMultiSegmentContractReportPorMesEAno(WebDriver driver, WebDriverWait wait, String contrato, int mesMasterActive, int anoMasterActive) throws Exception {
		
		String mes = "";
		String ano = String.valueOf(anoMasterActive);
		if (Util.getValor("termos.em.ingles").equals("S")) {
			mes = retornaMesPorExtensoIngles(mesMasterActive);
		} else {
			mes = retornaMesPorExtensoEspanhol(mesMasterActive);
		}
		
		RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport = new RelatorioMultiSegmentContractReport();
		relatorioMultiSegmentContractReport.setContrato(contrato);
		relatorioMultiSegmentContractReport.setAnoMasterActive(ano);
		relatorioMultiSegmentContractReport.setMesMasterActive(String.valueOf(mesMasterActive));
		listaMesAnoMultiSegmentContractReport.add(relatorioMultiSegmentContractReport);
		
		escolheDataCalendarioEExportaPlanilhaMultiSegmentContractReport(driver, wait, mes, ano);
		
		// Movo todos os arquivos baixados para o diret�rio corrente de relat�rios
		// Tamb�m l� os relat�rios baixados e armazena todas as linhas das planilhas na listaRelatorio
		listaEMovePlanilhasMultiSegmentContractReport(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, contrato);
		
	}
	
	public static int retornaNumeroMesMasterActiveResourceTrend(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Abrindo o menu Selection
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_ImgBtnTree']"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Resource Planning and Management 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Resource Planning and Management')]"))).click();
		Thread.sleep(2000);	
		
		// Abrindo a op��o Resource Trend 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Resource Trend')]"))).click();
		Thread.sleep(5000);
			
		// Combo Forecast Version
		String idComboForecastVersion = "//select[@id='ctl00_ctl00_c_dc_cntlForecastVersionAndChangeOrder_fvc_FV_ForecastVersion_dd']";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboForecastVersion)));
		WebElement comboForecastVersion = driver.findElement(By.xpath(idComboForecastVersion));
		// Elementos do combo
		Select elementosComboForecastVersion = new Select(comboForecastVersion);
		String textoMasterActive = "Master Active";
		String mesAnoMasterActive = "";
		boolean encontrouMasterActive = false;
		for (WebElement elemento : elementosComboForecastVersion.getOptions()) {
			if (elemento.getText().trim().contains(textoMasterActive)) {
				elementosComboForecastVersion.selectByVisibleText(elemento.getText());
				mesAnoMasterActive = elemento.getText();
				encontrouMasterActive = true;
				break;
			}
		}
		
		int mesMasterActive = 0;
		
		if (encontrouMasterActive) {
			mesMasterActive = retornaNumeroDoMes(mesAnoMasterActive);
		}
		return mesMasterActive;
		
	}
	
	public static void extracaoPlanilhaResourceTrend(WebDriver driver, WebDriverWait wait, String contrato) throws Exception {
		
		Thread.sleep(10000);
		
		// Abrindo o menu Selection
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_ImgBtnTree']"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Resource Planning and Management 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Resource Planning and Management')]"))).click();
		Thread.sleep(2000);	
		
		// Abrindo a op��o Resource Trend 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Resource Trend')]"))).click();
		Thread.sleep(5000);
			
		// Combo Forecast Version
		String idComboForecastVersion = "//select[@id='ctl00_ctl00_c_dc_cntlForecastVersionAndChangeOrder_fvc_FV_ForecastVersion_dd']";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboForecastVersion)));
		WebElement comboForecastVersion = driver.findElement(By.xpath(idComboForecastVersion));
		// Elementos do combo
		Select elementosComboForecastVersion = new Select(comboForecastVersion);
		String textoMasterActive = "Master Active";
		String mesAnoMasterActive = "";
		boolean encontrouMasterActive = false;
		for (WebElement elemento : elementosComboForecastVersion.getOptions()) {
			if (elemento.getText().trim().contains(textoMasterActive)) {
				elementosComboForecastVersion.selectByVisibleText(elemento.getText());
				mesAnoMasterActive = elemento.getText();
				encontrouMasterActive = true;
				break;
			}
		}
		
		if (encontrouMasterActive) {
			
			String mes = "";
			int mesMasterActive = retornaNumeroDoMes(mesAnoMasterActive);
			String ano = String.valueOf(retornaNumeroDoAno(mesMasterActive));
			if (Util.getValor("termos.em.ingles").equals("S")) {
				mes = retornaMesPorExtensoIngles(mesMasterActive);
			} else {
				mes = retornaMesPorExtensoEspanhol(mesMasterActive);
			}
			
			RelatorioResourceTrend relatorioResourceTrend = new RelatorioResourceTrend();
			relatorioResourceTrend.setContrato(contrato);
			relatorioResourceTrend.setAno(ano);
			relatorioResourceTrend.setMes(String.valueOf(mesMasterActive));
			listaMesAnoResourceTrend.add(relatorioResourceTrend);
			
			escolheDataCalendarioEExportaPlanilhaResourceTrend(driver, wait, mes, ano);
			
			// Movo todos os arquivos baixados para o diret�rio corrente de relat�rios
			// Tamb�m l� os relat�rios baixados e armazena todas as linhas das planilhas na listaRelatorio
			listaEMovePlanilhasResourceTrend(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, contrato, ano, String.valueOf(mesMasterActive));

		}

	}

	public static void extracaoPlanilhaCostTransactionExtract(WebDriver driver, WebDriverWait wait, String contrato) throws Exception {
		
		Thread.sleep(10000);
		
		// Recupera o n�mero do m�s do relat�rio Resource Trend
		//int mesMasterActiveResourceTrend = retornaNumeroMesMasterActiveResourceTrend(driver, wait);
		//Thread.sleep(2000);
		
		// Abrindo o menu Selection
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_ImgBtnTree']"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Cost Planning and Management 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cost Planning and Management')]"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Cost Transaction Extract(Everything Report) 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cost Transaction Extract(Everything Report)')]"))).click();
		Thread.sleep(5000);
		
		if ("S".equals(Util.getValor("exporta.anos.meses.anteriores"))) {
		
			exportarAnosMesesAnteriores(driver, wait, contrato);
		
		} else {
			
			// Exportando a planilha com o m�s atual
			escolheDataCalendarioEExportaPlanilhaCostTransactionExtract(driver, wait, contrato, mesAtual, anoAtual);
			
			// Exportando a planilha com o m�s anterior
			escolheDataCalendarioEExportaPlanilhaCostTransactionExtract(driver, wait, contrato, mesAnterior, anoCorreto);

		}
		
	}
	
	public static void exportarAnosMesesAnteriores (WebDriver driver, WebDriverWait wait, String contrato) throws Exception {
		
		// Exportando a planilha com os anos anteriores e com todos os meses
		
		String[] partesAnos = Util.getValor("anos.exportacao.planilha").split(",");
		
		if (partesAnos.length > 0) {
			
			for (String ano : partesAnos) {
				
				// Todos os meses
				for (int mes = 1; mes <=12; mes++) {
					
					// Existem alguns meses de alguns anos de alguns contratos que d�o erro na hora de extrair do Mme
					// Ent�o os ignoro
					// Tamb�m ignoro alguns anos que n�o existem para alguns contratos
					boolean excecao = excecoesExtracaoRelatorios(contrato, mes, Integer.parseInt(ano));
					
					if (excecao) {
						continue;
					}
					
					escolheDataCalendarioEExportaPlanilhaCostTransactionExtract(driver, wait, contrato, mes, Integer.parseInt(ano));
				}
				
			}
			
		}
	
	}
	
	public static boolean excecoesExtracaoRelatorios (String contrato, int mes, int ano) throws Exception {
		
		boolean excecoesExtracaoRelatorios = false;
		
		boolean excecao1 = "GVT Proforma".equals(contrato) && mes == 2 && ano == 2013;

		boolean excecao2 = "9940191116 SW Factories".equals(contrato) && mes == 2 && ano == 2015;
		boolean excecao3 = "9940191116 SW Factories".equals(contrato) && (ano == 2013 || ano == 2014);
		
		boolean excecao4 = "AM Faturamento".equals(contrato) && (ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017 || ano == 2018 || ano == 2019);
		boolean excecao5 = "AM Faturamento".equals(contrato) && ano == 2020 && ( mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7 || mes == 8 || mes == 9);

		boolean excecao6 = "B2C SFA - Salesforce".equals(contrato) && (ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017);
		boolean excecao7 = "B2C SFA - Salesforce".equals(contrato) && ano == 2018 && ( mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5);
		
		boolean excecao8 = "Callidus".equals(contrato) && (ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017 || ano == 2018);
		boolean excecao9 = "Callidus".equals(contrato) && ano == 2019 && ( mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7 || mes == 8 || mes == 9 || mes == 10);
		
		boolean excecao10 = "Hybris - eCommerce".equals(contrato) && mes == 8 && ano == 2014;
		boolean excecao11 = "Hybris - eCommerce".equals(contrato) && (ano == 2013);
		boolean excecao12 = "Hybris - eCommerce".equals(contrato) && ano == 2014 && ( mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7);
		
		boolean excecao13 = "Digital Factory".equals(contrato) && (ano == 2013 || ano == 2014 || ano == 2015);
		boolean excecao14 = "Digital Factory".equals(contrato) && ano == 2016 && ( mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7 || mes == 8 || mes == 9 || mes == 10);
		
		boolean excecao15 = "Nova Fabrica Design".equals(contrato) && (ano == 2009 || ano == 2010 || ano == 2011 || ano == 2012 || ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017 || ano == 2018 || ano == 2019);
		boolean excecao16 = "Nova Fabrica Design".equals(contrato) && ano == 2020 && ( mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7 || mes == 8 || mes == 9);
		
		boolean excecao17 = "Desligue Do Atis".equals(contrato) && (ano == 2009 || ano == 2010 || ano == 2011 || ano == 2012 || ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017 || ano == 2018 || ano == 2019 || ano == 2020);
		boolean excecao18 = "Desligue Do Atis".equals(contrato) && ano == 2021 && ( mes == 1 || mes == 2);
		
		boolean excecao19 = "Portal Terra".equals(contrato) && (ano == 2009 || ano == 2010 || ano == 2011 || ano == 2012 || ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017 || ano == 2018 || ano == 2019);
		boolean excecao20 = "Portal Terra".equals(contrato) && ano == 2020 && (  mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7  || mes == 8);
		
		boolean excecao21 = "Sustentacao VIVO GO".equals(contrato) && (ano == 2009 || ano == 2010 || ano == 2011 || ano == 2012 || ano == 2013 || ano == 2014 || ano == 2015 || ano == 2016 || ano == 2017 || ano == 2018 || ano == 2019);
		boolean excecao22 = "Sustentacao VIVO GO".equals(contrato) && ano == 2020 && (  mes == 1 || mes == 2 || mes == 3 || mes == 4 || mes == 5 || mes == 6 || mes == 7 || mes == 8 );
		
		boolean excecao23 = "AM Latam Brasil Fija - Contrato Local".equals(contrato) && ano == 2009 && (  mes == 1 || mes == 2 || mes == 3 || mes == 4 );
		boolean excecao24 = "FiberCo Imp. Arquitet.BSS/OSS".equals(contrato) && ano == 2021 && (  mes == 1 || mes == 2 || mes == 3 || mes == 4  || mes == 5  || mes == 6  || mes == 7 );
		boolean excecao25 = "RPA - Blue Prism".equals(contrato) && ano == 2021 && (  mes == 1 || mes == 2 || mes == 3 || mes == 4  || mes == 5  || mes == 6  || mes == 7  || mes == 8  || mes == 9 );
		
		// Ano Atual
		boolean excecao26 = ano == 2022 && (mes == 6 || mes == 7 || mes == 8 || mes == 9 || mes == 10 || mes == 11 || mes == 12);
		
		if (excecao1 || excecao2 || excecao3 || excecao4 || excecao5 || excecao6 || excecao7 || excecao8 || excecao9 || excecao10 || excecao11 || excecao12 || excecao13 || excecao14 || excecao15 || excecao16 || excecao17 || excecao18 || excecao19 || excecao20 || excecao21 || excecao22 || excecao23 || excecao24 || excecao25 || excecao26) {
		
			excecoesExtracaoRelatorios = true;
		}
		
		return excecoesExtracaoRelatorios;

	}
	
    public static void clickNoCalendarioCostTransactionExtract (WebDriver driver, WebDriverWait wait) throws Exception {
    	
    	try {
    		
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Calendar']"))).click();
    		Thread.sleep(2000);    		
    	
    	} catch (Exception e) {
    		
    		contadorErrosCalendario ++;
			
            // Tento escolher clicar no calend�rio por 20 vezes
            if (contadorErrosCalendario <= 20) {
            	
            	System.out.println("Erro ao clicar no calendario.Tentativa de numero: " + contadorErrosCalendario);
            	clickNoCalendarioCostTransactionExtract(driver, wait);
            
            } else {
         	   throw new Exception("Erro ao clicar no calendario: " + e);
            }
    		
    	}
    	
    }

	public static void escolheDataCalendarioEExportaPlanilhaCostTransactionExtract(WebDriver driver, WebDriverWait wait, String contrato, int mes, int ano) throws Exception {
		
		fecharMensagemBarraSuperior(driver);
		
		fecharBarraInferiorComInformacoesDaAccenture(driver);
		
		rolagemParaBaixo(driver);
		
		// Clicando no calend�rio
		clickNoCalendarioCostTransactionExtract(driver, wait);
		
		// Ano
		String idComboAno = "//div[@id='ui-datepicker-div']/div/div/select[2]";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboAno)));
		WebElement comboAno = driver.findElement(By.xpath(idComboAno));
		// Elementos do combo
		Select elementoscomboAno = new Select(comboAno);
		String anoPorExtenso = String.valueOf(ano);
		boolean encontrouAno = false;
		for (WebElement elemento : elementoscomboAno.getOptions()) {
			if (anoPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementoscomboAno.selectByVisibleText(anoPorExtenso);
				encontrouAno = true;
				break;
			}
		}
		Thread.sleep(1000);
		
		// M�s
		String idComboMes = "//div[@id='ui-datepicker-div']/div/div/select";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboMes)));
		WebElement comboMes = driver.findElement(By.xpath(idComboMes));
		// Elementos do combo
		Select elementosComboMes = new Select(comboMes);
		String mesPorExtenso = "";
		if (Util.getValor("termos.em.ingles").equals("S")) {
			mesPorExtenso = retornaMesPorExtensoIngles(mes);
   		} else {
   			mesPorExtenso = retornaMesPorExtensoEspanhol(mes);
   		}

		boolean encontrouMes = false;
		for (WebElement elemento : elementosComboMes.getOptions()) {
			if (mesPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementosComboMes.selectByVisibleText(mesPorExtenso);
				encontrouMes = true;
				break;
			}
		}
		Thread.sleep(1000);
		
		// Clicando no bot�o Fechar do calend�rio
		botaoFecharCalendario(wait);
		
		String periodo = "Actual";
		
		if (encontrouAno && encontrouMes) {
			
			// M�s e ano atual
			if (mes == mesAtual && ano == anoAtual) {
			
				// M�s ainda n�o est� fechado
				periodo = "Forecast";
			
			} else if (mes == mesAnterior && ano == anoCorreto) {
				
				Calendar cal = Calendar.getInstance();
				cal.setTime(new Date());
						 
				// Se o m�s a ser extra�do for o anterior e se o dia atual for menor ao 8� dia �til do m�s atual,
				// ent�o a extra��o do m�s anterior ainda n�o j� fechou, logo � Forecast
				// N�o estou considerando os dias �teis, somente o dia do m�s mesmo
				if (cal.get(Calendar.DAY_OF_MONTH) < 8 ) { // dia menor que o dia 8
					
					periodo = "Forecast";	
				
				} else {

					periodo = "Actual";
				
				}
				
			}

			// Clicando no bot�o Export
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_btnRunReport']"))).click();
			Thread.sleep(Integer.parseInt(Util.getValor("tempo.segundos.download.planilha")) * 1000);
			
			// Movo todos os arquivos baixados para o diret�rio corrente de relat�rios
			// Tamb�m l� os relat�rios baixados e armazena todas as linhas das planilhas na listaRelatorio
			listaEMovePlanilhasCostTransactionExtract(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, contrato, mes, ano, periodo);
			
		}
		
	}
	
	 public static void botaoFecharCalendario (WebDriverWait wait) throws Exception {
		 
		// Clicando no bot�o Fechar do calend�rio
		 if (Util.getValor("termos.em.ingles").equals("S")) {
			 wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(.,'Done')]"))).click();
		 } else {
			 wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(.,'Fechar')]"))).click();
		 }
		 
		Thread.sleep(2000);
	 }
	
    public static boolean calendarioResourceTrendTo (WebDriver driver, WebDriverWait wait, String mes, String ano) throws Exception {
    	
    	boolean encontrouAno_E_EncontrouMes = false;
    	
    	// Clicando no calend�rio com label To
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_cntlTimeFrameControl_tf_Time_To_cal_dateBox']"))).click();
    	Thread.sleep(2000);
    	
		// Ano
		String idComboAno = "//div[@id='ui-datepicker-div']/div/div/select[2]";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboAno)));
		WebElement comboAno = driver.findElement(By.xpath(idComboAno));
		// Elementos do combo
		Select elementoscomboAno = new Select(comboAno);
		String anoPorExtenso = String.valueOf(ano);
		boolean encontrouAno = false;
		for (WebElement elemento : elementoscomboAno.getOptions()) {
			if (anoPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementoscomboAno.selectByVisibleText(anoPorExtenso);
				encontrouAno = true;
				break;
			}
		}
		Thread.sleep(1000);
		
		// M�s
		String idComboMes = "//div[@id='ui-datepicker-div']/div/div/select";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboMes)));
		WebElement comboMes = driver.findElement(By.xpath(idComboMes));
		// Elementos do combo
		Select elementosComboMes = new Select(comboMes);
		String mesPorExtenso = mes;
		boolean encontrouMes = false;
		for (WebElement elemento : elementosComboMes.getOptions()) {
			if (mesPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementosComboMes.selectByVisibleText(mesPorExtenso);
				encontrouMes = true;
				break;
			}
		}
		Thread.sleep(1000);

		// Clicando no bot�o Fechar do calend�rio
		botaoFecharCalendario(wait);		
		
		if (encontrouAno && encontrouMes) {
			encontrouAno_E_EncontrouMes = true;
		}
		
		return encontrouAno_E_EncontrouMes;
    	
    }
    
    public static boolean calendarioResourceTrendFrom (WebDriver driver, WebDriverWait wait, String mes, String ano) throws Exception {
    	
    	boolean encontrouAno_E_EncontrouMes = false;
    	
    	// Clicando no calend�rio com label From
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_cntlTimeFrameControl_tf_Time_From_cal_dateBox']"))).click();
    	Thread.sleep(2000);
    	
		// Ano
		String idComboAno = "//div[@id='ui-datepicker-div']/div/div/select[2]";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboAno)));
		WebElement comboAno = driver.findElement(By.xpath(idComboAno));
		// Elementos do combo
		Select elementoscomboAno = new Select(comboAno);
		String anoPorExtenso = String.valueOf(ano);
		boolean encontrouAno = false;
		for (WebElement elemento : elementoscomboAno.getOptions()) {
			if (anoPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementoscomboAno.selectByVisibleText(anoPorExtenso);
				encontrouAno = true;
				break;
			}
		}
		Thread.sleep(1000);
		
		// M�s
		String idComboMes = "//div[@id='ui-datepicker-div']/div/div/select";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboMes)));
		WebElement comboMes = driver.findElement(By.xpath(idComboMes));
		// Elementos do combo
		Select elementosComboMes = new Select(comboMes);
		String mesPorExtenso = mes;
		boolean encontrouMes = false;
		for (WebElement elemento : elementosComboMes.getOptions()) {
			if (mesPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementosComboMes.selectByVisibleText(mesPorExtenso);
				encontrouMes = true;
				break;
			}
		}
		Thread.sleep(1000);

		// Clicando no bot�o Fechar do calend�rio
		botaoFecharCalendario(wait);		
		
		if (encontrouAno && encontrouMes) {
			encontrouAno_E_EncontrouMes = true;
		}
		
		return encontrouAno_E_EncontrouMes;
    	
    }

	public static void escolheDataCalendarioEExportaPlanilhaResourceTrend(WebDriver driver, WebDriverWait wait, String mes, String ano) throws Exception {
		
		fecharMensagemBarraSuperior(driver);
		
		fecharBarraInferiorComInformacoesDaAccenture(driver);
		
		rolagemParaBaixo(driver);
		
		// Calend�rio do label To
		boolean encontrouMesAnoCalendarioResourceTrendTo = calendarioResourceTrendTo(driver, wait, mes, ano);
		
		// Calend�rio do label From
		boolean encontrouMesAnoCalendarioResourceTrendFrom = calendarioResourceTrendFrom(driver, wait, mes, ano);
		
		if (encontrouMesAnoCalendarioResourceTrendTo && encontrouMesAnoCalendarioResourceTrendFrom) {
			// Clicando no bot�o Export
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_btnRunReport']"))).click();
			Thread.sleep(Integer.parseInt(Util.getValor("tempo.segundos.download.planilha")) * 1000);
		}
		
	}
	
    public static boolean calendarioMultiSegmentContractReportTo (WebDriver driver, WebDriverWait wait, String mes, String ano) throws Exception {
    	
    	boolean encontrouAno_E_EncontrouMes = false;
    	
    	// Clicando no calend�rio com label To
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_cntlTimeFrameControl_tf_Time_To_cal_dateBox']"))).click();
    	Thread.sleep(2000);
    	
		// Ano
		String idComboAno = "//div[@id='ui-datepicker-div']/div/div/select[2]";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboAno)));
		WebElement comboAno = driver.findElement(By.xpath(idComboAno));
		// Elementos do combo
		Select elementoscomboAno = new Select(comboAno);
		String anoPorExtenso = String.valueOf(ano);
		boolean encontrouAno = false;
		for (WebElement elemento : elementoscomboAno.getOptions()) {
			if (anoPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementoscomboAno.selectByVisibleText(anoPorExtenso);
				encontrouAno = true;
				break;
			}
		}
		Thread.sleep(1000);
		
		// M�s
		String idComboMes = "//div[@id='ui-datepicker-div']/div/div/select";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboMes)));
		WebElement comboMes = driver.findElement(By.xpath(idComboMes));
		// Elementos do combo
		Select elementosComboMes = new Select(comboMes);
		String mesPorExtenso = mes;
		boolean encontrouMes = false;
		for (WebElement elemento : elementosComboMes.getOptions()) {
			if (mesPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementosComboMes.selectByVisibleText(mesPorExtenso);
				encontrouMes = true;
				break;
			}
		}
		Thread.sleep(1000);

		// Clicando no bot�o Fechar do calend�rio
		botaoFecharCalendario(wait);		
		
		if (encontrouAno && encontrouMes) {
			encontrouAno_E_EncontrouMes = true;
		}
		
		return encontrouAno_E_EncontrouMes;
    	
    }
    
    public static boolean calendarioMultiSegmentContractReportFrom (WebDriver driver, WebDriverWait wait, String mes, String ano) throws Exception {
    	
    	boolean encontrouAno_E_EncontrouMes = false;
    	
    	// Clicando no calend�rio com label From
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_cntlTimeFrameControl_tf_Time_From_cal_dateBox']"))).click();
    	Thread.sleep(2000);
    	
		// Ano
		String idComboAno = "//div[@id='ui-datepicker-div']/div/div/select[2]";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboAno)));
		WebElement comboAno = driver.findElement(By.xpath(idComboAno));
		// Elementos do combo
		Select elementoscomboAno = new Select(comboAno);
		String anoPorExtenso = String.valueOf(ano);
		boolean encontrouAno = false;
		for (WebElement elemento : elementoscomboAno.getOptions()) {
			if (anoPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementoscomboAno.selectByVisibleText(anoPorExtenso);
				encontrouAno = true;
				break;
			}
		}
		Thread.sleep(1000);
		
		// M�s
		String idComboMes = "//div[@id='ui-datepicker-div']/div/div/select";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboMes)));
		WebElement comboMes = driver.findElement(By.xpath(idComboMes));
		// Elementos do combo
		Select elementosComboMes = new Select(comboMes);
		String mesPorExtenso = mes;
		boolean encontrouMes = false;
		for (WebElement elemento : elementosComboMes.getOptions()) {
			if (mesPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementosComboMes.selectByVisibleText(mesPorExtenso);
				encontrouMes = true;
				break;
			}
		}
		Thread.sleep(1000);

		// Clicando no bot�o Fechar do calend�rio
		botaoFecharCalendario(wait);		
		
		if (encontrouAno && encontrouMes) {
			encontrouAno_E_EncontrouMes = true;
		}
		
		return encontrouAno_E_EncontrouMes;
    	
    }
	
	public static void escolheDataCalendarioEExportaPlanilhaMultiSegmentContractReport(WebDriver driver, WebDriverWait wait, String mes, String ano) throws Exception {
		
		fecharMensagemBarraSuperior(driver);
		
		fecharBarraInferiorComInformacoesDaAccenture(driver);
		
		rolagemParaBaixo(driver);
		
		// Calend�rio do label From
		boolean encontrouMesAnoCalendarioMultiSegmentContractReportFrom = calendarioMultiSegmentContractReportFrom(driver, wait, mes, ano);

		// Calend�rio do label To
		// Nova regra para os relat�rios Multi Segment Contract Report
		// Iremos extrair no campo "To" tr�s anos para frente
		int anoInteiro = Integer.parseInt(ano) + 3;
		boolean encontrouMesAnoCalendarioMultiSegmentContractReportTo = calendarioMultiSegmentContractReportTo(driver, wait, mes, String.valueOf(anoInteiro));
		
		if (encontrouMesAnoCalendarioMultiSegmentContractReportTo && encontrouMesAnoCalendarioMultiSegmentContractReportFrom) {
			// Clicando no bot�o Export
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_btnRunReport']"))).click();
			Thread.sleep(Integer.parseInt(Util.getValor("tempo.segundos.download.planilha")) * 1000);
		}
		
	}

	public static String retornaMesPorExtensoEspanhol(int mes){
		
		String mesExtenso = null;
		
		switch(mes) {
	 	 
		 case 1:  mesExtenso = "Jan"; break;
	 	 case 2:  mesExtenso = "Fev"; break;
	 	 case 3:  mesExtenso = "Mar"; break;
	 	 case 4:  mesExtenso = "Abr"; break;
	 	 case 5:  mesExtenso = "Mai"; break;
	 	 case 6:  mesExtenso = "Jun"; break;
	 	 case 7:  mesExtenso = "Jul"; break;
	 	 case 8:  mesExtenso = "Ago"; break;
	 	 case 9:  mesExtenso = "Set"; break;
	 	 case 10: mesExtenso = "Out"; break;
	 	 case 11: mesExtenso = "Nov"; break;
	 	 case 12: mesExtenso = "Dez"; break;
		
		}
		
		return mesExtenso;
	}
	
	public static String retornaMesPorExtensoIngles(int mes){
		
		String mesExtenso = null;
		
		switch(mes) {
	 	 
		 case 1:  mesExtenso = "Jan"; break;
	 	 case 2:  mesExtenso = "Feb"; break;
	 	 case 3:  mesExtenso = "Mar"; break;
	 	 case 4:  mesExtenso = "Apr"; break;
	 	 case 5:  mesExtenso = "May"; break;
	 	 case 6:  mesExtenso = "Jun"; break;
	 	 case 7:  mesExtenso = "Jul"; break;
	 	 case 8:  mesExtenso = "Aug"; break;
	 	 case 9:  mesExtenso = "Sep"; break;
	 	 case 10: mesExtenso = "Oct"; break;
	 	 case 11: mesExtenso = "Nov"; break;
	 	 case 12: mesExtenso = "Dec"; break;
		
		}
		
		return mesExtenso;
	}
	
	public static String retornaMesPorExtenso(String mes){
		
		String mesExtenso = null;
		
		if (mes.contains("Enero") || mes.contains("January") || mes.contains("01") || mes.contains("/01/")) {
			mesExtenso = "Jan";	
		} else if (mes.contains("Febrero") || mes.contains("February") || mes.contains("02") || mes.contains("/02/")) {
			mesExtenso = "Fev";
		} else if (mes.contains("Marzo") || mes.contains("March") || mes.contains("03") || mes.contains("/03/")) {
			mesExtenso = "Mar";	
		} else if (mes.contains("Abril") || mes.contains("April") || mes.contains("04") || mes.contains("/04/")) {
			mesExtenso = "Abr";
		} else if (mes.contains("Mayo") || mes.contains("May") || mes.contains("05") || mes.contains("/05/")) {
			mesExtenso = "Mai";
		} else if (mes.contains("Junio") || mes.contains("June") || mes.contains("06") || mes.contains("/06/")) {
			mesExtenso = "Jun";
		} else if (mes.contains("Julio") || mes.contains("July") || mes.contains("07") || mes.contains("/07/")) {
			mesExtenso = "Jul";
		} else if (mes.contains("Agosto") || mes.contains("August") || mes.contains("08") || mes.contains("/08/")) {
			mesExtenso = "Ago";
		} else if (mes.contains("Septiembre") || mes.contains("September") || mes.contains("09") || mes.contains("/09/")) {
			mesExtenso = "Set";
		} else if (mes.contains("Octubre") || mes.contains("October") || mes.contains("10") || mes.contains("/10/")) {
			mesExtenso = "Out";
		} else if (mes.contains("Noviembre") || mes.contains("November") || mes.contains("11") || mes.contains("/11/")) {
			mesExtenso = "Nov";
		} else if (mes.contains("Diciembre") || mes.contains("December") || mes.contains("12") || mes.contains("/12/")) {
			mesExtenso = "Dez";
		}
		
		return mesExtenso;
	}
	
	public static int retornaNumeroDoMes(String mes){
		
		int numeroMes = mesAtual;
		
		if (mes.contains(String.valueOf(anoAtual))) {
			
			mes = mes.replace(String.valueOf(anoAtual), "");
		
		}
		
		if (mes.contains("Jan") || mes.contains("Jan") || mes.contains("Enero") || mes.contains("January") || mes.contains("01")) {
			numeroMes = 1;	
		} else if (mes.contains("Feb") || mes.contains("Fev") || mes.contains("Febrero") || mes.contains("February") || mes.contains("02")) {
			numeroMes = 2;
		} else if (mes.contains("Mar") || mes.contains("Mar") || mes.contains("Marzo") || mes.contains("March") || mes.contains("03")) {
			numeroMes = 3;	
		} else if (mes.contains("Apr") || mes.contains("Abr") || mes.contains("Abril") || mes.contains("April") || mes.contains("04")) {
			numeroMes = 4;
		} else if (mes.contains("May") || mes.contains("Mai") || mes.contains("Mayo") || mes.contains("May") || mes.contains("05")) {
			numeroMes = 5;
		} else if (mes.contains("Jun") || mes.contains("Jun") || mes.contains("Junio") || mes.contains("June") || mes.contains("06")) {
			numeroMes = 6;
		} else if (mes.contains("Jul") || mes.contains("Jul") || mes.contains("Julio") || mes.contains("July") || mes.contains("07")) {
			numeroMes = 7;
		} else if (mes.contains("Aug") || mes.contains("Ago") || mes.contains("Agost") || mes.contains("August") || mes.contains("08")) {
			numeroMes = 8;
		} else if (mes.contains("Sep") || mes.contains("Set") || mes.contains("Septiembre") || mes.contains("September") || mes.contains("09")) {
			numeroMes = 9;
		} else if (mes.contains("Oct") || mes.contains("Out") || mes.contains("Octubre") || mes.contains("October") || mes.contains("10")) {
			numeroMes = 10;
		} else if (mes.contains("Nov") || mes.contains("Nov") || mes.contains("Noviembre") || mes.contains("November") || mes.contains("11")) {
			numeroMes = 11;
		} else if (mes.contains("Dec") || mes.contains("Dez") || mes.contains("Diciembre") || mes.contains("December") || mes.contains("12")) {
			numeroMes = 12;
		}
		
		return numeroMes;
	}
	
	public static int retornaNumeroDoAno(int mesMasterActive){
		
		int ano = anoAtual;
		
		if (mesAtual == 1 && mesMasterActive == 12) {
			ano = anoAtual - 1;
		}
		
		return ano;
	}

	public static String retornaNumeroMes(String mes){
		
		String numeroMes = "";
		
		switch(mes) {
	 	 
		 case "Jan":  numeroMes = "01"; break;
	 	 case "Fev":  numeroMes = "02"; break;
	 	 case "Mar":  numeroMes = "03"; break;
	 	 case "Abr":  numeroMes = "04"; break;
	 	 case "Mai":  numeroMes = "05"; break;
	 	 case "Jun":  numeroMes = "06"; break;
	 	 case "Jul":  numeroMes = "07"; break;
	 	 case "Ago":  numeroMes = "08"; break;
	 	 case "Set":  numeroMes = "09"; break;
	 	 case "Out":  numeroMes = "10"; break;
	 	 case "Nov":  numeroMes = "11"; break;
	 	 case "Dez":  numeroMes = "12"; break;
		
		}
		
		return numeroMes;
	}
	
	
	public static String retornaNumeroMes2(String mesActive){
		
		String numeroMes = "";
		
		if (mesActive != null && !mesActive.isEmpty()) {
			
			if (mesActive.contains("Jan") || mesActive.contains("Jan") || mesActive.contains("Enero") || mesActive.contains("January")) {
			
				numeroMes = "01";
				
			} else if (mesActive.contains("Feb") || mesActive.contains("Fev") || mesActive.contains("Febrero") || mesActive.contains("February")) {
			
				numeroMes = "02";
			
			} else if (mesActive.contains("Mar") || mesActive.contains("Mar") || mesActive.contains("Marzo") || mesActive.contains("March")) {
				
				numeroMes = "03";
			
			} else if (mesActive.contains("Apr") || mesActive.contains("Abr") || mesActive.contains("Abril") || mesActive.contains("April")) {
				
				numeroMes = "04";
			
			} else if (mesActive.contains("May") || mesActive.contains("Mai") || mesActive.contains("Mayo") || mesActive.contains("May")) {
				
				numeroMes = "05";
			
			} else if (mesActive.contains("Jun") || mesActive.contains("Jun") || mesActive.contains("Junio") || mesActive.contains("June")) {
				
				numeroMes = "06";
			
			} else if (mesActive.contains("Jul") || mesActive.contains("Jul") || mesActive.contains("Julio") || mesActive.contains("July")) {
				
				numeroMes = "07";
			
			} else if (mesActive.contains("Aug") || mesActive.contains("Ago") || mesActive.contains("Agost") || mesActive.contains("August")) {
				
				numeroMes = "08";
			
			} else if (mesActive.contains("Sep") || mesActive.contains("Set") || mesActive.contains("Septiembre") || mesActive.contains("September")) {
				
				numeroMes = "09";
			
			} else if (mesActive.contains("Oct") || mesActive.contains("Out") || mesActive.contains("Octubre") || mesActive.contains("October")) {
				
				numeroMes = "10";
			
			} else if (mesActive.contains("Nov") || mesActive.contains("Nov") || mesActive.contains("Noviembre") || mesActive.contains("November")) {
				
				numeroMes = "11";
			
			} else if (mesActive.contains("Dec") || mesActive.contains("Dez") || mesActive.contains("Diciembre") || mesActive.contains("December")) {
				
				numeroMes = "12";
			}		
			
		}
		
		return numeroMes;
	}

	
	
	
	public static int retornaMesAnterior(int mesAtual){
		
		int mesAnterior =  mesAtual - 1;
		
		if (mesAtual == 1) {
			mesAnterior = 12;
		}
		
		return mesAnterior;
	}
	
	public static int retornaAnoCorreto(int mesAnterior, int anoAtual){
		
		int anoCorreto =  anoAtual;
		
		if (mesAnterior == 12) {
			anoCorreto = anoAtual - 1;
		}
		
		return anoCorreto;
	}
	
	public static String recuperaNomeContrato(String contrato){
		
		try {
			
			if (contrato != null && !contrato.isEmpty()) {
				
				String[] partes = contrato.split("\\n");
				
				if (partes != null && partes.length > 0) {
					contrato = partes[2];
				}
				
			}
			
			return contrato;

		} catch (Exception e) {
			return contrato;
		}
		
	}
    
    public static void criaDiretorio(String caminhoDiretorio){
        File diretorio = new File(caminhoDiretorio);
        if (!diretorio.exists()) {
        	diretorio.mkdirs();
        }
    }
    
	public static void criaDiretorioTemp(){
	    	
	   	String str1 = "echo %temp%";
	   	String command = "C:\\WINDOWS\\system32\\cmd.exe /y /c " + str1;
	    	
	   	try {
		    		
	   		Process processo = Runtime.getRuntime().exec(command);
	   		String line;
	   		String caminhoTemp = "";
		    		
	   		//pega o retorno do processo
	   		BufferedReader stdInput = new BufferedReader(new 
	   				InputStreamReader(processo.getInputStream()));
		    		
	   		//printa o retorno
	   		while ((line = stdInput.readLine()) != null) {
	   			caminhoTemp = line;
	   		}
		    		
	   		criaDiretorio(caminhoTemp);
		    		
	   	} catch (Exception e) {
	   		System.out.println("Deu erro na cria��o do diret�rio Temp: " + e.getMessage());
	   	}

	}
    
    // Propiedades do driver para abrir no IE, Chrome ou Firefox
    public static WebDriver getWebDriver() throws InterruptedException {
    	
    	WebDriver driver = null;
    		
            try {
				
            	if ("Chrome".equals(Util.getValor("navegador"))) {
				    
					File file = new File(Util.getValor("driver.Chrome.selenium"));
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.Chrome.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.chrome();
				    caps.setJavascriptEnabled(true);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
				    ChromeOptions chromeOptions = new ChromeOptions(); 
				    Map<String, Object> chromePreferences = new HashMap<String, Object>();
					chromePreferences.put("profile.default_content_settings.popups", 0);
				    chromePreferences.put("download.default_directory",Util.getValor("caminho.download.relatorios"));
				    chromePreferences.put("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
				    chromeOptions.setExperimentalOption("prefs", chromePreferences);
				    
				    // Argumento que faz com que o navegador use os dados do usu�rio salvos
				    // Com isso n�o ser� necess�rio digitar os dados de login no sharepoint, pois ele pegar� as informa��es do usu�rio salvas na m�quina
				    // Um ponto importante � que n�o poderemos ter mais de uma sess�o do Chrome aberta
				    // Outro ponto importante � que a op��o acima browser.helperApps.neverAsk.saveToDisk que permite que o browser salve um arquivo sem perguntar aonde salvar,
				    // n�o funcionar� por conta do trecho abaixo.
				    // Neste caso deveremos setar manualmente essa op��o no Chrome antes de rodar o rob�
				    // Ser� necess�rio fazer aparecer essa pasta no explorer do usu�rio
				    chromeOptions.addArguments("user-data-dir=" + Util.getValor("caminho.dados.usuario.Chrome"));
				    chromeOptions.addArguments("--lang=pt");
				    //chromeOptions.addArguments("--start-fullscreen");
				    driver = new ChromeDriver(chromeOptions);
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				
				} else if ("internetExplorer".equals(Util.getValor("navegador"))) {
				
					File file = new File(Util.getValor("driver.internetExplorer.selenium"));
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.internetExplorer.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.internetExplorer();
				    caps.setJavascriptEnabled(true);
				    //caps.setPlatform(org.openqa.selenium.Platform.WINDOWS);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
					InternetExplorerOptions ieOptions = new InternetExplorerOptions();
					ieOptions.setCapability("ignoreZoomSetting", true);
					ieOptions.setCapability("nativeEvents",false);
					ieOptions.setCapability("browser.download.folderList", 2);
					ieOptions.setCapability("browser.helperApps.alwaysAsk.force", false);
					ieOptions.setCapability("browser.download.manager.showWhenStarting",false);
					//ieOptions.setCapability("browser.download.dir",getValor("caminho.download.relatorios"));
					//ieOptions.setCapability("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
					//ieOptions.setCapability("browser.helperApps.alwaysAsk.force", true);
					// Limpando o cache com propriedades do Internet Explorer
					//caps.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION,true);
					//driver = new InternetExplorerDriver(caps);
					//driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
					driver = new InternetExplorerDriver(ieOptions);
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				
				} else if ("Firefox".equals(Util.getValor("navegador"))) {
					
					File file = new File(Util.getValor("driver.Firefox.selenium"));
					System.setProperty(Util.getValor("propriedade.binario.Firefox.selenium"),Util.getValor("binario.Firefox")); 
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.Firefox.selenium"),file.getAbsolutePath());
     				File profileDirectory = new File(Util.getValor("caminho.dados.usuario.Firefox"));
				    FirefoxProfile fxProfile = new FirefoxProfile(profileDirectory);
				    fxProfile.setPreference("browser.download.folderList",2);
				    fxProfile.setPreference("browser.download.manager.showWhenStarting",false);
				    fxProfile.setPreference("browser.download.dir",Util.getValor("caminho.download.relatorios"));
				    fxProfile.setPreference("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, application/zip, text/csv, text/comma-separated-values, application/octet-stream");
				    // Limpando o cache com propriedades do Firefox 
				    /*
				    fxProfile.setPreference("browser.cache.disk.enable", false);
				    fxProfile.setPreference("browser.cache.memory.enable", false);
				    fxProfile.setPreference("browser.cache.offline.enable", false);
				    fxProfile.setPreference("network.http.use-cache", false);
				    fxProfile.setPreference("network.cookie.cookieBehavior", 2);
				    */
				    FirefoxOptions fxOptions = new FirefoxOptions();
				    fxOptions.setProfile(fxProfile);
				    driver = new FirefoxDriver(fxOptions);
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				}
			
            } catch (IOException e) {
				System.out.println("Ocorreu um erro no metodo getWebDriver: " + e.getMessage());
			}
            
            return driver;
    } 
    
    public static WebDriver getHandleToWindow(String title, WebDriver driver){

        WebDriver popup = null;
        Set<String> windowIterator = driver.getWindowHandles();
        System.err.println("No of windows :  " + windowIterator.size());
        for (String s : windowIterator) {
          String windowHandle = s; 
          popup = driver.switchTo().window(windowHandle);
          System.out.println("Window Title : " + popup.getTitle());
          System.out.println("Window Url : " + popup.getCurrentUrl());
          if (popup.getTitle().equals(title) ){
              System.out.println("Selected Window Title : " + popup.getTitle());
              return popup;
          }

        }
          System.out.println("Window Title :" + popup.getTitle());
          System.out.println();
          return popup;
        }
    
    
    public static void gravarArquivo(String caminhoDiretorio, String nomeArquivo, String extensaoArquivo, String conteudoArquivo, String mensagem) throws IOException {
    	
    	String arquivo = caminhoDiretorio + "/" + nomeArquivo + extensaoArquivo; 
    	File file = new File(arquivo);
		BufferedWriter writer = new BufferedWriter(new FileWriter(file));
		writer.write(mensagem + conteudoArquivo);
		writer.newLine();
		//Criando o conte�do do arquivo
		writer.flush();
		//Fechando conex�o e escrita do arquivo.
		writer.close();
		
    }

}