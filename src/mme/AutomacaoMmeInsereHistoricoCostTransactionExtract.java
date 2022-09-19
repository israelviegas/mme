package mme;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
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

public class AutomacaoMmeInsereHistoricoCostTransactionExtract {
	
	private static List<RelatorioCostTransactionExtract> listaRelatorio = null;
	private static String dataAtual = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss").format(new Date());
	private static String dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
	private static SimpleDateFormat formatoDataReferencia = new SimpleDateFormat("yyyy-MM-dd"); 
	private static int mesAtual = Integer.parseInt(new SimpleDateFormat("MM").format(new Date()));
	private static int mesAnterior = retornaMesAnterior(mesAtual);
	private static int anoAtual = Integer.parseInt(new SimpleDateFormat("yyyy").format(new Date()));
	private static int anoCorreto = retornaAnoCorreto(mesAnterior, anoAtual);
	private static java.sql.Date dataAtualBancoFinal = new java.sql.Date(new Date().getTime());
	private static Set<String> listaNomesDeContratosDistintos = null;
	private static String diretorioLogs = "";
	private static String subdiretorioRelatoriosBaixados = "";
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
		}
        
    }
	
    public static void executaAutomacaoMme(WebDriver driver) throws Exception{
    	
    	try {
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		
    		listaRelatorio = new ArrayList<RelatorioCostTransactionExtract>();
    		listaNomesDeContratosDistintos = new HashSet<String>();
    		
    		String mensagemResultadoMme = "";
    		
    		System.out.println("In�cio: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		driver = getWebDriver();
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		WebDriverWait wait = new WebDriverWait(driver, 60);
    		
        	// Deleto arquivos que existirem no diret�rio relat�rio
        	verificaArquivosDiretorioDeRelatorios(Util.getValor("caminho.diretorio.relatorios"));
        	
        	// Deleto arquivos que existirem no sub-diret�rio de relat�rios baixados
        	verificaArquivosDiretorioDeRelatorios(subdiretorioRelatoriosBaixados);
        	
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
   			
			// Contrato AM Latam Brasil Fija - Contrato Local
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
			
			// Insere no banco a lista contendo todos os relat�rios baixados
			 if (listaRelatorio != null && !listaRelatorio.isEmpty()) {

				 int contador = 0;
				 
				 if (listaNomesDeContratosDistintos != null && !listaNomesDeContratosDistintos.isEmpty()) {
					 
					 for (String nomeContrato : listaNomesDeContratosDistintos) {
						 
						 // Apagando o m�s atual do contrato
						 //MmeDao mmeDaoDeletarMesAtual = new MmeDao();
						 //mmeDaoDeletarMesAtual.deletarRelatorios(mesAtual, anoAtual, nomeContrato);
						 //System.out.println("Apaguei todos os dados do banco do m�s atual do contrato " + nomeContrato);
						 
						 // Apagando o m�s anterior do contrato
						 //MmeDao mmeDaoDeletarMesAnterior = new MmeDao();
						 //mmeDaoDeletarMesAnterior.deletarRelatorios(mesAnterior, anoCorreto, nomeContrato);
						 //System.out.println("Apaguei todos os dados do banco do m�s anterior do contrato " + nomeContrato);
						
					}
					 
				 }
				 
				 for (RelatorioCostTransactionExtract relatorioMme : listaRelatorio) {
					 
		        	  // Converto os valores de null para espa�o em branco para gravar branco no banco e n�o dar problema
		        	  // no relat�rio do sharepoint da Accenture
		        	  // Esse relat�rio de sharepoint da Accenture � gerado atrav�s do banco
		             Util.converteValorNullParaEspacoEmBrancoRelatorioCostTransactionExtract(relatorioMme);
					 MmeDao mmeDaoInserirRelatorio = new MmeDao();
					 mmeDaoInserirRelatorio.inserirRelatoriosCostTransactionExtract(relatorioMme);
		             contador++;
		             System.out.println("Inseri o relat�rio de n�mero: " + contador + " de um total de " + listaRelatorio.size() + " nome do contrato: " + relatorioMme.getContrato());

				 }
				 
				mensagemResultadoMme = "Relat�rios do Mme gravados no banco com sucesso";

			 } else {
				 
				 mensagemResultadoMme = "N�o foram encontrados relat�rios do Mme para os contratos processados";
			 
			 }

			
            gravarArquivo(diretorioLogs, "Resultado Mme" + " " + dataAtual, ".txt", "", mensagemResultadoMme);
            
            inserirStatusExecucaoNoBanco("Mme", dataAtualPlanilhaFinal, "Sucesso");
            
            System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    	} catch (Exception e) {
			executaAutomacaoMme ++;
			// Executo at� 20 vezes se der erro no executaAutomacaoMme
			if (executaAutomacaoMme <= 20) {
				
				System.out.println("Deu erro no m�todo executaAutomacaoMme, tentativa de acerto: " + executaAutomacaoMme + "\n" + "Erro: " + e.getMessage());
				executaAutomacaoMme(driver);
			
			} else {
				throw new Exception("Ocorreu um erro no m�todo executaAutomacaoMme: " + e);
		    }

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
	public static void lerPlanilhaRelatorioMme(String planilha, String contrato, int mes, int ano) throws Exception {

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
		    	   
		    	   RelatorioCostTransactionExtract relatorioMme = new RelatorioCostTransactionExtract();
		    	   
		    	   if (row.getRowNum() == 0 || row.getRowNum() == 1) {
		    		   continue;
		    	   }
		    	   
		           int indice = 0;
		           int contador = 0;
		           
		           // PostingPeriod
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setPostingPeriod(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYearPeriod
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setFiscalYearPeriod(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYearQuarter
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setFiscalYearQuarter(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // FiscalYear
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setFiscalYear(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // PostingDate
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setPostingDate(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // DocumentDate
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setDocumentDate(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Category Group
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCategoryGroup(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Category
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCategory(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // DocumentType
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setDocumentType(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // DocumentNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setDocumentNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Description
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setDescription(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ReferenceNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setReferenceNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // AccountNbr
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setAccountNbr(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // AccountNbr Description
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setAccountNbrDescription(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Quantity Amount/Hours
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setQuantityAmountHours(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // GlobalAmount
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setGlobalAmount(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ObjectAmount
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setObjectAmount(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ObjectCurrency
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setObjectCurrency(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // TransactionalAmt
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setTransactionalAmt(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // TransactionalCurrency
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setTransactionalCurrency(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ReportingAmount
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setReportingAmount(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ReportingCurrency
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setReportingCurrency(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Enterprise ID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setEnterpriseID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Resource Name
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setResourceName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Personnel number
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setPersonnelNumber(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Current Worklocation
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCurrentWorklocation(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Current Home Office
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCurrentHomeOffice(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCenterID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCostCenterID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCenterName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCostCenterName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCostCollectorID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // CostCollectorName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setCostCollectorName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWBS(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS Description
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWBSDescription(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS ProfitCenter
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWBSProfitCenter(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS ProfitCenterName
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWBSProfitCenterName(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Parent WBS
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setParentWBS(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // ContractID
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setContractID(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WBS RaKey
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWBSRaKey(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel1
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel1(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel2
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel2(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel3
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel3(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel4
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel4(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel5
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel5(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel6
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel6(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel7
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel7(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel8
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel8(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel9
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel9(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // WMULevel10
		           indice = contador++;
		           if (formatter.formatCellValue(row.getCell(indice)) != null && !formatter.formatCellValue(row.getCell(indice)).isEmpty()) {
		        	   relatorioMme.setWMULevel10(formatter.formatCellValue(row.getCell(indice)));
		           }
		           
		           // Contrato
		           relatorioMme.setContrato(contrato);

		           // Data Extra��o
		           relatorioMme.setDataExtracao(dataAtualBancoFinal);
		           
		           // Data Refer�ncia
		           Date dataReferencia = formatoDataReferencia.parse(ano + "-" + mes + "-" + "01");
		           java.sql.Date dataReferenciaFinal = new java.sql.Date(dataReferencia.getTime());
		           relatorioMme.setDataReferencia(dataReferenciaFinal);
		           
		           // Armazeno os per�odos de extra��o distintos
		           //listaPeridosDeExtracaoDistintos.add(relatorioMme.getPostingPeriod());
		           
		           listaNomesDeContratosDistintos.add(relatorioMme.getContrato());
		           
		           if (relatorioMme != null && !Util.relatorioCostTransactionExtractPossuiTodosCamposNulos(relatorioMme)) {
		        	   listaRelatorio.add(relatorioMme);
		           }
		           
		       }
		       
		   }
		   
		   // Existem alguns registros da planilha que possuem o campo PostingPeriod em branco
		   // Preencho esses campos com o valor dos outros PostingPeriod que possuem valor, por�m
		   // os salvo no campo Data_Referencia
		   //preencheListaRelatorioTemporariaComDataReferencia(listaPeridosDeExtracaoDistintos, listaRelatorioTemporaria);
		   
		   // Por fim adiciono a listaRelatorioTemporaria na listaRelatorio
		   //listaRelatorio.addAll(listaRelatorioTemporaria);
		   
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de relat�rio n�o encontrado!");
			   throw new Exception("Arquivo Excel de relat�rio n�o encontrado!");
			
			} finally {
			
				if (arquivo != null) {
					arquivo.close();
				}
				
			}
		
			if (listaRelatorio.size() == 0) {
			   //Pode ser que existam arquivos vazios, ent�o n�o posso lan�ar exce��o aqui
				//throw new Exception("Lista de projetos est� vazia");
			}
			
		}
	
	   public static void preencheListaRelatorioTemporariaComDataReferencia(Set<String> listaPeridosDeExtracaoDistintos, List<RelatorioCostTransactionExtract> listaRelatorioTemporaria) throws Exception{
		   
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

	
	   public static void listaEMoveArquivosEntreDiretorios(String caminhoDiretorioOrigem, String caminhoDiretorioDestino, String contrato, int mes, int ano) throws Exception{
	    	
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
									listaEMoveArquivosEntreDiretorios(caminhoDiretorioOrigem, caminhoDiretorioDestino, contrato, mes, ano);
								} else {
									throw new Exception("Arquivo invalido: " + item );
								}
							}
							
							// L� os relat�rios baixados e armazena todas as linhas das planilhas em uma lista
				    		lerPlanilhaRelatorioMme(caminhoArquivo, contrato, mes, ano);
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
	   
	  public static void verificaArquivosDiretorioDeRelatorios(String caminhoDiretorio) throws Exception{
	    	
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
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='top-nav-breadcrumb-1']/div/div/button"))).click();
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
		
		abrirCaixaSelecaoContratos(driver, wait);
		
		passosIniciaisCaixaSelecaoContratos(driver, wait);
		
	}
	
	public static void passosFinaisPesquisaDeContratos(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Clicando no bot�o Select
		String textoSelect= "Select";
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
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);

		// Expandindo a op��o Global Village Telecom da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[16]/li/div/div"))).click();
		Thread.sleep(3000);

		// Expandindo a op��o GVT SW Factories da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[16]/li/ul/ez-node[2]/li/div/div"))).click();
		Thread.sleep(3000);

		// Clicando na op��o 9940191116 SW Factories da �rvores de op��es
		String id9940191116SwFactories = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[16]/li/ul/ez-node[2]/li/ul/ez-node/li/div[2]/span";
		//String SWFactories = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(id9940191116SwFactories))).getText();
		//String contrato = recuperaNomeContrato(SWFactories);
		String contrato = "9940191116 SW Factories";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(id9940191116SwFactories))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_AM_Faturamento(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o AM Faturamento da �rvores de op��es
		String idAMFaturamento = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[6]/li/div/span";
		//String AMFaturamento = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idAMFaturamento))).getText();
		//String contrato = recuperaNomeContrato(AMFaturamento);
		String contrato = "AM Faturamento";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idAMFaturamento))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_B2C_SFA_Salesforce(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o B2C SFA - Salesforce da �rvores de op��es
		String idB2C_SFA_Salesforce = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[7]/li/div[2]/span";
		//String B2C_SFA_Salesforce = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idB2C_SFA_Salesforce))).getText();
		//String contrato = recuperaNomeContrato(B2C_SFA_Salesforce);
		String contrato = "B2C SFA - Salesforce";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idB2C_SFA_Salesforce))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_Callidus(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Callidus da �rvores de op��es
		String idCallidus = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[8]/li/div[2]/span";
		//String Callidus = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idCallidus))).getText();
		//String contrato = recuperaNomeContrato(Callidus);
		String contrato = "Callidus";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idCallidus))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_GVT_Proforma(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o Global Village Telecom da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o GVT Proforma da �rvores de op��es
		String idGVTProforma = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/ul/ez-node/li/div[2]/span";
		//String GVT_Proforma = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idGVTProforma))).getText();
		//String contrato = recuperaNomeContrato(GVT_Proforma);
		String contrato = "GVT Proforma";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idGVTProforma))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}

	public static void contrato_Hybris_eCommerce(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o Hybris - eCommerce da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[12]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o SCO - Ecommerce da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[12]/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Hybris - eCommerce da �rvores de op��es
		String idHybris_eCommerce = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[12]/li/ul/ez-node/li/ul/ez-node/li/div/span";
		//String Hybris_eCommerce = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idHybris_eCommerce))).getText();
		//String contrato = recuperaNomeContrato(Hybris_eCommerce);
		String contrato = "Hybris - eCommerce";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idHybris_eCommerce))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_Digital_Factory(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o Telef�nica - XBD Digital Faactorye da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[18]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Digital Factory da �rvores de op��es
		String idDigitalFactory = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[18]/li/ul/ez-node/li/div[2]/span";
		//String DigitalFactory = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idDigitalFactory))).getText();
		//String contrato = recuperaNomeContrato(DigitalFactory);
		String contrato = "Digital Factory";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idDigitalFactory))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_AM_Latam_Brasil_Fija_Contrato_Local(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o APOLLO LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[7]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o AM LATAM - Master Contract 9940089177 da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[7]/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o AM Latam Brasil Fija - Contrato Local da �rvores de op��es
		String idAmLatamBrasilFija = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[7]/li/ul/ez-node/li/ul/ez-node[4]/li/div[2]/span";
		String contrato = "AM Latam Brasil Fija - Contrato Local";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idAmLatamBrasilFija))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_Nova_Fabrica_Design(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Nova Fabrica Design da �rvores de op��es
		String idNovaFabricaDesign = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[13]/li/div[2]/span";
		String contrato = "Nova Fabrica Design";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idNovaFabricaDesign))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_Desligue_Do_Atis(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Desligue Do Atis da �rvores de op��es
		String idDesligueDoAtis = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[9]/li/div[2]/span";
		String contrato = "Desligue Do Atis";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idDesligueDoAtis))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_Portal_Terra(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Portal Terra da �rvores de op��es
		String idPortalTerra = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[14]/li/div[2]/span";
		String contrato = "Portal Terra";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idPortalTerra))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}
	
	public static void contrato_Sustentacao_VIVO_GO(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a op��o Accenture da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o 5. LATAM da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a op��o BRASIL - Non OLGA da �rvores de op��es
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Clicando na op��o Sustentacao VIVO GO da �rvores de op��es
		String idSustentacaoVIVOGO = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[17]/li/div/span";
		String contrato = "Sustentacao VIVO GO";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idSustentacaoVIVOGO))).click();
		Thread.sleep(5000);

		passosFinaisPesquisaDeContratos(driver, wait);
		
		extracaoPlanilha(driver, wait, contrato);
		
	}

	
	public static void extracaoPlanilha(WebDriver driver, WebDriverWait wait, String contrato) throws Exception {
		
		Thread.sleep(10000);
		
		// Clicando no link Update Forecast
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Update Forecast')]"))).click();
		Thread.sleep(2000);
		
		// Clicando na op��o Advanced Reporting do menu
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Advanced Reporting')]"))).click();
		Thread.sleep(2000);
		
		// Abrindo o menu Selection
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_ImgBtnTree']"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Cost Planning and Management 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cost Planning and Management')]"))).click();
		Thread.sleep(2000);
		
		// Abrindo a op��o Cost Transaction Extract(Everything Report) 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cost Transaction Extract(Everything Report)')]"))).click();
		Thread.sleep(2000);
		
		
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
					
					escolheDataCalendarioEExportaPlanilha(driver, wait, contrato, mes, Integer.parseInt(ano));
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
		
		boolean excecao24 = ano == 2021 && (mes == 9 || mes == 10 || mes == 11 || mes == 12);
		
		if (excecao1 || excecao2 || excecao3 || excecao4 || excecao5 || excecao6 || excecao7 || excecao8 || excecao9 || excecao10 || excecao11 || excecao12 || excecao13 || excecao14 || excecao15 || excecao16 || excecao17 || excecao18 || excecao19 || excecao20 || excecao21 || excecao22 || excecao23 || excecao24) {
		
			excecoesExtracaoRelatorios = true;
		}
		
		return excecoesExtracaoRelatorios;

	}
	
    public static void clickNoCalendario (WebDriver driver, WebDriverWait wait) throws Exception {
    	
    	try {
    		
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Calendar']"))).click();
    		Thread.sleep(2000);    		
    	
    	} catch (Exception e) {
    		
    		contadorErrosCalendario ++;
			
            // Tento escolher clicar no calend�rio por 20 vezes
            if (contadorErrosCalendario <= 20) {
            	
            	System.out.println("Erro ao clicar no calendario.Tentativa de numero: " + contadorErrosCalendario);
            	clickNoCalendario(driver, wait);
            
            } else {
         	   throw new Exception("Erro ao clicar no calendario: " + e);
            }
    		
    	}
    	
    }

	
	public static void escolheDataCalendarioEExportaPlanilha(WebDriver driver, WebDriverWait wait, String contrato, int mes, int ano) throws Exception {
		
		fecharBarraInferiorComInformacoesDaAccenture(driver);
		
		rolagemParaBaixo(driver);
		
		// Clicando no calend�rio
		clickNoCalendario(driver, wait);
		
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
		
		String idComboMes = "//div[@id='ui-datepicker-div']/div/div/select";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboMes)));
		WebElement comboMes = driver.findElement(By.xpath(idComboMes));
		// Elementos do combo
		Select elementosComboMes = new Select(comboMes);
		String mesPorExtenso = retornaMesPorExtenso(mes);
		boolean encontrouMes = false;
		for (WebElement elemento : elementosComboMes.getOptions()) {
			if (mesPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementosComboMes.selectByVisibleText(mesPorExtenso);
				encontrouMes = true;
				break;
			}
		}
		
		// Clicando no bot�o Fechar do calend�rio
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(.,'Fechar')]"))).click();
		Thread.sleep(2000);
		
		if (encontrouAno && encontrouMes) {
			// Clicando no bot�o Export
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_btnRunReport']"))).click();
			Thread.sleep(Integer.parseInt(Util.getValor("tempo.segundos.download.planilha")) * 1000);
			
		    // Movo todos os arquivos baixados para o diret�rio corrente de relat�rios
		    // Tamb�m l� os relat�rios baixados e armazena todas as linhas das planilhas na listaRelatorio
		    listaEMoveArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, contrato, mes, ano);

		}
		
	}
	
	public static String retornaMesPorExtenso(int mes){
		
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
				    FirefoxProfile fxProfile = new FirefoxProfile();
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