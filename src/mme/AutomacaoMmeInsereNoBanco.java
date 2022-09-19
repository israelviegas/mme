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

public class AutomacaoMmeInsereNoBanco {
	
	private static List<RelatorioCostTransactionExtract> listaRelatorio = null;
	private static String dataAtual = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss").format(new Date());
	private static String dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
	private static java.sql.Date dataAtualBancoFinal = new java.sql.Date(new Date().getTime());
	private static String diretorioLogs = "";
	private static String subdiretorioRelatoriosBaixados = "";
	private static int contadorErrosCalendario = 0;
	private static int contadorErrosArquivoInvalido = 0;
	private static int executaAutomacaoMme = 0;
	
	public static void main(String[] args) throws InterruptedException, IOException, SQLException {
		
		WebDriver driver = null;

    		try {
    			
    		diretorioLogs = Util.getValor("caminho.diretorio.relatorios") + "/" + dataAtual;
    		String diretorioRelatorio = Util.getValor("caminho.download.relatorios") + "\\" + dataAtual;
    		subdiretorioRelatoriosBaixados = diretorioRelatorio + "\\" + "relatorios baixados " + dataAtual;
    		criaDiretorio(subdiretorioRelatoriosBaixados);
    		
    		//executaAutomacaoMme(driver);
    		
    		System.out.println("Início: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		listaRelatorio = new ArrayList<RelatorioCostTransactionExtract>();
    		
    	    // Movo todos os arquivos baixados para o diretório corrente de relatórios
    		// Também lê os relatórios baixados e armazena todas as linhas das planilhas em uma lista
    		String nomeContrato = Util.getValor("nome.contrato");
    		listaEMoveArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, nomeContrato);
    		int contador = 0;
    		
			// Insere no banco a lista contendo todos os relatórios baixados
			 if (listaRelatorio != null && !listaRelatorio.isEmpty()) {
				 
				 for (RelatorioCostTransactionExtract relatorioMme : listaRelatorio) {
					 
		        	  // Converto os valores de null para espaço em branco para gravar branco no banco e não dar problema
		        	  // no relatório do sharepoint da Accenture
		        	  // Esse relatório de sharepoint da Accenture é gerado através do banco
		             Util.converteValorNullParaEspacoEmBrancoRelatorioCostTransactionExtract(relatorioMme);
					 
					 MmeDao powerBiDao2 = new MmeDao();
					 powerBiDao2.inserirRelatoriosCostTransactionExtract(relatorioMme);
		             
		             contador++;
		             System.out.println("Inseri o relatório de número: " + contador + " de um total de " + listaRelatorio.size() + " nome do relatorio: " + relatorioMme.getContrato());

				 }
				 
			 }

    	    String mensagemResultadoMme = "Relatórios do Mme gravados no banco com sucesso";
			
           gravarArquivo(diretorioLogs, "Resultado Mme" + " " + dataAtual, ".txt", "", mensagemResultadoMme);
           
           inserirStatusExecucaoNoBanco("Mme", dataAtualPlanilhaFinal, "Sucesso");
           
           System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));

            
		} catch (Exception e) {
			gravarArquivo(diretorioLogs, "Erro Mme" + " " + dataAtual, ".txt", e.getMessage(), "Ocorreu um erro no Mme: ");
			inserirStatusExecucaoNoBanco("Mme", dataAtualPlanilhaFinal, "Erro");

		} finally {
			
			//mensagemErro("Houve um problema na extração dos pedidos no Mme\n");
			//fazerLogout(wait);
			if (driver != null) {
				//driver.quit();
			}
			
			//mataProcessosGoogle();
		}
        
    }
	
    public static void executaAutomacaoMme(WebDriver driver) throws Exception{
    	
    	try {
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		
    		listaRelatorio = new ArrayList<RelatorioCostTransactionExtract>();
    		
    		System.out.println("Início: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		driver = getWebDriver();
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		WebDriverWait wait = new WebDriverWait(driver, 30);
    		
        	// Deleto arquivos que existirem no diretório relatório
        	verificaArquivosDiretorioDeRelatorios(Util.getValor("caminho.diretorio.relatorios"));
        	
			// Contrato 9940191116 SW Factories
			contrato_9940191116_SW_Factories(driver, wait);
			
			// Contrato 9940191116 SW Factories
			//contrato_9940191116_SW_Factories_teste(driver, wait);
			
			// Apagando o mês atual e o mês anterior
			MmeDao mmeDao = new MmeDao();
			System.out.println("Apagando todos os dados do banco do mês atual e do mês anterior ...");
			int mesAtual = Integer.parseInt(new SimpleDateFormat("MM").format(new Date()));
			int mesAnterior = retornaMesAnterior(mesAtual);
			//mmeDao.deletarRelatorios(mesAtual, mesAnterior);
			System.out.println("Apaguei todos os dados do banco do mês atual e do mës anterior");
			int contador = 0;
			
			// Insere no banco a lista contendo todos os relatórios baixados
			 if (listaRelatorio != null && !listaRelatorio.isEmpty()) {
				 
				 for (RelatorioCostTransactionExtract relatorioMme : listaRelatorio) {
					 
		        	  // Converto os valores de null para espaço em branco para gravar branco no banco e não dar problema
		        	  // no relatório do sharepoint da Accenture
		        	  // Esse relatório de sharepoint da Accenture é gerado através do banco
		             Util.converteValorNullParaEspacoEmBrancoRelatorioCostTransactionExtract(relatorioMme);
					 
					 MmeDao powerBiDao2 = new MmeDao();
					 powerBiDao2.inserirRelatoriosCostTransactionExtract(relatorioMme);
		             
		             contador++;
		             System.out.println("Inseri o relatório de número: " + contador + " de um total de " + listaRelatorio.size());

				 }
				 
			 }

     	    String mensagemResultadoMme = "Relatórios do Mme gravados no banco com sucesso";
			
            gravarArquivo(diretorioLogs, "Resultado Mme" + " " + dataAtual, ".txt", "", mensagemResultadoMme);
            
            inserirStatusExecucaoNoBanco("Mme", dataAtualPlanilhaFinal, "Sucesso");
            
            System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    	} catch (Exception e) {
			executaAutomacaoMme ++;
			// Executo até 20 vezes se der erro no executaAutomacaoMme
			if (executaAutomacaoMme <= 20) {
				
				System.out.println("Deu erro no método executaAutomacaoMme, tentativa de acerto: " + executaAutomacaoMme);
				executaAutomacaoMme(driver);
			
			} else {
				throw new Exception("Ocorreu um erro no método executaAutomacaoMme: " + e);
		    }

		}
    	
    }

	
    public static void moverArquivosEntreDiretorios(String caminhoArquivoOrigem, String caminhoDiretorioDestino) throws Exception{
    	
    	boolean sucesso = true;
    	File arquivoOrigem = new File(caminhoArquivoOrigem);
        File diretorioDestino = new File(caminhoDiretorioDestino);
        if (arquivoOrigem.exists() && diretorioDestino.exists()) {
        	sucesso = arquivoOrigem.renameTo(new File(diretorioDestino, arquivoOrigem.getName()));
        }
        
        if (!sucesso) {
        	throw new Exception("Ocorreu um erro no momento de mover o relatório " + caminhoArquivoOrigem + " para " + caminhoDiretorioDestino);
        }
        
    }

	
	@SuppressWarnings({ "resource" })
	public static void lerPlanilhaRelatorioMme(String planilha, String contrato) throws Exception {

		try {
		   FileInputStream arquivo = new FileInputStream(new File(
				   planilha));
		
		   OPCPackage pkg = OPCPackage.open(new File(planilha));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetRelatorio = workbook.getSheet("RawData");
		   
		   // Uso o DataFormatter para deixar todos os campos como String, inclusive
		   // os que tem números
		   DataFormatter formatter = new DataFormatter();
		   String nomeContrato = "";
		   
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
		           
		           // Data Extração
		           relatorioMme.setDataExtracao(dataAtualBancoFinal);
		           
		           // Contrato
		           // Informação do nome do contrato
		           // Pode ser que para algumas linhas essa informação mude, então estou pegando só da primeira linha 
		           if (row.getRowNum() == 2) {
		        	   nomeContrato = relatorioMme.getWMULevel1();
		           }
		           relatorioMme.setContrato(nomeContrato);
		           
		           if (relatorioMme != null && !Util.relatorioCostTransactionExtractPossuiTodosCamposNulos(relatorioMme)) {
		        	   //System.out.println("Nome Contrato: " + relatorioMme.getContrato());
		        	   listaRelatorio.add(relatorioMme);
		           }
		           
		       }
		       
		   }
		   
		   arquivo.close();
		
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de relatório não encontrado!");
			   throw new Exception("Arquivo Excel de relatório não encontrado!");
			}
		
			if (listaRelatorio.size() == 0) {
			   //Pode ser que existam arquivos vazios, então não posso lançar exceção aqui
				//throw new Exception("Lista de projetos está vazia");
			}
			
		}
	
	   public static boolean verificaArquivoValido(String item) throws Exception{
		   
			// O arquivo está demorando para baixar, então precio garantir que só irei
			// processá-lo quando ele estiver íntegro e com o nome completo baixado
		   
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

	
	   public static void listaEMoveArquivosEntreDiretorios(String caminhoDiretorioOrigem, String caminhoDiretorioDestino, String contrato) throws Exception{
	    	
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
						// Se existirem arquivos, os movo para a pasta de saída
						if (arquivo.exists() && arquivo.isFile()) {
							Thread.sleep(1000);
							// As vezes o arquivo está demorando para baixar, então precio garantir que só irei
							// processá-lo quando ele estiver íntegro e com o nome completo baixado
							// Então fico tentando até ele ter baixado completamente 
							int indiceArquivoValido = item.indexOf("Cost");
							if (indiceArquivoValido == -1) {
								contadorErrosArquivoInvalido++;
								if (contadorErrosArquivoInvalido < 30) {
									System.out.println("Arquivo " + item + " ainda não está pronto.  Tentativa de numero: " + contadorErrosArquivoInvalido);
									Thread.sleep(3000);
									listaEMoveArquivosEntreDiretorios(caminhoDiretorioOrigem, caminhoDiretorioDestino, contrato);
								} else {
									throw new Exception("Arquivo invalido: " + item );
								}
							}
							
							// Lê os relatórios baixados e armazena todas as linhas das planilhas em uma lista
				    		lerPlanilhaRelatorioMme(caminhoArquivo, contrato);
		            		//Move o relatório baixado do diretório relatorios para o diretório correto
		            		moverArquivosEntreDiretorios(caminhoArquivo, caminhoDiretorioDestino);
						}

					}
				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("Não existe o diretório: " + caminhoDiretorioOrigem);
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
		driver.manage().window().maximize();
		driver.get(Util.getValor("url.mme"));
		Thread.sleep(2000);
		
	}
	
	public static void abrirCaixaSelecaoContratos(WebDriver driver, WebDriverWait wait) throws Exception {
		
		try {
			
			// Abrindo a caixa de seleção de contratos através do breadcrumb que fica salvo no histórico do navegador
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='top-nav-breadcrumb-1']/div/div/button"))).click();
			Thread.sleep(2000);

		} catch (Exception e) {
			//Se não achou a caixa de seleção de contratos, é porque ela abriu automaticamente no acesso à url
			System.out.println("Não achou a caixa de seleção de contratos");
		}
		
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
		
		// Seleciona a opção TELEFONICA GROUP
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
	
	public static void contrato_9940191116_SW_Factories(WebDriver driver, WebDriverWait wait) throws Exception {
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a opção Accenture da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a opção 5. LATAM da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a opção BRASIL - Non OLGA da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);

		// Expandindo a opção Global Village Telecom da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);

		// Expandindo a opção GVT SW Factories da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/ul/ez-node[2]/li/div/div"))).click();
		Thread.sleep(3000);

		// Clicando na opção 9940191116 SW Factories da árvores de opções
		String id9940191116SwFactories = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/ul/ez-node[2]/li/ul/ez-node/li/div[2]/span";
		String SWFactories = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(id9940191116SwFactories))).getText();
		String contrato = recuperaNomeContrato(SWFactories);
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(id9940191116SwFactories))).click();
		Thread.sleep(3000);

		// Clicando no botão Select
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//mat-dialog-container[@id='mat-dialog-0']/mme-orientation-component/div[2]/button[2]/span"))).click();
		Thread.sleep(3000);
		
		WebDriverWait waitCancel = new WebDriverWait(driver, 2);
		try {
			// O botão Select está provavelmente desabilitado porquê está pesquisando no projeto já selecionado anteriormente
			// Porém mesmo desabilitado, o Selenium encontra o botão Select
			// Então clico no botão Cancel para o modal sumir
			waitCancel.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Cancel')]"))).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			System.out.println("O botão Select está provavelmente desabilitado porquê está pesquisando no projeto já selecionado anteriormente. Então clico no botão Cancel para o modal sumir");
		}
		
		extracaoPlanilha(driver, wait);
		
	    // Movo todos os arquivos baixados para o diretório corrente de relatórios
	   // Também lê os relatórios baixados e armazena todas as linhas das planilhas em uma lista
	    listaEMoveArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, contrato);
		
	}
	
	public static void contrato_9940191116_SW_Factories_teste(WebDriver driver, WebDriverWait wait) throws Exception {
		
		System.out.println("Passando no segundo método");
		
    	//Passos iniciais para abrir pesquisa de contratos
    	passosIniciaisParaAbrirPesquisaDeContratos(driver, wait);
		
		// Expandindo a opção Accenture da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a opção 5. LATAM da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/div/div"))).click();
		Thread.sleep(3000);
		
		// Expandindo a opção BRASIL - Non OLGA da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/div/div"))).click();
		Thread.sleep(3000);

		// Expandindo a opção Global Village Telecom da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/div/div"))).click();
		Thread.sleep(3000);

		// Expandindo a opção GVT SW Factories da árvores de opções
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/ul/ez-node[2]/li/div/div"))).click();
		Thread.sleep(3000);

		// Clicando na opção 9940191116 SW Factories da árvores de opções
		String id9940191116SwFactories = "//ez-tree[@id='tree']/ez-node/li/ul/ez-node/li/ul/ez-node[13]/li/ul/ez-node[8]/li/ul/ez-node[11]/li/ul/ez-node[2]/li/ul/ez-node/li/div[2]/span";
		String SWFactories = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(id9940191116SwFactories))).getText();
		String contrato = recuperaNomeContrato(SWFactories);
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(id9940191116SwFactories))).click();
		Thread.sleep(3000);

		// Clicando no botão Select
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//mat-dialog-container[@id='mat-dialog-0']/mme-orientation-component/div[2]/button[2]/span"))).click();
		Thread.sleep(3000);
		
		WebDriverWait waitCancel = new WebDriverWait(driver, 2);
		try {
			// O botão Select está provavelmente desabilitado porquê está pesquisando no projeto já selecionado anteriormente
			// Porém mesmo desabilitado, o Selenium encontra o botão Select
			// Então clico no botão Cancel para o modal sumir
			waitCancel.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Cancel')]"))).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			System.out.println("O botão Select está provavelmente desabilitado porquê está pesquisando no projeto já selecionado anteriormente. Então clico no botão Cancel para o modal sumir");
		}
		
		extracaoPlanilha(driver, wait);
		
	    // Movo todos os arquivos baixados para o diretório corrente de relatórios
	   // Também lê os relatórios baixados e armazena todas as linhas das planilhas em uma lista
	    listaEMoveArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios"), subdiretorioRelatoriosBaixados, contrato);
		
	}

	
	public static void extracaoPlanilha(WebDriver driver, WebDriverWait wait) throws Exception {
		
		// Clicando no link Update Forecast
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(.,'Update Forecast')]"))).click();
		Thread.sleep(2000);
		
		// Clicando na opção Advanced Reporting do menu
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Advanced Reporting')]"))).click();
		Thread.sleep(2000);
		
		// Abrindo o menu Selection
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_ImgBtnTree']"))).click();
		Thread.sleep(2000);
		
		// Abrindo a opção Cost Planning and Management 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cost Planning and Management')]"))).click();
		Thread.sleep(2000);
		
		// Abrindo a opção Cost Transaction Extract(Everything Report) 
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cost Transaction Extract(Everything Report)')]"))).click();
		Thread.sleep(2000);
		
		// Exportando a planilha com os anos anteriores e com todos os meses
		
		String[] partesAnos = Util.getValor("anos.exportacao.planilha").split(",");
		
		if (partesAnos.length > 0) {
			
			for (String ano : partesAnos) {
				
				// Todos os meses
				for (int i = 1; i <=12; i++) {
					
					// O mes de fevereiro de 2015 não está fazendo
					boolean excecao = (i == 2 && ano.equals("2015"));
					
					if (excecao) {
						continue;
					}
					
					escolheDataCalendarioEExportaPlanilha(driver, wait, i, ano );
				}
				
			}
			
		}
		
	}
	
    public static void clickNoCalendario (WebDriver driver, WebDriverWait wait) throws Exception {
    	
    	try {
    		
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Calendar']"))).click();
    		Thread.sleep(2000);    		
    	
    	} catch (Exception e) {
    		
    		contadorErrosCalendario ++;
			
            // Tento escolher clicar no calendário por 10 vezes
            if (contadorErrosCalendario <= 10) {
            	
            	System.out.println("Erro ao clicar no calendario.Tentativa de numero: " + contadorErrosCalendario);
            	clickNoCalendario(driver, wait);
            
            } else {
         	   throw new Exception("Erro ao clicar no calendario: " + e);
            }
    		
    	}
    	
    }

	
	public static void escolheDataCalendarioEExportaPlanilha(WebDriver driver, WebDriverWait wait, int mes, String ano) throws Exception {
		
		// Clicando no calendário
		clickNoCalendario(driver, wait);
		
		// Ano
		String idComboAno = "//div[@id='ui-datepicker-div']/div/div/select[2]";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idComboAno)));
		WebElement comboAno = driver.findElement(By.xpath(idComboAno));
		// Elementos do combo
		Select elementoscomboAno = new Select(comboAno);
		String anoPorExtenso = ano;
		for (WebElement elemento : elementoscomboAno.getOptions()) {
			if (anoPorExtenso.equalsIgnoreCase(elemento.getText().trim())) {
				elementoscomboAno.selectByVisibleText(anoPorExtenso);
				break;
			}
		}
		Thread.sleep(1000);
		
		// Mês
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
		
		// Clicando no botão Fechar do calendário
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(.,'Fechar')]"))).click();
		Thread.sleep(2000);
		
		if (encontrouMes) {
			// Clicando no botão Export
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='ctl00_ctl00_c_dc_btnRunReport']"))).click();
			Thread.sleep(20000);
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
				    
				    // Argumento que faz com que o navegador use os dados do usuário salvos
				    // Com isso não será necessário digitar os dados de login no sharepoint, pois ele pegará as informações do usuário salvas na máquina
				    // Um ponto importante é que não poderemos ter mais de uma sessão do Chrome aberta
				    // Outro ponto importante é que a opção acima browser.helperApps.neverAsk.saveToDisk que permite que o browser salve um arquivo sem perguntar aonde salvar,
				    // não funcionará por conta do trecho abaixo.
				    // Neste caso deveremos setar manualmente essa opção no Chrome antes de rodar o robô
				    // Será necessário fazer aparecer essa pasta no explorer do usuário
				    chromeOptions.addArguments("user-data-dir=" + Util.getValor("caminho.dados.usuario.Chrome"));
				    chromeOptions.addArguments("--lang=pt");
				    
				    driver = new ChromeDriver(chromeOptions);
				    // Limpa o cache usando método do driver
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
				    // Limpa o cache usando método do driver
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
				    // Limpa o cache usando método do driver
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
		//Criando o conteúdo do arquivo
		writer.flush();
		//Fechando conexão e escrita do arquivo.
		writer.close();
		
    }

}