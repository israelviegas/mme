package mme;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;

import mme.util.ConnectionFactory;
import mme.util.Util;

public class MmeDao {
	
	private Connection connection;
	
	 public MmeDao() throws IOException {
		 
		 String sid = Util.getValor("sid");
	
		 this.connection = new ConnectionFactory().getConnection(sid);
	 
	 }
	 
    
	 public void deletarRelatoriosCostTransactionExtract(int mes, int ano, String contrato) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "DELETE from Tb_Cost_Planning where Contrato = " + "'" + contrato + "'" + " and (MONTH(Data_Referencia) = " + mes + " AND YEAR(Data_Referencia) = " + ano + ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 throw new RuntimeException(e);
		 } finally {
			 this.connection.close();
		 }
		 
	 }

    
	 public void inserirRelatoriosCostTransactionExtract(RelatorioCostTransactionExtract relatorioMme) throws SQLException, IOException {
		 
		 try {
			 
			  String sql = "INSERT INTO Tb_Cost_Planning (          ";
					 sql += "Contrato,                    ";
					 sql += "PostingPeriod,               ";
					 sql += "FiscalYearPeriod,            ";
					 sql += "FiscalYearQuarter,           "; 
					 sql += "FiscalYear,                  "; 
					 sql += "PostingDate,                 "; 
					 sql += "DocumentDate,                "; 
					 sql += "CategoryGroup,               "; 
					 sql += "Category,                    "; 
					 sql += "DocumentType,                "; 
					 sql += "DocumentNbr,                 "; 
					 sql += "Description,                 "; 
					 sql += "ReferenceNbr,                "; 
					 sql += "AccountNbr,                  "; 
					 sql += "AccountNbrDescription,       "; 
					 sql += "QuantityAmountHours,         "; 
					 sql += "GlobalAmount,                "; 
					 sql += "ObjectAmount,                "; 
					 sql += "ObjectCurrency,              "; 
					 sql += "TransactionalAmt,            "; 
					 sql += "TransactionalCurrency,       "; 
					 sql += "ReportingAmount,             "; 
					 sql += "ReportingCurrency,           "; 
					 sql += "EnterpriseID,                "; 
					 sql += "ResourceName,                "; 
					 sql += "PersonnelNumber,             "; 
					 sql += "CurrentWorklocation,         "; 
					 sql += "CurrentHomeOffice,           "; 
					 sql += "CostCenterID,                "; 
					 sql += "CostCenterName,              "; 
					 sql += "CostCollectorID,             "; 
					 sql += "CostCollectorName,           "; 
					 sql += "WBS,                         ";
					 sql += "WBSDescription,              "; 
					 sql += "WBSProfitCenter,             "; 
					 sql += "WBSProfitCenterName,         "; 
					 sql += "ParentWBS,                   "; 
					 sql += "ContractID,                  "; 
					 sql += "WBSRaKey,                    "; 
					 sql += "WMULevel1,                   "; 
					 sql += "WMULevel2,                   "; 
					 sql += "WMULevel3,                   "; 
					 sql += "WMULevel4,                   "; 
					 sql += "WMULevel5,                   "; 
					 sql += "WMULevel6,                   "; 
					 sql += "WMULevel7,                   "; 
					 sql += "WMULevel8,                   "; 
					 sql += "WMULevel9,                   "; 
					 sql += "WMULevel10,                  ";
					 sql += "Data_Extracao,               ";
					 sql += "Data_Referencia,             ";
					 sql += "Periodo                      ";
					 sql += ") VALUES (                   ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?,                           ";
					 sql += "?                            ";
					 sql += ")";

			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 int contador = 1;
			 
			 statement.setString(contador++, relatorioMme.getContrato());                   
			 statement.setString(contador++, relatorioMme.getPostingPeriod());              
			 statement.setString(contador++, relatorioMme.getFiscalYearPeriod());           
			 statement.setString(contador++, relatorioMme.getFiscalYearQuarter());          
			 statement.setString(contador++, relatorioMme.getFiscalYear());                 
			 statement.setString(contador++, relatorioMme.getPostingDate());                
			 statement.setString(contador++, relatorioMme.getDocumentDate());               
			 statement.setString(contador++, relatorioMme.getCategoryGroup());              
			 statement.setString(contador++, relatorioMme.getCategory());                   
			 statement.setString(contador++, relatorioMme.getDocumentType());               
			 statement.setString(contador++, relatorioMme.getDocumentNbr());                
			 statement.setString(contador++, relatorioMme.getDescription());                
			 statement.setString(contador++, relatorioMme.getReferenceNbr());               
			 statement.setString(contador++, relatorioMme.getAccountNbr());                 
			 statement.setString(contador++, relatorioMme.getAccountNbrDescription());      
			 statement.setString(contador++, relatorioMme.getQuantityAmountHours());        
			 statement.setString(contador++, relatorioMme.getGlobalAmount());               
			 statement.setString(contador++, relatorioMme.getObjectAmount());               
			 statement.setString(contador++, relatorioMme.getObjectCurrency());             
			 statement.setString(contador++, relatorioMme.getTransactionalAmt());           
			 statement.setString(contador++, relatorioMme.getTransactionalCurrency());      
			 statement.setString(contador++, relatorioMme.getReportingAmount());            
			 statement.setString(contador++, relatorioMme.getReportingCurrency());          
			 statement.setString(contador++, relatorioMme.getEnterpriseID());               
			 statement.setString(contador++, relatorioMme.getResourceName());               
			 statement.setString(contador++, relatorioMme.getPersonnelNumber());            
			 statement.setString(contador++, relatorioMme.getCurrentWorklocation());        
			 statement.setString(contador++, relatorioMme.getCurrentHomeOffice());          
			 statement.setString(contador++, relatorioMme.getCostCenterID());               
			 statement.setString(contador++, relatorioMme.getCostCenterName());             
			 statement.setString(contador++, relatorioMme.getCostCollectorID());            
			 statement.setString(contador++, relatorioMme.getCostCollectorName());          
			 statement.setString(contador++, relatorioMme.getWBS());                        
			 statement.setString(contador++, relatorioMme.getWBSDescription());             
			 statement.setString(contador++, relatorioMme.getWBSProfitCenter());            
			 statement.setString(contador++, relatorioMme.getWBSProfitCenterName());        
			 statement.setString(contador++, relatorioMme.getParentWBS());                  
			 statement.setString(contador++, relatorioMme.getContractID());                 
			 statement.setString(contador++, relatorioMme.getWBSRaKey());                   
			 statement.setString(contador++, relatorioMme.getWMULevel1());                  
			 statement.setString(contador++, relatorioMme.getWMULevel2());                  
			 statement.setString(contador++, relatorioMme.getWMULevel3());                  
			 statement.setString(contador++, relatorioMme.getWMULevel4());                  
			 statement.setString(contador++, relatorioMme.getWMULevel5());                  
			 statement.setString(contador++, relatorioMme.getWMULevel6());                  
			 statement.setString(contador++, relatorioMme.getWMULevel7());                  
			 statement.setString(contador++, relatorioMme.getWMULevel8());                  
			 statement.setString(contador++, relatorioMme.getWMULevel9());                  
			 statement.setString(contador++, relatorioMme.getWMULevel10());                 
			 statement.setDate(contador++, relatorioMme.getDataExtracao());
			 statement.setDate(contador++, relatorioMme.getDataReferencia());
			 statement.setString(contador++, relatorioMme.getPeriodo()); 
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 
			 String mensagem1 = "Houve algum problema no momento da inserção deste registro: " + "\n";
			 
			 mensagem1 += relatorioMme.getContrato() + ", " + relatorioMme.getPostingPeriod() + ", " + relatorioMme.getFiscalYearPeriod() + ", "  + relatorioMme.getFiscalYearQuarter() + ", "  + relatorioMme.getFiscalYear() + "\n";
			 
			 mensagem1 += "Esta é a mensagem de erro: "  + "\n";
			 
			 mensagem1 += e.getMessage() + "\n";
			 
			 if (mensagem1.contains("String or binary data would be truncated")) {
				 
				 mensagem1 += "A mensagem acima - String or binary data would be truncated - significa que algum campo do registro estourou o limite de tamanho permitido para o seu respectivo campo na tabela do banco." ;
			 }
			 
			 throw new RuntimeException(mensagem1);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
	 
	 public void deletarRelatoriosResourceTrend(int mes, int ano, String contrato) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "DELETE from Tb_Resource_Planning where Contrato = " + "'" + contrato + "'" + " and (MONTH(Data_Referencia) = " + mes + " AND YEAR(Data_Referencia) = " + ano + ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 throw new RuntimeException(e);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
	 public void inserirRelatoriosResourceTrend(RelatorioResourceTrend relatorioResourceTrend) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "INSERT INTO Tb_Resource_Planning (          ";
			 sql += "Contrato,                    		 			   ";
			 sql += "ForecastVersion,                    			   ";
			 sql += "TypeCostCenterCareerTrack,        		   		   ";
			 sql += "Level,                    						   ";
			 sql += "WMUID,                    					       ";
			 sql += "WMUName,                    					   ";
			 sql += "WMUOwner,                    					   ";
			 sql += "WBS,                    						   ";
			 sql += "CostCollector,                    			       ";
			 sql += "CostCollectorName,                    		       ";
			 sql += "RoleName,                    					   ";
			 sql += "Name,                    						   ";
			 sql += "PersonnelNumber,                    			   ";
			 sql += "HomeLocation,                    				   ";
			 sql += "WorkLocation,                    				   ";
			 sql += "Category,                    					   ";
			 sql += "Quantity,                    					   ";
			 sql += "ActualForecast,                    			   ";
			 sql += "Date,                    						   ";
			 sql += "OrderBy,                    					   ";
			 sql += "BillRateCardId,                    			   ";
			 sql += "RateName,                    					   ";
			 sql += "BillRateCardName,                    		       ";
			 sql += "Data_Extracao,                    				   ";
			 sql += "Data_Referencia                    			   ";
			 sql += ") VALUES (                   					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?                           					   ";
			 sql += ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 int contador = 1;
			 
			 statement.setString(contador++, relatorioResourceTrend.getContrato());                 
			 statement.setString(contador++, relatorioResourceTrend.getForecastVersion());			
			 statement.setString(contador++, relatorioResourceTrend.getTypeCostCenterCareerTrack());
			 statement.setString(contador++, relatorioResourceTrend.getLevel());                    
			 statement.setString(contador++, relatorioResourceTrend.getWMUID());                    
			 statement.setString(contador++, relatorioResourceTrend.getWMUName());                  
			 statement.setString(contador++, relatorioResourceTrend.getWMUOwner());                 
			 statement.setString(contador++, relatorioResourceTrend.getWBS());                      
			 statement.setString(contador++, relatorioResourceTrend.getCostCollector());            
			 statement.setString(contador++, relatorioResourceTrend.getCostCollectorName());        
			 statement.setString(contador++, relatorioResourceTrend.getRoleName());                 
			 statement.setString(contador++, relatorioResourceTrend.getName());                     
			 statement.setString(contador++, relatorioResourceTrend.getPersonnelNumber());          
			 statement.setString(contador++, relatorioResourceTrend.getHomeLocation());             
			 statement.setString(contador++, relatorioResourceTrend.getWorkLocation());             
			 statement.setString(contador++, relatorioResourceTrend.getCategory());                 
			 statement.setString(contador++, relatorioResourceTrend.getQuantity());                 
			 statement.setString(contador++, relatorioResourceTrend.getActualForecast());           
			 statement.setString(contador++, relatorioResourceTrend.getDate());                     
			 statement.setString(contador++, relatorioResourceTrend.getOrderBy());                  
			 statement.setString(contador++, relatorioResourceTrend.getBillRateCardId());           
			 statement.setString(contador++, relatorioResourceTrend.getRateName());                 
			 statement.setString(contador++, relatorioResourceTrend.getBillRateCardName());
			 statement.setDate(contador++, relatorioResourceTrend.getDataExtracao());                 
			 statement.setDate(contador++, relatorioResourceTrend.getDataReferencia());  			 
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 
			 String mensagem1 = "Houve algum problema no momento da inserção deste registro: " + "\n";
			 
			 mensagem1 += relatorioResourceTrend.getContrato() + ", " + relatorioResourceTrend.getForecastVersion() + ", " + relatorioResourceTrend.getTypeCostCenterCareerTrack() + ", "  + relatorioResourceTrend.getLevel() + ", "  + relatorioResourceTrend.getWMUID() + "\n";
			 
			 mensagem1 += "Esta é a mensagem de erro: "  + "\n";
			 
			 mensagem1 += e.getMessage() + "\n";
			 
			 if (mensagem1.contains("String or binary data would be truncated")) {
				 
				 mensagem1 += "A mensagem acima - String or binary data would be truncated - significa que algum campo do registro estourou o limite de tamanho permitido para o seu respectivo campo na tabela do banco." ;
			 }
			 
			 throw new RuntimeException(mensagem1);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
	 public void imprimirInsertResourceTrend(RelatorioResourceTrend relatorioResourceTrend) throws SQLException, IOException {
		 
		 String sql = "INSERT INTO Tb_Resource_Planning (          ";
		 sql += "Contrato,                    		 			   ";
		 sql += "ForecastVersion,                    			   ";
		 sql += "TypeCostCenterCareerTrack,        		   		   ";
		 sql += "Level,                    						   ";
		 sql += "WMUID,                    					       ";
		 sql += "WMUName,                    					   ";
		 sql += "WMUOwner,                    					   ";
		 sql += "WBS,                    						   ";
		 sql += "CostCollector,                    			       ";
		 sql += "CostCollectorName,                    		       ";
		 sql += "RoleName,                    					   ";
		 sql += "Name,                    						   ";
		 sql += "PersonnelNumber,                    			   ";
		 sql += "HomeLocation,                    				   ";
		 sql += "WorkLocation,                    				   ";
		 sql += "Category,                    					   ";
		 sql += "Quantity,                    					   ";
		 sql += "ActualForecast,                    			   ";
		 sql += "Date,                    						   ";
		 sql += "OrderBy,                    					   ";
		 sql += "BillRateCardId,                    			   ";
		 sql += "RateName,                    					   ";
		 sql += "BillRateCardName,                    		       ";
		 sql += "Data_Extracao,                    				   ";
		 sql += "Data_Referencia                    			   ";
		 sql += ") VALUES (                   					   ";
		 sql +=	 "'" + relatorioResourceTrend.getContrato()+ "',";                 
		 sql +=	 "'" + relatorioResourceTrend.getForecastVersion()+ "',";			
		 sql +=	 "'" + relatorioResourceTrend.getTypeCostCenterCareerTrack()+ "',";
		 sql +=	 "'" + relatorioResourceTrend.getLevel()+ "',";                    
		 sql +=	 "'" + relatorioResourceTrend.getWMUID()+ "',";                    
		 sql +=	 "'" + relatorioResourceTrend.getWMUName()+ "',";                  
		 sql +=	 "'" + relatorioResourceTrend.getWMUOwner()+ "',";                 
		 sql +=	 "'" + relatorioResourceTrend.getWBS()+ "',";                      
		 sql +=	 "'" + relatorioResourceTrend.getCostCollector()+ "',";            
		 sql +=	 "'" + relatorioResourceTrend.getCostCollectorName()+ "',";        
		 sql +=	 "'" + relatorioResourceTrend.getRoleName()+ "',";                 
		 sql +=	 "'" + relatorioResourceTrend.getName()+ "',";                     
		 sql +=	 "'" + relatorioResourceTrend.getPersonnelNumber()+ "',";          
		 sql +=	 "'" + relatorioResourceTrend.getHomeLocation()+ "',";             
		 sql +=	 "'" + relatorioResourceTrend.getWorkLocation()+ "',";             
		 sql +=	 "'" + relatorioResourceTrend.getCategory()+ "',";                 
		 sql +=	 "'" + relatorioResourceTrend.getQuantity()+ "',";                 
		 sql +=	 "'" + relatorioResourceTrend.getActualForecast()+ "',";           
		 sql +=	 "'" + relatorioResourceTrend.getDate()+ "',";                     
		 sql +=	 "'" + relatorioResourceTrend.getOrderBy()+ "',";                  
		 sql +=	 "'" + relatorioResourceTrend.getBillRateCardId()+ "',";           
		 sql +=	 "'" + relatorioResourceTrend.getRateName()+ "',";                 
		 sql +=	 "'" + relatorioResourceTrend.getBillRateCardName()+ "',";
		 sql +=	 "'" + relatorioResourceTrend.getDataExtracao()+ "',";                
		 sql +=	 "'" + relatorioResourceTrend.getDataReferencia()+ "'";
		 sql += ");";
		 
		 System.out.println(sql);
	
	 }
	 
	 public void deletarRelatoriosMultiSegmentContractReport(int mesMasterActive, int anoMasterActive, String contrato) throws SQLException, IOException {
		 
		 try {
			 
			 //String sql = "DELETE from Tb_Forecast_Analysis where Contrato = " + "'" + contrato + "'" + " and (MONTH(Data_Referencia) = " + mesMasterActive + " AND YEAR(Data_Referencia) = " + anoMasterActive + ")";
			 
			 String sql = "DELETE from Tb_Forecast_Analysis where Contrato = " + "'" + contrato + "'" + " and Data_Referencia >= " + "'" + anoMasterActive + "-" + mesMasterActive + "-" + "01" + "'";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 throw new RuntimeException(e);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
	 public void inserirRelatoriosMultiSegmentContractReport(RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "INSERT INTO Tb_Forecast_Analysis (          ";
			 sql += "Contrato,                    		 			   ";
			 sql += "Level01,                    			           ";
			 sql += "Level02,        		   		                   ";
			 sql += "Level03,                    					   ";
			 sql += "Level04,                    					   ";
			 sql += "Level05,                    					   ";
			 sql += "Level06,                    					   ";
			 sql += "Level07,                    					   ";
			 sql += "Level08,                    			           ";
			 sql += "Level09,                    		               ";
			 sql += "Level10,                    					   ";
			 sql += "Level11,                    					   ";
			 sql += "Level12,                    			           ";
			 sql += "WBSNbr,                    				       ";
			 sql += "WBSName,                    				       ";
			 sql += "CostCollectorNbr,                    			   ";
			 sql += "CostCollectorName,                    			   ";
			 sql += "FiscalYear,                    			       ";
			 sql += "FiscalQuarter,                    				   ";
			 // Como Month e Group são palavras reservadas do SQL Server, elas devem ser colocadas entre []
			 sql += "[Month],                    					   ";
			 sql += "[Group],                    			           ";
			 sql += "Category,                    					   ";
			 sql += "Actual,                    		               ";
			 sql += "Forecast,                    		       		   ";
			 sql += "ComparisonEAC,                    		           ";
			 sql += "Data_Extracao,                    				   ";
			 sql += "Data_Referencia                    			   ";
			 sql += ") VALUES (                   					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?,                           					   ";
			 sql += "?                           					   ";
			 sql += ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 int contador = 1;
			 
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getContrato()); 
	 		 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel01());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel02());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel03());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel04());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel05());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel06());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel07());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel08());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel09());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel10());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel11());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getLevel12());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getWBSNbr());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getWBSName());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getCostCollectorNbr());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getCostCollectorName());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getFiscalYear());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getFiscalQuarter());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getMonth());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getGroup());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getCategory());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getActual());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getForecast());
			 statement.setString(contador++, relatorioMultiSegmentContractReport.getComparisonEAC());
	 		 statement.setDate(contador++,	 relatorioMultiSegmentContractReport.getDataExtracao());
			 statement.setDate(contador++, relatorioMultiSegmentContractReport.getDataReferencia());
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 
			 String mensagem1 = "Houve algum problema no momento da inserção deste registro: " + "\n";
			 
			 mensagem1 += relatorioMultiSegmentContractReport.getContrato() + ", " + relatorioMultiSegmentContractReport.getLevel01() + ", " + relatorioMultiSegmentContractReport.getLevel02() + ", "  + relatorioMultiSegmentContractReport.getLevel03() + ", "  + relatorioMultiSegmentContractReport.getLevel04() + "\n";
			 
			 mensagem1 += "Esta é a mensagem de erro: "  + "\n";
			 
			 mensagem1 += e.getMessage() + "\n";
			 
			 if (mensagem1.contains("String or binary data would be truncated")) {
				 
				 mensagem1 += "A mensagem acima - String or binary data would be truncated - significa que algum campo do registro estourou o limite de tamanho permitido para o seu respectivo campo na tabela do banco." ;
			 }
			 
			 throw new RuntimeException(mensagem1);
		 } finally {
			 this.connection.close();
		 }
		 
	 }

}
