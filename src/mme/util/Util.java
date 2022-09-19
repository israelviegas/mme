package mme.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import mme.RelatorioCostTransactionExtract;
import mme.RelatorioMultiSegmentContractReport;
import mme.RelatorioResourceTrend;

public class Util {
	
    public static String getValor(String chave) throws IOException{
    	Properties props = getProp();
        return (String)props.getProperty(chave);
    }
    
    // Arquivo de properties sendo usado fora do projeto
    public static Properties getProp() throws IOException {
        Properties props = new Properties();
        
        // String arquivoProperties = "D:/JOBS/AutomacaoMme/configuracoes/propriedadesMme.properties";  
        // String arquivoProperties = "D:/JOBS/Mme/configuracoes/gerar relatorio pedidos.properties"; 
        // String arquivoProperties = "C:/AutomacaoMme/configuracoes/propriedadesMme.properties"; 
        // Quando a integração do Sharepoint estiver pronta, usar o arquivo propriedadesAdquira com Sharepoint.properties
         String arquivoProperties = "C:/Viegas/desenvolvimento/Selenium/arquivos propriedades/propriedadesMme.properties";
       // String arquivoProperties = "C:/Users/d.paschoaloni.jaques/OneDrive - Avanade/10 - AutomacaoMme/AutomacaoMme/configuracoes/propriedadesMme.properties";

        FileInputStream file = new FileInputStream(arquivoProperties);
        props.load(file);
        return props;
    }
    
	public static void converteValorNullParaEspacoEmBrancoRelatorioCostTransactionExtract(RelatorioCostTransactionExtract relatoriCostTransactionExtract) {
		
		 if(relatoriCostTransactionExtract.getContrato()==null					){                  relatoriCostTransactionExtract.setContrato(" ");   					}             
		 if(relatoriCostTransactionExtract.getPostingPeriod()==null					){              relatoriCostTransactionExtract.setPostingPeriod(" ");   					}         
		 if(relatoriCostTransactionExtract.getFiscalYearPeriod()==null					){          relatoriCostTransactionExtract.setFiscalYearPeriod(" ");   					}     
		 if(relatoriCostTransactionExtract.getFiscalYearQuarter()==null					){          relatoriCostTransactionExtract.setFiscalYearQuarter(" ");   					}     
		 if(relatoriCostTransactionExtract.getFiscalYear()==null					){                  relatoriCostTransactionExtract.setFiscalYear(" ");   					}         
		 if(relatoriCostTransactionExtract.getPostingDate()==null					){                  relatoriCostTransactionExtract.setPostingDate(" ");   					}         
		 if(relatoriCostTransactionExtract.getDocumentDate()==null					){              relatoriCostTransactionExtract.setDocumentDate(" ");   					}         
		 if(relatoriCostTransactionExtract.getCategoryGroup()==null					){              relatoriCostTransactionExtract.setCategoryGroup(" ");   					}         
		 if(relatoriCostTransactionExtract.getCategory()==null					){                  relatoriCostTransactionExtract.setCategory(" ");   					}             
		 if(relatoriCostTransactionExtract.getDocumentType()==null					){              relatoriCostTransactionExtract.setDocumentType(" ");   					}         
		 if(relatoriCostTransactionExtract.getDocumentNbr()==null					){                  relatoriCostTransactionExtract.setDocumentNbr(" ");   					}         
		 if(relatoriCostTransactionExtract.getDescription()==null					){                  relatoriCostTransactionExtract.setDescription(" ");   					}         
		 if(relatoriCostTransactionExtract.getReferenceNbr()==null					){              relatoriCostTransactionExtract.setReferenceNbr(" ");   					}         
		 if(relatoriCostTransactionExtract.getAccountNbr()==null					){                  relatoriCostTransactionExtract.setAccountNbr(" ");   					}         
		 if(relatoriCostTransactionExtract.getAccountNbrDescription()==null					){      relatoriCostTransactionExtract.setAccountNbrDescription(" ");   					} 
		 if(relatoriCostTransactionExtract.getQuantityAmountHours()==null					){          relatoriCostTransactionExtract.setQuantityAmountHours(" ");   					} 
		 if(relatoriCostTransactionExtract.getGlobalAmount()==null					){              relatoriCostTransactionExtract.setGlobalAmount(" ");   					}         
		 if(relatoriCostTransactionExtract.getObjectAmount()==null					){              relatoriCostTransactionExtract.setObjectAmount(" ");   					}         
		 if(relatoriCostTransactionExtract.getObjectCurrency()==null					){              relatoriCostTransactionExtract.setObjectCurrency(" ");   					}     
		 if(relatoriCostTransactionExtract.getTransactionalAmt()==null					){          relatoriCostTransactionExtract.setTransactionalAmt(" ");   					}     
		 if(relatoriCostTransactionExtract.getTransactionalCurrency()==null					){      relatoriCostTransactionExtract.setTransactionalCurrency(" ");   					} 
		 if(relatoriCostTransactionExtract.getReportingAmount()==null					){              relatoriCostTransactionExtract.setReportingAmount(" ");   					}     
		 if(relatoriCostTransactionExtract.getReportingCurrency()==null					){          relatoriCostTransactionExtract.setReportingCurrency(" ");   					}     
		 if(relatoriCostTransactionExtract.getEnterpriseID()==null					){              relatoriCostTransactionExtract.setEnterpriseID(" ");   					}         
		 if(relatoriCostTransactionExtract.getResourceName()==null					){              relatoriCostTransactionExtract.setResourceName(" ");   					}         
		 if(relatoriCostTransactionExtract.getPersonnelNumber()==null					){              relatoriCostTransactionExtract.setPersonnelNumber(" ");   					}     
		 if(relatoriCostTransactionExtract.getCurrentWorklocation()==null					){          relatoriCostTransactionExtract.setCurrentWorklocation(" ");   					} 
		 if(relatoriCostTransactionExtract.getCurrentHomeOffice()==null					){          relatoriCostTransactionExtract.setCurrentHomeOffice(" ");   					}     
		 if(relatoriCostTransactionExtract.getCostCenterID()==null					){              relatoriCostTransactionExtract.setCostCenterID(" ");   					}         
		 if(relatoriCostTransactionExtract.getCostCenterName()==null					){              relatoriCostTransactionExtract.setCostCenterName(" ");   					}     
		 if(relatoriCostTransactionExtract.getCostCollectorID()==null					){              relatoriCostTransactionExtract.setCostCollectorID(" ");   					}     
		 if(relatoriCostTransactionExtract.getCostCollectorName()==null					){          relatoriCostTransactionExtract.setCostCollectorName(" ");   					}     
		 if(relatoriCostTransactionExtract.getWBS()==null					){                          relatoriCostTransactionExtract.setWBS(" ");   					}                 
		 if(relatoriCostTransactionExtract.getWBSDescription()==null					){              relatoriCostTransactionExtract.setWBSDescription(" ");   					}     
		 if(relatoriCostTransactionExtract.getWBSProfitCenter()==null					){              relatoriCostTransactionExtract.setWBSProfitCenter(" ");   					}     
		 if(relatoriCostTransactionExtract.getWBSProfitCenterName()==null					){          relatoriCostTransactionExtract.setWBSProfitCenterName(" ");   					} 
		 if(relatoriCostTransactionExtract.getParentWBS()==null					){                  relatoriCostTransactionExtract.setParentWBS(" ");   					}             
		 if(relatoriCostTransactionExtract.getContractID()==null					){                  relatoriCostTransactionExtract.setContractID(" ");   					}         
		 if(relatoriCostTransactionExtract.getWBSRaKey()==null					){                  relatoriCostTransactionExtract.setWBSRaKey(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel1()==null					){                  relatoriCostTransactionExtract.setWMULevel1(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel2()==null					){                  relatoriCostTransactionExtract.setWMULevel2(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel3()==null					){                  relatoriCostTransactionExtract.setWMULevel3(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel4()==null					){                  relatoriCostTransactionExtract.setWMULevel4(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel5()==null					){                  relatoriCostTransactionExtract.setWMULevel5(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel6()==null					){                  relatoriCostTransactionExtract.setWMULevel6(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel7()==null					){                  relatoriCostTransactionExtract.setWMULevel7(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel8()==null					){                  relatoriCostTransactionExtract.setWMULevel8(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel9()==null					){                  relatoriCostTransactionExtract.setWMULevel9(" ");   					}             
		 if(relatoriCostTransactionExtract.getWMULevel10()==null					){                  relatoriCostTransactionExtract.setWMULevel10(" ");   					}         
	
	}
	
	public static boolean relatorioCostTransactionExtractPossuiTodosCamposNulos(RelatorioCostTransactionExtract relatorioCostTransactionExtract) {
		
		boolean relatorioPossuiTodosCamposNulos = false; 
		if (
			 relatorioCostTransactionExtract.getContrato()==null 						&&
			 relatorioCostTransactionExtract.getPostingPeriod()==null					&&
			 relatorioCostTransactionExtract.getFiscalYearPeriod()==null				&&
			 relatorioCostTransactionExtract.getFiscalYearQuarter()==null				&&
			 relatorioCostTransactionExtract.getFiscalYear()==null						&&
			 relatorioCostTransactionExtract.getPostingDate()==null					&&
			 relatorioCostTransactionExtract.getDocumentDate()==null					&&
			 relatorioCostTransactionExtract.getCategoryGroup()==null					&&
			 relatorioCostTransactionExtract.getCategory()==null						&&
			 relatorioCostTransactionExtract.getDocumentType()==null					&&
			 relatorioCostTransactionExtract.getDocumentNbr()==null					&&
			 relatorioCostTransactionExtract.getDescription()==null					&&
			 relatorioCostTransactionExtract.getReferenceNbr()==null					&&
			 relatorioCostTransactionExtract.getAccountNbr()==null						&&
			 relatorioCostTransactionExtract.getAccountNbrDescription()==null			&&
			 relatorioCostTransactionExtract.getQuantityAmountHours()==null			&&
			 relatorioCostTransactionExtract.getGlobalAmount()==null					&&
			 relatorioCostTransactionExtract.getObjectAmount()==null					&&
			 relatorioCostTransactionExtract.getObjectCurrency()==null					&&
			 relatorioCostTransactionExtract.getTransactionalAmt()==null			    &&
			 relatorioCostTransactionExtract.getTransactionalCurrency()==null			&&
			 relatorioCostTransactionExtract.getReportingAmount()==null				&&
			 relatorioCostTransactionExtract.getReportingCurrency()==null				&&
			 relatorioCostTransactionExtract.getEnterpriseID()==null					&&
			 relatorioCostTransactionExtract.getResourceName()==null					&&
			 relatorioCostTransactionExtract.getPersonnelNumber()==null				&&
			 relatorioCostTransactionExtract.getCurrentWorklocation()==null			&&
			 relatorioCostTransactionExtract.getCurrentHomeOffice()==null				&&
			 relatorioCostTransactionExtract.getCostCenterID()==null					&&
			 relatorioCostTransactionExtract.getCostCenterName()==null					&&
			 relatorioCostTransactionExtract.getCostCollectorID()==null				&&
			 relatorioCostTransactionExtract.getCostCollectorName()==null				&&
			 relatorioCostTransactionExtract.getWBS()==null							&&
			 relatorioCostTransactionExtract.getWBSDescription()==null					&&
			 relatorioCostTransactionExtract.getWBSProfitCenter()==null				&&
			 relatorioCostTransactionExtract.getWBSProfitCenterName()==null			&&
			 relatorioCostTransactionExtract.getParentWBS()==null						&&
			 relatorioCostTransactionExtract.getContractID()==null						&&
			 relatorioCostTransactionExtract.getWBSRaKey()==null						&&
			 relatorioCostTransactionExtract.getWMULevel1()==null						&&
			 relatorioCostTransactionExtract.getWMULevel2()==null						&&
			 relatorioCostTransactionExtract.getWMULevel3()==null						&&
			 relatorioCostTransactionExtract.getWMULevel4()==null						&&
			 relatorioCostTransactionExtract.getWMULevel5()==null						&&
			 relatorioCostTransactionExtract.getWMULevel6()==null						&&
			 relatorioCostTransactionExtract.getWMULevel7()==null						&&
			 relatorioCostTransactionExtract.getWMULevel8()==null						&&
			 relatorioCostTransactionExtract.getWMULevel9()==null						&&
			 relatorioCostTransactionExtract.getWMULevel10()==null						) {
			
			relatorioPossuiTodosCamposNulos = true;
		 }
		
		return relatorioPossuiTodosCamposNulos;	 
	
	}
	
	public static void converteValorNullParaEspacoEmBrancoRelatorioResourceTrend(RelatorioResourceTrend relatorioResourceTrend) {
		
		 if(relatorioResourceTrend.getContrato()                        == null 	){  relatorioResourceTrend.setContrato(" ");   					}                 
		 if(relatorioResourceTrend.getForecastVersion()					== null 	){  relatorioResourceTrend.setForecastVersion(" ");				}
		 if(relatorioResourceTrend.getTypeCostCenterCareerTrack()       == null 	){  relatorioResourceTrend.setTypeCostCenterCareerTrack(" ");   }
		 if(relatorioResourceTrend.getLevel()                           == null 	){  relatorioResourceTrend.setLevel(" ");						}
		 if(relatorioResourceTrend.getWMUID()                           == null 	){  relatorioResourceTrend.setWMUID(" ");   					}                 
		 if(relatorioResourceTrend.getWMUName()                         == null 	){  relatorioResourceTrend.setWMUName(" ");   					}                 
		 if(relatorioResourceTrend.getWMUOwner()                        == null 	){  relatorioResourceTrend.setWMUOwner(" ");   					}               
		 if(relatorioResourceTrend.getWBS()                             == null 	){  relatorioResourceTrend.setWBS(" ");   					    }              
		 if(relatorioResourceTrend.getCostCollector()                   == null 	){  relatorioResourceTrend.setCostCollector(" ");				}
		 if(relatorioResourceTrend.getCostCollectorName()               == null 	){  relatorioResourceTrend.setCostCollectorName(" ");           }
		 if(relatorioResourceTrend.getRoleName()                        == null 	){  relatorioResourceTrend.setRoleName(" ");   					}     
		 if(relatorioResourceTrend.getName()                            == null 	){  relatorioResourceTrend.setName(" ");   						}              
		 if(relatorioResourceTrend.getPersonnelNumber()                 == null 	){  relatorioResourceTrend.setPersonnelNumber(" ");				}
		 if(relatorioResourceTrend.getHomeLocation()                    == null 	){  relatorioResourceTrend.setHomeLocation(" ");   				}       
		 if(relatorioResourceTrend.getWorkLocation()                    == null 	){  relatorioResourceTrend.setWorkLocation(" ");   				}          
		 if(relatorioResourceTrend.getCategory()                        == null 	){  relatorioResourceTrend.setCategory(" ");   					}          
		 if(relatorioResourceTrend.getQuantity()                        == null 	){  relatorioResourceTrend.setQuantity(" ");   					}              
		 if(relatorioResourceTrend.getActualForecast()                  == null 	){  relatorioResourceTrend.setActualForecast(" ");				}
		 if(relatorioResourceTrend.getDate()                            == null 	){  relatorioResourceTrend.setDate(" ");   						}	        
		 if(relatorioResourceTrend.getOrderBy()                         == null 	){  relatorioResourceTrend.setOrderBy(" ");   					}                  
		 if(relatorioResourceTrend.getBillRateCardId()                  == null 	){  relatorioResourceTrend.setBillRateCardId(" ");				}
		 if(relatorioResourceTrend.getRateName()                        == null 	){  relatorioResourceTrend.setRateName(" ");   					}        
		 if(relatorioResourceTrend.getBillRateCardName()                == null 	){  relatorioResourceTrend.setBillRateCardName(" ");			}
		
	}
	
	public static boolean relatorioResourceTrendPossuiTodosCamposNulos(RelatorioResourceTrend relatorioResourceTrend) {
		
		boolean relatorioPossuiTodosCamposNulos = false; 
		if (
			 relatorioResourceTrend.getContrato()                        == null 						&&
			 relatorioResourceTrend.getForecastVersion()				 == null 						&&
			 relatorioResourceTrend.getTypeCostCenterCareerTrack()       == null 						&&
			 relatorioResourceTrend.getLevel()                           == null 						&&
			 relatorioResourceTrend.getWMUID()                           == null 						&&
			 relatorioResourceTrend.getWMUName()                         == null 						&&
			 relatorioResourceTrend.getWMUOwner()                        == null 						&&
			 relatorioResourceTrend.getWBS()                             == null 						&&
			 relatorioResourceTrend.getCostCollector()                   == null 						&&
			 relatorioResourceTrend.getCostCollectorName()               == null 						&&
			 relatorioResourceTrend.getRoleName()                        == null 						&&
			 relatorioResourceTrend.getName()                            == null 						&&
			 relatorioResourceTrend.getPersonnelNumber()                 == null 						&&
			 relatorioResourceTrend.getHomeLocation()                    == null 						&&
			 relatorioResourceTrend.getWorkLocation()                    == null 						&&
			 relatorioResourceTrend.getCategory()                        == null 						&&
			 relatorioResourceTrend.getQuantity()                        == null 						&&
			 relatorioResourceTrend.getActualForecast()                  == null 						&&
			 relatorioResourceTrend.getDate()                            == null 						&&
			 relatorioResourceTrend.getOrderBy()                         == null 						&&
			 relatorioResourceTrend.getBillRateCardId()                  == null 						&&
			 relatorioResourceTrend.getRateName()                        == null 						&&
			 relatorioResourceTrend.getBillRateCardName()                == null						) {
			
			relatorioPossuiTodosCamposNulos = true;
		 }
		
		return relatorioPossuiTodosCamposNulos;	 
	
	}
	
	public static void converteValorNullParaEspacoEmBrancoRelatorioMultiSegmentContractReport(RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport) {
		
		if(relatorioMultiSegmentContractReport.getContrato()       	     == null 	){ relatorioMultiSegmentContractReport.setContrato(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel01()       	     == null 	){ relatorioMultiSegmentContractReport.setLevel01(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel02()	             == null 	){ relatorioMultiSegmentContractReport.setLevel02(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel03()              == null 	){ relatorioMultiSegmentContractReport.setLevel03(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel04()              == null 	){ relatorioMultiSegmentContractReport.setLevel04(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel05()              == null 	){ relatorioMultiSegmentContractReport.setLevel05(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel06()              == null 	){ relatorioMultiSegmentContractReport.setLevel06(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel07()              == null 	){ relatorioMultiSegmentContractReport.setLevel07(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel08()              == null 	){ relatorioMultiSegmentContractReport.setLevel08(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel09()              == null 	){ relatorioMultiSegmentContractReport.setLevel09(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel10()              == null 	){ relatorioMultiSegmentContractReport.setLevel10(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel11()              == null 	){ relatorioMultiSegmentContractReport.setLevel11(" ");   			}
		if(relatorioMultiSegmentContractReport.getLevel12()              == null 	){ relatorioMultiSegmentContractReport.setLevel12(" ");   			}
		if(relatorioMultiSegmentContractReport.getWBSNbr()               == null 	){ relatorioMultiSegmentContractReport.setWBSNbr(" ");   			}
		if(relatorioMultiSegmentContractReport.getWBSName()       		 == null 	){ relatorioMultiSegmentContractReport.setWBSName(" ");   			}
		if(relatorioMultiSegmentContractReport.getCostCollectorNbr() 	 == null 	){ relatorioMultiSegmentContractReport.setCostCollectorNbr(" ");    }
		if(relatorioMultiSegmentContractReport.getCostCollectorName()    == null 	){ relatorioMultiSegmentContractReport.setCostCollectorName(" ");   }
		if(relatorioMultiSegmentContractReport.getFiscalYear()           == null 	){ relatorioMultiSegmentContractReport.setFiscalYear(" ");   		}
		if(relatorioMultiSegmentContractReport.getFiscalQuarter()        == null 	){ relatorioMultiSegmentContractReport.setFiscalQuarter(" ");   	}
		if(relatorioMultiSegmentContractReport.getMonth()                == null 	){ relatorioMultiSegmentContractReport.setMonth(" ");   			}
		if(relatorioMultiSegmentContractReport.getGroup()                == null 	){ relatorioMultiSegmentContractReport.setGroup(" ");   			}
		if(relatorioMultiSegmentContractReport.getCategory()             == null 	){ relatorioMultiSegmentContractReport.setCategory(" ");   			}
		if(relatorioMultiSegmentContractReport.getActual()               == null 	){ relatorioMultiSegmentContractReport.setActual(" ");   			}
		if(relatorioMultiSegmentContractReport.getForecast()             == null 	){ relatorioMultiSegmentContractReport.setForecast(" ");   			}
		if(relatorioMultiSegmentContractReport.getComparisonEAC()        == null 	){ relatorioMultiSegmentContractReport.setComparisonEAC(" ");   	}
		
	}
	
	public static boolean relatorioMultiSegmentContractReportPossuiTodosCamposNulos(RelatorioMultiSegmentContractReport relatorioMultiSegmentContractReport) {
		
		boolean relatorioPossuiTodosCamposNulos = false; 
		if (
			 relatorioMultiSegmentContractReport.getContrato()       	             == null 						&& 
 			 relatorioMultiSegmentContractReport.getLevel01()       	             == null 						&&
			 relatorioMultiSegmentContractReport.getLevel02()	                     == null 						&&
			 relatorioMultiSegmentContractReport.getLevel03()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel04()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel05()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel06()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel07()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel08()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel09()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel10()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel11()                        == null 						&&
			 relatorioMultiSegmentContractReport.getLevel12()                        == null 						&&
			 relatorioMultiSegmentContractReport.getWBSNbr()                         == null 						&&
			 relatorioMultiSegmentContractReport.getWBSName()       				 == null 						&&
			 relatorioMultiSegmentContractReport.getCostCollectorNbr()       		 == null 						&&
			 relatorioMultiSegmentContractReport.getCostCollectorName()              == null 						&&
			 relatorioMultiSegmentContractReport.getFiscalYear()                     == null 						&&
			 relatorioMultiSegmentContractReport.getFiscalQuarter()                  == null 						&&
			 relatorioMultiSegmentContractReport.getMonth()                          == null 						&&
			 relatorioMultiSegmentContractReport.getGroup()                          == null 						&&
			 relatorioMultiSegmentContractReport.getCategory()                       == null 						&&
			 relatorioMultiSegmentContractReport.getActual()                         == null 						&&
			 relatorioMultiSegmentContractReport.getForecast()                       == null 						&&
			 relatorioMultiSegmentContractReport.getComparisonEAC()                  == null 						) {
			
			relatorioPossuiTodosCamposNulos = true;
		 }
		
		return relatorioPossuiTodosCamposNulos;	 
	
	}
	
}