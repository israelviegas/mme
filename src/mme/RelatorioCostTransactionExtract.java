package mme;

import java.sql.Date;

public class RelatorioCostTransactionExtract {
	
	String	Contrato;
	String	PostingPeriod;
	String	FiscalYearPeriod;
	String	FiscalYearQuarter;
	String	FiscalYear;
	String	PostingDate;
	String	DocumentDate;
	String	CategoryGroup;
	String	Category;
	String	DocumentType;
	String	DocumentNbr;
	String	Description;
	String	ReferenceNbr;
	String	AccountNbr;
	String	AccountNbrDescription;
	String	QuantityAmountHours;
	String	GlobalAmount;
	String	ObjectAmount;
	String	ObjectCurrency;
	String	TransactionalAmt;
	String	TransactionalCurrency;
	String	ReportingAmount;
	String	ReportingCurrency;
	String	EnterpriseID;
	String	ResourceName;
	String	PersonnelNumber;
	String	CurrentWorklocation;
	String	CurrentHomeOffice;
	String	CostCenterID;
	String	CostCenterName;
	String	CostCollectorID;
	String	CostCollectorName;
	String	WBS;
	String	WBSDescription;
	String	WBSProfitCenter;
	String	WBSProfitCenterName;
	String	ParentWBS;
	String	ContractID;
	String	WBSRaKey;
	String	WMULevel1;
	String	WMULevel2;
	String	WMULevel3;
	String	WMULevel4;
	String	WMULevel5;
	String	WMULevel6;
	String	WMULevel7;
	String	WMULevel8;
	String	WMULevel9;
	String	WMULevel10;
	Date	dataExtracao;
	Date    dataReferencia;
	String	periodo;
	
	public String getContrato() {
		return Contrato;
	}
	public void setContrato(String contrato) {
		Contrato = contrato;
	}
	public String getPostingPeriod() {
		return PostingPeriod;
	}
	public void setPostingPeriod(String postingPeriod) {
		PostingPeriod = postingPeriod;
	}
	public String getFiscalYearPeriod() {
		return FiscalYearPeriod;
	}
	public void setFiscalYearPeriod(String fiscalYearPeriod) {
		FiscalYearPeriod = fiscalYearPeriod;
	}
	public String getFiscalYearQuarter() {
		return FiscalYearQuarter;
	}
	public void setFiscalYearQuarter(String fiscalYearQuarter) {
		FiscalYearQuarter = fiscalYearQuarter;
	}
	public String getFiscalYear() {
		return FiscalYear;
	}
	public void setFiscalYear(String fiscalYear) {
		FiscalYear = fiscalYear;
	}
	public String getPostingDate() {
		return PostingDate;
	}
	public void setPostingDate(String postingDate) {
		PostingDate = postingDate;
	}
	public String getDocumentDate() {
		return DocumentDate;
	}
	public void setDocumentDate(String documentDate) {
		DocumentDate = documentDate;
	}
	public String getCategoryGroup() {
		return CategoryGroup;
	}
	public void setCategoryGroup(String categoryGroup) {
		CategoryGroup = categoryGroup;
	}
	public String getCategory() {
		return Category;
	}
	public void setCategory(String category) {
		Category = category;
	}
	public String getDocumentType() {
		return DocumentType;
	}
	public void setDocumentType(String documentType) {
		DocumentType = documentType;
	}
	public String getDocumentNbr() {
		return DocumentNbr;
	}
	public void setDocumentNbr(String documentNbr) {
		DocumentNbr = documentNbr;
	}
	public String getDescription() {
		return Description;
	}
	public void setDescription(String description) {
		Description = description;
	}
	public String getReferenceNbr() {
		return ReferenceNbr;
	}
	public void setReferenceNbr(String referenceNbr) {
		ReferenceNbr = referenceNbr;
	}
	public String getAccountNbr() {
		return AccountNbr;
	}
	public void setAccountNbr(String accountNbr) {
		AccountNbr = accountNbr;
	}
	public String getAccountNbrDescription() {
		return AccountNbrDescription;
	}
	public void setAccountNbrDescription(String accountNbrDescription) {
		AccountNbrDescription = accountNbrDescription;
	}
	public String getQuantityAmountHours() {
		return QuantityAmountHours;
	}
	public void setQuantityAmountHours(String quantityAmountHours) {
		QuantityAmountHours = quantityAmountHours;
	}
	public String getGlobalAmount() {
		return GlobalAmount;
	}
	public void setGlobalAmount(String globalAmount) {
		GlobalAmount = globalAmount;
	}
	public String getObjectAmount() {
		return ObjectAmount;
	}
	public void setObjectAmount(String objectAmount) {
		ObjectAmount = objectAmount;
	}
	public String getObjectCurrency() {
		return ObjectCurrency;
	}
	public void setObjectCurrency(String objectCurrency) {
		ObjectCurrency = objectCurrency;
	}
	public String getTransactionalAmt() {
		return TransactionalAmt;
	}
	public void setTransactionalAmt(String transactionalAmt) {
		TransactionalAmt = transactionalAmt;
	}
	public String getTransactionalCurrency() {
		return TransactionalCurrency;
	}
	public void setTransactionalCurrency(String transactionalCurrency) {
		TransactionalCurrency = transactionalCurrency;
	}
	public String getReportingAmount() {
		return ReportingAmount;
	}
	public void setReportingAmount(String reportingAmount) {
		ReportingAmount = reportingAmount;
	}
	public String getReportingCurrency() {
		return ReportingCurrency;
	}
	public void setReportingCurrency(String reportingCurrency) {
		ReportingCurrency = reportingCurrency;
	}
	public String getEnterpriseID() {
		return EnterpriseID;
	}
	public void setEnterpriseID(String enterpriseID) {
		EnterpriseID = enterpriseID;
	}
	public String getResourceName() {
		return ResourceName;
	}
	public void setResourceName(String resourceName) {
		ResourceName = resourceName;
	}
	public String getPersonnelNumber() {
		return PersonnelNumber;
	}
	public void setPersonnelNumber(String personnelNumber) {
		PersonnelNumber = personnelNumber;
	}
	public String getCurrentWorklocation() {
		return CurrentWorklocation;
	}
	public void setCurrentWorklocation(String currentWorklocation) {
		CurrentWorklocation = currentWorklocation;
	}
	public String getCurrentHomeOffice() {
		return CurrentHomeOffice;
	}
	public void setCurrentHomeOffice(String currentHomeOffice) {
		CurrentHomeOffice = currentHomeOffice;
	}
	public String getCostCenterID() {
		return CostCenterID;
	}
	public void setCostCenterID(String costCenterID) {
		CostCenterID = costCenterID;
	}
	public String getCostCenterName() {
		return CostCenterName;
	}
	public void setCostCenterName(String costCenterName) {
		CostCenterName = costCenterName;
	}
	public String getCostCollectorID() {
		return CostCollectorID;
	}
	public void setCostCollectorID(String costCollectorID) {
		CostCollectorID = costCollectorID;
	}
	public String getCostCollectorName() {
		return CostCollectorName;
	}
	public void setCostCollectorName(String costCollectorName) {
		CostCollectorName = costCollectorName;
	}
	public String getWBS() {
		return WBS;
	}
	public void setWBS(String wBS) {
		WBS = wBS;
	}
	public String getWBSDescription() {
		return WBSDescription;
	}
	public void setWBSDescription(String wBSDescription) {
		WBSDescription = wBSDescription;
	}
	public String getWBSProfitCenter() {
		return WBSProfitCenter;
	}
	public void setWBSProfitCenter(String wBSProfitCenter) {
		WBSProfitCenter = wBSProfitCenter;
	}
	public String getWBSProfitCenterName() {
		return WBSProfitCenterName;
	}
	public void setWBSProfitCenterName(String wBSProfitCenterName) {
		WBSProfitCenterName = wBSProfitCenterName;
	}
	public String getParentWBS() {
		return ParentWBS;
	}
	public void setParentWBS(String parentWBS) {
		ParentWBS = parentWBS;
	}
	public String getContractID() {
		return ContractID;
	}
	public void setContractID(String contractID) {
		ContractID = contractID;
	}
	public String getWBSRaKey() {
		return WBSRaKey;
	}
	public void setWBSRaKey(String wBSRaKey) {
		WBSRaKey = wBSRaKey;
	}
	public String getWMULevel1() {
		return WMULevel1;
	}
	public void setWMULevel1(String wMULevel1) {
		WMULevel1 = wMULevel1;
	}
	public String getWMULevel2() {
		return WMULevel2;
	}
	public void setWMULevel2(String wMULevel2) {
		WMULevel2 = wMULevel2;
	}
	public String getWMULevel3() {
		return WMULevel3;
	}
	public void setWMULevel3(String wMULevel3) {
		WMULevel3 = wMULevel3;
	}
	public String getWMULevel4() {
		return WMULevel4;
	}
	public void setWMULevel4(String wMULevel4) {
		WMULevel4 = wMULevel4;
	}
	public String getWMULevel5() {
		return WMULevel5;
	}
	public void setWMULevel5(String wMULevel5) {
		WMULevel5 = wMULevel5;
	}
	public String getWMULevel6() {
		return WMULevel6;
	}
	public void setWMULevel6(String wMULevel6) {
		WMULevel6 = wMULevel6;
	}
	public String getWMULevel7() {
		return WMULevel7;
	}
	public void setWMULevel7(String wMULevel7) {
		WMULevel7 = wMULevel7;
	}
	public String getWMULevel8() {
		return WMULevel8;
	}
	public void setWMULevel8(String wMULevel8) {
		WMULevel8 = wMULevel8;
	}
	public String getWMULevel9() {
		return WMULevel9;
	}
	public void setWMULevel9(String wMULevel9) {
		WMULevel9 = wMULevel9;
	}
	public String getWMULevel10() {
		return WMULevel10;
	}
	public void setWMULevel10(String wMULevel10) {
		WMULevel10 = wMULevel10;
	}
	public Date getDataExtracao() {
		return dataExtracao;
	}
	public void setDataExtracao(Date dataExtracao) {
		this.dataExtracao = dataExtracao;
	}
	public Date getDataReferencia() {
		return dataReferencia;
	}
	public void setDataReferencia(Date dataReferencia) {
		this.dataReferencia = dataReferencia;
	}
	public String getPeriodo() {
		return periodo;
	}
	public void setPeriodo(String periodo) {
		this.periodo = periodo;
	}
	
}
