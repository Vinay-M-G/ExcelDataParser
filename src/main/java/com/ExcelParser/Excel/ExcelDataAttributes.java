package com.ExcelParser.Excel;

import java.util.Arrays;
import java.util.List;

public class ExcelDataAttributes {
	
	private String partyName;
	private String invoiceNo;
	private double invoiceAmount;
	private String pvn;
	private double deductedAmount;
	private String deductionReason;
	private String paymentDate;
	private String bankReference;
	private String paymentType;
	private double transferAmount;
	
	public String getPartyName() {
		return partyName;
	}

	public void setPartyName(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 1;
		int rowDist = 0;
		String partyName = excelUtilities.getCellValueInString(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.partyName = partyName;
	}

	public String getInvoiceNo() {
		return invoiceNo;
	}

	public void setInvoiceNo(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 2;
		int rowDist = 2;
		String invoiceNo = excelUtilities.getCellValueInString(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.invoiceNo = invoiceNo;
	}

	public double getInvoiceAmount() {
		return invoiceAmount;
	}

	public void setInvoiceAmount(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 3;
		int rowDist = 2;
		double invoiceAmount = excelUtilities.getCellValueInDouble(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.invoiceAmount = invoiceAmount;
	}

	public String getPvn() {
		return pvn;
	}

	public void setPvn(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 4;
		int rowDist = 0;
		String pvn = excelUtilities.getCellValueInString(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.pvn = pvn;
	}

	public double getDeductedAmount() {
		return deductedAmount;
	}

	public void setDeductedAmount(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 3;
		int rowDist = 1;
		double deductedAmount = excelUtilities.getCellValueInDouble(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.deductedAmount = deductedAmount;
	}

	public String getDeductionReason() {
		return deductionReason;
	}

	public void setDeductionReason(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 2;
		int rowDist = 1;
		String deductionReason = excelUtilities.getCellValueInString(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.deductionReason = deductionReason;
	}

	public String getPaymentDate() {
		return paymentDate;
	}

	public void setPaymentDate(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 3;
		int rowDist = 4;
		String paymentDate = excelUtilities.getCellValueInDateFormat(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.paymentDate = paymentDate;
	}

	public String getBankReference() {
		return bankReference;
	}

	public void setBankReference(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 2;
		int rowDist = 4;
		String bankReference = excelUtilities.getCellValueInString(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.bankReference = bankReference;
	}

	public String getPaymentType() {
		return paymentType;
	}

	public void setPaymentType(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 2;
		int rowDist = 4;
		String paymentType = excelUtilities.getCellValueInString(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.paymentType = paymentType;
	}

	public double getTransferAmount() {
		return transferAmount;
	}

	public void setTransferAmount(ExcelUtilities excelUtilities, int referenceRowId, int referenceColumnId) {
		int columnDist = 5;
		int rowDist = 0;
		double transferAmount = excelUtilities.getCellValueInDouble(referenceColumnId + columnDist, referenceRowId + rowDist);
		this.transferAmount = transferAmount;
	}

	public void setDataAttributes(List<String> dataAttributes) {
		this.dataAttributes = dataAttributes;
	}

	private static List<String> dataAttributes = Arrays.asList("Party Name", "Invoice No.", "Invoice Amount", "PVN (Tally)",
			"Amt. Deducted", "Reason For Deduction", "Payment Date (DDMMYYYY)", "Bank Ref.",
			"Payment Type", "Bank Transfer Amount");
	
	public List<String> getDataAttributes(){
		return dataAttributes;
	}
	
	
}
  