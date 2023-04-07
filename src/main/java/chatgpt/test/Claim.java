package chatgpt.test;

import java.util.Date;
import java.util.List;

public class Claim {
	private String claimNumber;
	private String claimCategory; 
	private String employeeCode; 
	private String claimDesc; 
	private Date receiptDate; 
	private Double claimedAmount; 
	private Date submissionDate; 
	private String claimStatus; 
	private Double approvedAmount; 
	private Date approvalDate; 
	private String approvedBy;
	
	public String getClaimNumber() {
		return claimNumber;
	}
	public void setClaimNumber(String claimNumber) {
		this.claimNumber = claimNumber;
	}
	public String getClaimCategory() {
		return claimCategory;
	}
	public void setClaimCategory(String claimCategory) {
		this.claimCategory = claimCategory;
	}
	public String getEmployeeCode() {
		return employeeCode;
	}
	public void setEmployeeCode(String employeeCode) {
		this.employeeCode = employeeCode;
	}
	public String getClaimDesc() {
		return claimDesc;
	}
	public void setClaimDesc(String claimDesc) {
		this.claimDesc = claimDesc;
	}
	public Date getReceiptDate() {
		return receiptDate;
	}
	public void setReceiptDate(Date receiptDate) {
		this.receiptDate = receiptDate;
	}
	public Double getClaimedAmount() {
		return claimedAmount;
	}
	public void setClaimedAmount(Double claimedAmount) {
		this.claimedAmount = claimedAmount;
	}
	public Date getSubmissionDate() {
		return submissionDate;
	}
	public void setSubmissionDate(Date submissionDate) {
		this.submissionDate = submissionDate;
	}
	public String getClaimStatus() {
		return claimStatus;
	}
	public void setClaimStatus(String claimStatus) {
		this.claimStatus = claimStatus;
	}
	public Double getApprovedAmount() {
		return approvedAmount;
	}
	public void setApprovedAmount(Double approvedAmount) {
		this.approvedAmount = approvedAmount;
	}
	public Date getApprovalDate() {
		return approvalDate;
	}
	public void setApprovalDate(Date approvalDate) {
		this.approvalDate = approvalDate;
	}
	public String getApprovedBy() {
		return approvedBy;
	}
	public void setApprovedBy(String approvedBy) {
		this.approvedBy = approvedBy;
	}
//class for evaluation
}
