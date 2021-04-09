package com.sprint.sox.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class RmaReceipt implements Serializable {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private List<String> header;
	private List<String> transactionNo;
	private List<String> skuId;
	private List<String> referenceOne;
	private List<String> referenceTwo;
	private List<String> invAdjustmentQty;
	private List<String> creationDate;
	private List<String> day;
	private List<String> hour;
	private List<String> minute;
	private List<String> status;
	private List<String> comment;
	private List<String> soxResearch;
	
	
	public RmaReceipt() {
		header = new ArrayList<>();
		transactionNo = new ArrayList<>();
		skuId = new ArrayList<>();
		referenceOne = new ArrayList<>();
		referenceTwo = new ArrayList<>();
		invAdjustmentQty = new ArrayList<>();
		creationDate = new ArrayList<>();
		day = new ArrayList<>();
		hour = new ArrayList<>();
		minute = new ArrayList<>();
		status = new ArrayList<>();
		comment = new ArrayList<>();
		soxResearch = new ArrayList<>();	
	}
	

	public List<String> getHeader() {
		return header;
	}

	public void setHeader(String header) {
		this.header.add(header);
	}

	public List<String> getTransactionNo() {
		return transactionNo;
	}

	public void setTransactionNo(String transactionNo) {
		String pattern = "^0+(?!$)";
		this.transactionNo.add(transactionNo.replaceAll(pattern, ""));
	}

	public List<String> getSkuId() {
		return skuId;
	}

	public void setSkuId(String skuId) {
		this.skuId.add(skuId);
	}

	public List<String> getReferenceOne() {
		return referenceOne;
	}

	public void setReferenceOne(String referenceOne) {
		this.referenceOne.add(referenceOne);
	}

	public List<String> getReferenceTwo() {
		return referenceTwo;
	}

	public void setReferenceTwo(String referenceTwo) {
		this.referenceTwo.add(referenceTwo);
	}
	
	public List<String> getInvAdjustmentQty() {
		return invAdjustmentQty;
	}

	public void setInvAdjustmentQty(String invAdjustmentQty) {
		this.invAdjustmentQty.add(invAdjustmentQty);
	}


	public List<String> getCreationDate() {
		return creationDate;
	}

	public void setCreationDate(String creationDate) {
		this.creationDate.add(creationDate);
	}

	public List<String> getDay() {
		return day;
	}

	public void setDay(String day) {
		this.day.add(day);
	}

	public List<String> getHour() {
		return hour;
	}

	public void setHour(String hour) {
		this.hour.add(hour);
	}

	public List<String> getMinute() {
		return minute;
	}

	public void setMinute(String minute) {
		this.minute.add(minute);
	}


	public List<String> getStatus() {
		return status;
	}


	public void setStatus(String status) {
		this.status.add(status);
	}


	public List<String> getComment() {
		return comment;
	}


	public void setComment(String comment) {
		this.comment.add(comment);
	}


	public List<String> getSoxResearch() {
		return soxResearch;
	}


	public void setSoxResearch(String soxResearch) {
		this.soxResearch.add(soxResearch);
	}
	

}
