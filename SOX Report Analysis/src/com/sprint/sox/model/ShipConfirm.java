package com.sprint.sox.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class ShipConfirm implements Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private List<String> header;
	private List<String> controlNumber;
	private List<String> shippedQty;
	private List<String> batchId; 
	private List<String> errorCode;
	private List<String> creationDate;
	private List<String> day;
	private List<String> hour;
	private List<String> minute;
	private List<String> status;
	private List<String> comment;
	private List<String> soxResearch;
	
	
	public ShipConfirm() {
		header = new ArrayList<>();
		controlNumber = new ArrayList<>();
		shippedQty = new ArrayList<>();
		batchId = new ArrayList<>();
		errorCode = new ArrayList<>();
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
	

	public List<String> getControlNumber() {
		return controlNumber;
	}


	public void setControlNumber(String controlNumber) {
		String pattern = "^0+(?!$)";
		this.controlNumber.add(controlNumber.replaceAll(pattern, ""));
	}


	public List<String> getShippedQty() {
		return shippedQty;
	}


	public void setShippedQty(String shippedQty) {
		this.shippedQty.add(shippedQty);
	}


	public List<String> getBatchId() {
		return batchId;
	}


	public void setBatchId(String batchId) {
		this.batchId.add(batchId);
	}


	public List<String> getErrorCode() {
		return errorCode;
	}


	public void setErrorCode(String errorCode) {
		this.errorCode.add(errorCode);
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
		this.status.add(status) ;
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
