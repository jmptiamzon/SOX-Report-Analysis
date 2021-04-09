package com.sprint.sox.model;

import java.util.ArrayList;
import java.util.List;

public class Resku {
	
	private List<String> header;
	private List<String> trxNo;
	private List<String> skuId;
	private List<String> warehouse;
	private List<String> location;
	private List<String> invadjTypeQty;
	private List<String> errorCode;
	private List<String> creationDate;
	private List<String> day;
	private List<String> hour;
	private List<String> minute;
	private List<String> status;
	private List<String> comment;
	private List<String> soxResearch;
	
	
	public Resku() {
		header = new ArrayList<>();
		trxNo = new ArrayList<>();
		skuId = new ArrayList<>();
		warehouse = new ArrayList<>();
		location = new ArrayList<>();
		invadjTypeQty = new ArrayList<>();
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
	

	public List<String> getTrxNo() {
		return trxNo;
	}


	public void setTrxNo(String trxNo) {
		String pattern = "^0+(?!$)";
		this.trxNo.add(trxNo.replaceAll(pattern, ""));
	}


	public List<String> getSkuId() {
		return skuId;
	}


	public void setSkuId(String skuId) {
		this.skuId.add(skuId);
	}


	public List<String> getWarehouse() {
		return warehouse;
	}


	public void setWarehouse(String warehouse) {
		this.warehouse.add(warehouse);
	}


	public List<String> getLocation() {
		return location;
	}


	public void setLocation(String location) {
		this.location.add(location);
	}


	public List<String> getInvadjTypeQty() {
		return invadjTypeQty;
	}


	public void setInvadjTypeQty(String invadjTypeQty) {
		this.invadjTypeQty.add(invadjTypeQty);
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
