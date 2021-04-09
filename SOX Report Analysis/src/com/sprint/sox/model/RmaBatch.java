package com.sprint.sox.model;

import java.util.ArrayList;
import java.util.List;

public class RmaBatch {
	private List<String> header;
	private List<String> upsRefNo;
	private List<String> origOrdNo;
	private List<String> ordLineNo;
	private List<String> itemId;
	private List<String> qty;
	private List<String> errorCode;
	private List<String> creationDate;
	private List<String> day;
	private List<String> hour;
	private List<String> minute;
	private List<String> status;
	private List<String> comment;
	private List<String> soxResearch;
	
	
	public RmaBatch() {
		header = new ArrayList<>();
		upsRefNo = new ArrayList<>();
		origOrdNo = new ArrayList<>();
		ordLineNo = new ArrayList<>();
		itemId = new ArrayList<>();
		qty = new ArrayList<>();
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
	
	
	public List<String> getUpsRefNo() {
		return upsRefNo;
	}


	public void setUpsRefNo(String upsRefNo) {
		String pattern = "^0+(?!$)";
		this.upsRefNo.add(upsRefNo.replaceAll(pattern, ""));
	}


	public List<String> getOrigOrdNo() {
		return origOrdNo;
	}


	public void setOrigOrdNo(String origOrdNo) {
		this.origOrdNo.add(origOrdNo);
	}


	public List<String> getOrdLineNo() {
		return ordLineNo;
	}


	public void setOrdLineNo(String ordLineNo) {
		this.ordLineNo.add(ordLineNo);
	}


	public List<String> getItemId() {
		return itemId;
	}


	public void setItemId(String itemId) {
		this.itemId.add(itemId);
	}


	public List<String> getQty() {
		return qty;
	}


	public void setQty(String qty) {
		this.qty.add(qty);
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
