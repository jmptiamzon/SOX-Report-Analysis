package com.sprint.sox.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import javax.swing.JLabel;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sprint.sox.model.Model;
import com.sprint.sox.model.MovesAndPo;
import com.sprint.sox.model.Resku;
import com.sprint.sox.model.RmaBatch;
import com.sprint.sox.model.RmaReceipt;
import com.sprint.sox.model.ShipConfirm;
import com.sprint.sox.view.View;

public class Controller {
	private View view;
	private Model model;
	private JLabel statusLbl;
	private Map<String, String[]> queryRecord;
	private String parameters;
	private static final String []SOX_RESEARCH_SHEETS = new String[] {"Moves and PO", "RMA Batch", "RMA Receipt", "Resku", "Ship Confirm"};	
	private static final String SOX_RES_FILE_DIR = System.getProperty("user.home") + "\\Documents\\sox_files\\sox_research.xlsx";
	
	public Controller(View view) {
		statusLbl = view.getStatusLbl();
		statusLbl.setText("Program has started.");
		statusLbl.paintImmediately(statusLbl.getVisibleRect());
		this.view = view;
		model = new Model(this);
	}
	
	public void startApp() {
		queryRecord = new HashMap<>();
		Date start, end;
		long diff, diffMinutes;
		String soxFileDir = "", soxFileDir2 = "", soxFileDir3 = "", soxFileDir4 = "", soxFileDir5 = "";
		String []soxFilePath = view.getSoxFilePath();
		String fileDate = Paths.get(soxFilePath[0]).getFileName().toString().split(",")[1] + "," + Paths.get(soxFilePath[0]).getFileName().toString().split(",")[2];
		String date = "";
		String []dateArr;
		try {
			date = new SimpleDateFormat("dd-MMM-yyyy").format(new SimpleDateFormat("MMMdd,yyyy").parse(fileDate));
		} catch (ParseException e) {
			System.out.println("Date format error: " + e.getMessage());
		}
		
		dateArr = date.split("-");
		date = dateArr[0] + getOrdinalIndicator(Integer.parseInt(dateArr[0])) + "-" + dateArr[1] + "-" + dateArr[2];
		Map<String, Map<String, String>> soxResearch = new HashMap<>();
		MovesAndPo movesAndPo = new MovesAndPo();
		RmaBatch rmaBatch = new RmaBatch();
		RmaReceipt rmaReceipt = new RmaReceipt();
		Resku reskuReport = new Resku();
		ShipConfirm shipConfirm = new ShipConfirm();
		
		
		for (int ctr = 0; ctr < soxFilePath.length; ctr++) {
			if (Paths.get(soxFilePath[ctr]).getFileName().toString().toLowerCase().contains("moves_and_po")) {
				soxFileDir = soxFilePath[ctr];
			}
			
			else if (Paths.get(soxFilePath[ctr]).getFileName().toString().toLowerCase().contains("rma_batch")) {
				soxFileDir2 = soxFilePath[ctr];
			}
			
			else if (Paths.get(soxFilePath[ctr]).getFileName().toString().toLowerCase().contains("rma_receipts")) {
				soxFileDir3 = soxFilePath[ctr];
			}
			
			else if (Paths.get(soxFilePath[ctr]).getFileName().toString().toLowerCase().contains("rsku_report")) {
				soxFileDir4 = soxFilePath[ctr];
			}
			
			else if (Paths.get(soxFilePath[ctr]).getFileName().toString().toLowerCase().contains("ship_confirm")) {
				soxFileDir5 = soxFilePath[ctr];
			}
		}
		
		start = new Date();
		
		readSoxResearch(soxResearch);

		readFiles(movesAndPo, soxFileDir);
		readFiles(rmaBatch, soxFileDir2);
		readFiles(rmaReceipt, soxFileDir3);
		readFiles(reskuReport, soxFileDir4);
		readFiles(shipConfirm, soxFileDir5);
		
		querySheetsData(movesAndPo, queryRecord);
		querySheetsData(rmaBatch, queryRecord);
		querySheetsData(rmaReceipt, queryRecord);
		querySheetsData(reskuReport, queryRecord);
		querySheetsData(shipConfirm, queryRecord);
		
		finalizeMovesAndPoData(soxResearch, movesAndPo);
		finalizeRmaBatchData(soxResearch, rmaBatch);
		finalizeRmaReceiptData(soxResearch, rmaReceipt);
		finalizeReskuData( soxResearch, reskuReport);
		finalizeShipConfirmData(soxResearch, shipConfirm);
		
		writeSoxReport(movesAndPo, rmaBatch, rmaReceipt, reskuReport, shipConfirm, date);
		
		end = new Date();
		diff = start.getTime() - end.getTime();
		diffMinutes = diff / (60 * 1000) % 60;
		
		setStatusLabel("File generation done! Runtime: " + Math.abs(diffMinutes) + " minutes");
	}
	
	public String getOrdinalIndicator(int day) {
		String ordinalIndicator = "";
		
		if (day == 1) {
			ordinalIndicator = "st";
		} else if (day == 2) {
			ordinalIndicator = "nd";
		} else if (day == 3) {
			ordinalIndicator = "rd";
		} else {
			if (day >= 4 && day <= 20) {
				ordinalIndicator = "th";
			} else {
				switch(day % 10) {
					case 1:
						ordinalIndicator = "st";
						break;
					case 2:
						ordinalIndicator = "nd";
						break;
					case 3:
						ordinalIndicator = "rd";
						break;
					default:
						ordinalIndicator = "th";
						break;
				}
				
			}
		}
		
		return ordinalIndicator;
	}
	
	public void setStatusLabel(String status) {
		statusLbl.setText(status);
		statusLbl.paintImmediately(statusLbl.getVisibleRect());
	}
	
	
	public void readSoxResearch(Map<String, Map<String, String>> soxResearch) {
		setStatusLabel("Reading sox_research.xlsx");
		
		try {
			FileInputStream xlsxFile = new FileInputStream(new File(SOX_RES_FILE_DIR));
			Workbook workbook = new XSSFWorkbook(xlsxFile);
			
			for (int ctr = 0; ctr < SOX_RESEARCH_SHEETS.length; ctr++) {
				Sheet xlsxSheet = workbook.getSheet(SOX_RESEARCH_SHEETS[ctr]);
				soxResearch.put(SOX_RESEARCH_SHEETS[ctr], new HashMap<>());
			    Row row = null;
	
				for (int inCtr = 1; inCtr <= xlsxSheet.getLastRowNum(); inCtr++) {
					
					row = xlsxSheet.getRow(inCtr);
					soxResearch.get(SOX_RESEARCH_SHEETS[ctr]).put(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue());
						
				}
			}
		    
			workbook.close();
			xlsxFile.close();

		} catch (IOException e) {
			System.out.println("File error: " + e.getMessage());
			
		}
		
		setStatusLabel("Reading sox_research.xlsx. Done!");
	}
	
	public void readFiles(Object myObject, String soxFileDir) {
		
		try {
			Reader reader = Files.newBufferedReader(Paths.get(soxFileDir));
			CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT
					.withFirstRecordAsHeader()
					.withIgnoreHeaderCase()
					.withTrim()
			);
			
			if (myObject instanceof MovesAndPo) {
				setStatusLabel("Reading Moves and PO");
				((MovesAndPo) myObject).getHeader().addAll(csvParser.getHeaderNames());
				((MovesAndPo) myObject).setHeader("STATUS");
				((MovesAndPo) myObject).setHeader("COMMENT");
				((MovesAndPo) myObject).setHeader("SOX RESEARCH");
				
				for (CSVRecord csvRecord : csvParser) {
					((MovesAndPo) myObject).setTrxNo(csvRecord.get(0));		
					((MovesAndPo) myObject).setTrxType(csvRecord.get(1));
					((MovesAndPo) myObject).setSkuId(csvRecord.get(2));
					((MovesAndPo) myObject).setWarehouse(csvRecord.get(3));
					((MovesAndPo) myObject).setLocation(csvRecord.get(4));
					((MovesAndPo) myObject).setTrxQty(csvRecord.get(5));
					((MovesAndPo) myObject).setErrorCode(csvRecord.get(6));
					((MovesAndPo) myObject).setCreationDate(csvRecord.get(7));
					((MovesAndPo) myObject).setDay(csvRecord.get(8));
					((MovesAndPo) myObject).setHour(csvRecord.get(9));
					((MovesAndPo) myObject).setMinute(csvRecord.get(10));
				}
			}
			
			if (myObject instanceof RmaBatch) {
				setStatusLabel("Reading RMA Batch");
				((RmaBatch) myObject).getHeader().addAll(csvParser.getHeaderNames());
				((RmaBatch) myObject).setHeader("STATUS");
				((RmaBatch) myObject).setHeader("COMMENT");
				((RmaBatch) myObject).setHeader("SOX RESEARCH");
				
				for (CSVRecord csvRecord : csvParser) {
					((RmaBatch) myObject).setUpsRefNo(csvRecord.get(0));		
					((RmaBatch) myObject).setOrigOrdNo(csvRecord.get(1));
					((RmaBatch) myObject).setOrdLineNo(csvRecord.get(2));
					((RmaBatch) myObject).setItemId(csvRecord.get(3));
					((RmaBatch) myObject).setQty(csvRecord.get(4));
					((RmaBatch) myObject).setErrorCode(csvRecord.get(5));
					((RmaBatch) myObject).setCreationDate(csvRecord.get(6));
					((RmaBatch) myObject).setDay(csvRecord.get(7));
					((RmaBatch) myObject).setHour(csvRecord.get(8));
					((RmaBatch) myObject).setMinute(csvRecord.get(9));
				}
			}
			
			if (myObject instanceof RmaReceipt) {
				setStatusLabel("Reading RMA Receipts");
				((RmaReceipt) myObject).getHeader().addAll(csvParser.getHeaderNames());
				((RmaReceipt) myObject).setHeader("STATUS");
				((RmaReceipt) myObject).setHeader("COMMENT");
				((RmaReceipt) myObject).setHeader("SOX RESEARCH");
				
				for (CSVRecord csvRecord : csvParser) {
					((RmaReceipt) myObject).setTransactionNo(csvRecord.get(0));		
					((RmaReceipt) myObject).setSkuId(csvRecord.get(1));
					((RmaReceipt) myObject).setReferenceOne(csvRecord.get(2));
					((RmaReceipt) myObject).setReferenceTwo(csvRecord.get(3));
					((RmaReceipt) myObject).setInvAdjustmentQty(csvRecord.get(4));
					((RmaReceipt) myObject).setCreationDate(csvRecord.get(5));
					((RmaReceipt) myObject).setDay(csvRecord.get(6));
					((RmaReceipt) myObject).setHour(csvRecord.get(7));
					((RmaReceipt) myObject).setMinute(csvRecord.get(8));
				}
			}
			
			if (myObject instanceof Resku) {
				setStatusLabel("Reading RESKU");
				((Resku) myObject).getHeader().addAll(csvParser.getHeaderNames());
				((Resku) myObject).setHeader("STATUS");
				((Resku) myObject).setHeader("COMMENT");
				((Resku) myObject).setHeader("SOX RESEARCH");
				
				for (CSVRecord csvRecord : csvParser) {
					((Resku) myObject).setTrxNo(csvRecord.get(0));		
					((Resku) myObject).setSkuId(csvRecord.get(1));
					((Resku) myObject).setWarehouse(csvRecord.get(2));
					((Resku) myObject).setLocation(csvRecord.get(3));
					((Resku) myObject).setInvadjTypeQty(csvRecord.get(4));
					((Resku) myObject).setErrorCode(csvRecord.get(5));
					((Resku) myObject).setCreationDate(csvRecord.get(6));
					((Resku) myObject).setDay(csvRecord.get(7));
					((Resku) myObject).setHour(csvRecord.get(8));
					((Resku) myObject).setMinute(csvRecord.get(9));
				}
			}
			
			if (myObject instanceof ShipConfirm) {
				setStatusLabel("Reading Ship Confirm");
				((ShipConfirm) myObject).getHeader().addAll(csvParser.getHeaderNames());
				((ShipConfirm) myObject).setHeader("STATUS");
				((ShipConfirm) myObject).setHeader("COMMENT");
				((ShipConfirm) myObject).setHeader("SOX RESEARCH");
				
				for (CSVRecord csvRecord : csvParser) {
					((ShipConfirm) myObject).setControlNumber(csvRecord.get(0));
					((ShipConfirm) myObject).setShippedQty(csvRecord.get(1));
					((ShipConfirm) myObject).setBatchId(csvRecord.get(2));
					((ShipConfirm) myObject).setErrorCode(csvRecord.get(3));
					((ShipConfirm) myObject).setCreationDate(csvRecord.get(4));
					((ShipConfirm) myObject).setDay(csvRecord.get(5));
					((ShipConfirm) myObject).setHour(csvRecord.get(6));
					((ShipConfirm) myObject).setMinute(csvRecord.get(7));
				}
			}
			
			csvParser.close();
			reader.close();
			
		} catch (IOException e) {
			System.out.println("File error: " + e.getMessage());
		}
	}
	
	public void querySheetsData(Object myObject, Map<String, String[]> queryRecord) {
		setStatusLabel("Querying data");
		parameters = "";
		int rows = 0;
		
		if (myObject instanceof MovesAndPo) rows = ((MovesAndPo) myObject).getTrxNo().size();
		if (myObject instanceof RmaBatch) rows = ((RmaBatch) myObject).getUpsRefNo().size();
		if (myObject instanceof RmaReceipt) rows = ((RmaReceipt) myObject).getTransactionNo().size();
		if (myObject instanceof Resku) rows = ((Resku) myObject).getTrxNo().size();
		if (myObject instanceof ShipConfirm) rows = ((ShipConfirm) myObject).getControlNumber().size();
		
		
		for (int ctr = 0; ctr < rows; ctr++) {
			if (myObject instanceof MovesAndPo) parameters += "'" +  ((MovesAndPo) myObject).getTrxNo().get(ctr) + "',";
			if (myObject instanceof RmaBatch) parameters += "'" +  ((RmaBatch) myObject).getUpsRefNo().get(ctr) + "',";
			if (myObject instanceof RmaReceipt) parameters += "'" +  ((RmaReceipt) myObject).getTransactionNo().get(ctr) + "',";
			if (myObject instanceof Resku) parameters += "'" +  ((Resku) myObject).getTrxNo().get(ctr) + "',";
			if (myObject instanceof ShipConfirm) parameters += "'0" +  ((ShipConfirm) myObject).getControlNumber().get(ctr) + "',";
			
					
			if ((ctr % 998) == 0) {
				parameters = parameters.substring(0, parameters.length() - 1);
				if (myObject instanceof MovesAndPo) {
					setStatusLabel("Querying Moves and PO");
					model.executeMovesAndPoQuery();
				}
				
				if (myObject instanceof RmaBatch) { 
					setStatusLabel("Querying RMA Batch");
					model.executeRmaBatchQuery();
				}
				
				if (myObject instanceof RmaReceipt) {
					setStatusLabel("Querying RMA Receipts");
					model.executeRmaReceiptQuery();
				}
				
				if (myObject instanceof Resku) {
					setStatusLabel("Querying RESKU");
					model.executeReskuQuery();
				}
				
				if (myObject instanceof ShipConfirm) {
					setStatusLabel("Querying Ship Confirm");
					model.executeShipConfirmQuery();
				}
				
				parameters = "";
				
			} 
			
			else if ((ctr + 1) == rows) {
				parameters = parameters.substring(0, parameters.length() - 1);	
				if (myObject instanceof MovesAndPo) {
					setStatusLabel("Querying Moves and PO");
					model.executeMovesAndPoQuery();
				}
				
				if (myObject instanceof RmaBatch) { 
					setStatusLabel("Querying RMA Batch");
					model.executeRmaBatchQuery();
				}
				
				if (myObject instanceof RmaReceipt) {
					setStatusLabel("Querying RMA Receipts");
					model.executeRmaReceiptQuery();
				}
				
				if (myObject instanceof Resku) {
					setStatusLabel("Querying RESKU");
					model.executeReskuQuery();
				}
				
				if (myObject instanceof ShipConfirm) {
					setStatusLabel("Querying Ship Confirm");
					model.executeShipConfirmQuery();
				}
				
				parameters = "";
						
			}
			
		} 
		
		setStatusLabel("Querying done!");
	
	}
	
	public void finalizeMovesAndPoData(Map<String, Map<String, String>> soxResearch, MovesAndPo movesAndPo) {
		setStatusLabel("Finalizing Moves and PO data");
		
		String temp = "";
		String keyMpo = "";
		String errorCode = "";
		
		for (int ctr = 0; ctr < movesAndPo.getTrxNo().size(); ctr++) {
			keyMpo = movesAndPo.getTrxNo().get(ctr) + movesAndPo.getTrxType().get(ctr) + movesAndPo.getSkuId().get(ctr) + movesAndPo.getWarehouse().get(ctr) + movesAndPo.getLocation().get(ctr);
			errorCode = movesAndPo.getErrorCode().get(ctr);
			
			
			if (queryRecord.containsKey(keyMpo)) { // IF RAW EXCEL ID IS IN QUERY RESULT
				
				movesAndPo.setStatus(queryRecord.get(keyMpo)[0]); // status
				
				if (queryRecord.get(keyMpo)[0].trim().equalsIgnoreCase("success") || queryRecord.get(keyMpo)[0].trim().equalsIgnoreCase("success1")) { // GET STATUS, CHECK IF SUCCESS
					movesAndPo.setComment("");
					movesAndPo.getErrorCode().set(ctr, "");
					
				} else { // IF ANYTHING ELSE
					
					if (movesAndPo.getStatus().get(ctr).trim().isEmpty()) {
						movesAndPo.getStatus().set(ctr, "ERROR");
					
					} 
					
					if (!queryRecord.get(keyMpo)[1].trim().isEmpty()) { // IF NOT EMPTY COMMENT IN QUERY RESULT
						
						movesAndPo.setComment(queryRecord.get(keyMpo)[1]);
						
						if (movesAndPo.getErrorCode().get(ctr).trim().isEmpty()) {
							movesAndPo.getErrorCode().set(ctr, queryRecord.get(keyMpo)[1]);
						}
						
					} else { // IF NO COMMENT IN RESULT QUERY
						
						// IF NOT EMPTY ERROR CODE, COPY COMMENT
						if (!errorCode.isEmpty()) {

							movesAndPo.setComment(errorCode);
							
						} else {
							// IF EMPTY ERROR CODE
							movesAndPo.getErrorCode().set(ctr, "SPRN_API_FAILED");
							movesAndPo.setComment("SPRN_API_FAILED");
							
						}
						
					}
					
				}
				
			} 
			
			else {  // IF RAW EXCEL ID IS NOT IN QUERY RESULT
				movesAndPo.setStatus("ERROR");
				
				if (!errorCode.trim().isEmpty()) {
					movesAndPo.setComment(errorCode);
				} else {
					movesAndPo.getErrorCode().set(ctr, "SPRN_API_FAILED");
					movesAndPo.setComment("SPRN_API_FAILED");
				}
				
				 // INSERT ERROR
			}
			
			if (!movesAndPo.getComment().get(ctr).isEmpty()) { // IF NOT EMPTY COMMENT, PUT SOX RESEARCH
				temp = movesAndPo.getComment().get(ctr).trim().toLowerCase();
				
				if (temp.contains(",")) {
					for (String mpoComment : temp.split(",")) {
						if (!mpoComment.trim().isEmpty()) {
							temp = mpoComment;
							break;
						}
					}
				}
				
				
				// SET SOX RESEARCH
				for (String key : soxResearch.get("Moves and PO").keySet()) {
					// IF COMMENT IS AVAILABLE IN SOX RESEARCH
					
					if (temp.equalsIgnoreCase(key.trim().toLowerCase())) {
						movesAndPo.setSoxResearch(soxResearch.get("Moves and PO").get(key));
						break;
					}
					
				}
				
				try { // IF COMMENT NOT FOUND IN SOX RESEARCH, SET AS EMPTY
					
					movesAndPo.getSoxResearch().get(ctr);
					
				} catch (IndexOutOfBoundsException e) {
					
					movesAndPo.setSoxResearch("");
					
				}
				
			} else { // IF EMPTY COMMENT, SET EMPTY AS SOX RESEARCH
				
				movesAndPo.setSoxResearch("");
				
			}

			
		}

		setStatusLabel("Moves and PO finalization done!");
	}
	
	public void finalizeRmaBatchData(Map<String, Map<String, String>> soxResearch, RmaBatch rmaBatch) {
		setStatusLabel("Finalizing RMA Batch data");
		
		String temp = "";
		String upsRefNo = "";
		String origOrdNo = "";
		String itemID = "";
		String errorCode = "";
		
		for (int ctr = 0; ctr < rmaBatch.getUpsRefNo().size(); ctr++) {
			upsRefNo = rmaBatch.getUpsRefNo().get(ctr);
			origOrdNo = rmaBatch.getOrigOrdNo().get(ctr);
			itemID = rmaBatch.getItemId().get(ctr);
			errorCode = rmaBatch.getErrorCode().get(ctr);
			
			
			if (queryRecord.containsKey(upsRefNo + origOrdNo + itemID)) { // IF RAW EXCEL ID IS IN QUERY RESULT
				
				rmaBatch.setStatus(queryRecord.get(upsRefNo + origOrdNo + itemID)[0]); // status
				
				if (queryRecord.get(upsRefNo + origOrdNo + itemID)[0].trim().equalsIgnoreCase("success") || queryRecord.get(upsRefNo + origOrdNo + itemID)[0].trim().equalsIgnoreCase("success1")) { // GET STATUS, CHECK IF SUCCESS
					rmaBatch.setComment("");
					rmaBatch.getErrorCode().set(ctr, "");
					
				} else { // IF ANYTHING ELSE
					
					if (rmaBatch.getStatus().get(ctr).trim().isEmpty()) {
						rmaBatch.getStatus().set(ctr, "ERROR");
					
					} 
					
					if (!queryRecord.get(upsRefNo + origOrdNo + itemID)[1].trim().isEmpty()) { // IF NOT EMPTY COMMENT IN QUERY RESULT
						
						rmaBatch.setComment(queryRecord.get(upsRefNo + origOrdNo + itemID)[1]);
						
						if (rmaBatch.getErrorCode().get(ctr).trim().isEmpty()) {
							rmaBatch.getErrorCode().set(ctr, queryRecord.get(upsRefNo + origOrdNo + itemID)[1]);
						}
						
					} else { // IF NO COMMENT IN RESULT QUERY
						
						// IF NOT EMPTY ERROR CODE, COPY COMMENT
						if (!errorCode.isEmpty()) {

							rmaBatch.setComment(errorCode);
							
						} else {
							// IF EMPTY ERROR CODE
							rmaBatch.getErrorCode().set(ctr, "SPRN_API_FAILED");
							rmaBatch.setComment("SPRN_API_FAILED");
							
						}
						
					}
					
				}
				
			} 
			
			else {  // IF RAW EXCEL ID IS NOT IN QUERY RESULT
				rmaBatch.setStatus("ERROR"); // INSERT ERROR
				
				if (!errorCode.trim().isEmpty()) {
					rmaBatch.setComment(errorCode);
				} else {
					rmaBatch.getErrorCode().set(ctr, "SPRN_API_FAILED");
					rmaBatch.setComment("SPRN_API_FAILED");
				}
					
			}
			
			if (!rmaBatch.getComment().get(ctr).isEmpty()) { // IF NOT EMPTY COMMENT, PUT SOX RESEARCH
				temp = rmaBatch.getComment().get(ctr).trim().toLowerCase();
				
				if (temp.contains(",")) {
					for (String rmaComment : temp.split(",")) {
						if (!rmaComment.trim().isEmpty()) {
							temp = rmaComment;
							break;
						}
					}
				}
				
				
				// SET SOX RESEARCH
				for (String key : soxResearch.get("RMA Batch").keySet()) {
					// IF COMMENT IS AVAILABLE IN SOX RESEARCH
					
					if (temp.equalsIgnoreCase(key.trim().toLowerCase())) {
						rmaBatch.setSoxResearch(soxResearch.get("RMA Batch").get(key));
						break;
					}
					
				}
				
				try { // IF COMMENT NOT FOUND IN SOX RESEARCH, SET AS EMPTY
					
					rmaBatch.getSoxResearch().get(ctr);
					
				} catch (IndexOutOfBoundsException e) {
					
					rmaBatch.setSoxResearch("");
					
				}
				
			} else { // IF EMPTY COMMENT, SET EMPTY AS SOX RESEARCH
				
				rmaBatch.setSoxResearch("");
				
			}

			
		}

		setStatusLabel("RMA Batch finalization done!");
	}
	
	
	public void finalizeRmaReceiptData(Map<String, Map<String, String>> soxResearch, RmaReceipt rmaReceipt) {
		setStatusLabel("Finalizing RMA Receipts data");
		
		String temp = "";
		String txnNo = "";
		String skuID = "";
		String reference1 = "";
		
		for (int ctr = 0; ctr < rmaReceipt.getTransactionNo().size(); ctr++) {
			txnNo = rmaReceipt.getTransactionNo().get(ctr);
			skuID = rmaReceipt.getSkuId().get(ctr);
			reference1 = rmaReceipt.getReferenceOne().get(ctr);
			
			
			if (queryRecord.containsKey(txnNo + skuID + reference1)) { // IF RAW EXCEL ID IS IN QUERY RESULT
				
				rmaReceipt.setStatus(queryRecord.get(txnNo + skuID + reference1)[0]); // status
				
				if (queryRecord.get(txnNo + skuID + reference1)[0].trim().equalsIgnoreCase("success") || queryRecord.get(txnNo + skuID + reference1)[0].trim().equalsIgnoreCase("success1")) { // GET STATUS, CHECK IF SUCCESS
					rmaReceipt.setComment("");
					
				} else { // IF ANYTHING ELSE
					
					if (rmaReceipt.getStatus().get(ctr).trim().isEmpty()) {
						rmaReceipt.getStatus().set(ctr, "ERROR");
					
					} 
					
					if (!queryRecord.get(txnNo + skuID + reference1)[1].trim().isEmpty()) { // IF NOT EMPTY COMMENT IN QUERY RESULT
						
						rmaReceipt.setComment(queryRecord.get(txnNo + skuID + reference1)[1]);
						
					} else { // IF NO COMMENT IN RESULT QUERY
						
						rmaReceipt.setComment("SPRN_API_FAILED");
						
					}
					
				}
				
			} 
			
			else {  // IF RAW EXCEL ID IS NOT IN QUERY RESULT
				rmaReceipt.setStatus("ERROR"); // INSERT ERROR
				rmaReceipt.setComment("SPRN_API_FAILED");
					
			}
			
			if (!rmaReceipt.getComment().get(ctr).isEmpty()) { // IF NOT EMPTY COMMENT, PUT SOX RESEARCH
				temp = rmaReceipt.getComment().get(ctr).trim().toLowerCase();
				
				if (temp.contains(",")) {
					for (String rmaComment : temp.split(",")) {
						if (!rmaComment.trim().isEmpty()) {
							temp = rmaComment;
							break;
						}
					}
				}
				
				
				// SET SOX RESEARCH
				for (String key : soxResearch.get("RMA Receipt").keySet()) {
					// IF COMMENT IS AVAILABLE IN SOX RESEARCH
					
					if (temp.equalsIgnoreCase(key.trim().toLowerCase())) {
						rmaReceipt.setSoxResearch(soxResearch.get("RMA Receipt").get(key));
						break;
					}
					
				}
				
				try { // IF COMMENT NOT FOUND IN SOX RESEARCH, SET AS EMPTY
					
					rmaReceipt.getSoxResearch().get(ctr);
					
				} catch (IndexOutOfBoundsException e) {
					
					rmaReceipt.setSoxResearch("");
					
				}
				
			} else { // IF EMPTY COMMENT, SET EMPTY AS SOX RESEARCH
				
				rmaReceipt.setSoxResearch("");
				
			}

		}
		
		setStatusLabel("RMA Receipts finalization done!");
	}
	
		
	public void finalizeReskuData(Map<String, Map<String, String>> soxResearch, Resku reskuReport) {
		setStatusLabel("Finalizing Resku data");
		
		String trxNo = "";
		String skuID = "";
		String warehouse = "";
		String location = "";
		String errorCode = "";
		String temp = "";
		
		for (int ctr = 0; ctr < reskuReport.getTrxNo().size(); ctr++) {
			trxNo = reskuReport.getTrxNo().get(ctr);
			skuID = reskuReport.getSkuId().get(ctr);
			warehouse = reskuReport.getWarehouse().get(ctr);
			location = reskuReport.getLocation().get(ctr);
			errorCode = reskuReport.getErrorCode().get(ctr);
			
			if (queryRecord.containsKey(trxNo + skuID + warehouse + location)) { // IF RAW EXCEL ID IS IN QUERY RESULT
				reskuReport.setStatus(queryRecord.get(trxNo + skuID + warehouse + location)[0]); // status
				
				if (queryRecord.get(trxNo + skuID + warehouse + location)[0].trim().equalsIgnoreCase("success") || queryRecord.get(trxNo + skuID + warehouse + location)[0].trim().equalsIgnoreCase("success1")) { // GET STATUS, CHECK IF SUCCESS
					//IF SUCCESS REMOVE ERROR CODE
					reskuReport.getErrorCode().set(ctr, "");
					reskuReport.setComment("");
					
				} else { // IF ANYTHING ELSE
					
					if (queryRecord.get(trxNo + skuID + warehouse + location)[0].trim().isEmpty()) {
						// COPY STATUS
						reskuReport.getStatus().set(ctr, "ERROR");
					
					} 
					
					if (!queryRecord.get(trxNo + skuID + warehouse + location)[1].trim().isEmpty()) { // IF NOT EMPTY COMMENT IN QUERY RESULT
						
						reskuReport.setComment(queryRecord.get(trxNo + skuID + warehouse + location)[1]);
						
						if (reskuReport.getErrorCode().get(ctr).trim().isEmpty()) {
							reskuReport.getErrorCode().set(ctr, queryRecord.get(trxNo + skuID + warehouse + location)[1]);
						}
						
					} else { // IF NO COMMENT IN RESULT QUERY
						
						// IF NOT EMPTY ERROR CODE, COPY COMMENT
						if (!errorCode.isEmpty()) {

							reskuReport.setComment(errorCode);
							
						} else {
							// IF EMPTY ERROR CODE
							reskuReport.getErrorCode().set(ctr, "SPRN_API_FAILED");
							reskuReport.setComment("SPRN_API_FAILED");
							
						}
						
					}
					
				}
				
			} 
			
			else {  // IF RAW EXCEL ID IS NOT IN QUERY RESULT
				reskuReport.setStatus("ERROR"); // INSERT ERROR
				
				if (!errorCode.trim().isEmpty()) { // CHECK IF ERROR CODE IS NOT EMPTY
					// COPY COMMENT FROM ERROR CODE
					reskuReport.setComment(errorCode);
					
				}
				
				else { // IF ERROR CODE IS EMPTY
					
					reskuReport.getErrorCode().set(ctr, "SPRN_API_FAILED");
					reskuReport.setComment("SPRN_API_FAILED");
					
				}
				
			}
			
			if (!reskuReport.getComment().get(ctr).isEmpty()) { // IF NOT EMPTY COMMENT, PUT SOX RESEARCH
				temp = reskuReport.getComment().get(ctr).trim();
				
				if (temp.contains(",")) {
					for (String reskuComment : temp.split(",")) {
						if (!reskuComment.trim().isEmpty()) {
							temp = reskuComment;
							break;
						}
					}
				}
				
				// SET SOX RESEARCH
				for (String key : soxResearch.get("Resku").keySet()) {
					// IF COMMENT IS AVAILABLE IN SOX RESEARCH
									
					if (temp.equalsIgnoreCase(key.trim())) {
						reskuReport.setSoxResearch(soxResearch.get("Resku").get(key));
					}
					
				}
				
				try { // IF COMMENT NOT FOUND IN SOX RESEARCH, SET AS EMPTY
					
					reskuReport.getSoxResearch().get(ctr);
					
				} catch (IndexOutOfBoundsException e) {
					
					reskuReport.setSoxResearch("");
					
				}
				
			} else { // IF EMPTY COMMENT, SET EMPTY AS SOX RESEARCH
				
				reskuReport.setSoxResearch("");
				
			}

			
		}

		setStatusLabel("Resku finalization done!");
	}
	
	
	public void finalizeShipConfirmData(Map<String, Map<String, String>> soxResearch, ShipConfirm shipConfirm) {
		setStatusLabel("Finalizing Ship Confirm data");
		
		String controlNumber = "";
		String batchID = "";
		String errorCode = "";
		String temp = "";
		
		for (int ctr = 0; ctr < shipConfirm.getControlNumber().size(); ctr++) {
			controlNumber = shipConfirm.getControlNumber().get(ctr);
			batchID = shipConfirm.getBatchId().get(ctr);
			errorCode = shipConfirm.getErrorCode().get(ctr);
			
			if (queryRecord.containsKey(controlNumber + batchID)) { // IF RAW EXCEL ID IS IN QUERY RESULT
				
				shipConfirm.setStatus(queryRecord.get(controlNumber + batchID)[0]); // status
				
				if (queryRecord.get(controlNumber + batchID)[0].trim().equalsIgnoreCase("success") || queryRecord.get(controlNumber + batchID)[0].trim().equalsIgnoreCase("success1")) { // GET STATUS, CHECK IF SUCCESS
					//IF SUCCESS REMOVE ERROR CODE
					shipConfirm.getErrorCode().set(ctr, "");
					shipConfirm.setComment("");
					
				} else { // IF ANYTHING ELSE
					
					if (queryRecord.get(controlNumber + batchID)[0].trim().isEmpty()) {
						// COPY STATUS
						shipConfirm.getStatus().set(ctr, "ERROR");
					
					} 
					
					if (!queryRecord.get(controlNumber + batchID)[1].trim().isEmpty()) { // IF NOT EMPTY COMMENT IN QUERY RESULT
						
							shipConfirm.setComment(queryRecord.get(controlNumber + batchID)[1]);
							
							if (shipConfirm.getErrorCode().get(ctr).trim().isEmpty()) {
								shipConfirm.getErrorCode().set(ctr, queryRecord.get(controlNumber + batchID)[1]);
							}
						
					} else { // IF NO COMMENT IN RESULT QUERY
						
						// IF NOT EMPTY ERROR CODE, COPY COMMENT
						if (!errorCode.isEmpty()) {

							shipConfirm.setComment(errorCode);
							
						} else {
							// IF EMPTY ERROR CODE
							shipConfirm.getErrorCode().set(ctr, "SPRN_API_FAILED");
							shipConfirm.setComment("SPRN_API_FAILED");
							
						}
						
					}
					
				}
				
			} 
			
			else {  // IF RAW EXCEL ID IS NOT IN QUERY RESULT
				shipConfirm.setStatus("ERROR"); // INSERT ERROR
				
				if (!errorCode.trim().isEmpty()) { //CHECK IF ERROR CODE IS NOT EMPTY
					// COPY COMMENT FROM ERROR CODE
					shipConfirm.setComment(errorCode);
					
				}
				
				else { // IF ERROR CODE IS EMPTY
					
					shipConfirm.getErrorCode().set(ctr, "SPRN_API_FAILED");
					shipConfirm.setComment("SPRN_API_FAILED");
					
				}
				
			}
			
			if (!shipConfirm.getComment().get(ctr).isEmpty()) { // IF NOT EMPTY COMMENT, PUT SOX RESEARCH
				// SET SOX RESEARCH
				temp = shipConfirm.getComment().get(ctr).trim().toLowerCase();
				
				if (temp.contains(",")) {
					for (String shipComment : temp.split(",")) {
						if (!shipComment.trim().isEmpty()) {
							temp = shipComment;
							break;
						}
					}
				}
				
				
				for (String key : soxResearch.get("Ship Confirm").keySet()) {
					// IF COMMENT IS AVAILABLE IN SOX RESEARCH					
					if (temp.equalsIgnoreCase(key.trim().toLowerCase())) {
						shipConfirm.setSoxResearch(soxResearch.get("Ship Confirm").get(key));
						break;
					}
					
				}
				
				try { // IF COMMENT NOT FOUND IN SOX RESEARCH, SET AS EMPTY
					
					shipConfirm.getSoxResearch().get(ctr);
					
				} catch (IndexOutOfBoundsException e) {
					
					shipConfirm.setSoxResearch("");
					
				}
				
			} else { // IF EMPTY COMMENT, SET EMPTY AS SOX RESEARCH
				
				shipConfirm.setSoxResearch("");
				
			}
	
		}

		setStatusLabel("Ship Confirm finalization done!");
	}
	
	
	public void writeSoxReport(MovesAndPo movesAndPo, RmaBatch rmaBatch, RmaReceipt rmaReceipt, Resku reskuReport, ShipConfirm shipConfirm, String date) {
		setStatusLabel("Generating SOX Report");
		
		String soxFileOutput = System.getProperty("user.home") + "\\Documents\\sox_files\\SOX Report " + date + ".xlsx";

		SimpleDateFormat formatter = new SimpleDateFormat("dd-MMM-yyyy H:mm:ss", Locale.ENGLISH);
		SimpleDateFormat formatter2 = new SimpleDateFormat("M/dd/yyyy H:mm", Locale.ENGLISH);
		Workbook workbook = new SXSSFWorkbook(100);
		Sheet sheet = (SXSSFSheet) workbook.createSheet("Moves and PO");
		((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		
		Row row = sheet.createRow(0);
		Cell cell = null;
		Cell dateCell = null;
		Cell decimalCell= null;
		
		// SET HEADER STYLE
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// STYLE FOR BODY
		CellStyle bodyStyle = workbook.createCellStyle();
		
		XSSFFont headerFont = (XSSFFont) workbook.createFont();
		headerFont.setFontName("Calibri");
		headerFont.setBold(true);
		headerFont.setItalic(true);
		headerStyle.setFont(headerFont);
		
		XSSFFont bodyFont = (XSSFFont) workbook.createFont();
		bodyFont.setFontName("Calibri");
		bodyStyle.setFont(bodyFont);
		
		short dataStyle = workbook.createDataFormat().getFormat("dd-mm-yyyy hh:mm:ss");
		CellStyle dateStyle = workbook.createCellStyle();
		dateStyle.setFont(bodyFont);
		dateStyle.setDataFormat(dataStyle);
		
		CellStyle decimalStyle = workbook.createCellStyle();
		decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.0"));
		
		// MOVES AND PO
		for (int ctr = 0; ctr < movesAndPo.getHeader().size(); ctr++) {
			
			cell = row.createCell(ctr);
			cell.setCellValue(movesAndPo.getHeader().get(ctr));
			cell.setCellStyle(headerStyle);
			
		}
		
		for (int rowCtr = 0; rowCtr < movesAndPo.getTrxNo().size(); rowCtr++) {
			row = sheet.createRow(rowCtr + 1);
			
			cell = row.createCell(0);				
			cell.setCellValue(Long.parseLong(movesAndPo.getTrxNo().get(rowCtr)));
				
			cell = row.createCell(1);
			cell.setCellValue(movesAndPo.getTrxType().get(rowCtr));
			
			cell = row.createCell(2);
			try {
				Long.parseLong(movesAndPo.getSkuId().get(rowCtr));
				cell.setCellValue(Long.parseLong(movesAndPo.getSkuId().get(rowCtr)));
			} catch (NumberFormatException e) {
				cell.setCellValue(movesAndPo.getSkuId().get(rowCtr));
			}
			
			cell = row.createCell(3);
			cell.setCellValue(movesAndPo.getWarehouse().get(rowCtr));
			
			cell = row.createCell(4);
			cell.setCellValue(movesAndPo.getLocation().get(rowCtr));

			cell = row.createCell(5);
			cell.setCellValue(Integer.parseInt(movesAndPo.getTrxQty().get(rowCtr)));
			
			cell = row.createCell(6);
			cell.setCellValue(movesAndPo.getErrorCode().get(rowCtr));
			
			dateCell = row.createCell(7);
			try {
				dateCell.setCellValue(formatter.parse(movesAndPo.getCreationDate().get(rowCtr)));
				
			} catch (ParseException e) {
				try {
					dateCell.setCellValue(formatter2.parse(movesAndPo.getCreationDate().get(rowCtr)));
					
				} catch (ParseException e1) {
					System.out.println("Date parsing error: " + e1.getMessage());
					
				}
					
			}
			
			cell = row.createCell(8);
			cell.setCellValue(Integer.parseInt(movesAndPo.getDay().get(rowCtr)));
			
			cell = row.createCell(9);
			cell.setCellValue(Integer.parseInt(movesAndPo.getHour().get(rowCtr)));
			
			cell = row.createCell(10);
			cell.setCellValue(Integer.parseInt(movesAndPo.getMinute().get(rowCtr)));
			
			cell = row.createCell(11);
			cell.setCellValue(movesAndPo.getStatus().get(rowCtr));
			
			cell = row.createCell(12);
			cell.setCellValue(movesAndPo.getComment().get(rowCtr));
			
			cell = row.createCell(13);
			cell.setCellValue(movesAndPo.getSoxResearch().get(rowCtr));
				
			dateCell.setCellStyle(dateStyle);
			cell.setCellStyle(bodyStyle);
			
		}
		
		for (int ctr = 0; ctr < 6; ctr++) {
			sheet.autoSizeColumn(ctr);
		}
		sheet.setColumnWidth(6, 30 * 256);
		sheet.autoSizeColumn(7);
		sheet.autoSizeColumn(11);
		sheet.setColumnWidth(12, 30 * 256);
		sheet.autoSizeColumn(13);
		
		
		// RMA BATCH
		sheet = (SXSSFSheet) workbook.createSheet("RMA Batch Report");
		((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		
		row = sheet.createRow(0);
		
		
		for (int ctr = 0; ctr < rmaBatch.getHeader().size(); ctr++) {
			
			cell = row.createCell(ctr);
			cell.setCellValue(rmaBatch.getHeader().get(ctr));
			cell.setCellStyle(headerStyle);
			
		}
		
		for (int rowCtr = 0; rowCtr < rmaBatch.getUpsRefNo().size(); rowCtr++) {
			row = sheet.createRow(rowCtr + 1);
			
			cell = row.createCell(0);
			try {
				Long.parseLong(rmaBatch.getUpsRefNo().get(rowCtr));
				cell.setCellValue(Long.parseLong(rmaBatch.getUpsRefNo().get(rowCtr)));
				
			} catch (NumberFormatException e) {
				cell.setCellValue(rmaBatch.getUpsRefNo().get(rowCtr));
				
			}
				
			cell = row.createCell(1);
			try {
				cell.setCellValue(Long.parseLong(rmaBatch.getOrigOrdNo().get(rowCtr)));
				
			} catch (NumberFormatException e) {
				cell.setCellValue(rmaBatch.getOrigOrdNo().get(rowCtr));
			}
			
			if (rmaBatch.getOrdLineNo().get(rowCtr).contains(".")) {
				decimalCell = row.createCell(2);
				decimalCell.setCellValue(Float.parseFloat(rmaBatch.getOrdLineNo().get(rowCtr)));
				decimalCell.setCellStyle(decimalStyle);
			} else {
				cell = row.createCell(2);
				try {
					cell.setCellValue(Integer.parseInt(rmaBatch.getOrdLineNo().get(rowCtr)));
				} catch (NumberFormatException e) {
					cell.setCellValue(rmaBatch.getOrdLineNo().get(rowCtr));
				}
			}
			
			cell = row.createCell(3);
			try {
				Long.parseLong(rmaBatch.getItemId().get(rowCtr));
				cell.setCellValue(Long.parseLong(rmaBatch.getItemId().get(rowCtr)));
				
			} catch (NumberFormatException e) {
				cell.setCellValue(rmaBatch.getItemId().get(rowCtr));
				
			}
			
			cell = row.createCell(4);
			cell.setCellValue(Integer.parseInt(rmaBatch.getQty().get(rowCtr)));
			
			cell = row.createCell(5);
			cell.setCellValue(rmaBatch.getErrorCode().get(rowCtr));
			
			dateCell = row.createCell(6);
			try {
				dateCell.setCellValue(formatter.parse(rmaBatch.getCreationDate().get(rowCtr)));
				
			} catch (ParseException e) {
				try {
					dateCell.setCellValue(formatter2.parse(rmaBatch.getCreationDate().get(rowCtr)));
					
				} catch (ParseException e1) {
					System.out.println("Date parsing error: " + e1.getMessage());
					
				}
					
			}
			
			cell = row.createCell(7);
			cell.setCellValue(Integer.parseInt(rmaBatch.getDay().get(rowCtr)));
			
			cell = row.createCell(8);
			cell.setCellValue(Integer.parseInt(rmaBatch.getHour().get(rowCtr)));
			
			cell = row.createCell(9);
			cell.setCellValue(Integer.parseInt(rmaBatch.getMinute().get(rowCtr)));
			
			cell = row.createCell(10);
			cell.setCellValue(rmaBatch.getStatus().get(rowCtr));
			
			cell = row.createCell(11);
			cell.setCellValue(rmaBatch.getComment().get(rowCtr));
			
			cell = row.createCell(12);
			cell.setCellValue(rmaBatch.getSoxResearch().get(rowCtr));
				
			dateCell.setCellStyle(dateStyle);
			cell.setCellStyle(bodyStyle);
			
		}
		
		for (int ctr = 0; ctr < 4; ctr++) {
			sheet.autoSizeColumn(ctr);
		}
		sheet.setColumnWidth(5, 30 * 256);
		sheet.autoSizeColumn(6);
		sheet.autoSizeColumn(10);
		sheet.setColumnWidth(11, 30 * 256);
		sheet.autoSizeColumn(12);
		
		
		// RMA RECEIPT
		sheet = (SXSSFSheet) workbook.createSheet("RMA Receipts Report");
		((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		
		row = sheet.createRow(0);
		
		for (int ctr = 0; ctr < rmaReceipt.getHeader().size(); ctr++) {
			
			cell = row.createCell(ctr);
			cell.setCellValue(rmaReceipt.getHeader().get(ctr));
			cell.setCellStyle(headerStyle);
			
		}
		
		for (int rowCtr = 0; rowCtr < rmaReceipt.getTransactionNo().size(); rowCtr++) {
			row = sheet.createRow(rowCtr + 1);
		
			cell = row.createCell(0);
			cell.setCellValue(Long.parseLong(rmaReceipt.getTransactionNo().get(rowCtr)));
			
			cell = row.createCell(1);
			try {
				cell.setCellValue(Long.parseLong(rmaReceipt.getSkuId().get(rowCtr)));
				
			} catch (NumberFormatException e) {
				cell.setCellValue(rmaReceipt.getSkuId().get(rowCtr));
			}
			
			cell = row.createCell(2);
			cell.setCellValue(rmaReceipt.getReferenceOne().get(rowCtr));
			
			
			if (rmaReceipt.getReferenceTwo().get(rowCtr).contains(".")) {
				Float.parseFloat(rmaReceipt.getReferenceTwo().get(rowCtr));
				decimalCell = row.createCell(3);
				decimalCell.setCellValue(Float.parseFloat(rmaReceipt.getReferenceTwo().get(rowCtr)));
				decimalCell.setCellStyle(decimalStyle);

			} else {
				try {
					cell = row.createCell(3);
					cell.setCellValue(Integer.parseInt(rmaReceipt.getReferenceTwo().get(rowCtr)));
					
				} catch (NumberFormatException e) {
					cell = row.createCell(3);
					cell.setCellValue(rmaReceipt.getReferenceTwo().get(rowCtr));
					
				}
				
			}
			
			

			cell = row.createCell(4);
			cell.setCellValue(Integer.parseInt(rmaReceipt.getInvAdjustmentQty().get(rowCtr)));
			
			dateCell = row.createCell(5);
			try {
				dateCell.setCellValue(formatter.parse(rmaReceipt.getCreationDate().get(rowCtr)));
				
			} catch (ParseException e) {
				try {
					dateCell.setCellValue(formatter2.parse(rmaReceipt.getCreationDate().get(rowCtr)));
					
				} catch (ParseException e1) {
					System.out.println("Date parsing error: " + e1.getMessage());
					
				}
					
			}
			
			cell = row.createCell(6);
			cell.setCellValue(Integer.parseInt(rmaReceipt.getDay().get(rowCtr)));
			
			cell = row.createCell(7);
			cell.setCellValue(Integer.parseInt(rmaReceipt.getHour().get(rowCtr)));
			
			cell = row.createCell(8);
			cell.setCellValue(Integer.parseInt(rmaReceipt.getMinute().get(rowCtr)));
			
			cell = row.createCell(9);
			cell.setCellValue(rmaReceipt.getStatus().get(rowCtr));
			
			cell = row.createCell(10);
			cell.setCellValue(rmaReceipt.getComment().get(rowCtr));
			
			cell = row.createCell(11);
			cell.setCellValue(rmaReceipt.getSoxResearch().get(rowCtr));
				
			dateCell.setCellStyle(dateStyle);
			cell.setCellStyle(bodyStyle);
			
		}
		
		for (int ctr = 0; ctr < 6; ctr++) {
			sheet.autoSizeColumn(ctr);
		}
		sheet.autoSizeColumn(9);
		sheet.autoSizeColumn(11);
		sheet.setColumnWidth(10, 30 * 256);
		
		
		
		// RESKU
		sheet = (SXSSFSheet) workbook.createSheet("RESKU Report");
		((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		
		row = sheet.createRow(0);
		
		for (int ctr = 0; ctr < reskuReport.getHeader().size(); ctr++) {
			
			cell = row.createCell(ctr);
			cell.setCellValue(reskuReport.getHeader().get(ctr));
			cell.setCellStyle(headerStyle);
			
		}
		
		for (int rowCtr = 0; rowCtr < reskuReport.getTrxNo().size(); rowCtr++) {
			row = sheet.createRow(rowCtr + 1);
		
			cell = row.createCell(0);
			cell.setCellValue(Long.parseLong(reskuReport.getTrxNo().get(rowCtr)));
			
			cell = row.createCell(1);
			try {
				cell.setCellValue(Long.parseLong(reskuReport.getSkuId().get(rowCtr)));
				
			} catch (NumberFormatException e) {
				cell.setCellValue(reskuReport.getSkuId().get(rowCtr));
			}
			
			cell = row.createCell(2);
			cell.setCellValue(reskuReport.getWarehouse().get(rowCtr));
			
			cell = row.createCell(3);
			cell.setCellValue(reskuReport.getLocation().get(rowCtr));
			
			cell = row.createCell(4);
			cell.setCellValue(Integer.parseInt(reskuReport.getInvadjTypeQty().get(rowCtr)));
			
			cell = row.createCell(5);
			cell.setCellValue(reskuReport.getErrorCode().get(rowCtr));
			
			dateCell = row.createCell(6);
			try {
				dateCell.setCellValue(formatter.parse(reskuReport.getCreationDate().get(rowCtr)));
				
			} catch (ParseException e) {
				try {
					dateCell.setCellValue(formatter2.parse(reskuReport.getCreationDate().get(rowCtr)));
					
				} catch (ParseException e1) {
					System.out.println("Date parsing error: " + e1.getMessage());
					
				}
					
			}
			
			cell = row.createCell(7);
			cell.setCellValue(Integer.parseInt(reskuReport.getDay().get(rowCtr)));
			
			cell = row.createCell(8);
			cell.setCellValue(Integer.parseInt(reskuReport.getHour().get(rowCtr)));
			
			cell = row.createCell(9);
			cell.setCellValue(Integer.parseInt(reskuReport.getMinute().get(rowCtr)));
			
			cell = row.createCell(10);
			cell.setCellValue(reskuReport.getStatus().get(rowCtr));
			
			cell = row.createCell(11);
			cell.setCellValue(reskuReport.getComment().get(rowCtr));
			
			cell = row.createCell(12);
			cell.setCellValue(reskuReport.getSoxResearch().get(rowCtr));
				
			dateCell.setCellStyle(dateStyle);
			cell.setCellStyle(bodyStyle);
			
		}
		
		for (int ctr = 0; ctr < 5; ctr++) {
			sheet.autoSizeColumn(ctr);
		}

		sheet.setColumnWidth(5, 30 * 256);
		sheet.autoSizeColumn(6);
		sheet.autoSizeColumn(10);
		sheet.setColumnWidth(11, 30 * 256);
		sheet.autoSizeColumn(12);
		
		
		
		// SHIP CONFIRM
		sheet = (SXSSFSheet) workbook.createSheet("Ship Confirm");
		((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
		
		row = sheet.createRow(0);
		
		for (int ctr = 0; ctr < shipConfirm.getHeader().size(); ctr++) {
			
			cell = row.createCell(ctr);
			cell.setCellValue(shipConfirm.getHeader().get(ctr));
			cell.setCellStyle(headerStyle);
			
		}
		
		for (int rowCtr = 0; rowCtr < shipConfirm.getControlNumber().size(); rowCtr++) {
			row = sheet.createRow(rowCtr + 1);
		
			cell = row.createCell(0);
			cell.setCellValue(Long.valueOf(shipConfirm.getControlNumber().get(rowCtr)));
			
			cell = row.createCell(1);
			cell.setCellValue(Long.valueOf(shipConfirm.getShippedQty().get(rowCtr)));
			
			cell = row.createCell(2);
			cell.setCellValue(shipConfirm.getBatchId().get(rowCtr));
			
			cell = row.createCell(3);
			cell.setCellValue(shipConfirm.getErrorCode().get(rowCtr));
			
			dateCell = row.createCell(4);
			try {
				dateCell.setCellValue(formatter.parse(shipConfirm.getCreationDate().get(rowCtr)));
				
			} catch (ParseException e) {
				try {
					dateCell.setCellValue(formatter2.parse(shipConfirm.getCreationDate().get(rowCtr)));
					
				} catch (ParseException e1) {
					System.out.println("Date parsing error: " + e1.getMessage());
					
				}
					
			}
			
			cell = row.createCell(5);
			cell.setCellValue(Integer.parseInt(shipConfirm.getDay().get(rowCtr)));
			
			cell = row.createCell(6);
			cell.setCellValue(Integer.parseInt(shipConfirm.getHour().get(rowCtr)));
			
			cell = row.createCell(7);
			cell.setCellValue(Integer.parseInt(shipConfirm.getMinute().get(rowCtr)));
			
			cell = row.createCell(8);
			cell.setCellValue(shipConfirm.getStatus().get(rowCtr));
			
			cell = row.createCell(9);
			cell.setCellValue(shipConfirm.getComment().get(rowCtr));
			
			cell = row.createCell(10);
			cell.setCellValue(shipConfirm.getSoxResearch().get(rowCtr));
				
			dateCell.setCellStyle(dateStyle);
			cell.setCellStyle(bodyStyle);
			
		}
		
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		sheet.autoSizeColumn(4);
		sheet.autoSizeColumn(8);
		sheet.autoSizeColumn(10);
		
		sheet.setColumnWidth(3, 30 * 256);
		sheet.setColumnWidth(9, 30 * 256);
		

		// WRITE EXCEL
		try {
			FileOutputStream outputStream = new FileOutputStream(soxFileOutput);
			workbook.write(outputStream);
			workbook.close();

			//JOptionPane.showMessageDialog(null, "SOX Report generation done.");
			
		} catch (IOException e) {
			System.out.println("File error: " + e.getMessage());
			
		}
		
		
	}
	
	public void setQueryRecord(String key, String []value) {
		queryRecord.put(key, value);
	} 
	
	public String getParameters() {
		return parameters;
	}

}
