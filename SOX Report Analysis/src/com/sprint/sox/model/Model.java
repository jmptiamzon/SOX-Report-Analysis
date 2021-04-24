package com.sprint.sox.model;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import com.sprint.sox.controller.Controller;

public class Model {
	private Controller controller;
	
	public Model(Controller controller) {
		this.controller = controller;
	}
	
	public void executeMovesAndPoQuery() {
		String trxNo = "";
		String trxType = "";
		String skuID = "";
		String warehouse = "";
		String location = "";
		String status = "";
		String errorCode = "";
		String parameters = controller.getParameters();
		
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection conn = 
					DriverManager.getConnection("", "", "");
			
			PreparedStatement statement = conn.prepareStatement("SELECT * FROM (SELECT DISTINCT A.TRANSACTION_NUMBER, A.TRANSACTION_TYPE, A.SKUID, A.WAREHOUSE, A.LOCATION, A.STATUS, A.ERROR_CODE, B.ERROR_CODE||'-'||B.ERROR_MESSAGE COMMENTS, ROW_NUMBER() OVER(PARTITION BY A.TRANSACTION_NUMBER, A.TRANSACTION_TYPE, A.SKUID, A.WAREHOUSE, A.LOCATION ORDER BY A.LAST_UPDATE_DATE desc)row1 FROM  APPS.SPRN_IM_WM_TRX_IN_INT A, SPRN.SPRN_COMMON_ERROR_TXN B WHERE A.RECORD_ID = B.RECORD_ID (+) AND A.TRANSACTION_NUMBER IN (" + parameters + ")) WHERE ROW1 = 1");
			
			ResultSet rs = statement.executeQuery();
			
			while (rs.next()) {
				trxNo = rs.getString(1);
				trxType = rs.getString(2) == null ? "" : rs.getString(2);
				skuID = rs.getString(3) == null ? "" : rs.getString(3);
				warehouse = rs.getString(4) == null ? "" : rs.getString(4);
				location = rs.getString(5) == null ? "" : rs.getString(5);
				status = rs.getString(6) == null ? "" : rs.getString(6);
				errorCode = (rs.getString(7) == null || rs.getString(7).trim().equals("-")) ? "" : rs.getString(7);

				controller.setQueryRecord(trxNo + trxType + skuID + warehouse + location, new String[] {status, errorCode});
				 //Control Number
				
			}
			
			conn.close();
			
		} catch (ClassNotFoundException e) {
			System.out.println("Driver not found: " + e.getMessage());
			
		} catch (SQLException e) {
			System.out.println("SQL error: " + e.getMessage());
			
		}
		
	}
	
	public void executeRmaBatchQuery() {
		String upsRefNo = "";
		String origOrdNo = "";
		String itemID = "";
		String status = "";
		String errorCode = "";
		String parameters = controller.getParameters();
		
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection conn = 
					DriverManager.getConnection("", "", "");
			
			PreparedStatement statement = conn.prepareStatement("SELECT * FROM (SELECT DISTINCT A.UPS_REF_NO, A.ORIG_ORD_NUMBER, A.ITEM_ID, A.STATUS, A.ERROR_CODE, A.LAST_UPDATE_DATE, ROW_NUMBER() OVER(PARTITION BY A.UPS_REF_NO, A.ORIG_ORD_NUMBER, A.ITEM_ID ORDER BY A.LAST_UPDATE_DATE desc)row1 FROM SPRN.SPRN_RMA_BATCH_IN_INT A WHERE 1=1 AND A.UPS_REF_NO IN(" + parameters + ")) WHERE ROW1 = 1");
			
			ResultSet rs = statement.executeQuery();
			
			while (rs.next()) {
				upsRefNo = rs.getString(1);
				origOrdNo = rs.getString(2) == null ? "" : rs.getString(2);
				itemID = rs.getString(3) == null ? "" : rs.getString(3);
				status = rs.getString(4) == null ? "" : rs.getString(4);
				errorCode = rs.getString(5) == null ? "" : rs.getString(5);
				
				controller.setQueryRecord(upsRefNo + origOrdNo + itemID, new String[] {status, errorCode});
				 //Control Number
				
			}
			
			conn.close();
			
		} catch (ClassNotFoundException e) {
			System.out.println("Driver not found: " + e.getMessage());
			
		} catch (SQLException e) {
			System.out.println("SQL error: " + e.getMessage());
			
		}
		
	}
	
	public void executeRmaReceiptQuery() {
		String txnNo = "";
		String skuID = "";
		String reference1 = "";
		String status = "";
		String errorCode = "";
		String parameters = controller.getParameters();
		
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection conn = 
					DriverManager.getConnection("", "", "");
			
			PreparedStatement statement = conn.prepareStatement("SELECT * FROM (SELECT DISTINCT a.TRANSACTION_NUMBER, a.skuid, a.reference1, a.STATUS, a.error_code, b.error_code||'-'|| b.error_message \"COMMENTS\", a.last_update_date, ROW_NUMBER() OVER(PARTITION BY A.transaction_number, a.skuid, a.reference1 ORDER BY A.LAST_UPDATE_DATE desc)row1 FROM apps.sprn_rcv_trans_in_int a, SPRN.SPRN_COMMON_ERROR_TXN b WHERE a.record_id = b.record_id(+) and a.TRANSACTION_NUMBER IN(" + parameters + ")) WHERE ROW1 = 1");
			
			ResultSet rs = statement.executeQuery();
			
			while (rs.next()) {
				txnNo = String.valueOf(rs.getLong(1));
				skuID = rs.getString(2) == null ? "" : rs.getString(2);
				reference1 = rs.getString(3) == null ? "" : rs.getString(3);
				status = rs.getString(4) == null ? "" : rs.getString(4);
				errorCode = (rs.getString(5) == null || rs.getString(5).trim().equals("-")) ? "" : rs.getString(5);

				controller.setQueryRecord(txnNo + skuID + reference1, new String[] {status, errorCode});
				 //Control Number
				
			}
			
			conn.close();
			
		} catch (ClassNotFoundException e) {
			System.out.println("Driver not found: " + e.getMessage());
			
		} catch (SQLException e) {
			System.out.println("SQL error: " + e.getMessage());
			
		}
		
	}
	
	public void executeReskuQuery() {
		String txnNo = "";
		String skuID = "";
		String warehouse = "";
		String location = "";
		String status = "";
		String errorCode = "";
		String parameters = controller.getParameters();
		
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection conn = 
					DriverManager.getConnection("", "", "");
			
			PreparedStatement statement = conn.prepareStatement("SELECT * FROM (SELECT DISTINCT A.TRANSACTION_NUMBER, A.SKU_ID, A.WAREHOUSE, A.LOCATION, A.STATUS, B.ERROR_CODE||'-'||B.ERROR_MESSAGE, a.last_update_date, ROW_NUMBER() OVER(PARTITION BY A.TRANSACTION_NUMBER, A.SKU_ID, A.WAREHOUSE, A.LOCATION ORDER BY A.LAST_UPDATE_DATE desc)row1 FROM APPS.SPRN_WM_TRX_RESKU_TXN_IN_INT A, SPRN.SPRN_COMMON_ERROR_TXN B  WHERE A.RECORD_ID = B.RECORD_ID (+) AND A.TRANSACTION_NUMBER IN(" + parameters + ")) WHERE ROW1 = 1");
			
			ResultSet rs = statement.executeQuery();
			
			while (rs.next()) {
				txnNo = String.valueOf(rs.getLong(1));
				skuID = rs.getString(2) == null ? "" : rs.getString(2);
				warehouse = rs.getString(3) == null ? "" : rs.getString(3);
				location = rs.getString(4) == null ? "" : rs.getString(4);
				status = rs.getString(5) == null ? "" : rs.getString(5);
				errorCode = (rs.getString(6) == null || rs.getString(6).trim().equals("-")) ? "" : rs.getString(6);
				
				controller.setQueryRecord(txnNo + skuID + warehouse + location, new String[] {status, errorCode});
				 //Control Number
			}
			
			conn.close();
			
		} catch (ClassNotFoundException e) {
			System.out.println("Driver not found: " + e.getMessage());
			
		} catch (SQLException e) {
			System.out.println("SQL error: " + e.getMessage());
			
		}
		
	}
	
	public void executeShipConfirmQuery() {
		String batchID = "";
		String errorCode = "";
		String status = "";
		String parameters = controller.getParameters();
		
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			Connection conn = 
					DriverManager.getConnection("", "", "");
			
			PreparedStatement statement = conn.prepareStatement("SELECT * FROM (SELECT DISTINCT A.PICKTICKET_CONTROL_NUMBER, A.BATCH_ID, A.STATUS, A.ERROR_CODE, B.ERROR_CODE || '-' || B.ERROR_MESSAGE \"COMMENTS\" , A.LAST_UPDATE_DATE, ROW_NUMBER() OVER(PARTITION BY A.PICKTICKET_CONTROL_NUMBER, A.BATCH_ID ORDER BY A.LAST_UPDATE_DATE desc)row1 FROM APPS.SPRN_OM_SHIP_HDR_IN_INT A, SPRN.SPRN_COMMON_ERROR_TXN B WHERE A.RECORD_ID = B.RECORD_ID (+) AND A.PICKTICKET_CONTROL_NUMBER IN(" + parameters + ")) WHERE ROW1 = 1");
			
			ResultSet rs = statement.executeQuery();
			
			while (rs.next()) {
				batchID = rs.getString(2) == null ? "" : rs.getString(2);
				status = rs.getString(3) == null ? "" : rs.getString(3);
				errorCode = (rs.getString(4) == null || rs.getString(4).trim().equals("-")) ? "" : rs.getString(4);
						
				controller.setQueryRecord(String.valueOf(rs.getInt(1)) + batchID, new String[] {status, errorCode});
				 //Control Number
			}
			
			conn.close();
			
		} catch (ClassNotFoundException e) {
			System.out.println("Driver not found: " + e.getMessage());
			
		} catch (SQLException e) {
			System.out.println("SQL error: " + e.getMessage());
			
		}
		
	}
	
	
}
