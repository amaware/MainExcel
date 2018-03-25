package net.amaware.apps.mainexcel;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;

//import net.amaware.serv.SourceServProperty;
import net.amaware.app.MainAppDataStore;
import net.amaware.autil.*;
//
/**
 * @author AMAWARE - Angelo M Adduci
 * 
 */

public class MainExcel {
	// set Properties file key names to being used
	//Properties file 
    final static String propFileName   = "MainFile.properties";
    //
    final static String marketsPropFileName   = "markets-amaware.properties";    
	//Architecture Common communication Class 
	static ACommDb acomm;
	//Architecture Framework Class
	static MainAppDataStore mainApp;
	//Application Classes
	//static ExcelProcess _fileProcess = new ExcelProcess();
	static AFileExcelPOI aFileExcelPOI = new AFileExcelPOI(); 
	Sheet aSheetRequest;	 
	Sheet aSheetResult;    	
    //	
	//
		public static void main(String[] args) {
			final String thisClassName = "MainExcel";
			//
			try { 
				acomm = new ACommDb(propFileName, args);
				
				ADatabaseAccess appADatabaseAccess = new ADatabaseAccess(acomm, marketsPropFileName
		                , "sym_fundamental", true, 250); //use this number to resolve timeout of update data_track
				

			    String outExcelFileName=AComm.getOutFileDirectory()+AComm.getFileDirSeperator()+thisClassName+AComm.getAppClassFileSep()
		        +AComm.getArgFileName()+".report.xls";

		         acomm.addPageMsgsLineOut(thisClassName+ "=>Output Excel File Name{" +outExcelFileName +"}");
				//
				
		        aFileExcelPOI = new AFileExcelPOI(acomm, outExcelFileName);
		       
		        
		        //				
				
		        
				appADatabaseAccess.doQueryRsExcel(aFileExcelPOI
						//thisDataTrackAccess.doQueryRsExcel(aFileExcelPOI
				                , "doQueryRsExcel data_track"
				                , "Select *"
				                  +" from data_track " 
				                 //+ " Where field_nme  = '" + ufieldname +"'" 
				                 //+ " order by tab_name"
				                 + " order by subject, topic, item"
				                 );
				        //
						//thisDataTrackAccess.doDbMetadataExcelSheet(aFileExcelPOI,"DataTrack MetaData");
						//
				appADatabaseAccess.doQueryRsExcel(aFileExcelPOI
						//thisDataTrackAccess.childDataTrackStoreAccess.doQueryRsExcel(aFileExcelPOI
				                , "doQueryRsExcel data_track_store"
				                , "Select *"
				                  +" from data_track_store " 
				                 //+ " Where source_nme  = '" + thisDataTrackAccess.getTrackFileName() +"'"
				                 //+ " order by tab_name"
				                 + " order by  run_start_ts desc, source_nme, source_mod_ts, data_track_id"
				                 );		        
		        
		        
		   		try {
					aFileExcelPOI.doOutputEnd(acomm);
				} catch (IOException e) {
					throw new AException(acomm, e, " Close of outFileExcel");
				}		        
		        
		        
				/*
				if (AComm.getArgFileName().toLowerCase().startsWith("maps")) {
					
					mainApp = new MainAppDataStore(acomm, new PMaps(), args, acomm.getFileTextDelimTab());
					mainApp.setSourceHeadRowStart(1);
					mainApp.setSourceDataHeadRowStart(2);
					//mainApp.setSourceDataHeadRowEnd(1);
					mainApp.setSourceDataRowStart(3);
					//mainApp.setSourceDataRowEnd(10);
					
				} else {
					mainApp = new MainAppDataStore(acomm, new ExcelProcess(), args, acomm.getFileTextDelimTab());
					mainApp.setSourceHeadRowStart(1);
					mainApp.setSourceDataHeadRowStart(3);
					//mainApp.setSourceDataHeadRowEnd(1);
					mainApp.setSourceDataRowStart(4);
					//mainApp.setSourceDataRowEnd(10);
				}
				
				mainApp.doProcess(acomm, thisClassName);
				*/
				//mainApp.getHtmlServ().outPageLine(acomm, thisClassName+" completed ");
				acomm.end();
				
			} catch (AExceptionSql e1) {
				
				acomm.addPageMsgsLineOut("MainExcel AExceptionSql msg{"+e1.getMessage()+e1.getExceptionMsg()+"}");
				
				throw e1;
				
			} catch (AException e1) {
				throw e1;
			}
			
		}
//
// END CLASS
//	
}
