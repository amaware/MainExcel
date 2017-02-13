package net.amaware.apps.mainexcel;
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
	//Architecture Common communication Class 
	static ACommDb acomm;
	//Architecture Framework Class
	static MainAppDataStore mainApp;
	//Application Classes
	static ExcelProcess _fileProcess = new ExcelProcess();
    //	
	//
        //
		public static void main(String[] args) {
			final String thisClassName = "MainExcel";
			//
			try { 
				acomm = new ACommDb(propFileName);
				
				mainApp = new MainAppDataStore(acomm, _fileProcess, args, acomm.getFileTextDelimTab());
				mainApp.setSourceHeadRowStart(1);
				mainApp.setSourceDataHeadRowStart(3);
				//mainApp.setSourceDataHeadRowEnd(1);
				mainApp.setSourceDataRowStart(4);
				//mainApp.setSourceDataRowEnd(10);
				
				
				mainApp.doProcess(acomm, thisClassName);
				
				//mainApp.getHtmlServ().outPageLine(acomm, thisClassName+" completed ");
				acomm.end();
				
			} catch (AException e1) {
				throw e1;
			}
			
		}
//
// END CLASS
//	
}
