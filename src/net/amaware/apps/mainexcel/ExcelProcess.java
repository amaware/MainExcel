/**
 * 
 */
package net.amaware.apps.mainexcel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Vector;

import org.apache.poi.ss.usermodel.CellStyle;

import net.amaware.app.DataStoreReport;
import net.amaware.autil.ACommDb;
import net.amaware.autil.ADataColResult;
import net.amaware.autil.AException;
import net.amaware.autil.AFileExcelPOI;
import net.amaware.autil.AFileO;
import net.amaware.autil.AProperties;

//import net.amaware.serv.DataStore;
import net.amaware.serv.HtmlTargetServ;
import net.amaware.serv.SourceProperty;
//import net.amaware.autil.ASqlStatements;
//import net.amaware.serv.SourceServProperty;



/**
 * @author PSDAA88 - Angelo M Adduci - Sep 6, 2005 3:02:12 PM
 *
 */

public class ExcelProcess extends DataStoreReport {
	final String thisClassName = this.getClass().getName();
	//
	//Field map
	ADataColResult fid = mapDataCol("id");
	ADataColResult ftab_name = mapDataCol("tab_name");
	ADataColResult fcode_name = mapDataCol("code_name");
	ADataColResult fcode_value = mapDataCol("code_value");
	ADataColResult fuser_mod_id = mapDataCol("user_mod_id");
 	ADataColResult fuser_mod_ts = mapDataCol("user_mod_ts");
//	ADataColResult fCol6 = mapDataCol("SixCol");
    //
	//
    protected String outFileNamePrefix = "";
	protected AFileO outXmlTxtFile = new AFileO();	
    protected String outExcelFileName = "";
	//
	AFileExcelPOI aFileExcelPOI = new AFileExcelPOI();    
    //
	/**
	 * 
	 */
	class TabTwo {
        String  tab_name ="";
        String  code_name ="";
        String  code_value ="";
        
        TabTwo(String itn, String icn, String icv) {
        	tab_name=itn;
        	code_name=icn;
        	code_value=itn;
        }
    }	
	TabTwo aTabTwo; 
	List<TabTwo> aTabTwoList = new ArrayList<TabTwo> ();
	/**
	 * 
	 */
	public ExcelProcess() {
		super();

	}
	
	

	public DataStoreReport processThis(ACommDb acomm
			, SourceProperty _aProperty, HtmlTargetServ _aHtmlServ) {
		super.processThis(acomm, _aProperty, _aHtmlServ); // always call this first
		
		getThisHtmlServ().outPageLine(acomm,  thisClassName+"=>processThis");
	
		_aProperty.displayProperties(acomm);
		
		if (!outXmlTxtFile.isFileOpen()) {

			outFileNamePrefix = acomm.getOutFileDirectoryWithSep()+acomm.getArgFileName().replace(".xls", ".out");
			
    		outExcelFileName = outFileNamePrefix+".xls";
	   		//
			_aHtmlServ.outTargetLine(acomm,
					"outExcelFile Opened Name=" + outExcelFileName);
            //
		}
		return this;
	}
	
	

	/*
	 * 
	 * 
	 */
	
	public boolean  doDataHead(ACommDb acomm, int rowNum) throws AException {
		super.doDataHead(acomm, rowNum);
		
		try {
			aFileExcelPOI.doOutputStart(outExcelFileName, "From Input Excel");
		} catch (IOException e) {
			throw new AException(acomm, e, "exportFileExcel");
		}
		aFileExcelPOI.getaWorkBookSheet().createFreezePane(0,2);
   		aFileExcelPOI.doOutputHeader(acomm,
   				  (new ArrayList<String>(getSourceHeadVector()))
   				  , getSourceDataHeadList()
   				  );
		
		return true;

	}	
	
    
	public boolean doDataRowsNotFound(ACommDb acomm)
	throws AException {
		super.doDataRowsNotFound(acomm);
		
		//getThisHtmlServ().outPageLine(acomm,  "DataRowsNotFound");
		
		return true; 

	}

	/*
	 * 
	 * 
	 */
    
    
	public boolean doDataRow(ACommDb acomm, AException _exceptionSql, boolean _isRowBreak)
	throws AException {
		
		super.doDataRow(acomm, _exceptionSql, _isRowBreak); // sends pout row
		
		
		//getThisHtmlServ().outPageLine(acomm,  "DataRowFound");

		
		int _currRowNum = getSourceRowNum();
	    acomm.addPageMsgsLineOut(thisClassName+"=>Row#{" + _currRowNum +  "}");

   		aFileExcelPOI.doOutputRowNext(acomm, 
			       //(Arrays.asList("colRed", "Blue", "Green") )
   				getDataRowColsToList()
			     );
   		//aTabTwoList.add(new TabTwo(ftab_name.getColumnValue(), ftab_name.getColumnValue(),ftab_name.getColumnValue()));
   		aTabTwoList.add(new TabTwo(ftab_name.getColumnValue(), fcode_name.getColumnValue(),fcode_value.getColumnValue()));
		
		return true; //or false to stop processing of file

	}

	/**
	 * @return 
	 */


	public boolean doDataRowsEnded(ACommDb acomm)
	throws AException {
		
		super.doDataRowsEnded(acomm);
        //
		aFileExcelPOI.doCreateNewSheet("New");
		aFileExcelPOI.getaWorkBookSheet().createFreezePane(0,2);
		
   		aFileExcelPOI.doOutputHeader(acomm,
 				  (Arrays.asList(acomm.getDbURL()
 					  	        , "table_name"
 						        , "options"
 						        , acomm.getCurrTimestampOld()
 						        ) 
 				  )
				  ,(Arrays.asList(ftab_name.getColumnTitle()
				  	        , fcode_name.getColumnTitle()
					        , fcode_value.getColumnTitle()
					        , fuser_mod_id.getColumnTitle()
					        , fuser_mod_ts.getColumnTitle()
				            ) 
                    )
 				);
		
   		
		 for (TabTwo aTabTwoList : aTabTwoList) {
	   		 aFileExcelPOI.doOutputRowNext(acomm, 
				       //(Arrays.asList("colRed", "Blue", "Green") )
					  (Arrays.asList(aTabTwoList.tab_name
					  	        , aTabTwoList.code_name
					  	        //, aTabTwoList.code_name
						        , aTabTwoList.code_value
						        //, acomm.getDbUserID()
						        , this.getClass().getName()
						        , acomm.getCurrTimestampOld()
						        ) 
				     ));
	     }
		//
   		try {
			aFileExcelPOI.doOutputEnd();
		} catch (IOException e) {
			throw new AException(acomm, e, " Close of outFileExcel");
		}

		
		return true; //or false to stop processing of file

	}

	

//
//END
//	
}
