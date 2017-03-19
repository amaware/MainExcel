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
import org.apache.poi.ss.usermodel.Sheet;

import net.amaware.app.DataStoreReport;
import net.amaware.autil.AComm;
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
 	ADataColResult fother = mapDataCol("other");
//	ADataColResult fCol6 = mapDataCol("SixCol");
    //
	//
    protected String outFileNamePrefix = "";
	protected AFileO outXmlTxtFile = new AFileO();	
    protected String outExcelFileName = "";
	//
	AFileExcelPOI aFileExcelPOI = new AFileExcelPOI();   
	Sheet aSheetSummary;
	Sheet aSheetDetail;
	Sheet aSheetLog;
    //
	/**
	 * 
	 */
	class SummaryData {
        String  tab_name ="";
        String  code_name ="";
        String  code_value ="";
        
        SummaryData(String itn, String icn, String icv) {
        	tab_name=itn;
        	code_name=icn;
        	code_value=icv;
        }
    }	
	SummaryData aTabTwo; 
	List<SummaryData> aTabTwoList = new ArrayList<SummaryData> ();
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

			//outFileNamePrefix = acomm.getOutFileDirectoryWithSep()+acomm.getArgFileName().replace(".xls", ".out");
			outFileNamePrefix = acomm.getOutFileDirectoryWithClassName()+AComm.getArgFileName();
			
			
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
	 */
	
	public boolean  doDataHead(ACommDb acomm, int rowNum) throws AException {
		super.doDataHead(acomm, rowNum);
		
		try {
			aSheetSummary = aFileExcelPOI.doOutputStart(outExcelFileName, "Summary");
		} catch (IOException e) {
			throw new AException(acomm, e, "exportFileExcel");
		}
		
		aSheetSummary.createFreezePane(0,2);
   		aFileExcelPOI.doOutputHeader(acomm,
   				  //(new ArrayList<String>(getSourceHeadVector()))
   		          (new ArrayList<String>(
			         Arrays.asList(aSheetSummary.getSheetName()
			                       , AComm.getArgFileName()
			                       , acomm.getCurrTimestampAny()
			        		      )
			         )
   				  )
				  ,(Arrays.asList(ftab_name.getColumnTitle()
				  	        //, fcode_name.getColumnTitle()
					        //, fcode_value.getColumnTitle()
					        , "Count Tab"
					        //, "Count Code"
					        //, "Count Value"
				            ) 
                  )

   			  );
		
   		//
   		aSheetDetail = aFileExcelPOI.doCreateNewSheet("Detail");
   		aSheetDetail.createFreezePane(0,2);
   		aFileExcelPOI.doOutputHeader(acomm,
 				  (Arrays.asList(aSheetDetail.getSheetName()))
				  , getSourceDataHeadList()
 				);
   		//
		aSheetLog = aFileExcelPOI.doCreateNewSheet("Log");
		aSheetLog.createFreezePane(0,2);
   		aFileExcelPOI.doOutputHeader(acomm,
				  (Arrays.asList(aSheetLog.getSheetName()))
				  ,(Arrays.asList("SourceRow#"
				  	        , "Item"
				  	        , "Msg"
				            ) 
                  )
				);
   		//
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

   		aFileExcelPOI.doOutputRowNext(acomm 
   			         , aSheetDetail
   				     , getDataRowColsToList()
			     );
	    
   		//aTabTwoList.add(new TabTwo(ftab_name.getColumnValue(), ftab_name.getColumnValue(),ftab_name.getColumnValue()));
   		aTabTwoList.add(new SummaryData(ftab_name.getColumnValue(), fcode_name.getColumnValue(),fcode_value.getColumnValue()));
   		//
   		
   		if (_currRowNum == getSourceDataRowStartNum()) {
   	  		aFileExcelPOI.doOutputRowNext(acomm 
  			      , aSheetLog
  				  , (Arrays.asList(
  						    ""+getSourceRowNum()
  						    ,"count"
  				  	        , "#Cols{"+getDataRowColsToList().size() +"}"
  					        ) 
  			        )
    		);
   			
   		}
		
		return true; //or false to stop processing of file

	}

	/**
	 * @return 
	 */


	public boolean doDataRowsEnded(ACommDb acomm)
	throws AException {
        //		
		super.doDataRowsEnded(acomm);
        //
		String tab_namePrev="";
		String code_namePrev="";
		String code_valuePrev="";
		
		int tab_nameCnt=0;
		int code_nameCnt=0;
		int code_valueCnt=0;
		
		int itemcnt=0;
		for (SummaryData aTabTwoList : aTabTwoList) {
			
			++itemcnt;
			
			if (itemcnt > 1) { 
				if (!aTabTwoList.tab_name.contentEquals(tab_namePrev)) {
			   		 aFileExcelPOI.doOutputRowNext(acomm 
						      , aSheetSummary
							  , (Arrays.asList(tab_namePrev
							  	       // , code_namePrev
								       // , code_valuePrev
								        , ""+tab_nameCnt
								       // , ""+code_nameCnt
								      //  , ""+code_valueCnt
								        ) 
						        )
							  );
				}
				/*
				if (!aTabTwoList.code_name.contentEquals(code_namePrev)) {
			   		 aFileExcelPOI.doOutputRowNext(acomm 
						      , aSheetSummary
							  , (Arrays.asList(tab_namePrev
							  	        , code_namePrev
								        , code_valuePrev
								        , ""+tab_nameCnt
								        , ""+code_nameCnt
								        , ""+code_valueCnt
								        ) 
						        )
							  );
				}
				if (!aTabTwoList.code_value.contentEquals(code_valuePrev)) {
			   		 aFileExcelPOI.doOutputRowNext(acomm 
						      , aSheetSummary
							  , (Arrays.asList(tab_namePrev
							  	        , code_namePrev
								        , code_valuePrev
								        , ""+tab_nameCnt
								        , ""+code_nameCnt
								        , ""+code_valueCnt
								        ) 
						        )
							  );
				}
				*/
				
			}
            //			
			if (aTabTwoList.tab_name.contentEquals(tab_namePrev)) {
				++tab_nameCnt;
			}
			if (aTabTwoList.code_name.contentEquals(code_namePrev)) {
				++code_nameCnt;
			}
			if (aTabTwoList.code_value.contentEquals(code_valuePrev)) {
				++code_valueCnt;
			}
	   		 

	 		tab_namePrev=aTabTwoList.tab_name;
			code_namePrev=aTabTwoList.code_name;
			code_valuePrev=aTabTwoList.code_value;
			
			
	    }
		//
  		aFileExcelPOI.doOutputRowNext(acomm 
			      , aSheetLog
				  , (Arrays.asList(""+getSourceRowNum()
						    , "At End"
						    , "#SummaryRows{"+aSheetSummary.getLastRowNum() +"}"
						      + " |#DetailRows{"+aSheetDetail.getLastRowNum() +"}"
					        ) 
			        )
		);
		
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
