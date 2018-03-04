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
import net.amaware.autil.ADatabaseAccess;
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
    final static String propFileDbLOGS          = "dna-amaware.properties";
    final static String propFileDbTABLE_CODES   = "dna-amaware.properties";
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
	AFileExcelPOI aFileExcelPOI;   
	Sheet aSheetSummary;
	Sheet aSheetDetail;
	Sheet aSheetLog;
	Sheet aSheetLog2;
    //
	ADatabaseAccess thisADatabaseAccess;	
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
/*		
		try {
			aSheetSummary = aFileExcelPOI.doOutputStart(outExcelFileName, "Summary");
		} catch (IOException e) {
			throw new AException(acomm, e, "exportFileExcel");
		}
*/		
		aFileExcelPOI = new AFileExcelPOI(acomm, outExcelFileName);
   		//
		aSheetSummary = aFileExcelPOI.doCreateNewSheet("Summary",   2				
   				  //(new ArrayList<String>(getSourceHeadVector()))
   		          , (Arrays.asList(AComm.getArgFileName()
			                       , acomm.getCurrTimestampAny()
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
   		aSheetDetail =  aFileExcelPOI.doCreateNewSheet("Detail", 2
 				  ,(Arrays.asList(acomm.getCurrTimestampNew()))
				  , getSourceDataHeadList()
 				);
   		//
		aSheetLog = aFileExcelPOI.doCreateNewSheet("Log", 2
				  ,(Arrays.asList(acomm.getCurrTimestampNew()))
				  ,(Arrays.asList("SourceRow#"
				  	        , "Item"
				  	        , "Msg"
				            ) 
                  )
				);
		aSheetLog2 = aFileExcelPOI.doCreateNewSheet("Log2", 2
				  ,(Arrays.asList(acomm.getCurrTimestampNew()))
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
   				     , getDataRowColsValueList()
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
  				  	        , "#Cols{"+getDataRowColsValueList().size() +"}"
  					        ) 
  			        )
    		);
   	  		aFileExcelPOI.doOutputRowNext(acomm 
    			      , aSheetLog2
    				  , (Arrays.asList(
    						    ""+getSourceRowNum()
    						    ,"count"
    				  	        , "#Cols{"+getDataRowColsValueList().size() +"}"
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
		
		String tabName="table_codes";
		
        thisADatabaseAccess = new ADatabaseAccess(acomm, propFileDbTABLE_CODES);
        thisADatabaseAccess.doQueryRsExcel(aFileExcelPOI
                                              , tabName
                                              , "Select *"
                                          +" from " + tabName 
                            //+ " Where field_nme  = '" + ufieldname +"'" 
                                          
                            //+ " order by tab_name"
                            + " order by tab_name, code_name, code_value"
                                              );     
		
		//
        thisADatabaseAccess = new ADatabaseAccess(acomm, propFileDbLOGS);
        thisADatabaseAccess.doQueryRsExcel(aFileExcelPOI
                                              , "logs"
                                              , "Select * from logs " 
                                                + " order by entry_type, entry_subject, entry_topic"
                                              );     
        thisADatabaseAccess = new ADatabaseAccess(acomm, propFileDbLOGS);
        thisADatabaseAccess.doQueryRsExcel(aFileExcelPOI
                                          , "logs Again"
                                          , "Select *"
                                            +" from logs " 
            
                                            + " order by entry_type, entry_subject, entry_topic"
                );     
        
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
  		aFileExcelPOI.doOutputRowNext(acomm 
			      , aSheetLog2
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
