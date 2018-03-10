/**
 * 
 */
package net.amaware.apps.mainexcel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
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

public class PMaps extends DataStoreReport {
	final String thisClassName = this.getClass().getName();
	//
	//Field map
	ADataColResult flevel = mapDataCol("level");
	ADataColResult fname = mapDataCol("name");
	ADataColResult foccurs = mapDataCol("occurs");
	ADataColResult fusage = mapDataCol("usage");
 	ADataColResult fother = mapDataCol("other");
//	ADataColResult fCol6 = mapDataCol("SixCol");
    //
	//
    protected String outFileNamePrefix = "";
	protected AFileO outXmlTxtFile = new AFileO();	
    protected String outExcelFileName = "";
	//
    AFileExcelPOI aFileExcelPOI = new AFileExcelPOI();   
	Sheet aSheetRequest;
	Sheet aSheetLinkedHashMap;
	Sheet aSheetLog;
    //
	/**
	 * 
	 */
	class FieldData	{
		
	   	
       int level      = 0;
       String name    = "";
       String occurs  = "";
       String usage  = "";
       
       List<Integer> occurGroups = new ArrayList<Integer>();
       
	   FieldData () {
	   }
	   
	}
	List<FieldData> aFieldDataList = new ArrayList<FieldData>();
	
	
	/**
	 * 
	 */
	
	
    //
	//This class extends HashMap and maintains a linked list of the entries in the map
	//, in the order in which they were inserted.
	// If you assign two different values to a key the second value will simply 
	//    replace the first value assigned to that key.
	Map<Integer, String> aLinkedHashMap = new LinkedHashMap<Integer, String>();
	/**
	 * 
	 */
	public PMaps() {
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
			aSheetRequest = aFileExcelPOI.doOutputStart(outExcelFileName, "Request");
		} catch (IOException e) {
			throw new AException(acomm, e, "exportFileExcel");
		}
		
		aSheetRequest.createFreezePane(0,2);
   		aFileExcelPOI.doOutputHeader(acomm,
   				  //(new ArrayList<String>(getSourceHeadVector()))
   		          (new ArrayList<String>(
			         Arrays.asList(aSheetRequest.getSheetName()
			                       , AComm.getArgFileName()
			                       , acomm.getCurrTimestampAny()
			        		      )
			         )
   				  )
				  , getSourceDataHeadList()

   			  );
		
   		//
   		aSheetLinkedHashMap = aFileExcelPOI.doCreateNewSheet("LinkedHashMap");
   		aSheetLinkedHashMap.createFreezePane(0,2);
   		aFileExcelPOI.doOutputHeader(acomm,
 				  (Arrays.asList(aSheetLinkedHashMap.getSheetName()))
				  , outLinkedHashMapHeader()
				  
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
	
	
	public List<String> outLinkedHashMapHeader() {
		List<String> outlidt = new ArrayList<String>(); 
		outlidt.addAll(getSourceDataHeadList()); 
		outlidt.addAll((Arrays.asList("=>Map"
				  	        , "Key"
				  	        , "Value"
				            ) 
                 )
				
				); 

		
		return outlidt;
		
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

   		aFileExcelPOI.doOutputRowNext(acomm, aSheetRequest, outInputList());
   		
   		aFileExcelPOI.doOutputRowNext(acomm, aSheetLog, (Arrays.asList(" ")));
   		
   		FieldData aFieldData = new FieldData();
   		
   		doLinkedHashMap(acomm, aFieldData);
   		
   		aFieldDataList.add(aFieldData);
   		
   		outMapList(acomm);
   		
   		//
		
		return true; //or false to stop processing of file

	}

	public List<String> doLinkedHashMap(ACommDb acomm, FieldData aFieldData) {
		
		aFieldData.level=0;
		aFieldData.name="";
		aFieldData.occurs="";
		aFieldData.usage="";
		
		List<String> output = new ArrayList<String>();
		
   		int inLevelInt = doFieldValidateInt(acomm, flevel.getColumnValue(),0);
   		//int inOccursInt = doFieldValidateInt(acomm, foccurs.getColumnValue(),0);
		
   		List<Integer> removeitemList = new ArrayList<Integer>();
   		if (!foccurs.getColumnValue().trim().isEmpty()) {
   			
   			for (Integer keylevel : aLinkedHashMap.keySet()) { //remove all pevious items with Levle <= this one
   				if (keylevel >= inLevelInt) {
   					removeitemList.add(keylevel);
   			   	  	aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
	   					      , (Arrays.asList(""+getSourceRowNum()
	   							    , "doDataRow:Loop To Remove inlevel-level{" + inLevelInt + "}"
	   							      + "{" + keylevel + "}"
	   							    //, "#aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
	   							      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
	   						        ) 
	   				        ));   				
   					
   				} else {
   		   	  		aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
     					      , (Arrays.asList(""+getSourceRowNum()
     							    , "doDataRow:Loop To keep/put inlevel-level{" + inLevelInt + "}"
     							      + "{" + keylevel + "}"
   								      + " #aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
   							    , " "

     							      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
     						        ) 
     				        ));
   				}
   				
   	     	}
   			
   		} else { //no occurs...need to re
   			for (Integer keylevel : aLinkedHashMap.keySet()) { //remove all pevious items with Levle <= this one
   				if (keylevel >= inLevelInt) {
   					removeitemList.add(keylevel);
   			   	  	aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
	   					      , (Arrays.asList(""+getSourceRowNum()
	   							    , "doDataRow:Loop To Remove inlevel-level{" + inLevelInt + "}"
	   							      + "{" + keylevel + "}"
	   							    //, "#aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
	   							      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
	   						        ) 
	   				        ));   				
   				} 
   			}	
   		}
		for (Integer keylevel : removeitemList) { 
   				aLinkedHashMap.remove(keylevel);
		   	  	aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
	   					      , (Arrays.asList(""+getSourceRowNum()
	   							    , "doDataRow:Remove inlevel-level{" + inLevelInt + "}"
	   							      + "{" + keylevel + "}"
   								      + " #aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
   							    , " "

	   							      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
	   						        ) 
	   				        ));   				
		}
   		

   		if (!foccurs.getColumnValue().trim().isEmpty()) {			
   			aLinkedHashMap.put(Integer.valueOf(flevel.getColumnValue()), foccurs.getColumnValue());
   	  		aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
				      , (Arrays.asList(""+getSourceRowNum()
						    , "doDataRow:Put level-occurs{" + flevel.getColumnValue() + "}"
						    	+ "{" + foccurs.getColumnValue() +  "}"
								      + " #aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
							    , " "

						      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
					        ) 
			        ));   		
   		}
   		
   		if (foccurs.getColumnValue().trim().isEmpty()) { //no occurs
   			int occursValCnt=0;
   			for (Integer keylevel : aLinkedHashMap.keySet()) {
   				 occursValCnt += doFieldValidateInt(acomm, aLinkedHashMap.get(keylevel),0);
   				 
 	   	  	    aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
 					      , (Arrays.asList(""+getSourceRowNum()
 							    , "doDataRow: Item occured{" + inLevelInt + ""
 							      + " " + fname.getColumnValue() +"-"+ occursValCnt + ""
 							      + " " + fusage.getColumnValue() 
								      + " #aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
							    , " "
 							      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
 						        ) 
 				        ));   				
   				 
   			}
   			
   			if (occursValCnt==0) { //field is not occursed
   			
	   	  	    aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
				      , (Arrays.asList(""+getSourceRowNum()
					    , "doDataRow:Item not occured{" + inLevelInt + ""
							      + " " + fname.getColumnValue()
							      + " " + fusage.getColumnValue() 
							      + "}"
								      + " #aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
							    , " "
						      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
					        ) 
			        ));   				
   		     }
   			
   		} else { //occurs and pic?
   			
   			if (fusage!=null || !fusage.getColumnValue().trim().isEmpty()) {
   				int occursValCnt=0;
   	   			for (Integer keylevel : aLinkedHashMap.keySet()) {
   	   			    occursValCnt += doFieldValidateInt(acomm, aLinkedHashMap.get(keylevel),0);
   		   	  	    aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
   					      , (Arrays.asList(""+getSourceRowNum()
   						    , "doDataRow:Item Occurs With Usage{" + inLevelInt + ""
   								      + " " + fname.getColumnValue() +"-"+ occursValCnt + ""
   								      + " " + fusage.getColumnValue() 
   								      + "}"
   								      + " #aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
   							    , " "
   							      
   						        ) 
   				        ));   				
   	   				
   	   			}
   				
   			}
   		}
        //        
		aFieldData.level=0;
		aFieldData.name=fname.getColumnValue();
		aFieldData.occurs=foccurs.getColumnValue();
		aFieldData.usage=fusage.getColumnValue();
   		
        for (Integer keylevel : aLinkedHashMap.keySet()) {
	   		   int occursValCnt = doFieldValidateInt(acomm, aLinkedHashMap.get(keylevel),0);
		   	   aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
	   				      , (Arrays.asList(""+getSourceRowNum()
	   					    , "element"
	   		 			    , "" + inLevelInt + ""
	   						     + " " + fname.getColumnValue() +"-"+ occursValCnt + ""
	   						     + " " + fusage.getColumnValue() 
	   						     + ""
	   						) 
	   				        ));  
		   	   
		   	aFieldData.occurGroups.add(occursValCnt);
		   	   
	   	}
        
		return output;
		
	}
	
	
	
	public List<String> outInputList() {
		
		List<String> output = new ArrayList<String>();
		
		StringBuffer sb = new StringBuffer();
		
		int levint =  Integer.valueOf(flevel.getColumnValue());
		for (int i=0;i<levint;i++) {
			sb.append(".");
		}
		sb.append(flevel.getColumnValue());
		
		output.add(sb.toString());
		output.add(fname.getColumnValue());
		output.add(foccurs.getColumnValue());
		output.add(fusage.getColumnValue());
		
		return output;
		
	}
	
	
	public List<String> outMapList(ACommDb acomm) {
		//aFileExcelPOI.doOutputRowNext(acomm, aSheetResult, outMapList());
		List<String> output = new ArrayList<String>();
		aFileExcelPOI.doOutputRowNext(acomm, aSheetLinkedHashMap, (Arrays.asList(" ")));
		
		
		aFileExcelPOI.doOutputRowNext(acomm, aSheetLinkedHashMap, outInputList());
		
		for (Integer key : aLinkedHashMap.keySet()) {
			//System.out.println(key + ":\t" + aLinkedHashMap.get(key));
			
			aFileExcelPOI.doOutputRowNext(acomm, aSheetLinkedHashMap
					                     , (Arrays.asList(" ", " ", " ", " ", " "
					   					                 , ""+key
					  						             , aLinkedHashMap.get(key)
					  						             )
 				                           ));
     	}
		
		return output;
		
	}
	
	/**
	 * @return 
	 */


	public boolean doDataRowsEnded(ACommDb acomm)
	throws AException {
        //		
		super.doDataRowsEnded(acomm);
        //
		aFileExcelPOI.doOutputRowNext(acomm, aSheetRequest, (Arrays.asList(" ")));
		aFileExcelPOI.doOutputRowNext(acomm, aSheetRequest, (Arrays.asList("Result Fields ")));
		
		for (FieldData thisFieldData : aFieldDataList) {
			if (thisFieldData.occurGroups.size() > 0) {
				int occuritem=0;
				for (int thisOccurs : thisFieldData.occurGroups) {
					++occuritem;
					String outnamesuff="";
					for (int i=0;i<thisOccurs;i++) {
						outnamesuff="-"+occuritem+"-"+(i+1);
				        aFileExcelPOI.doOutputRowNext(acomm, aSheetRequest
						      , (Arrays.asList(" "
								    , thisFieldData.name+outnamesuff
								    , thisFieldData.occurs
								    , thisFieldData.usage
							        ) 
					             ));
					}
					
				}
				
			} else {
			
			    aFileExcelPOI.doOutputRowNext(acomm, aSheetRequest
				      , (Arrays.asList(" "
						    , thisFieldData.name
						    , thisFieldData.occurs
						    , thisFieldData.usage
					        ) 
			             ));
			}
			
		}
		//
  		aFileExcelPOI.doOutputRowNext(acomm, aSheetLog
				      , (Arrays.asList(""+getSourceRowNum()
						    , "At End"
						    , "#aLinkedHashMapItems{"+aLinkedHashMap.size() +"}"
						      //+ " |#DetailRows{"+aSheetLinkedHashMap.getLastRowNum() +"}"
					        ) 
			             ));
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
