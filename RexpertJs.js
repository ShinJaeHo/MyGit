﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿
/*******************************************************************************
 * 기  능 : 렉스퍼트 화면 조회
 * param  : FormatType
 * return :
 * 수정자 : 백동원
 * 예  제 : ifn_RexPreView(Obj , pRptName,pDataType,pRptParams);
 * PARMETER 설명
 * Obj --> Rexpert Viewer 객체
 * pRptName --> 서버에 배포된 Rexpert 파일명 (.Rex 는 제외)
 * pDataType --> 데이터 타입 XML로 넘기면 됨
 * pRptParams --> Parameter (예 : A:2^B:가나다)
 ******************************************************************************/
function ifn_RexPreView(Obj, pRptDB, pRptName,pDataType,pRptParams)
{
	

		var strRptParam;
		var arrRptParam;
		var oReport;
		var oSubReport;
		var oConnection;
		var oDataSet;
		var oSQL;

		var pRptDB    = pRptDB;
		var pDataType = pDataType;
		var pRptName = pRptName;
		var pRptParams = pRptParams;

		//파라미터 단위 string 값을 배열로 변환
		strRptParam = pRptParams.split("^");
	
		//RexCtl.EnableHotKey("all=0");
		oReport = Obj.OpenReport("http://ekp.kolon.com/ReportWeb/rexfiles/KEC/CCK/" + pRptName +".reb");
		
		Obj.ShowParameterDialog = false;

		if(oReport == null)
		{
		    alert("can't open report file");
		    return;
		}
		
		if (strRptParam.length >= 1)
		{
		    for (i = 0; i < strRptParam.length; i++)
		    {	
				arrRptParam = strRptParam[i].split(":");
		        oReport.SetParameterFieldValue(arrRptParam[0], arrRptParam[1]);
		    }
		}
	    
		oSQL = oReport.GetSQLControl();
		
		if (pDataType=="XML" || pDataType=="")
		{
			oConnection = Obj.CreateConnection("http.post");
		}
		else if (pDataType=="CSV")
		{
			oConnection = Obj.CreateConnection("http.csv");
		}

		oConnection.AddParameter("sql", oSQL.GetSQL());
		
		for (k = 1; k <= oReport.GetReportCount(); k++)
		{
			var arrSubRptParam;
			
			oSubReport = oReport.OpenReport(k - 1);
			
			if (oSubReport == null)
			{
			   alert("can't open report file");
			   return;
			}

			if (strRptParam.length >= 1)
			{
				for (i = 0; i < strRptParam.length; i++)
				{	
					arrRptParam = strRptParam[i].split(":");
					oSubReport.SetParameterFieldValue(arrRptParam[0], arrRptParam[1]);
				}
			}
	            
			oSQL = oSubReport.GetSQLControl();
			oConnection.AddParameter("sql", oSQL.GetSQL());
		}
		
		oConnection.AddParameter("datatype", pDataType);
		oConnection.AddParameter("userservice", pRptDB);
		oConnection.Path = "http://ekp.kolon.com/ReportWeb/RexService.aspx";
		oConnection.Send();

		if ((pDataType=="XML" || pDataType==""))
		{
			oDataSet = oReport.CreateDataSetXML(oConnection, "root/main/rpt1/rexdataset/rexrow", 0);
		}
		else if (pDataType=="CSV")
		{
			oDataSet = oReport.CreateDataSetCSV(oConnection, 0, "|@|", "", "|*|", "");
		}
		    
		for (i = 1; i <= oReport.GetReportCount(); i++)
		{		
			oSubReport = oReport.OpenReport(i - 1);
			
			if (pDataType=="XML" || pDataType=="")
			{
				oDataSet = oSubReport.CreateDataSetXML(oConnection, "root/main/rpt" + (i + 1) + "/rexdataset/rexrow", 0);
			}
			else if (pDataType=="CSV")
			{
				oDataSet = oSubReport.CreateDataSetCSV(oConnection, 1, "|@|", "", "|*|", "");		
			}
			
		}
    
		Obj.Run();
}

function ifn_RexPreView30(Obj , pRptDB, pRptName,pDataType,pRptParams,pPrint)
{

		var pRptDB     = pRptDB;
		var pDataType  = pDataType;
		var pRptName   = pRptName;
		var pRptParams = pRptParams;
		var strRptParam;
		var arrRptParam;
		
		if(pPrint = null) pPrint = 0;

		//파라미터 단위 string 값을 배열로 변환
		strRptParam = pRptParams.split("^");
		
		var sOOF = "";
		sOOF += "<?xml version='1.0' encoding='euc-kr'?>";
		sOOF += "<oof version='3.0'>";
		
		if(pPrint == "0" ){// 일반뷰어
			sOOF += "<document titile='test'>";
		} else if(pPrint == "1" ){ //인쇄
			sOOF += "<document titile='test' enable-thread='0'>";
		}

		sOOF += "<connection-list>";

		sOOF += "<connection type='http' namespace='*'>";
		sOOF += "  <config-param-list>";
		sOOF += " 	   <config-param name='path'>http://ekp.kolon.com/ReportWeb/RexService.aspx</config-param>";
		sOOF += "	   <config-param name='method'>post</config-param>";
		sOOF += "</config-param-list>	";
		sOOF += "<http-param-list>";
		sOOF += "   <http-param name='presql'>{%dataset.ado.pre.sql%}</http-param>";
		sOOF += "   <http-param name='sql'>{%dataset.ado.sql%}</http-param>";
		sOOF += "   <http-param name='xmldata'>{%dataset.ado.sql.xml.prepared%}</http-param>";
		sOOF += "   <http-param name='postsql'>{%dataset.ado.post.sql%}</http-param>";
		sOOF += "   <http-param name='datatype'>" + pDataType + "</http-param>";

		if(pRptDB != "")
			sOOF += "   <http-param name='userservice'>" + pRptDB + "</http-param>";

		sOOF += "</http-param-list>";
		
		if(pDataType == "XML" || pDataType == "" ){
			sOOF += "<content content-type='xml'>";
			sOOF += " 	   <content-param name='root'>root/main/rpt1/rexdataset/rexrow</content-param>";
			sOOF += " 	   <content-param name='encoding'>utf-8</content-param>";
			sOOF += "</content>";
		} else if (pDataType=="CSV"){
			sOOF += "<content content-type='csv'>";
			//sOOF += " 	   <content-param name='row-delim'>|^|</content-param>";
			sOOF += " 	   <content-param name='col-delim'>|*|</content-param>";
			sOOF += " 	   <content-param name='encoding'>utf-8</content-param>";
			sOOF += "</content>";
		}

		sOOF += " </connection>";

		sOOF += "</connection-list>";
       	sOOF += "<field-list>";

		if (strRptParam.length >= 1)
		{
		    for (i = 0; i < strRptParam.length; i++)
		    {	
				arrRptParam = strRptParam[i].split(":");
				sOOF += "	<field name='" + arrRptParam[0] + "'>" + arrRptParam[1] + "</field>";
		    }
		}
		sOOF += "</field-list>";

		sOOF += "<file-list>";
		//sOOF += "   <file type='reb'  path='http://localhost:8080/rexpert30/rebfiles/" + pRptName + ".reb'/>";
		sOOF += "   <file type='reb'  path='http://ekp.kolon.com/ReportWeb/rexfiles/KEC/CCK/" + pRptName +".reb'/>";
		sOOF += "</file-list>";
		sOOF += "</document>";
		sOOF += "</oof>";

		Obj.OpenOOF(sOOF);//.Run();
		
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.sheetoption=2"); //-> 엑셀 저장 버튼->옵션 설정
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.pagesectiononceprint=1"); //->엑셀 저장 버튼-> 페이지 머리글/바닥글 한번만 출력 설정
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.autocellmerge=1"); //->엑셀 저장 버튼->자동 셀 머지
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.emptytextautocellmerge=1"); //->엑셀 저장 버튼->빈 텍스트 자동 셀 머지
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.errorrange=10"); //엑셀 저장 버튼->X,Y 좌표 오차 범위
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.zoomrate=100"); //엑셀 저장 버튼->확대/축소 비율
		Obj.SetCSS("appearance.toolbar.button.exportxls.option.95format=0"); //엑셀 저장 버튼->엑셀95포맷으로저장
		Obj.SetCSS("accessibility.enable=0");  //시각장애인용 리더 프로그램이 읽어주는  컨트롤을 표시
		Obj.UpdateCSS(); 

		
		if(pPrint == "1" ){
		RexCtl.Print(true, 1,-1,1,"");
		}
}

function ifn_Print_log(strBNS_SEQN_ID , strBASE_YYMM_ID , strBNS_ID , strRPT_CD , strType)
{
	var str_url = "";
 	var str_arg = "";
 	var str_outds = ""; 
 	var str_inds = "";
 	 	 	
 	str_url   = "server::COMMON/Public_Common.aspx";  
 	
 	str_inds  = "";
	str_arg  = "workgu=" + quote("Printlog"); 	
	str_arg  += " BNS_SEQN_ID=" + quote(strBNS_SEQN_ID); 
	str_arg  += " BASE_YYMM_ID=" + quote(strBASE_YYMM_ID); 
    str_arg  += " BNS_ID=" + quote(strBNS_ID);
	str_arg  += " IKENID=" + quote(gsIkenID);  		
	str_arg  += " RPT_CD=" + quote(strRPT_CD);  
	str_arg  += " TYPE=" + quote(strType); 
	 
	Transaction("Printlog", str_url, str_inds, str_outds, str_arg,  "fn_callback");
 	
}