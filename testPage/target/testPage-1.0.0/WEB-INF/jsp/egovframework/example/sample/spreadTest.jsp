<%@ page contentType="text/html; charset=utf-8" pageEncoding="utf-8"%>
<%@ taglib prefix="c"         uri="http://java.sun.com/jsp/jstl/core" %>
<%@ taglib prefix="form"      uri="http://www.springframework.org/tags/form" %>
<%@ taglib prefix="validator" uri="http://www.springmodules.org/tags/commons-validator" %>
<%@ taglib prefix="spring"    uri="http://www.springframework.org/tags"%>
<%
  /**
  * @Class Name : egovSampleRegister.jsp
  * @Description : Sample Register 화면
  * @Modification Information
  *
  *   수정일         수정자                   수정내용
  *  -------    --------    ---------------------------
  *  2009.02.01            최초 생성
  *
  * author 실행환경 개발팀
  * since 2009.02.01
  *
  * Copyright (C) 2009 by MOPAS  All right reserved.
  */
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <c:set var="registerFlag" value="${empty sampleVO.id ? 'create' : 'modify'}"/>
    <title>Sample <c:if test="${registerFlag == 'create'}"><spring:message code="button.create" /></c:if>
                  <c:if test="${registerFlag == 'modify'}"><spring:message code="button.modify" /></c:if>
    </title>
    <link type="text/css" rel="stylesheet" href="<c:url value='/css/egovframework/sample.css'/>"/>
    
    <!-- spread JS 관련 시작 -->
    <link type="text/css" rel="stylesheet" href="<c:url value='/css/spread/gc.spread.sheets.excel2016colorful.12.2.5.css'/>"/>
    <script type="text/javascript" language="JavaScript" src="/js/spread/gc.spread.sheets.all.12.2.5.min.js"></script>
    <script type="text/javascript" language="JavaScript" src="/js/spread/gc.spread.sheets.resources.ko.12.2.5.min.js"></script>
    <script type="text/javascript" language="JavaScript" src="/js/spread/gc.spread.excelio.12.2.5.min.js"></script>
    <!-- spread JS 관련 종료 -->
    
    <!--For Commons Validator Client Side-->
    <script type="text/javascript" src="<c:url value='/cmmn/validator.do'/>"></script>
    <validator:javascript formName="sampleVO" staticJavascript="false" xhtml="true" cdata="false"/>
    
    <script type="text/javaScript" language="javascript" defer="defer">
   
        // 전역 변수 선언
        var excelIo = "";
        var sheet = "";
        var spread = "";
        
        var sheet2 = "";
        var spread2 = "";        
        var SheetArea = "";

        // 엑셀 업로드
        function fn_import(){
        	   /*
	           spread.suspendPaint();
	           spread.suspendCalcService();
	           spread.suspendEvent();
	           */
        	   var excelFile = document.getElementById("fileDemo").files[0];
               // here is excel IO API
               excelIo.open(excelFile, function (json) {
                   var workbookObj = json;
                   spread.fromJSON(workbookObj);
                   
                   sheet = spread.getActiveSheet();
                   //spread.options.scrollbarShowMax = false;  //스크롤 바를 전체 크기에 맞추기 않음
   	               //sheet.setRowCount(1048576); //row 및 column 갯수 설정
   	               //sheet.setColumnCount(1048576);
               }, function (e) {
                   // process error
                   alert(e.errorMessage);
               });
               /*
               spread.resumeEvent();
               spread.resumeCalcService();
               spread.resumePaint();
               */
        }         
        
        // 통계표 형태 전환
        function fn_transport(){
    		var jsonOptions = {
   				ignoreFormula: true,
   				ignoreStyle: false
    		};
    		
    		var serializationOption = {
   				ignoreFormula: true,
   				ignoreStyle: false
   			};
        	
    		//ToJson
    		var spread1 = GC.Spread.Sheets.findControl(document.getElementById('ss'));
    		var jsonStr = JSON.stringify(spread1.toJSON(serializationOption));

    		//FromJson
    		var spread2 = GC.Spread.Sheets.findControl(document.getElementById('ss1'));;
    		spread2.fromJSON(JSON.parse(jsonStr), jsonOptions);
    		
    		// 전환된 통계표 형태 
    		var style2 = new GC.Spread.Sheets.Style();
    		
    		var sheet2 = spread2.getActiveSheet();
    		/*
    		sheet2.deleteRows(0,7);
    		sheet2.deleteRows(sheet2.getRowCount(GC.Spread.Sheets.SheetArea.viewport)-1,1);
    		sheet2.deleteRows(sheet2.getRowCount(GC.Spread.Sheets.SheetArea.viewport)-1,1);
			
    		sheet2.suspendPaint();
			sheet2.options.gridline.showHorizontalGridline = true;
			sheet2.options.gridline.showVerticalGridline = true;
			sheet2.resumePaint();
			*/
    		// 보호 설정 및 셀 스타일 지정
    		//sheet2.options.isProtected = true;
    		//style2.locked = true;
        	//style2.backColor = 'lightGreen';
        	//sheet2.setStyle(11, 2, style2);
        	//sheet.setStyle(1, 13, style2);
        }
        
        //수치 데이터만 추출
        function fn_getData(){
        	var spreadNS = GC.Spread.Sheets;
    		var range = new GC.Spread.Sheets.Range(-1, 1, -1, 2);
    		var sheet2 = spread2.getActiveSheet();
    		
    		var rowFilter = new GC.Spread.Sheets.Filter.HideRowFilter(range);

    		sheet2.rowFilter(rowFilter);
    		rowFilter.filterButtonVisible(true);
        	
            var filter = sheet2.rowFilter();
            if (filter) {
                filter.removeFilterItems(1);
                //if (this.checked) {
                    var condition = new spreadNS.ConditionalFormatting.Condition(spreadNS.ConditionalFormatting.ConditionType.numberCondition, {
                        compareType: 3,
                        expected: 3,
                        formula : 3
                    });
                    filter.addFilterItem(1, condition);
                //}
                filter.filter(1);
                sheet2.invalidateLayout();
                sheet2.repaint();
            }
            
            //spread.getActiveSheet().getArray(11, 1, 2, 15);
            //getArray(시작row,시작col,row갯수,col갯수)
        }
        
        /*
        * spread JS 시작
        */
        window.onload = function() {
        	/*
            var spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'), {
                sheetCount: 1
            });
            */
            spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
            excelIo = new GC.Spread.Excel.IO();
            sheet = spread.getActiveSheet();
            
            spread2 = new GC.Spread.Sheets.Workbook(document.getElementById("ss1"));
            sheet2 = spread2.getActiveSheet();
            
			// 이벤트 테스트            
            var spreadNS = GC.Spread.Sheets;
            SheetArea = spreadNS.SheetArea;
            
            spread.bind(spreadNS.Events.CellClick, function (e, args) {
                var sheetArea = args.sheetArea === 0 ? 'sheetCorner' : args.sheetArea === 1 ? 'columnHeader' : args.sheetArea === 2 ? 'rowHeader' : 'viewPort';
                //alert('row: ' + args.row + 'col: ' + args.col );
            });
           
        };        
        
    </script>
</head>
<body style="text-align:center; margin:0 auto; display:inline; padding-top:100px;">

<form:form commandName="sampleVO" id="detailForm" name="detailForm">
    <div id="content_pop" style="width:100%">
    	<!-- 타이틀 -->
    	<div id="title" style="width:100%">
    		<ul>
    			<li style="width:100%"><img src="<c:url value='/images/egovframework/example/title_dot.gif'/>" alt=""/>통계표 수치입력 [통계표명 : 인구추이 / 통계표 ID : DT_21303_B000032]</li>
    		</ul>
    	</div>

	    <div id="sysbtn" style="margin-top:10px;margin-bottom:10px;float:left;">
	      <ul>
	   	      <li>
	   	           <input type="file" id="fileDemo" class="input" style="margin-left:5px;border:solid 1px"></input>
	   	           
	   	           <span class="btn_blue_l" style="float:right;margin-left:5px">
	   	              <a id="excelImport"  href="javascript:fn_import();">엑셀 업로드</a>
	   	              <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
	               </span>
             </li>
	      </ul>
	    </div>

	    <!--  excel 1 영역  -->
	    <div id="ss" style="width:100%; height:360px;border: 1px solid gray;"></div>
    
	    <div id="sysbtn" style="margin-top:85px; margin-bottom:10px;float:left">
	   	  <ul>
	   	      <li>
	   	          <span class="btn_blue_l">
	   	              <a id="excelImport"  href="javascript:fn_transport();">통계표 형태 전환</a>
	   	              <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
	                 </span>        	      
	             </li>

	   	         <li>
	   	          <span class="btn_blue_l">
	   	              <a id="excelImport"  href="javascript:fn_getData();">숫자만 추출</a>
	   	              <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
	                 </span>        	      
	             </li>	             
	         </ul>
	    </div>    
	    
	    <!--  excel 2 영역  -->
        <div id="ss1" style="width:100%; height:360px;border: 1px solid gray;margin-top:130px"></div>
    
  </div>
  
</form:form>
</body>
</html>