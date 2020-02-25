<%@ page contentType="text/html; charset=utf-8" pageEncoding="utf-8"%>
<%@ taglib prefix="c"      uri="http://java.sun.com/jsp/jstl/core" %>
<%@ taglib prefix="form"   uri="http://www.springframework.org/tags/form" %>
<%@ taglib prefix="ui"     uri="http://egovframework.gov/ctl/ui"%>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags"%>
<%
  /**
  * @Class Name : egovSampleList.jsp
  * @Description : Sample List 화면
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
    <title><spring:message code="title.sample" /></title>
    <link type="text/css" rel="stylesheet" href="<c:url value='/css/egovframework/sample.css'/>"/>
    <script src="http://code.jquery.com/jquery-latest.min.js"></script>
    <link type="text/css" rel="stylesheet" media="all" href="http://kosis.kr/openapi/ext/style/subCommon.css" />
    <link type="text/css" rel="stylesheet" media="all"href="http://kosis.kr/openapi/devGuide/devGuide01/css/openTmp.css" />
    <script type="text/javascript" language="JavaScript" src="http://kosis.kr/openapi/devGuide/devGuide01/js/dojo.js" ></script>
    <script type="text/javascript" language="JavaScript" src="http://kosis.kr/openapi/devGuide/devGuide01/js/json.js"></script>
    <script type="text/javascript" language="JavaScript" src="http://kosis.kr/openapi/devGuide/devGuide01/js/ajax.js"></script>

    
    <script type="text/javaScript" language="javascript" defer="defer">
        <!--
        /* 글 수정 화면 function */
        function fn_egov_select(id) {
        	document.listForm.selectedId.value = id;
           	document.listForm.action = "<c:url value='/updateSampleView.do'/>";
           	document.listForm.submit();
        }
        
        /* 글 등록 화면 function */
        function fn_egov_addView() {
           	document.listForm.action = "<c:url value='/addSample.do'/>";
           	document.listForm.submit();
        }
        
        /* 글 목록 화면 function */
        function fn_egov_selectList() {
        	document.listForm.action = "<c:url value='/egovSampleList.do'/>";
           	document.listForm.submit();
        }
        
        /* pagination 페이지 링크 function */
        function fn_egov_link_page(pageNo){
        	document.listForm.pageIndex.value = pageNo;
        	document.listForm.action = "<c:url value='/egovSampleList.do'/>";
           	document.listForm.submit();
        }
        
        //-->
        
        
        function fn_kosis_api(){
        	
        	var paraObj = {
       			 // jsp cross domain . (devGuidePop.jsp 임의의 페이지를 호출함으로써 제약을 우회할 수 있다 소스는 소스 하단에 제공)
       			url : "http://localhost:8080/devGuidePop.jsp?method=getList&key=ZjZjOTI3MjRjNmU1YzdhZTMwOWRjNjgxN2MzNDgwNmY=&vwcd=MT_ZTITLE&parentId=" + "" + "&type=json",
      			sync : true,
       			load : function(resObj, a, b) { mapData = resObj; },
       			error : function ( resObj, e ) { alert(dojo.toJson(resObj)); }
            }
       			
       		sendPost( paraObj );

        	$.ajax({
        		type: "GET",
        		url : "http://kosis.kr/openapi/Param/statisticsParameterData.do?method=getList&apiKey=ODYwNDBlYmI1YzcxNmMyZTRiN2EyOTlhMThjMjk1Y2Y=&itmId=I_3+&objL1=DATA&objL2=&objL3=&objL4=&objL5=&objL6=&objL7=&objL8=&format=json&jsonVD=Y&prdSe=Y&newEstPrdCnt=3&prdInterval=1&loadGubun=1&orgId=206&tblId=DT_20603_B001003", 
        		async : true,
        		dataType: "xml",
        		success:function(object){
					alert("success");
        		},  
        		error : function(xhr){
        	        alert(xhr);
        	    }
        	});      
        	
        }    
        
        function fn_move_srpead(){
        	
        	document.listForm.action = "<c:url value='/spreadJsTest.do'/>";
           	document.listForm.submit();
        	
           	/*
        	$.ajax({
        		type: "POST",
        		url : "<c:url value='/spreadJsTest.do'/>",
        		data : "",
        		dataType: "json",
        		success:function(object){
        		},  
        		error : function(xhr){
        	        alert(xhr);
        	    }
        	});   
        	*/
        }

    </script>
    
</head>

<body style="text-align:center; margin:0 auto; display:inline; padding-top:100px;">
    <form:form commandName="searchVO" id="listForm" name="listForm" method="post">
        <input type="hidden" name="selectedId" />
        <div id="content_pop">
        	<!-- 타이틀 -->
        	<div id="title">
        		<ul>
        			<li><img src="<c:url value='/images/egovframework/example/title_dot.gif'/>" alt=""/><spring:message code="list.sample" /></li>
        		</ul>
        	</div>
        	<!-- // 타이틀 -->
        	<div id="search">
        		<ul>
        			<li>
        			    <label for="searchCondition" style="visibility:hidden;"><spring:message code="search.choose" /></label>
        				<form:select path="searchCondition" cssClass="use">
        					<form:option value="1" label="Name" />
        					<form:option value="0" label="ID" />
        				</form:select>
        			</li>
        			<li><label for="searchKeyword" style="visibility:hidden;display:none;"><spring:message code="search.keyword" /></label>
                        <form:input path="searchKeyword" cssClass="txt"/>
                    </li>
        			<li>
        	            <span class="btn_blue_l">
        	                <a href="javascript:fn_egov_selectList();"><spring:message code="button.search" /></a>
        	                <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
        	            </span>
        	        </li>
                </ul>
        	</div>
        	<!-- List -->
        	<div id="table">
        		<table width="100%" border="0" cellpadding="0" cellspacing="0" summary="카테고리ID, 케테고리명, 사용여부, Description, 등록자 표시하는 테이블">
        			<caption style="visibility:hidden">카테고리ID, 케테고리명, 사용여부, Description, 등록자 표시하는 테이블</caption>
        			<colgroup>
        				<col width="40"/>
        				<col width="100"/>
        				<col width="150"/>
        				<col width="80"/>
        				<col width="?"/>
        				<col width="60"/>
        			</colgroup>
        			<tr>
        				<th align="center">No</th>
        				<th align="center"><spring:message code="title.sample.id" /></th>
        				<th align="center"><spring:message code="title.sample.name" /></th>
        				<th align="center"><spring:message code="title.sample.useYn" /></th>
        				<th align="center"><spring:message code="title.sample.description" /></th>
        				<th align="center"><spring:message code="title.sample.regUser" /></th>
        			</tr>
        			<c:forEach var="result" items="${resultList}" varStatus="status">
            			<tr>
            				<td align="center" class="listtd"><c:out value="${paginationInfo.totalRecordCount+1 - ((searchVO.pageIndex-1) * searchVO.pageSize + status.count)}"/></td>
            				<td align="center" class="listtd"><a href="javascript:fn_egov_select('<c:out value="${result.id}"/>')"><c:out value="${result.id}"/></a></td>
            				<td align="left" class="listtd"><c:out value="${result.name}"/>&nbsp;</td>
            				<td align="center" class="listtd"><c:out value="${result.useYn}"/>&nbsp;</td>
            				<td align="center" class="listtd"><c:out value="${result.description}"/>&nbsp;</td>
            				<td align="center" class="listtd"><c:out value="${result.regUser}"/>&nbsp;</td>
            			</tr>
        			</c:forEach>
        		</table>
        	</div>
        	<!-- /List -->
        	<div id="paging">
        		<ui:pagination paginationInfo = "${paginationInfo}" type="image" jsFunction="fn_egov_link_page" />
        		<form:hidden path="pageIndex" />
        	</div>
        	<div id="sysbtn">
        	  <ul>
        	      <li>
        	          <span class="btn_blue_l">
        	         	  <a id="kosisBtn"  href="javascript:fn_move_srpead();">spread JS 테스트</a>
                          <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
                      </span>        	      
        	          <span class="btn_blue_l">
        	         	  <a id="kosisBtn"  href="javascript:fn_kosis_api();">KOSIS API 테스트</a>
                          <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
                      </span>
        	          <span class="btn_blue_l" style="margin-left:10px">
        	         	  <a href="javascript:fn_egov_addView();"><spring:message code="button.create" /></a>
                          <img src="<c:url value='/images/egovframework/example/btn_bg_r.gif'/>" style="margin-left:6px;" alt=""/>
                      </span>                      
                  </li>
              </ul>
        	</div>
        </div>
    </form:form>
</body>
</html>
