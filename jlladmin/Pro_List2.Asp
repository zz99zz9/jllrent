<!--#include file="inc/xgheader.asp"-->
      <!--sidebar start-->
<!--#include file="inc/xgmenu.asp"-->
      <!--sidebar end-->
      <!--main content start-->
      <script language="JavaScript" type="text/javascript">  
        //第一种方法  
        function method1(tableid) {  
  
            var curTbl = document.getElementById(tableid);  
            var oXL = new ActiveXObject("Excel.Application");  
            var oWB = oXL.Workbooks.Add();  
            var oSheet = oWB.ActiveSheet;  
            var sel = document.body.createTextRange();  
            sel.moveToElementText(curTbl);  
            sel.select();  
            sel.execCommand("Copy");  
            oSheet.Paste();  
            oXL.Visible = true;  
  
        }  
        //第二种方法  
        function method2(tableid)  
        {  
  
            var curTbl = document.getElementById(tableid);  
            var oXL = new ActiveXObject("Excel.Application");  
            var oWB = oXL.Workbooks.Add();  
            var oSheet = oWB.ActiveSheet;  
            var Lenr = curTbl.rows.length;  
            for (i = 0; i < Lenr; i++)  
            {        var Lenc = curTbl.rows(i).cells.length;  
                for (j = 0; j < Lenc; j++)  
                {  
                    oSheet.Cells(i + 1, j + 1).value = curTbl.rows(i).cells(j).innerText;  
  
                }  
  
            }  
            oXL.Visible = true;  
        }  
        //第三种方法  
        function getXlsFromTbl(inTblId, inWindow){  
  
            try {  
                var allStr = "";  
                var curStr = "";  
                if (inTblId != null && inTblId != "" && inTblId != "null") {  
  
                    curStr = getTblData(inTblId, inWindow);  
  
                }  
                if (curStr != null) {  
                    allStr += curStr;  
                }  
  
                else {  
  
                    alert("你要导出的表不存在");  
                    return;  
                }  
                var fileName = getExcelFileName();  
                doFileExport(fileName, allStr);  
  
            }  
  
            catch(e) {  
  
                alert("导出发生异常:" + e.name + "->" + e.description + "!");  
  
            }  
  
        }  
  
        function getTblData(inTbl, inWindow) {  
  
            var rows = 0;  
            var tblDocument = document;  
            if (!!inWindow && inWindow != "") {  
  
                if (!document.all(inWindow)) {  
                    return null;  
                }  
  
                else {  
                    tblDocument = eval(inWindow).document;  
                }  
  
            }  
  
            var curTbl = tblDocument.getElementById(inTbl);  
            var outStr = "";  
            if (curTbl != null) {  
                for (var j = 0; j < curTbl.rows.length; j++) {  
                    for (var i = 0; i < curTbl.rows[j].cells.length; i++) {  
  
                        if (i == 0 && rows > 0) {  
                            outStr += " t";  
                            rows -= 1;  
                        }  
  
                        outStr += curTbl.rows[j].cells[i].innerText + "t";  
                        if (curTbl.rows[j].cells[i].colSpan > 1) {  
                            for (var k = 0; k < curTbl.rows[j].cells[i].colSpan - 1; k++) {  
                                outStr += " t";  
                            }  
                        }  
                        if (i == 0) {  
                            if (rows == 0 && curTbl.rows[j].cells[i].rowSpan > 1) {  
                                rows = curTbl.rows[j].cells[i].rowSpan - 1;  
                            }  
                        }  
                    }  
                    outStr += "rn";  
                }  
            }  
  
            else {  
                outStr = null;  
                alert(inTbl + "不存在 !");  
            }  
            return outStr;  
        }  
  
        function getExcelFileName() {  
            var d = new Date();  
            var curYear = d.getYear();  
            var curMonth = "" + (d.getMonth() + 1);  
            var curDate = "" + d.getDate();  
            var curHour = "" + d.getHours();  
            var curMinute = "" + d.getMinutes();  
            var curSecond = "" + d.getSeconds();  
            if (curMonth.length == 1) {  
                curMonth = "0" + curMonth;  
            }  
  
            if (curDate.length == 1) {  
                curDate = "0" + curDate;  
            }  
  
            if (curHour.length == 1) {  
                curHour = "0" + curHour;  
            }  
  
            if (curMinute.length == 1) {  
                curMinute = "0" + curMinute;  
            }  
  
            if (curSecond.length == 1) {  
                curSecond = "0" + curSecond;  
            }  
            var fileName = "table" + "_" + curYear + curMonth + curDate + "_"  
                    + curHour + curMinute + curSecond + ".csv";  
            return fileName;  
  
        }  
  
        function doFileExport(inName, inStr) {  
            var xlsWin = null;  
            if (!!document.all("glbHideFrm")) {  
                xlsWin = glbHideFrm;  
            }  
            else {  
                var width = 6;  
                var height = 4;  
                var openPara = "left=" + (window.screen.width / 2 - width / 2)  
                        + ",top=" + (window.screen.height / 2 - height / 2)  
                        + ",scrollbars=no,width=" + width + ",height=" + height;  
                xlsWin = window.open("", "_blank", openPara);  
            }  
            xlsWin.document.write(inStr);  
            xlsWin.document.close();  
            xlsWin.document.execCommand('Saveas', true, inName);  
            xlsWin.close();  
  
        }  
  
        //第四种  
        function method4(tableid){  
  
            var curTbl = document.getElementById(tableid);  
            var oXL;  
            try{  
                oXL = new ActiveXObject("Excel.Application"); //创建AX对象excel  
            }catch(e){  
                alert("无法启动Excel!\n\n如果您确信您的电脑中已经安装了Excel，"+"那么请调整IE的安全级别。\n\n具体操作：\n\n"+"工具 → Internet选项 → 安全 → 自定义级别 → 对没有标记为安全的ActiveX进行初始化和脚本运行 → 启用");  
                return false;  
            }  
            var oWB = oXL.Workbooks.Add(); //获取workbook对象  
            var oSheet = oWB.ActiveSheet;//激活当前sheet  
            var sel = document.body.createTextRange();  
            sel.moveToElementText(curTbl); //把表格中的内容移到TextRange中  
            sel.select(); //全选TextRange中内容  
            sel.execCommand("Copy");//复制TextRange中内容  
            oSheet.Paste();//粘贴到活动的EXCEL中  
            oXL.Visible = true; //设置excel可见属性  
            var fname = oXL.Application.GetSaveAsFilename("将table导出到excel.xls", "Excel Spreadsheets (*.xls), *.xls");  
            oWB.SaveAs(fname);  
            oWB.Close();  
            oXL.Quit();  
        }  
  
  
        //第五种方法  
        var idTmr;  
        function  getExplorer() {  
            var explorer = window.navigator.userAgent ;  
            //ie  
            if (explorer.indexOf("MSIE") >= 0) {  
                return 'ie';  
            }  
            //firefox  
            else if (explorer.indexOf("Firefox") >= 0) {  
                return 'Firefox';  
            }  
            //Chrome  
            else if(explorer.indexOf("Chrome") >= 0){  
                return 'Chrome';  
            }  
            //Opera  
            else if(explorer.indexOf("Opera") >= 0){  
                return 'Opera';  
            }  
            //Safari  
            else if(explorer.indexOf("Safari") >= 0){  
                return 'Safari';  
            }  
        }  
        function method5(tableid) {  
            if(getExplorer()=='ie')  
            {  
                var curTbl = document.getElementById(tableid);  
                var oXL = new ActiveXObject("Excel.Application");  
                var oWB = oXL.Workbooks.Add();  
                var xlsheet = oWB.Worksheets(1);  
                var sel = document.body.createTextRange();  
                sel.moveToElementText(curTbl);  
                sel.select();  
                sel.execCommand("Copy");  
                xlsheet.Paste();  
                oXL.Visible = true;  
  
                try {  
                    var fname = oXL.Application.GetSaveAsFilename("Excel.xls", "Excel Spreadsheets (*.xls), *.xls");  
                } catch (e) {  
                    print("Nested catch caught " + e);  
                } finally {  
                    oWB.SaveAs(fname);  
                    oWB.Close(savechanges = false);  
                    oXL.Quit();  
                    oXL = null;  
                    idTmr = window.setInterval("Cleanup();", 1);  
                }  
  
            }  
            else  
            {  
                tableToExcel(tableid)  
            }  
        }  
        function Cleanup() {  
            window.clearInterval(idTmr);  
            CollectGarbage();  
        }  
        var tableToExcel = (function() {  
            var uri = 'data:application/vnd.ms-excel;base64,',  
                    template = '<html><head><meta charset="UTF-8"></head><body><table>{table}</table></body></html>',  
                    base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) },  
                    format = function(s, c) {  
                        return s.replace(/{(\w+)}/g,  
                                function(m, p) { return c[p]; }) }  
            return function(table, name) {  
                if (!table.nodeType) table = document.getElementById(table)  
                var ctx = {worksheet: name || 'Worksheet', table: table.innerHTML}  
                window.location.href = uri + base64(format(template, ctx))  
            }  
        })()  
  
    </script> 
      <section id="main-content">
          <section class="wrapper">
              <!-- page start -->
<%
dim i
dim sql,rs

sql="select * From [Table_Product] where ArticleID>0"
sql=sql & " order by orderid desc,ArticleID desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
             <!--成功失败提示-->
<div class="alert" id="alert" style="display:none;">
                                  <button class="close close-sm" type="button">
                                      <i class="icon-remove"></i>
                                  </button>
                                  <strong>操作成功！</strong> <span>栏目内容修改成功。</span>
                              </div>
 
                          <!----> 
              <div class="row">
                  <div class="col-lg-12">
                      <section class="panel">
                          <header class="panel-heading">
                              产品列表
                          </header>
                          
<!--列表开始-->
<div >  
        <button type="button" onclick="method1('sample_1')">导出Excel方法一</button>  
        <button type="button" onclick="method2('sample_1')">导出Excel方法二</button>  
        <button type="button" onclick="getXlsFromTbl('sample_1','myDiv')">导出Excel方法三</button>  
        <button type="button" onclick="method4('sample_1')">导出Excel方法四</button>  
        <button type="button" onclick="method5('sample_1')">导出Excel方法五</button>  
    </div> 
<!--以后补一个类别切换-->
<form>
<table class="table table-striped border-top table-hover" id="sample_1">
<thead>
          <tr > 
          
            <th width="40" class="hidden-phone"><strong>ID</strong></td>
            <th width="60" class="hidden-phone"><strong>国家</strong></td>
            <th width="70" class="hidden-phone"><strong>城市</strong></td>
            <th width="200" ><strong>项目名称</strong></td>
                
                    <th width="200" ><strong>价格1</strong></td>
                <th width="200" ><strong>价格2</strong></td>
                    <th width="200" ><strong>面积1</strong></td>
                        <th width="200" ><strong>面积2</strong></td>
                            <th width="200" ><strong>豪宅类型</strong></td>
                                <th width="200" ><strong>房源类型</strong></td> 
                                    <th width="200" ><strong>坐标</strong></td> 
                                        <th width="200" ><strong>销售顾问</strong></td>
                                            <th width="200" ><strong>户型</strong></td>
                                                <th width="200" ><strong>特点</strong></td>
                                                    <th width="200" ><strong>项目简介</strong></td>
                                                        <th width="200" ><strong>户型图</strong></td>
            <th width="80" class="hidden-phone"><strong>加入时间</strong></td>
            <th width="40" class="hidden-phone"><strong>排序</strong></td>
             <th width="40" class="hidden-phone"><strong>JLL代理</strong></td>
            <th width="40" class="hidden-phone"><strong>审核</strong></td>
            <th width="40" class="hidden-phone"><strong>推荐</strong></td>
      
          </tr></thead>
          <tbody>
          <%do while not rs.eof%>
          <%if rs("gid")="" then gid=0 else gid=rs("gid") end if%>
          <tr > 
            <td class="hidden-phone"><%=rs("articleid")%></td>
            <td class="hidden-phone"><%=rs("bigclassname")%></td>
            <td class="hidden-phone"><%=rs("smallclassname")%></td>
            <td><%=rs("title")%> <!--[ <%=rs("entitle")%> ] --><%if rs("Passed")=false then%>(草稿)<%end if%></td>
            <th width="200" ><strong><%=rs("jgzj")%></strong></td>
                <th width="200" ><strong><%=rs("jgjj")%></strong></td>
                    <th width="200" ><strong><%=rs("mj1")%></strong></td>
                        <th width="200" ><strong><%=rs("mj2")%></strong></td>
                            <th width="200" ><strong><%
                                clbid=rs("clbid")
                                Set rs2= Server.CreateObject("ADODB.Recordset")
                                sql="select * from [class_lb] order by oid desc"
                                rs2.open sql,conn,1,1
                                if not(rs2.eof and rs2.bof) then
                                do while not rs2.eof
                                %>
                                           <%if int(rs2("cid"))=int(clbid) then%><%=rs2("CName")%><%end if%>
                                            <%
                                rs2.movenext
                                loop
                                end if
                                rs2.close
                                
                                %></strong></td>
                                <th width="200" ><strong><%
                                    ckfsid=rs("ckfsid")
                                    sql="select * from [class_kfs] order by oid desc"
                                    rs2.open sql,conn,1,1
                                    if not(rs2.eof and rs2.bof) then
                                    
                                    do while not rs2.eof
                                    %>
                                             <%if int(rs2("cid"))=int(ckfsid) then%><%=rs2("CName")%><%end if%>
                                                <%
                                    rs2.movenext
                                    loop
                                    end if
                                    rs2.close
                                    
                                    %></strong></td> 
                                    <th width="200" ><strong><%=Rs("qt")%></strong></td> 
                                        <th width="200" ><strong><%
                                            gid=rs("gid")
                                                        sql="select * from [team] order by oid desc,id desc"
                                                        rs2.open sql,conn,1,1
                                                        if not(rs2.eof and rs2.bof) then
                                            
                                                        do while not rs2.eof
                                                        %>
                                                                    <%if int(rs2("id"))=int(gid) then%><%=rs2("CName")%> <%end if%>
                                                                    <%
                                                        rs2.movenext
                                                        loop
                                                        end if
                                                        rs2.close
                                            
                                                        %></strong></td>
                                            <th width="200" ><strong><%
                                                
                                                sql="select * from [class_lx] order by oid desc"
                                                rs2.open sql,conn,1,1
                                                if not(rs2.eof and rs2.bof) then
                                                
                                                do while not rs2.eof
                                                %>
                                                      <%if instr(","&rs("clxid")&",",","&rs2("Cid")&",") then%><%=rs2("CName")%> <%end if%> <%
                                                rs2.movenext
                                                loop
                                                end if
                                                rs2.close
                                                
                                                %></strong></td>
                                                <th width="200" ><strong>
                                                        <%
                                                        
                                                        sql="select * from [class_td] order by oid desc"
                                                        rs2.open sql,conn,1,1
                                                        if not(rs2.eof and rs2.bof) then
                                                        
                                                        do while not rs2.eof
                                                        %>
                                                                <%if instr(rs("ctdid"),","&rs2("Cid")&",") then%><%=rs2("CName")%> <%end if%> <%
                                                        rs2.movenext
                                                        loop
                                                        end if
                                                        rs2.close
                                                        
                                                        %></strong></td>
                                                    <th width="200" ><strong><%=RemoveHTML(rs("content1"))%></strong></td>
                                                        <th width="200" ><strong><%=RemoveHTML(rs("content"))%></strong></td>
            <td class="hidden-phone"><%= FormatDateTime(rs("UpdateTime"),2) %></td>
            <td class="hidden-phone"><%=rs("orderid")%></td>
            <td class="hidden-phone"><%if gid<>0 then%><i class="icon-ok" style="color:#78cd51;"></i><%else%><i class="icon-remove" style="color:#ff6c60;"></i><%end if%></td>
            <td class="hidden-phone"><%if rs("Passed")=false then%><i class="icon-remove" style="color:#ff6c60;"></i><%else%><i class="icon-ok" style="color:#78cd51;"></i><%end if%></td>
                       <td class="hidden-phone"><%if rs("Elite")=false then%><i class="icon-remove" style="color:#ff6c60;"></i><%else%><i class="icon-ok" style="color:#78cd51;"></i><%end if%></td>
            

          </tr>
          <%
	i=i+1
	  
	      rs.movenext
	loop
%>
          </tbody>
          </table>
        </table>
</form>
<!--列表结束-->
                      </section>
                  </div>
              </div>

              <!--  page end  -->
          </section>
      </section>
      <!--main content end-->
  </section>

<!--#include file="inc/xgfooter.asp"-->
<!--#include file="inc/windows.asp"-->
<%
action=request.QueryString("Action")
delid=request.QueryString("ArticleID")
if action="Del" then
conn.execute "delete from [Table_Product] where ArticleId="&Trim(delid)
response.Redirect("Pro_List.Asp?sta=del")
end if
%>
  <script type="text/javascript">
//导航菜单
$(".nav3").addClass('active');
$(".nav3-1").addClass('active');

function ConfirmDel()
{
   if(confirm("确定要删除选中的产品吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
}
var sta='<%=request.QueryString("sta")%>';
if(sta=='add'){
			  		$("#alert").fadeIn();
					$("#alert").addClass("alert-success");
					$("#alert span").html("数据添加成功!");
}else if(sta=='mod'){
			  		$("#alert").fadeIn();
					$("#alert").addClass("alert-success");
					$("#alert span").html("数据修改成功!");
}else if(sta=='del'){
			  		$("#alert").fadeIn();
					$("#alert").addClass("alert-success");
					$("#alert span").html("数据删除成功!");
}
console.log(sta);
</script>
<script type="text/javascript" src="xgwl/assets/data-tables/jquery.dataTables.js"></script>
<script type="text/javascript" src="xgwl/assets/data-tables/DT_bootstrap.js"></script>
<script src="xgwl/js/dynamic-table.js"></script>
