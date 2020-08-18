<html>
<head>
<title> [ AspBB ] - Active Server Page Bulletin Board </title>
<style type="text/css">
A:link, A:visited { text-decoration: none; color: 'silver' }
</style>
</head>
<body bgcolor="#006699">
<p align=center>
<font face=arial color=white size=4><b>AspBB - Active Server Page Bulletin Board</b></font><br>
<%
 dim currentpage
 if request.querystring("currentpage") = "" OR request.querystring("currentpage") < 1 then
 currentpage = 1
 else
 currentpage = request.querystring("currentpage")
 end if
Session("currentpage")=currentpage
set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 
%> 
<div align="center">
<table border="0" width="90%">
<tr>
      <td width="69%"> 
        <% Response.Write "<font face=arial size=1 color=white>The time now is: </font><font face=arial size=1 color=silver>"&now()&"</font>"%>
      </td>
	  <td><font face=arial size=2><b><a href='search.asp'>Search</a></b></font></td>
      <td width="25%"> 
        <p align="right"><font face="arial" size="1"><b> 
<%Set rsot = Server.CreateObject("ADODB.Recordset")
rsot.Open "SELECT * FROM board where threadP=0",conn, 3%>
<%response.write "<font face='arial' size=2 color=white>Total Topics: </font><b><font face='arial' size=2 color=silver>"& rsot.recordcount &"</font></b>"%>
<% rsot.Close %>
<b><font face="arial" size="2" color=white> - </font></b>
<%Set rsott = Server.CreateObject("ADODB.Recordset")
rsott.Open "SELECT * FROM board",conn, 3%>
<%response.write "<font face='arial' size=2 color=white>Total Posts: </font><b><font face='arial' size=2 color=silver>"& rsott.recordcount &"</font></b>"%>
<% rsott.Close %>
	  </b></font></td>
</tr>
</table>
</div>
<div align=center>
<table border=1 width=90% cellpadding=0 cellspacing=0>
<tr>
<td width=40%><p align=left><font face=arial size=2 color=white><b>Thread</b></font></td>
<td width=17%><p align=center><font face=arial size=2 color=white><b>Thread Starter</b></font></td>
<td width=10%><p align=center><font face=arial size=2 color=white><b>Replies</b></font></td>
<td width=10%><p align=center><font face=arial size=2 color=white><b>Views</b></font></td>
<td width=23%><p align=left><font face=arial size=2 color=white><b>Last Post</b></font></td>
</tr>
</table>
</div>
<%
Set rsttrv = Server.CreateObject("ADODB.Recordset")

rsttrv.CursorLocation = 3 'set cursorlocation to aduserclient 
rsttrv.CacheSize = 10 'cache the number of record to display per page into cache
rsttrv.Open "SELECT * FROM board WHERE ThreadP=0 ORDER BY msgID DESC", conn, 3

if rsttrv.EOF then Response.Write "<p align=center><font face=arial size=2 color=white><b>No Post Yet</b></font>"
rsttrv.PageSize = 10 'Set that each page will list 10 records

Dim TotalPages, TotalRecords
TotalPages = rsttrv.PageCount 'Pagecount will count how many page will need if 10 record fill one page
TotalRecords = rsttrv.RecordCount

if rsttrv.RecordCount > 0 then
rsttrv.AbsolutePage = currentpage 'the currentpage value will be the rs.Absolute value
end if

Dim count 'Counting variable for our recordset
count = 0 
Do While Not rsttrv.EOF AND count < rsttrv.PageSize
%>
<div align=center>
<table border=0 width=90% cellpadding=0 cellspacing=0>
<tr>
      <td width=40%><font face=arial size=2><b><a href='view.asp?msgID=<%=rsttrv("msgID")%>'><%=rsttrv("topic")%></a></b></font>
<% msgID = rsttrv("msgid") 
Set rsnay = Server.CreateObject("ADODB.Recordset")
rsnay.CacheSize = 10 'cache the number of record to display per page into cache
rsnay.PageSize = 10 'Set that each page will list 10 records
rsnay.Open "SELECT * FROM board WHERE ThreadID="&msgID&" ORDER BY msgID", conn, 3
TotalPage = rsnay.PageCount 'Pagecount will count how many page will need if 10 record fill one page
if TotalPage > 1 then
Response.Write "<font face='arial' size='1' color=white>"
Response.Write "<b><< Page:</b> "
Response.Write "<select name='whatever' onChange='window.location=this.options[this.selectedIndex].value'>"
Response.Write "<option value='view.asp?msgID="&msgID&"&seepage=1 SELECTED'> - 1 - </option>"
for i = 1 to TotalPage
Response.Write "<option value='view.asp?msgID="&msgID&"&seepage="&i&"'> - "&i&" - </option>"
next
Response.Write "</select>"
Response.Write " >>"
end if
%>
</td>
	  <td width=17% align=center> 
        <% if instr(rsttrv("email") & "" , "@") > 0 then %>
        <font face=arial size=2><b><a href='mailto:<%=trim(rsttrv("email"))%>'><%=trim(rsttrv("name"))%></a></b></font> 
        <%else%>
        <font face=arial size=2 color=white><b><%=rsttrv("name")%></b></font>
        <% end if %>
      </td>
<%
Set rsr = Server.CreateObject("ADODB.Recordset")
rsr.Open "SELECT * FROM board WHERE ThreadID="&msgID&"AND ThreadP=1",conn, 3
%>
      <td width=10% align=center><font face=arial color=silver size=2><b><%=rsr.recordcount%></b></font></td>
      <td width=10% align=center><font face=arial color=silver size=2><b><%=rsttrv("count")%></b></font></td>
<%
if rsr.recordcount <> 0 then
Set rslp = Server.CreateObject("ADODB.Recordset")
rslp.Open "SELECT * FROM board WHERE ThreadID="&msgID&" ORDER BY msgID",conn, 3
rslp.MoveLast
%>
      <td width=23%><font face=arial color=white size=1><b><%=rslp("date")%> by </b></font> 
        <% if instr(rslp("email") & "" , "@") > 0 then %>
        <font face=arial color=silver size=1><b><a href='mailto:<%=trim(rslp("email"))%>'><%=trim(rslp("name"))%></a></b></font> 
        <% else %>
        <font face=arial color=white size=1><b><%=rslp("name")%></b></font> 
        <% end if %>
      </td>
 <% else %>
      <td width="23%"><font face=arial color=white size=1><b><%=rsttrv("date")%></b></font></td>
<% end if %>
</tr></table>
</div>
<%
count = count + 1
rsttrv.MoveNext
Loop
%>
<br>
<p align=center>
<font face=arial size=2><b><a href="post.asp">Post New Topic</a></b></font>
<br>
<br>
<%
Response.Write "<div align=center>"
Response.Write "<table border=0 width=840>"
Response.Write "<tr>"
Response.Write "<td width=210><p align=left><font face=arial size=2><b>"
	  if rsttrv.RecordCount > 0 then 
	  if currentpage > 1 then
	  Response.Write "<a href='default.asp?currentpage=1'>First</a>"
	  end if
	  end if
Response.Write "</b></font></td>"
Response.Write "<td width=210><p align=center><font face=arial size=2><b>"

	  if rsttrv.RecordCount > 0 then
	  if currentpage > 1 then
	  Response.Write "<a href='default.asp?currentpage=" & currentpage - 1 &"'><< Pervious</a>"
	  end if
	  end if
	
Response.Write "</b></font></td>"
Response.Write "<td width=210><p align=center><font face=arial size=2><b>"
	  
	  if rsttrv.RecordCount > 0 then
	  if CInt(currentpage) <> CInt(TotalPages) then
	  Response.Write "<a href='default.asp?currentpage=" & currentpage + 1 &"'>Next >></a>"
	  end if
	  end if
	 
	  Response.Write "</b></font></td>"
      Response.Write "<td width=210><p align=right><font face=arial size=2><b>"
	 
	  if rsttrv.RecordCount > 0 then 
	  if CInt(currentpage) <> CInt(TotalPages) then
	  Response.Write "<a href='default.asp?currentpage=" & TotalPages &"'>Last</a>"
	  end if
	  end if
	 
	  Response.Write "</b></font></td>"
Response.Write "</tr></table></div>"
%>
<%
if TotalPages > 1 then
Response.Write "<center>"
Response.Write "<font face='arial' size='1' color=white>"
Response.Write "<b>Page:</b> "
Response.Write "<select name='whatever' onChange='window.location=this.options[this.selectedIndex].value'>"
Response.Write "<option value='default.asp?currentpage="&currentpage&" SELECTED'> - "&currentpage&" - </option>"
for i = 1 to TotalPages
if cStr(currentpage) <> cStr(i) then
Response.Write "<option value='default.asp?currentpage="&i&"'> - "&i&" - </option>"
end if
next
Response.Write "</select>"
Response.Write "</font>"
Response.Write "</center>"

Response.Write "<font face=arial size=1 color=white>"
Response.Write "<p align=center>"
Response.Write("Page " & currentpage & " of " & TotalPages & "</p>")
Response.Write "</font>"
end if
%>
<%
conn.Close
Set conn = Nothing
%>
<center>
<font face=arial size=1 color=white>Powered by AspBB v1.0 - Programmed by <a href='mailto:doma111@yahoo.com'>Johnny Yu</a></font>
</center>
</body>
</html>
