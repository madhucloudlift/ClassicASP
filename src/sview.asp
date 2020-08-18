<html>
<head>
<title> [ AspBB ] - View </title>
<style type="text/css">
A:link, A:visited { text-decoration: none; color: 'silver' }
</style>
</head>

<body bgcolor="#006699">
<a name='#top'></a>
<%
author = Request.QueryString("author")
keyword = Request.QueryString("keyword")
style = Request.QueryString("style")
dim seepage
if request.querystring("seepage") = "" OR request.querystring("seepage") < 1 then
seepage = 1
else
seepage = request.querystring("seepage")
end if
 
id = Request.QueryString("msgID")
set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 
%>
<div align=center>
<table border=0 width=80%>
<tr>
<td width="67%"> 
<%
if author = "" then
 if Session("currentpage")="" then
 Response.Write "<font face=arial size=1><a href='searching.asp?keyword="&keyword&"&style="&style&"'>Back to Topics</a>"
 else
 Response.Write "<font face=arial size=1><a href='searching.asp?currentpage="&Session("currentpage")&"&keyword="&keyword&"&style="&style&"'>Back to Topics</a></font>"
 end if
else
 if Session("currentpage")="" then
 Response.Write "<font face=arial size=1><a href='searching.asp?author="&author&"'>Back to Topics</a>"
 else
 Response.Write "<font face=arial size=1><a href='searching.asp?currentpage="&Session("currentpage")&"&author="&author&"'>Back to Topics</a></font>"
 end if
end if
%>
</td>
<td align=right width="33%"> 
<%
Set rsp = Server.CreateObject("ADODB.Recordset")
if author = "" AND style = "post" then 
rsp.Open "SELECT * FROM board where msgID>"&id&" AND body LIKE '%" & Replace(keyword, "'", "''") & "%'ORDER BY msgID ASC", conn, 3
end if
if author = "" AND style = "topic" then
rsp.Open "SELECT * FROM board where msgID>"&id&" AND threadP = 0 AND topic LIKE '%" & Replace(keyword, "'", "''") & "%'ORDER BY msgID ASC", conn, 3
end if
if author <> "" then
rsp.Open "SELECT * FROM board where msgID>"&id&" AND name LIKE '%" & Replace(author, "'", "''") & "%'ORDER BY msgID ASC", conn, 3
end if
if rsp.EOF then
Response.Write "<font face=arial size=1 color=white> << View Pervious Thread :</font>"
else
 if author = "" then
 Response.Write "<font face=arial size=1 color=white><a href='sview.asp?msgID="&rsp("msgID")&"&keyword="&keyword&"&style="&style&"'> << View Pervious Thread</a> :</font>"
 else
 Response.Write "<font face=arial size=1 color=white><a href='sview.asp?msgID="&rsp("msgID")&"&author="&author&"'> << View Pervious Thread</a> :</font>"
 end if
end if

Set rsn = Server.CreateObject("ADODB.Recordset")
if author = "" AND style = "post" then 
rsn.Open "SELECT * FROM board where msgID<"&id&" AND body LIKE '%" & Replace(keyword, "'", "''") & "%'ORDER BY msgID ASC", conn, 3
end if
if author = "" AND style = "topic" then
rsn.Open "SELECT * FROM board where msgID<"&id&" AND threadP = 0 AND topic LIKE '%" & Replace(keyword, "'", "''") & "%'ORDER BY msgID ASC", conn, 3
end if
if author <> "" then
rsn.Open "SELECT * FROM board where msgID<"&id&" AND name LIKE '%" & Replace(author, "'", "''") & "%'ORDER BY msgID ASC", conn, 3
end if
if rsn.EOF then
Response.Write "<font face=arial size=1 color=white>: View Next Thread >> </font>"
else
 if author = "" then
 Response.Write "<font face=arial size=1 color=white>: <a href='sview.asp?msgid="&rsn("msgID")&"&keyword="&keyword&"&style="&style&"'>View Next Thread >> </a></font>"
 else
 Response.Write "<font face=arial size=1 color=white>: <a href='sview.asp?msgid="&rsn("msgID")&"&author="&author&"'>View Next Thread >> </a></font>"
 end if
end if
%>
</td>
</tr>
</table><br>
</div>
<% if author <> "" OR style = "post" then %>
<% Set rsview = Server.CreateObject("ADODB.Recordset") %>
<% rsview.Open "SELECT * FROM board where msgID="&id, conn, 3 %>
<% 
   if rsview("threadID") = 0 then
   Set rsadd = Server.CreateObject("ADODB.Recordset")
   rsadd.Open "SELECT * FROM board WHERE msgID="&id, conn, 1,2
   rsadd("threadID") = rsadd("msgID")
   rsadd.Update  
   end if
%>
<% While NOT rsview.EOF %>
<div align=center>
<table border=1 width=80%cellpadding=0 cellspacing=0>
<tr><td>
<font face=arial size=2 color=white><b><%=rsview("topic")%></b></font>
</td></tr>
</table> 
<table border=0 width=80% cellpadding=0 cellspacing=0>
<tr>
      <td valign=top width="25%">
	  <%  
	  if instr(rsview("email") & "" , "@") > 0 then
	  Response.Write "<font face=arial color=silver size=2><b><a href='mailto:" & trim(rsview("email")) & "'>" & trim(rsview("name")) & "</a></b></font>"
	  else
	  Response.Write "<b><font face=arial color=white size=2>"& rsview("name") &"</b></font>"
	  end if    
	  %>
	  </td>
      <td width="75%"> 
	 <font face=arial size=1 color=silver>Posted on <%=rsview("date")%></font><br>
	 <hr noshade>
	 <font face=arial color=white size=2><b><%=rsview("body")%></b></font><br><br>
     </td>
</tr>
</table>
<table border=0 width=80% cellpadding=0 cellspacing=0>
<tr>
      <td width="25%"><font face=arial size=1 color><b><a href='#top'>Back to Top</a></b></font></td>
      <td width="15%">
	  <%    
	  if instr(rsview("email") & "" , "@") > 0 then
	  Response.Write "<font face=arial color=silver size=2><b><a href='mailto:" & trim(rsview("email")) & "'>E-Mail</a></b></font>"
	  end if  
	  %>
	  </td>
	  <td width="15%">
	  <%    
	  if instr(rsview("http") & "", "http://") > 0 then
      Response.Write "<font face=arial color=silver size=2><b><a href='" & trim(rsview("http")) &"'>WebSite</a></b></font>"
	  end if
      %>
	  </td>
	  <td width="45%"></td>
</tr>
</table>
</div>
<br>
<%
rsview.MoveNext
WEND
%>
<% else %>

<% 
Set rsadd = Server.CreateObject("ADODB.Recordset")
rsadd.Open "SELECT * FROM board WHERE msgID="&id, conn, 1,2
rsadd("threadID") = rsadd("msgID")
rsadd.Update 
%>

<% Set rsviewtopic = Server.CreateObject("ADODB.Recordset") %>
<% 
rsviewtopic.CursorLocation = 3 'set cursorlocation to aduserclient 
rsviewtopic.CacheSize = 10 'cache the number of record to display per page into cache 
rsviewtopic.Open "SELECT * FROM board where threadID="&id, conn, 3 
rsviewtopic.PageSize = 10 'Set that each page will list 10 records 
%>

<% 
Dim TotalPages, TotalRecords
TotalPages = rsviewtopic.PageCount 'Pagecount will count how many page will need if 10 record fill one page
TotalRecords = rsviewtopic.RecordCount
rsviewtopic.AbsolutePage = seepage 'the currentpage value will be the rs.Absolute value
Dim count 'Counting variable for our recordset
count = 0 
Do While Not rsviewtopic.EOF AND count < rsviewtopic.PageSize 
%>
<div align=center>
<table border=1 width=80%cellpadding=0 cellspacing=0>
<tr><td>
<font face=arial size=2 color=white><b><%=rsviewtopic("topic")%></b></font>
</td></tr>
</table> 
<table border=0 width=80% cellpadding=0 cellspacing=0>
<tr>
      <td valign=top width="25%">
	  <%  
	  if instr(rsviewtopic("email") & "" , "@") > 0 then
	  Response.Write "<font face=arial color=silver size=2><b><a href='mailto:" & trim(rsviewtopic("email")) & "'>" & trim(rsviewtopic("name")) & "</a></b></font>"
	  else
	  Response.Write "<b><font face=arial color=white size=2>"& rsviewtopic("name") &"</b></font>"
	  end if    
	  %>
	  </td>
      <td width="75%"> 
	 <font face=arial size=1 color=silver>Posted on <%=rsviewtopic("date")%></font><br>
	 <hr noshade>
	 <font face=arial color=white size=2><b><%=rsviewtopic("body")%></b></font><br><br>
     </td>
</tr>
</table>
<table border=0 width=80% cellpadding=0 cellspacing=0>
<tr>
      <td width="25%"><font face=arial size=1 color><b><a href='#top'>Back to Top</a></b></font></td>
      <td width="15%">
	  <%    
	  if instr(rsviewtopic("email") & "" , "@") > 0 then
	  Response.Write "<font face=arial color=silver size=2><b><a href='mailto:" & trim(rsviewtopic("email")) & "'>E-Mail</a></b></font>"
	  end if  
	  %>
	  </td>
	  <td width="15%">
	  <%    
	  if instr(rsviewtopic("http") & "", "http://") > 0 then
      Response.Write "<font face=arial color=silver size=2><b><a href='" & trim(rsviewtopic("http")) &"'>WebSite</a></b></font>"
	  end if
      %>
	  </td>
	  <td width="45%"></td>
</tr>
</table>
</div>
<br>
<%
count = count + 1
rsviewtopic.MoveNext
Loop

Set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = conn
rs.Open "SELECT * FROM board where msgID="&id, conn, 1, 2
rs("count") = rs("count").OriginalValue + 1
rs.Update
%>
<div align=center>
<table border="0" width="84%">
<tr><td align=right>
<font face="arial" size="2">
<a href="repost.asp?msgID=<%=id%>"><b>Reply to this Topic</b></a>
</font>
</td></tr></table></div>
<br>
<br>
<div align="center">
<table border="0" width="84%">
<tr>
      <td width="21%"><p align="left"><font face="arial" size="2"><b>
	  <%
	  if seepage > 1 then
	  Response.Write "<a href='sview.asp?seepage=1&msgID="&id&"&keyword="&keyword&"&style="&style&"'>First</a>"
	  end if 
	  %>
	  </b></font></td>
      <td width="21%"><p align="center"><font face="arial" size="2"><b>
	 <% 
	  if seepage > 1 then
	  Response.Write "<a href='sview.asp?seepage="& seepage - 1 &"&msgID="&id&"&keyword="&keyword&"&style="&style&"'><< Pervious</a>"
	  end if
	 %>
	  </b></font></td>
      <td width="21%"><p align="center"><font face="arial" size="2"><b>
	  <% 
	  if CInt(seepage) <> CInt(TotalPages) then
	  Response.Write "<a href='sview.asp?seepage="& seepage + 1 &"&msgID="&id&"&keyword="&keyword&"&style="&style&"'>Next >></a>"
	  end if
	  %>
	  </b></font></td>
      <td width="21%"><p align="right"><font face="arial" size="2"><b>
	  <%
	  if CInt(seepage) <> CInt(TotalPages) then
	  Response.Write "<a href='sview.asp?seepage="& TotalPages &"&msgID="&id&"&keyword="&keyword&"&style="&style&"'>Last</a>"
	  end if
	  %>
	  </b></font></td>
</tr>
</table>
</div>

<%
if TotalPages > 1 then
Response.Write "<center>"
Response.Write "<font face='arial' size='1' color=white>"
Response.Write "<b>Page:</b> "
Response.Write "<select name='whatever' onChange='window.location=this.options[this.selectedIndex].value'>"
Response.Write "<option value='sview.asp?seepage="&seepage&"&msgID="&id&"&keyword="&keyword&"&style="&style&" SELECTED'> - "&seepage&" - </option>"
for i = 1 to TotalPages
if cStr(seepage) <> cStr(i) then
Response.Write "<option value='sview.asp?seepage="&i&"&msgID="&id&"&keyword="&keyword&"&style="&style&"'> - "&i&" - </option>"
end if
next
Response.Write "</select>"
Response.Write "</font>"
Response.Write "</center>"

  Response.Write "<font face=arial size=1 color=white>"
  Response.Write "<p align=center>"
  Response.Write("Page " & seepage & " of " & TotalPages & "</p>")
  Response.Write "</font>"
end if
%>
<% end if %>
<% conn.Close
Set conn = Nothing %>
</body>
</html>
