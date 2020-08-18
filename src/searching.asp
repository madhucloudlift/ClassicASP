<html>
<head>
<title> [ AspBB ] - Search Results </title>
<style type="text/css">
A:link, A:visited { text-decoration: none; color: 'silver' }
</style>
</head>
<body bgcolor="#006699">
<p align=center>
<font face=arial color=white size=4><b>AspBB - Search Results</b></font><br>
<%
dim currentpage
if request.QueryString("currentpage") = "" OR Request.QueryString("currentpage") < 1 then
currentpage = 1
else
currentpage = Request.QueryString("currentpage")
end if
Session("currentpage")=currentpage
%>
<%
if Request.Form("tosearch") = "" AND Request.Form("username") = "" then
Response.Redirect("search.asp?error=1")
end if
%>
<% 
if Request.Form("tosearch") <> "" then
keyword = Request.Form("tosearch") 
style = Request.Form("style")
end if
%>
<%
if Request.Form("username") <> "" then
author = Request.Form("username")
end if
%>
<%
if Request.QueryString("author") <> "" then
author = Request.QueryString("author")
end if
%>
<% 
if Request.QueryString("keyword") <> "" then
keyword = Request.QueryString("keyword") 
style = Request.QueryString("style")
end if
%>
<%
Set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 
%>
<div align="center">
<table border="0" width="90%">
<tr>
      <td width="60%"> 
        <% Response.Write "<font face=arial size=1 color=white>The time now is: </font><font face=arial size=1 color=silver>"&now()&"</font>"%>
      </td>
      <td width="8%"><font face=arial size=2><b><a href='search.asp'>Search</a></b></font></td>
      <td width="11%"><font face=arial size=2><b><a href='default.asp'>Back To 
        Main</a></b></font></td>
      <td width="21%"> 
        <p align="right"><font face="arial" size="1"><b> 
<%
Set rsott = Server.CreateObject("ADODB.Recordset")
		if author <> "" then
		rsott.Open "SELECT * FROM board WHERE name LIKE '%" & Replace(author, "'", "''") & "%'",conn, 3
		end if
		if style = "post" AND author = "" then
        rsott.Open "SELECT * FROM board WHERE body LIKE '%" & Replace(keyword, "'", "''") & "%'",conn, 3
		elseif style ="topic" AND author = "" then
		rsott.Open "SELECT * FROM board WHERE ThreadP=0 AND topic LIKE '%" & Replace(keyword, "'", "''") & "%'",conn, 3
		end if
%>
<% response.write "<font face='arial' size=2 color=white><b>Searched Total Results: </b></font><b><font face='arial' size=2 color=silver>"& rsott.recordcount &"</font></b>" %>
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
Set rsearch = Server.CreateObject("ADODB.Recordset")

rsearch.CursorLocation = 3 'set cursorlocation to aduserclient 
rsearch.CacheSize = 10 'cache the number of record to display per page into cache
if author <> "" then
rsearch.Open "SELECT * FROM board WHERE name LIKE '%" & Replace(author, "'", "''") & "%' ORDER BY msgID DESC",conn, 3
end if
if author = "" AND style = "post" then
rsearch.Open "SELECT * FROM board WHERE body LIKE '%" & Replace(keyword, "'", "''") & "%' ORDER BY msgID DESC",conn, 3
elseif author = "" AND style = "topic" then
rsearch.Open "SELECT * FROM board WHERE ThreadP = 0 AND topic LIKE '%" & Replace(keyword, "'", "''") & "%' ORDER BY msgID DESC",conn, 3
end if


if rsearch.EOF then Response.Write "<p align=center><font face=arial size=2 color=white><b>Search Not Found</b></font>"
rsearch.PageSize = 10 'Set that each page will list 10 records

Dim TotalPages, TotalRecords
TotalPages = rsearch.PageCount 'Pagecount will count how many page will need if 10 record fill one page
TotalRecords = rsearch.RecordCount

if rsearch.RecordCount > 0 then
rsearch.AbsolutePage = currentpage 'the currentpage value will be the rs.Absolute value
end if

Dim count 'Counting variable for our recordset
count = 0 
Do While Not rsearch.EOF AND count < rsearch.PageSize
%>
<div align=center>
<table border=0 width=90% cellpadding=0 cellspacing=0>
 <tr>
  <td width=40%><p align=left><font face=arial size=2 color=white><b>
  <% if author = "" then %>
  <a href='sview.asp?msgID=<%=rsearch("msgID")%>&keyword=<%=keyword%>&style=<%=style%>'><%=rsearch("topic")%></a>
     <%
	 if style = "topic" then
	  Set rsnay = Server.CreateObject("ADODB.Recordset")
      rsnay.CacheSize = 10 'cache the number of record to display per page into cache
      rsnay.PageSize = 10 'Set that each page will list 10 records
      rsnay.Open "SELECT * FROM board WHERE ThreadID="&rsearch("msgID")&" ORDER BY msgID", conn, 3
      TotalPage = rsnay.PageCount 'Pagecount will count how many page will need if 10 record fill one page
      if TotalPage > 1 then
      Response.Write "<font face='arial' size='1' color=white>"
      Response.Write "<b><< Page:</b> "
      Response.Write "<select name='whatever' onChange='window.location=this.options[this.selectedIndex].value'>"
      Response.Write "<option value='sview.asp?seepage=1&msgID="&rsearch("msgID")&"&keyword="&keyword&"&style="&style&" SELECTED'> - 1 - </option>"
      for i = 1 to TotalPage
      Response.Write "<option value='sview.asp?seepage="&i&"&msgID="&rsearch("msgID")&"&keyword="&keyword&"&style="&style&"'> - "&i&" - </option>"
      next
      Response.Write "</select>"
      Response.Write " >>"
      end if
     end if
	 %>
  <% else %>
  <a href='sview.asp?msgID=<%=rsearch("msgID")%>&author=<%=author%>'><%=rsearch("topic")%></a>
  <% end if %>
  </b></font></td>
  <% msgID = rsearch("msgID") %>
	  <td width=17% align=center> 
        <% if instr(rsearch("email") & "" , "@") > 0 then %>
        <font face=arial size=2><b><a href='mailto:<%=trim(rsearch("email"))%>'><%=trim(rsearch("name"))%></a></b></font> 
        <% else %>
        <font face=arial size=2 color=white><b><%=rsearch("name")%></b></font>
        <% end if %>
      </td>
<%
Set rsr = Server.CreateObject("ADODB.Recordset")
rsr.Open "SELECT * FROM board WHERE ThreadID="&msgID&"AND ThreadP=1",conn, 3
%>
      <td width=10% align=center><font face=arial color=silver size=2><b><%=rsr.recordcount%></b></font></td>
      <td width=10% align=center><font face=arial color=silver size=2><b><%=rsearch("count")%></b></font></td>
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
      <td width="23%"><font face=arial color=white size=1><b><%=rsearch("date")%></b></font></td>
<% end if %>
</tr></table>
</div>
<%
count = count + 1
rsearch.MoveNext
Loop
Response.Write "<br>"
Response.Write "<div align=center>"
Response.Write "<table border=0 width=840>"
Response.Write "<tr>"
Response.Write "<td width=210><p align=left><font face=arial size=2><b>"
	  if rsearch.RecordCount > 0 then 
	  if currentpage > 1 then
	   if author = "" then
	   Response.Write "<a href='searching.asp?currentpage=1&keyword="&keyword&"&style="&style&"'>First</a>"
	   else
	   Response.Write "<a href='searching.asp?currentpage=1&author="&author&"'>First</a>"
	   end if
	  end if
	  end if
Response.Write "</b></font></td>"
Response.Write "<td width=210><p align=center><font face=arial size=2><b>"

	  if rsearch.RecordCount > 0 then
	  if currentpage > 1 then
	   if author = "" then
	   Response.Write "<a href='searching.asp?currentpage="& currentpage - 1 &"&keyword="&keyword&"&style="&style&"'><< Pervious</a>"
	   else
	   Response.Write "<a href='searching.asp?currentpage="& currentpage - 1 &"&author="&author&"'><< Pervious</a>"
	   end if
	  end if
	  end if
	
Response.Write "</b></font></td>"
Response.Write "<td width=210><p align=center><font face=arial size=2><b>"
	  
	  if rsearch.RecordCount > 0 then
	  if CInt(currentpage) <> CInt(TotalPages) then
	   if author = "" then
	   Response.Write "<a href='searching.asp?currentpage="& currentpage + 1 &"&keyword="&keyword&"&style="&style&"'>Next >></a>"
	   else
	   Response.Write "<a href='searching.asp?currentpage="& currentpage + 1 &"&author="&author&"'>Next >></a>"
	   end if
	  end if
	  end if
	 
	  Response.Write "</b></font></td>"
      Response.Write "<td width=210><p align=right><font face=arial size=2><b>"
	 
	  if rsearch.RecordCount > 0 then 
	  if CInt(currentpage) <> CInt(TotalPages) then
	   if author = "" then
	   Response.Write "<a href='searching.asp?currentpage="& TotalPages &"&keyword="&keyword&"&style="&style&"'>Last</a>"
	   else
	   Response.Write "<a href='searching.asp?currentpage="& TotalPages &"&author="&author&"'>Last</a>"
	   end if
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
 if author = "" then
 Response.Write "<option value='searching.asp?currentpage="&currentpage&"&keyword="&keyword&"&style="&style&" SELECTED'> - "&currentpage&" - </option>"
 else
 Response.Write "<option value='searching.asp?currentpage="&currentpage&"&author="&author&" SELECTED'> - "&currentpage&" - </option>"
 end if
for i = 1 to TotalPages
if cStr(currentpage) <> cStr(i) then
 if author = "" then
 Response.Write "<option value='searching.asp?currentpage="&i&"&keyword="&keyword&"&style="&style&"'> - "&i&" - </option>"
 else
 Response.Write "<option value='searching.asp?currentpage="&i&"&author="&author&"'> - "&i&" - </option>"
 end if
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

conn.Close
Set conn = Nothing
%>
<br>
<center><font face=arial size=1 color=white>Powered by AspBB v1.0 - Programmed by <a href='mailto:doma111@yahoo.com'>Johnny Yu</a></font></center>
</body>
</html>
