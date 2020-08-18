<html>
<head>
<title> [ AspBB ] - Search </title>
<style type="text/css">
A:link, A:visited { text-decoration: none; color: 'silver' }
</style>
</head>

<body bgcolor="#006699">
<div align="center">
<table border=0 width=70%>
<tr>
<td>
<font face=arial color=white size=3><b>[ AspBB ] - Search</b></font><br>
<%
if Request.QueryString("error") = "" then
Response.Write " "
else
Response.Write "<center><font face=arial color=white size=2><b>You have to enter in one field</b></font></center>" 
end if
%>
<form name=searchy method=post action=searching.asp>
<center>
<table border=0 width=80%><tr><td valign=top>
<font face=arial color=white size=2><b>Search by Keywords</b></font><br>
<input type=text name=tosearch size=40><br>
<font face=arial color=white size=2>
<b>Search For:  
By Posts<input type="radio" name=style value=post CHECKED>
By Topics<input type=radio name=style value=topic>
</b>
</font></td>
<td valign=top><font face=arial color=white size=2><b>Search by Username</b></font><br>
<input type=text name=username size=25>
</td></tr></table>

</center>
<center>
<table border=0 width=50%>
<tr><td valign=right><p align=right><font face=arial size=1><a href="javascript:history.go(-1);" onMouseOver="window.status='Return Back';return true;" onMouseOut="window.status='';">Return Back</a></font></td></tr></table>
</center>
<center><input type=Submit Value=Search></center>
</form>
</td>
</tr>
</table>
</div>
<br>
<center><font face=arial size=1 color=white>Powered by AspBB v1.0 - Programmed by <a href='mailto:doma111@yahoo.com'>Johnny Yu</a></font></center>
</body>
</html>
