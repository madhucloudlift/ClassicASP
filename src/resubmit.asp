<%
response.buffer = true

id = Request.Form("msgID")
topic = Replace(Server.HTMLEncode(Request.Form("topic")),"'","''")
name = Replace(Server.HTMLEncode(Request.Form("name")),"'","''")
email = Replace(Server.HTMLEncode(Request.Form("email")),"'","''")
http = Replace(Server.HTMLEncode(Request.Form("http")),"'", "''")
body =  Request.Form("body")
body = Replace(body , "'", "''")
body = Replace(body, vbcrlf, "<br>")

if name = "" OR body = "" then
response.redirect("repost.asp?msgID="&id&"&error=1")
end if
if instr(LCase(body), "<script") <> 0 then 
response.redirect("repost.asp?msgID="&id&"&error=2")
end if
if email <> "" then 
  if instr(email, "@") = 0 OR instr(email, ".") = 0 then
  response.redirect("repost.asp?msgID="&id&"&error=3")
  end if
end if
if http <> "" then
  if instr(http, "http://") = 0 OR instr(http, ".") = 0 then
  response.redirect("repost.asp?msgID="&id&"&error=4")
  end if
end if

set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 

sqlstring = "INSERT INTO board ( [date], topic, name, email, http, body, threadID, threadP )  " &_ 
 "values ( # " &now()& " #, ' " &topic& " ' ,  ' " &name& " ' ,  ' " &email& " ', ' " &http& " ' , ' " &body& " ', ' " &id& " ', 1 )" 
conn.Execute(sqlstring)
conn.Close 
set conn = Nothing
Session.Abandon
Response.redirect "default.asp"
%>