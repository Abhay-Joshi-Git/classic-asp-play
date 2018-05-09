<%@ Language=VBScript %>
<%
Option Explicit
Dim message
Dim cnnTest
Dim connString
Dim rs
Dim i

%>
<link type="text/css" href="main.css" rel="stylesheet" />
<h1>Buyers Info : </h1>


<%

on error resume next
    
connString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=admin@123;Initial Catalog=test;Data Source=ultp_758\sqlexpress;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096"
Set cnnTest = Server.CreateObject("ADODB.Connection")
cnnTest.ConnectionString = connString
cnnTest.Open

set rs = Server.CreateObject("ADODB.recordset")
rs.Open "select * from dbo.Buyer", cnnTest

    If rs.EOF Then
        response.Write("no records found")
    End If

    response.Write("<br />")


%>
  
<div style="width:100%">
    <table id="buyers">
        <tr>
            <%
                For i = 0 To (rs.Fields.Count-1)
                    response.Write("<th>" & UCase(rs.Fields.Item(i).Name) & "</th>")
                Next
            %>
        </tr>
            <%
                do until rs.EOF
                    response.Write("<tr>")
                    For i = 0 To (rs.Fields.Count-1)
                        response.write("<td>" & rs.Fields.Item(i).Value & "</td>")
                    Next
                    response.Write("</tr>")
                    rs.MoveNext
                loop
            %>
    </table>
</div>

<%
do until rs.EOF
    For i = 0 To (rs.Fields.Count-1)
        response.Write(rs.Fields.Item(i).Name)    
        response.write(" = ")
        response.write(rs.Fields.Item(i).Value)
        response.Write("<br />")
    Next
    response.Write("<br />")
    rs.MoveNext
loop
    
%>
