<%@ Language=VBScript %>
<%
Option Explicit
Dim message
Dim cnnTest
Dim connString
Dim rs
Dim i

%><!-- #INCLUDE VIRTUAL = "Header.asp" -->
<link type="text/css" rel="stylesheet" href="bootstrap.min.css" />
<link type="text/css" href="main.css" rel="stylesheet" />

<div class="container-half container">

<h3>Buyers Info : </h3>


<%

on error resume next
    
Set cnnTest = Server.CreateObject("ADODB.Connection")
cnnTest.ConnectionString = Application("connectionString")
cnnTest.Open

set rs = Server.CreateObject("ADODB.recordset")
rs.Open "select * from dbo.Buyer", cnnTest

    If rs.EOF Then %>

    <div class="card p-3 text-warning bg-dark bg-gradient-light">
        <h3>There are no buyers</h3>
    </div>
        
    
<%
    Else


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

End If

   

set rs = Nothing
Set cnnTest = Nothing
%>

    <div style="text-align: right; padding-top: 20px">
        <button onClick="addBuyer()" class="btn btn-primary">Add Buyer</button>
    </div>
</div>

<script type="text/javascript">
    function addBuyer() {
        console.log('add buyer clicked');
        document.location = 'addBuyer.asp'
    }
</script>
