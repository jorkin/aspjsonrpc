<!--#INCLUDE file="adovbs.inc"-->
<%
dim DB_CONNECTION_STRING
DB_CONNECTION_STRING="Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
        "Data Source=" &  Server.MapPath("/AspJsonUtils/tests/db/test.mdb") 
%>