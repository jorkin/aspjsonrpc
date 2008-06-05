<% option explicit %>
<!--#INCLUDE file="../includes/config.asp"-->
<!--#INCLUDE file="../src/StoredProcedureAccessor.asp"-->
<!--#INCLUDE file="../src/RsToJson.asp"-->
<%
dim spa, rs
set spa = new StoredProcedureAccessor

spa.setStoredProcedure("selectAll")

set rs = spa.executeQueryRecordset()

do while not rs.eof
	
	response.write rs.fields("name") & "<br />"

	rs.moveNext
loop

rs.close

set spa = nothing

%>
<br />
<br />
<hr />
<%
dim jsonConverter, jsonText
set jsonConverter = new RsToJson
set spa = new StoredProcedureAccessor

spa.setStoredProcedure("selectAll")

jsonText = spa.executeQueryJson(jsonConverter)

response.write jsonText

set rs = nothing

%>

<br />
<br />
<hr />
<%
set spa = new StoredProcedureAccessor

spa.setStoredProcedure("getById")
spa.addIntegerInput "id", 1

jsonText = spa.executeQueryJson(jsonConverter)

response.write jsonText


%>