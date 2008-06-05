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

class numvalid 
	public function validate(exp)
		if exp = 1 then
			validate = true
		else
			validate = false
		end if
	end function
end class
spa.addValidator "id", "3"

spa.addIntegerInput "id", 2

spa.validate()
response.write spa.getInvalidFields

jsonText = spa.executeQueryJson(jsonConverter)

response.write jsonText


%>