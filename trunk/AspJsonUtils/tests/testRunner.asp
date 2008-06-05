<%
Option Explicit
%>
<!--#INCLUDE file="../includes/config.asp"-->
<!-- #include file="../asp-lib/aspunit/ASPUnitRunner.asp"-->
<!--#INCLUDE file="../src/StoredProcedureAccessor.asp"-->
<!--#INCLUDE file="../src/RsToJson.asp"-->
<!-- #include file="testStoredProcedureAccessor.asp"-->
<!-- #include file="testRsToJson.asp"-->
<%
	Dim oRunner
	Set oRunner = New UnitRunner
	oRunner.pathToThis = "/AspJsonUtils/asp-lib/aspunit/"
	oRunner.AddTestContainer New StoredProcedureAccessorTest
	oRunner.AddTestContainer New RsToJsonTest
	oRunner.Display()
%>