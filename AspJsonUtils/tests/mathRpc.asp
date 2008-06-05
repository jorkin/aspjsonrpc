<%@ Language=VBScript %>	
<% Option Explicit %> 
<!--#include file="jsonRpcServer.asp"-->
<%
	class MathClass
		function add(a, b)
			add = CInt(a) + CInt(b)
		end function
	end class
	
	dim rpc
	set rpc = new JsonRpcServer
	rpc.setDelegateClass(new MathClass)
	rpc.run
%>