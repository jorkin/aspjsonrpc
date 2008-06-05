<!--#include file="jsonParser.asp"-->
<%
class RpcRequest
	public id
	public method
	public params
	
	public sub initialize(value)
		dim json
		set json = new JsonParser
		
		json.json = value
		json.parse
		
		id = json.dictionary.item("id")
		method = json.dictionary.item("method")
		set params = json.dictionary.item("params")
	end sub
	
end class



'This is a simple Json-RPC Server for vbscript
' provide a delegateClass
'the delegate class must implement
'the following interface:
'	method( param0, param1, ...)
'the params will either be strings or dictionaries (you are responsible for knowing which and casting them to the types you really need!)
'if an error occurs,  you should call server.writeErrorResponse, or you can use err.raise, which will be handled gracefully
'(server is set with a call to set delegateClass.server = me)
'id should return a valid id for the response, it should store the set id. if null, it should set it as appropriate
'version should return a version of the class implementation
'the method itself should return null if there was an error
'otherwise you can return a scalar, or a json object string, or alternatively, return nothing
'and just call server.writeSuccessfulResponse and do your own thing!
class JsonRpcServer
	
	private sub class_initialize
		Response.Buffer = true
		'Response.ContentType = "application/json-rpc"
	end sub
	
	private delegateClass
	
	public sub setDelegateClass(value)
		set delegateClass = value
		on error resume next 'if there is no property for server, the delegate class doesn't care about the server! (it can always user Raise to send error back)
			set value.server = me
		on error goto 0
	end sub
	
	public sub writeErrorResponse(name, code, message) 
		dim id, version
		version = "unknown"
		on error resume next 'optional for delegateClass to provide this
			version = delegateClass.version
		on error goto 0
		response.write "/*{ version:""" & version & """, error:{ name:""" & name & """, code:""" & code & """, message:""" & Server.URLEncode(message) & """ }}*/"
		response.flush
		response.end
	end sub
	
	private function binaryToString(binary)
		  'Antonin Foller, http://www.motobit.com
		  'Optimized version of a simple BinaryToString algorithm.
		  
		  Dim cl1, cl2, cl3, pl1, pl2, pl3
		  Dim L
		  cl1 = 1
		  cl2 = 1
		  cl3 = 1
		  L = LenB(Binary)
		  
		  Do While cl1<=L
		    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		    cl1 = cl1 + 1
		    cl3 = cl3 + 1
		    If cl3>300 Then
		      pl2 = pl2 & pl3
		      pl3 = ""
		      cl3 = 1
		      cl2 = cl2 + 1
		      If cl2>200 Then
		        pl1 = pl1 & pl2
		        pl2 = ""
		        cl2 = 1
		      End If
		    End If
		  Loop
		  binaryToString = pl1 & pl2 & pl3
	end function

	
	private function getRequest()
		dim requestMethod, rawRequest, bytes
		requestMethod = request.serverVariables("REQUEST_METHOD")
		if requestMethod = "GET" then
			rawRequest = request.QueryString
		else 
			bytes = request.binaryRead(request.totalBytes) ' the only way i could figure out how to get the raw form data
			rawRequest = binaryToString(bytes)
		end if
		
		if NOT rawRequest <> "" then
			writeErrorResponse "Empty call", 2, "The request was empty using: " + requestMethod
		end if
		
		rawRequest = URLDecode(rawRequest)
		
		set getRequest = new RpcRequest
		
		'on error resume next
			getRequest.initialize(rawRequest)
			
			if err.number <> 0 then
				writeErrorResponse "Parse error", err.number, "Could not parse request: " & rawRequest & " ASPError: " & err.description
			end if
	'	on error goto 0
	end function

	public sub writeSuccessfulResponse(result, id, version)
		if result <> "" then
			dim strResult
			strResult = CStr(result)
			if NOT left(strResult, 1) = "{" and NOT left(strResult, 1) = "[" then
				result = """" & result & """"
			end if
		end if
	
		if NOT version <> "" then
			version = "unknown"
		end if
	
		response.write 	"/*{version:""" & version & """, "
		if id <> "" then
			response.write "	id:" & id & ", "
		else 
			response.write "	id:null, "
		end if
		response.write	"	result:" & result & "" &_
						"}*/"
		response.flush
		response.end
	end sub
	
	public function run() 
		dim req
		set req = getRequest()
		response.write "id: " & req.id & "<br />"
		response.write "method: " & req.method & "<br />"

		dim i
		for i=0 to req.params.count-1
			response.write "param " & i & ": " & req.params.item("" & i) & "<br />"
		next
		dim useReqId
		useReqId = false
		on error resume next 'optional. 
			delegateClass.id = id
			if err.number <> 0 then
				useReqId = true
			end if
		on error goto 0
		
		dim calledMethod
		calledMethod = req.method & "("
		dim params
		params = ""
		for i=0 to req.params.count-1
			if params <> "" then
				params = params & ", "
			end if
			params = params & "req.params.item(CStr(" & i & "))"
		next
		
		calledMethod = calledMethod & params & ")"
		
		dim result
		on error resume next
		result = eval("delegateClass." & calledMethod)
		if err.number <> 0 then
			writeErrorResponse "Internal Error", err.number, err.description
		end if
		on error goto 0
		
		dim id, version
		id = ""
		version = ""
		on error resume next 'optional for delegateClass to provide these
			id = delegateClass.id
			version = delegateClass.version
		on error goto 0
		
		if useReqId then
			id = req.id
		end if
		
		writeSuccessfulResponse result, id, version
		
	end function

end class

%>

