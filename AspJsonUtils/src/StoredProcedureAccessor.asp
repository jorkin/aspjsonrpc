<%
' A class for simplifying ADO via stored procedures
' You must have a global DB_CONNECTION_STRING, or call setConnectionString
' You must call setStoredProcedure to set the name of the stored procedure
' then, just add parameters:
' addStringInput(name, value, maxLength)
' addStringOutput(name, maxLength)
' addIntegerInput(name, value)
' addIntegerOutput(name)
' addIntegerReturnValue(name)
' addStringReturnValue(name, maxLength)
' or for more complicated input/ouput use
' addInputParameter (name, paramType, length, value)  or addOutputParameter(name, paramType, length, retVal)
' you can also get the raw command object if you so desire via getCommandObject
'
' finally call 
' executeUpdate -- which returns a dictionary of parameters for you to look for any return values/output values
' executeQueryRecordset -- which returns a recordset
' executeQueryJson -- to which you pass a RsToJson converter that you have already configured, and you get back the json string
'

	class StoredProcedureAccessor
		private storedProcedure
		private cmd
		private rs
		private connectionString
		
		public sub setConnectionString(str)
			connectionString = str
		end sub
		
		private sub class_initialize
			set cmd = server.createObject("ADODB.Command")
			set rs = server.createObject("ADODB.Recordset")
		end sub
		
		private sub class_terminate
			set cmd = nothing
			set rs = nothing
		end sub
		
		public sub addInputParameter(name, paramType, length, value)
			if not isNull(length) then
				cmd.parameters.append cmd.createParameter(name, paramType, adParamInput, length,  value)
			else
				cmd.parameters.append cmd.createParameter(name, paramType, adParamInput, ,  value)
			end if
		end sub
		
		public sub addOutputParameter(name, paramType, length, retVal)
			dim direction
			if retVal = true then
				direction = adParamReturnValue
			else
				direction = adParamOutput
			end if
			
			if not isNull(length) then
				cmd.parameters.append cmd.createParameter(name, paramType, direction, length)
			else
				cmd.parameters.append cmd.createParameter(name, paramType, direction)
			end if
		end sub
		
		public sub setStoredProcedure(sp)
			storedProcedure = sp
		end sub
		
		public sub addStringInput(name, value, maxLength)
			if maxLength > 255 then
				addInputParameter name, adLongVarChar, maxLength, value
			else
				addInputParameter name, adVarChar, maxLength, value
			end if
		end sub
		
		public sub addIntegerInput(name, value)
			addInputParameter name, adInteger, null, value
		end sub
		
		public sub addIntegerOutput(name)
			addOutputParameter name, adInteger, null, false
		end sub
		
		public sub addStringOutput(name, maxLength)
			addOutputParameter name, adVarChar, maxLength, false
		end sub
		
		public sub addIntegerReturnValue(name)
			addOutputParameter name, adInteger, null, true
		end sub
		
		public sub addStringReturnValue(name, maxLength)
			addOutputParameter name, adVarChar, maxLength, true
		end sub
		
		public function getCommandObject()
			set getCommandObject = cmd
		end function
	
		private sub prepare()
			cmd.CommandType = adCmdStoredProc
			if connectionString <> "" then
				cmd.activeConnection = connectionString
			else 
				cmd.activeConnection = DB_CONNECTION_STRING
			end if
			cmd.CommandText = storedProcedure
		end sub
	
		'execute an update and return the parameters dicitonary
		public function executeUpdate()
			prepare()
			cmd.execute , , adExecuteNoRecords
			set executeUpdate = cmd.parameters
		end function
	
		'Execute a query and return the recordset
		'you must be sure to close the recordset
		public function executeQueryRecordset()
			prepare()
			rs.open cmd
			set executeQueryRecordset = rs
		end function
		
		
		public function executeQueryJson(jsonConverter)
			executeQueryRecordset()
			jsonConverter.setRecordset(rs)
			executeQueryJson = jsonConverter.convert
			rs.close()
		end function
		
	end class
%>