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
		private invalidFields
		private validators
		private forceValidation_
		
		'if true, are rror will be thrown if invalid, default is true
		public property let forceValidation(value)
			forceValidation_ = value
		end property
		
		public sub setConnectionString(str)
			connectionString = str
		end sub
		
		private sub class_initialize
			forceValidation_ = true
			set cmd = server.createObject("ADODB.Command")
			cmd.commandtimeout = 10
			set rs = server.createObject("ADODB.Recordset")
			rs.cursorType = adOpenStatic
			rs.lockType = adLockReadOnly
			rs.cursorLocation = adUseClient
			set validators = server.createObject("Scripting.Dictionary")
		end sub
		
		private sub class_terminate
			set rs = nothing
			set cmd = nothing
			set validators = nothing
		end sub
		
		private sub addRegexValidator(fieldName, pattern)
			dim re
			set re = new RegexValidator
			re.setPattern(pattern)
			
			validators.add fieldName, ""
			set validators.item(fieldName) = re
		end sub
		
		public sub addValidator(fieldName, validator)
		
			if IsObject(validator) then
				validators.add fieldName, ""
				set validators.item(fieldName) = validator
			else
				addRegexValidator fieldName, validator
			end if
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
			if forceValidation_ then
				if not validate() then
					err.raise 99, "Invalid Field Value -- Validator Failed", invalidFields
				end if
			end if
			cmd.CommandType = adCmdStoredProc
			if connectionString <> "" then
				cmd.activeConnection = connectionString
			else 
				cmd.activeConnection = DB_CONNECTION_STRING
			end if
			cmd.CommandText = storedProcedure
		end sub
	
		public function getInvalidFields()
			getInvalidFields = invalidFields
		end function
	
		public function validate()
			invalidFields = ""
			dim param
			
			for each param in cmd.parameters
				if isobject(validators.item(param.name)) then
					dim validator
					set validator = validators.item(param.name)
					dim value
					value = param.value
					if not validator.validate(value) then
					
						if invalidFields <> "" then
							invalidFields = invalidFields & ", "
						end if
						invalidFields = invalidFields & param.name & ":" & value
					end if
				end if
			next
			if invalidFields <> "" then
				validate = false
			else 
				validate = true
			end if
		end function
	
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
	

	'a validator is any class that has a validate method that
	'returns true or fals.
	'this is a simple regex based validator
	class RegexValidator
		private re
		
		public property let globalMatch(val)
			re.global = val
		end property
		
		public property let ignoreCase(val)
			re.ignoreCase = val
		end property
		
		private sub class_initialize
			set re = new RegExp
		end sub
		
		private sub class_terminate
			set re = nothing
		end sub
		
		public sub setPattern(str)
			re.pattern = str
		end sub
		
		public function validate(expression)
			validate = re.test(expression)
		end function
		
	end class
%>