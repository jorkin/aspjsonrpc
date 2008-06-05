<%
'conver a recordset to json
'call setRecordset
'call addFieldName to add any fieldName conversion you would like to make (i.e. change the SQL field name to tje json object field name)
'call addJsonConverter to add a converter to any field
'set useAllRecordsetFields = true/false. if false, then only fields add via addfieldName are used in the resulting json, otherwise, everything is.
'call convert() to the resulting string -- this returns an array of items, or just the item if only one item is in the recordset
'a converter has the method "convert" that accepts the fieldName,  and recordset as an argument and returns a string -- quoted if a scalar
'at the least you should use the URLo_ENCODEo_CONVERTER (addJsonConverter(rsFieldName, URLo_ENCODEo_CONVERTER))
'(note that this will AUTOMATICALLY escape double quotes)
'if you expect quotes to break the json
	class RsToJson
		private rs
		private fieldNamesDictionary
		private o_useAllRsFields
		private converterDictionary
		private o_alwaysAsArray
		
		private sub class_initialize
			set fieldNamesDictionary = Server.CreateObject("Scripting.Dictionary")
			set converterDictionary = Server.CreateObject("Scripting.Dictionary")
			o_useAllRsFields = true
			o_alwaysAsArray = false
		end sub
		
		public property let useAllRecordsetFields(value)
			o_usAllRsFields = value
		end property
		
		public property let alwaysAsArray(value)
			o_alwaysAsArray = value
		end property
		
		public sub setRecordSet(recordset)
			set rs = recordset
		end sub
	
		public sub addFieldName(rsFieldName, jsonFieldName)
			fieldNamesDictionary.add rsFieldName, jsonFieldName
		end sub
		
		public sub addJsonConverter(rsFieldName, converter)
			converterDictionary.add rsFieldName, ""
			set converterDictionary.item(rsFieldName) = converter
		end sub
	
		public function convert()
			dim response, item, fieldName, field, jsonName, jsonConverter, value
			response = ""
			dim num
			num = 0
			if not rs.eof then	
				do while not rs.eof
					num = num + 1
					if len(response) > 0 then
						response = response & ", "
					end if
					
					item = "{"
					
					for each field in rs.fields
						fieldName = field.name
						jsonName = fieldNamesDictionary.item(fieldName)
						
						if NOT jsonName <> "" and o_useAllRsFields then
							jsonName = fieldName
						end if
						
						if isObject(converterDictionary.item(fieldName)) then
						
							set jsonConverter = converterDictionary.item(fieldName)
							value = jsonCoverter.convert(fieldName, rs)
						else
							value = field.value
							if InStr(value, """") then
								value = replace(value, """", "\""") 'escape quotes!
							end if
							value = """" & value & """"
						end if
						
						if not value <> "" then
							value = """"""
						end if
						
						if len(item) > 1 then
							item = item & ", "
						end if
						
						item = item & """" & jsonName & """:" & value 'note that value is ALREADY quoted if it needs to be!
						
					next
					
					item = item & "}"
					response = response & item
					rs.moveNext
				loop
			end if
			
			if num > 1 or o_alwaysAsArray then
				response = "[" & response & "]"
			end if
			convert = response
		end function
	end class
	
	class UrlEncodeConverter
		public function convert(fieldName, rs)
			convert = server.urlencode(rs.fields(fieldName))
			if convert <> "" then
				convert = """" & convert & """"
			end if
		end function
	end class
	
	dim URL_ENCODE_CONVERTER
	set URL_ENCODE_CONVERTER = new UrlEncodeConverter
%>
