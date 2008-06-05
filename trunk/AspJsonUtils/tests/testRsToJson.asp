<%
class RsToJsonTest
	private spa
	private jsonConverter
	
	public function testCaseNames()
		testCaseNames = Array("testRsToJsonSimple", "testArray", "testNotArray")
	end function

	public sub setUp()
		set spa = new StoredProcedureAccessor
		set jsonConverter = new RsToJson
	end sub

	public sub tearDown()
		set spa = nothing
		set jsonConverter = nothing
	end sub
	
	public sub testRsToJsonSimple(tester)
		spa.setStoredProcedure("selectAll")
		
		dim result
		result = spa.executeQueryJson(jsonConverter)
		
		tester.assertEquals "[{""id"":""1"", ""name"":""abc"", ""value"":""1""}, {""id"":""2"", ""name"":""def"", ""value"":""2""}, {""id"":""3"", ""name"":""ghi"", ""value"":""3""}]", result, "result not valid"
	end sub
	
	public sub testArray(tester)
		spa.setStoredProcedure("getById")
		spa.addIntegerInput "id", 1
		
		jsonConverter.alwaysAsArray = true
		
		dim result
		result = spa.executeQueryJson(jsonConverter)
		
		tester.assertEquals left(result, 1), "[", "should always return an array"
		tester.assertEquals "[{""id"":""1"", ""name"":""abc"", ""value"":""1""}]", result, "result not valid"
	end sub

	public sub testNotArray(tester)
		spa.setStoredProcedure("getById")
		spa.addIntegerInput "id", 1		
		dim result
		result = spa.executeQueryJson(jsonConverter)
		
		tester.assertEquals left(result, 1), "{", "should return an object for one return"
		tester.assertEquals "{""id"":""1"", ""name"":""abc"", ""value"":""1""}", result, "result not valid"
	end sub
	
end class


%>