<%
class StoredProcedureAccessorTest
	private spa
	private num
	
	public function testCaseNames()
		testCaseNames = Array("testValidators", "testExecuteQueryRecordset", "testExecuteUpdate")
	end function

	private sub class_initialize
		num = 0
	end sub
	
	public sub setUp()
		set spa = new StoredProcedureAccessor
	end sub

	public sub tearDown()
		set spa = nothing
	end sub
	
	public sub testExecuteQueryRecordset(tester)
		dim rs
		spa.setStoredProcedure("selectAll")
		set rs = spa.executeQueryRecordset()
		tester.assert rs.EOF = false, "should be open recordset"
		
		do while not rs.EOF
			tester.assertEquals 1, rs.fields("id"), "first id should be 1"
			exit do
		loop
		
		rs.close()
	end sub

	public sub testExecuteUpdate(tester)
		dim rs
		dim name
		name = "TEST NAME: " & Date
		spa.setStoredProcedure("addRecord")
		spa.addStringInput "name", name, 50
		spa.addIntegerInput "value", 100
		spa.executeUpdate()
		
		set spa = nothing
		set spa = new StoredProcedureAccessor
		spa.setStoredProcedure("getLastAddedRow")
		set rs = spa.executeQueryRecordset()
		
		tester.assertEquals name, rs.fields("name"), "should find the added record"
		tester.assertEquals 100, rs.fields("value"), "should find the added record"
		
		rs.close()
		
		set spa = nothing
		set spa = new StoredProcedureAccessor
		spa.setStoredProcedure("deleteAddedRows")
		spa.executeUpdate()
		
		set spa = nothing
		set spa = new StoredProcedureAccessor
		spa.setStoredProcedure("getLastAddedRow")
		set rs = spa.executeQueryRecordset()
		
		tester.assertEquals "ghi", rs.fields("name"), "should find the test record 3"
		tester.assertEquals 3, rs.fields("value"), "should find the test record 3"
		
		rs.close()
		
		
	end sub
	
	
	public sub testValidators(tester)
		
		spa.addValidator "name", "\w{4,5}"
		
		spa.addStringInput "name", "abc", 50
		spa.addIntegerInput "value", 100
		
		dim isvalid
		isvalid = spa.validate()
		tester.assert  (not isvalid), "validation should fail"
		
	
		
		set spa = nothing
		set spa = new StoredProcedureAccessor
		spa.addValidator "name", "\w{4,5}"
		
		spa.addStringInput "name", "abcd", 50
		spa.addIntegerInput "value", 100
		isvalid = spa.validate
		tester.assert isvalid, "validation should succeeed"
	end sub
	
end class


%>