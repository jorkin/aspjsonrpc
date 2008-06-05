<%
class StoredProcedureAccessorTest
	private spa

	public function testCaseNames()
		testCaseNames = Array("testExecuteQueryRecordset", "testExecuteUpdate")
	end function

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
	
end class


%>