xl := ComObjActive("Excel.Application")
xl.workbooks("Exec_Review_ACE_110823_Test.xlsx").sheets(1).range("a1:z23").copy(xl.workbooks("Dest.xls").sheets(1).range("c3"))