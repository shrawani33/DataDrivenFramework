Datatable.AddSheet "Action1" @@ hightlight id_;_134710_;_script infofile_;_ZIP::ssf17.xml_;_
Datatable.ImportSheet "C:\Temp\Keyword Driver framework\Organizer\Organizerr(1).xlsx",1,"Action1"
mrowcount=Datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For i = 1 To mrowcount  Step 1
Datatable.SetCurrentRow(i)
	Modexe=Datatable("ModuleExe","Action1")
	msgbox Modexe
	If Modexe= "Y" Then
		Modid=Datatable("ModuleID","Action1")
		msgbox Modid
	trowcount=Datatable.GetSheet("Action2").GetRowCount
	msgbox trowcount
	
For j = 1 To trowcount Step 1
Datatable.SetCurrentRow(j)
	If Modid =Datatable("ModuleID","Action2") and Datatable("TestCaseExecution","Action2")="Y" then
	testcaseid=Datatable("TestCaseID","Action2")
	msgbox testcaseid
	tsrowcount=Datatable.GetSheet("Action3").GetRowCount 
msgbox tsrowcount
For k = 1 To tsrowcount  Step 1
Datatable.SetCurrentRow(k)
If testcaseid=Datatable("TestCaseId","Action3") Then
	keyword=Datatable("Keyword","Action3")
	msgbox Keyword
	Select Case (keyword)
		Case "In"
		Call Login()
		Case "ca"
		Call CloseApp()
		Case "oo"
		Call OpenOrder()
		Case "uo"
		Call UpdateOrder()
	End Select
	
End If
	
Next
	
	End If
Next
	End If
	Next
	




