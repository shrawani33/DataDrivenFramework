Dim objuft

Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.Open("C:\Temp\Keyword Driver framework\Driver")
objuft.Test.run
objuft.Test.close
objuft.quit
Set objuft=nothing