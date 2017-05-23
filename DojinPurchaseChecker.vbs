'******************************
'* Input / Output file
'******************************
'Purchase checklist
Const PURCHASE_NO_FILE  = "Purchase_NO_list.txt"
'Purchasable list
Const PURCHASE_YES_FILE = "Purchase_YES_list.txt"

'******************************
'* 'Checkable doujin site
'******************************
Const SITE_URL_TORA = "http://www.toranoana.jp/"

Const SITE_NO_OTHER = 0
Const SITE_NO_TORA  = 1


'******************************
'* Process execution
'******************************
main



'******************************
'* Main
'******************************
Sub main()

	'******************************
	'* Confirmation
	'******************************
	Dim Button_no
	button_no = MsgBox( "Start checking available for purchase", vbOKCancel,"Check available for purchase" )
	
	If button_no = vbCancel then
		Exit Sub
	End If
	

	'******************************
	'* Preparation
	'******************************
	Dim old_no_list(), new_no_list(), new_yes_list()

	'Get number of lines of input file
	Get_File_NumberOfLines PURCHASE_NO_FILE, list_size

	ReDim Preserve old_no_list(list_size), new_no_list(list_size), new_yes_list(list_size)


	'******************************
	'* Load purchase checklist file
	'******************************
	Read_File PURCHASE_NO_FILE, old_no_list


	'******************************
	'* Sales check
	'******************************
	Dim new_no_ct, new_yes_ct
	new_no_ct  = 0
	new_yes_ct = 0

	For ii = 0 To UBound(old_no_list)

		Check_Saihan old_no_list(ii), result
		
		If result = 1 then 'Available
			new_yes_list(new_yes_ct) = old_no_list(ii)
			new_yes_ct = new_yes_ct + 1

		Else	'No sale
			new_no_list(new_no_ct) = old_no_list(ii)
			new_no_ct = new_no_ct + 1
		End If
	Next


	'******************************
	'* File output
	'******************************

	If new_yes_ct <> 0 then
		'Update purchase checklist file
		Write_File PURCHASE_NO_FILE, new_no_list
		
		'Update purchasable list
		Add_File PURCHASE_YES_FILE, new_yes_list
		
		MsgBox( "Found new purchasable items" )
	End If

End Sub



'******************************
'* Get number of file lines
'******************************
Sub Get_File_NumberOfLines(file_name, n_lines)

	n_lines = 0

	Dim myfso, myfile
	Set myfso = CreateObject("Scripting.FileSystemObject")
	Set myfile = myfso.OpenTextFile(file_name, 8, False)	'Append mode

	'Get line number of file pointer (end)
	n_lines = myfile.Line

	myfile.Close
	Set myfile = Nothing
	Set myfso = Nothing

End Sub



'******************************
'* Read file
'******************************
Sub Read_File(file_name, in_list)

	'Open the file read-only
	Dim myfso, myfile, line
	Set myfso = CreateObject("Scripting.FileSystemObject")
	Set myfile = myfso.OpenTextFile(file_name, 1, False)

	'Read one line at a time
	ii = 0
	Do Until myfile.AtEndOfStream
		in_list(ii) = myfile.ReadLine
		ii = ii + 1
	Loop

	'Close the file
	myfile.Close
	Set myfile = Nothing
	Set myfso = Nothing

End Sub



'******************************
'* Write to file (new)
'******************************
Sub Write_File(file_name, out_list)

	'Open file for writing only
	Dim myfso, myfile, line
	Set myfso = CreateObject("Scripting.FileSystemObject")
	Set myfile = myfso.OpenTextFile(file_name, 2, False)	'Overwrite


	'Write row by row
	For ii = 0 To UBound(out_list)
		If out_list(ii) <> "" then 
			myfile.WriteLine(out_list(ii))
		End If
	Next

	'Close the file
	myfile.Close
	Set myfile = Nothing
	Set myfso = Nothing
End Sub



'******************************
'* Write to file (append)
'******************************
Sub Add_File(file_name, out_list)

	'Open file for writing only
	Dim myfso, myfile, line
	Set myfso = CreateObject("Scripting.FileSystemObject")
	Set myfile = myfso.OpenTextFile(file_name, 8, True)	'append


	'Write row by row
	For ii = 0 To UBound(out_list)
		If out_list(ii) <> "" then 
			myfile.WriteLine(out_list(ii))
		End If
	Next

	'Close the file
	myfile.Close
	Set myfile = Nothing
	Set myfso = Nothing
End Sub


'******************************
'* Sales check
'******************************
Sub Check_Saihan( url, result )

	result = 0	'No sale

	'******************************
	'* URL check
	'******************************
	Dim site_no
	Check_URL url, site_no

	if site_no = SITE_NO_OTHER then 
		' Unsupported site
		Exit Sub
	End If


	'******************************
	'* Open IE
	'******************************
	Dim ie
	Set ie = CreateObject("InternetExplorer.Application")
	ie.Visible = True
	ie.Navigate url
	waitIE ie 	'Waiting for activation


	'******************************
	'* Breakthrough of age confirmation
	'******************************
	Age_Limit_Break ie , site_no


	'******************************
	'* Check if selling
	'******************************
	Check_Sale ie , site_no , sale_result
	If sale_result = 1 then
		result = 1	'Available
	End If


	'******************************
	'* Close IE
	'******************************
	ie.Quit

End Sub



'******************************
'* Check URL
'******************************
Sub Check_URL( url, site_no )

	site_no = SITE_NO_OTHER

	if Instr( url, SITE_URL_TORA ) > 0 then
		'Toranoana
		site_no = SITE_NO_TORA
	End If

End Sub	



'******************************
'* Breakthrough of age confirmation
'******************************
Sub Age_Limit_Break( ie, site_no )

	If site_no = SITE_NO_TORA then
		'Toranoana

		Set elem_adultcheck = ie.Document.getElementsByClassName("AdultCheck")
		If elem_adultcheck.Length = 0 then
			'It is not age confirmation page
			Exit Sub
		End If

		Set elem_input = ie.Document.getElementsByTagName("input")
		For ii = 0 To elem_input.Length - 1
			if Instr( elem_input(ii).Value,"‚Í‚¢" ) > 0 then
				elem_input(ii).Click
				Exit For
			End If
		Next
		waitIE ie

	End If

End Sub



'******************************
'* Check if you sell
'******************************
Sub Check_Sale( ie , site_no , sale_result)
	
	sale_result = 0

	If site_no = SITE_NO_TORA then
		'Toranoana

		Check_Sale_Tora ie, sale_result
		
	End If
End Sub



'******************************
'* Sales check of Toranoana
'******************************
Sub Check_Sale_Tora( ie, sale_result)

	sale_result = 0

	'******************************
	'* Sales check
	'******************************
	Set elem_cart_submit = ie.Document.getElementsByClassName("cart_submit")
	If elem_cart_submit.Length > 0 then
		'With order button
		sale_result = 1
		Exit Sub
	End If

	'******************************
	'* Search for resale button
	'******************************
	Set elem_input = ie.Document.getElementsByTagName("input")

	ii = 0
	Do While ii < elem_input.Length
		If elem_input(ii).name = "revoteitem" then
			'Button discovery
			Exit do
		End If
		ii = ii + 1
	Loop

	If ii >= elem_input.Length then 
		Exit Sub
	End If
	If elem_input(ii).disabled then
		'I can not push resale button
		Exit Sub
	End If

	'******************************
	'* Press the resale request button
	'******************************
	Dim myshell , swc, pop_ie
	Set myshell = CreateObject("Shell.Application")
	swc = myshell.Windows.Count

	elem_input(ii).Click

	if  myshell.Windows.Count > swc then
		Set pop_ie = myShell.Windows(swc -1 +1)
		waitIE pop_ie

		'close the window
		pop_ie.Quit
	End If
End Sub



'******************************
'* IE busy wait
'******************************
Sub waitIE(ie)

    Do While ie.Busy = True Or ie.readystate <> 4
        WScript.Sleep 100
    Loop

End Sub


