#CS----
	A group of functions to perform simple ADOdb operations upon AcuCOBOL data files.
#CE----


#include-once
#include <Array.au3>
#include ".\Includes\SimpleLogger.au3"

Opt("TrayIconDebug", 1)        ;0=no info, 1=debug line info
	
	
#cs----  _GetData  ----
	Description:  
		Get data values from COBOL structured files using ODBC. Table names within SQL statements must match the .xfd name of the COBOL file 
		you wish to access.
		
		
	_GetData ( Data Source Name, SQL Select command[, Display Results ])
	
	
	Parameters:
		Data Source Name (DSN) :   	The string value representing the reference to a relational database or otherwise structured data source.
		SQL Command :              	The string literal of the SQL statement to be used to retrieve a record set object from the data source connection.
		Display Results : 			A numeric value indicating that calling function or program wishes to see the results of the SQL command.
									0 = no display (assumed if none)
									1 = yes display
									
	Returns:
		Success:	A 2 dimensional array ( Arr[y][x] ) where sub-x represents columns and sub-y represents rows (row 0 being field names)  
		Fail:
			-1 	= 	COM error handled (example: DSN not found)
			1	=	Missing Required Parameters
			
	Remarks:
		Table names within SQL statements must match the .xfd name of the COBOL file you wish to access.
		
		The function used for logging should be updated or removed if the calling program has functional logging processes.
		
#ce----	
Func _GetData($DSN, $SQL, $Disp_Param = 0)
	$file = FileOpen("GetData_Results.txt",  2)
	
	; define an error handling object and something to check
	Global $getError = ObjEvent("AutoIt.Error", "ErrFunc")
	Global $Err = 0
	
	; Check validity of parameters / set defaults
	If ($DSN = '') or ($SQL = '') Then
		ConsoleWrite("Either no DSN, or no SQL was supplied")
		Return 1
	ElseIf $Disp_Param = '' Then
		$Disp_Param = 0 
	EndIf
	
	; Is $SQL a Select statement?
	If StringMid($SQL, 1, 6) <> "Select" Then
		ConsoleWrite("SQL not a Select statement.")
		Return 1
	EndIf
	
	; create the ADOdb connection and recordset objects
	Global $ODBC = ObjCreate("ADODB.connection")	
	Global $RecSelect = ObjCreate ("ADODB.Recordset") ; Create a Record Set to handles SQL Records - SELECT SQL
;~     $RecUpdate = ObjCreate ("ADODB.Recordset") ; Create a Record Set to handles SQL Records - UPDATE SQL
	
	; open the connection to the DSN provided
	$ODBC.open($DSN)
	if $Err = 0 Then ; check for errors opening connection
		FileWriteLine($file, "open connection successful")
	Else
		return -1
	EndIf
    
    $RecSelect.CursorType = 0  	; no scrolling
    $RecSelect.LockType = 3

	; define arrays for titles and values 
	dim $arrResults[1][1] 	; array-name[row][column]
	dim $arrTitle[1]
	
    ; open and operate the Record set object to populate the arrays for display.
	With $RecSelect
		; open the record set using the SQL statement and connection object created above
		.Open($SQL, $ODBC)
		If $Err <> 0 then 
			return -1 ; check for error, exit if bad SQL
		Else
			FileWriteLine($file, "SQL execute success:  " & $SQL)
		EndIf
		
		; log record count & get the size to redim the arrays.
		FileWriteLine($file, "Counting Records.")
		$RecCount = 0
		While not .EOF
			$RecCount +=1
			.movenext
		WEnd
		
		; return the bland array if count is 0
		If $RecCount > 0 Then
			.MoveFirst
		Else
			FileWriteLine($file, "RecCount = " & $RecCount & @LF)
			ConsoleWrite("RecCount = " & $RecCount & @lf)
			$arrResults[0][0] = ''
			Return $arrResults
		EndIf		
		
		$Size = .fields.count
		FileWriteLine($file, 'reccount = ' & $RecCount & ' Column count = ' & $Size & @LF)
		
		; ReDim the arrays according to $Size
			; Note: For _ArrayDisplay, $RecSelect.name is in the second dimension
			; in order to display .names  as the "column header".
		ReDim $arrResults[2][$Size]	; make room for the .name and the first .value
		ReDim $arrTitle[$Size]
		
		; populate title(s)
		FileWriteLine($file, 'Starting array populate of titles')
		If $Size > 1 Then
			For $i = 0 to $Size - 1
				$Title = .fields($i).Name	; get the name of each field in .fields collection of ADOdb.RecordSet record 1  (note: names are same for all records in RecordSet, values change)
				$arrTitle[$i] = $Title
				$arrResults[0][$i] = $Title								
			Next
		EndIf
		FileWriteLine($file, 'titles done')

		; prime the value loop and ReDim the array	
		$intRecCount = 2 ; count + 1 of current array records (rows):::(if count is 2, subs are 0 & 1, where 0 has already been populated above with .name
		$intPos = 1	; array index value of the next record of fields to be added
		
		; loop until no more RecordSet records to read
		FileWriteLine($file, 'Starting array populate of values')
		While Not .EOF
				; if titles (.name) are on row 0 then $i represents column numbers where $intPos represents row number
				; loop through the fields for the current RecordSet record and populate array row.
				for $i = 0 to $Size - 1
					$value = .fields( $i ).Value
					$arrResults[$intPos][$i] = $value	
				Next
				
				$intRecCount += 1	; count + 1 of current array records for sub $i
				$intPos += 1	; index value of the next record
				ReDim $arrResults[$intRecCount][$Size] ; make room for the next record in $RecSelect
			
			; move to next RecordSet record, giving a new collection of .fields
			.MoveNext
				
		WEnd	; EOF
		FileWriteLine($file, 'values done')
		
		If $Disp_param = 1 Then
			FileWriteLine($file, 'Displaying Results')
			; show the names of .fields pulled by the SQL
			_ArrayDisplay($arrTitle, "titles")
			
			; Show the 2D array of all results
			_ArrayDisplay($arrResults, "results")
		EndIf
			
	EndWith
	
	FileWriteLine($file, 'Returning results')
	Return $arrResults
	
EndFunc 

; error events come here
Func ErrFunc()
	$file = FileOpen("GetData_Err.txt",  2)
    dim $oRet[3]
	$line = $getError.scriptline
	
;~     $HexNumber = Hex($onError.number, 8)
;~     $oRet[0] = $HexNumber
	$HexNumber = $getError.number
	$oRet[0] = $getError.Number
    $oRet[1] = StringStripWS($getError.description,3)
	$oRet[2] = '' ;$RecSelect.SQL 
    ConsoleWrite("### Error !  Number: " & $HexNumber & "   ScriptLine: " & $line & "   Description:" & $oRet[1] & @LF & $oRet[2])
	FileWriteLine($file, "### Error !  Number: " & $HexNumber & "   ScriptLine: " & $line & "   Description:" & $oRet[1] & @LF & $oRet[2])
	MsgBox(0,"An Error Occurred", "### Error !  Number: " & $HexNumber & "   ScriptLine: " & $line & "   Description:" & $oRet[1] & @LF & $oRet[2])
;~     SetError(1); something to check for when this function returns
	$Err = 1
    Return $Err
EndFunc



#cs---- PutData ----
	Description: 	Performs an update on an AcuCOBOL data file using ADOdb and SQL.
	
	
	Parameters:
	
	Returns:
	
	Remarks:
	
#ce----
Func _PutData($dsn, $sql)
	
EndFunc