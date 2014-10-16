'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(""C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'----------FUNCTIONS----------
FUNCTION convert_letter_to_number(excel_col)
	alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	excel_col = ucase(excel_col)
	IF len(excel_col) = 1 THEN 
		excel_col = InStr(alphabet, excel_col)
	ELSEIF len(excel_col) = 2 THEN 
		excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
	END IF
END FUNCTION

'----------DIALOGS----------
BeginDialog bulk_case_note_dialog, 0, 0, 256, 210, "Case Note Information"
  EditBox 10, 25, 235, 15, excel_file_path
  EditBox 220, 45, 25, 15, excel_col
  EditBox 65, 65, 40, 15, excel_row
  EditBox 190, 65, 40, 15, end_row
  EditBox 10, 105, 235, 15, case_note_header
  EditBox 10, 145, 235, 15, case_note_body
  EditBox 120, 165, 90, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 65, 190, 55, 15
    CancelButton 125, 190, 60, 15
  Text 10, 10, 150, 10, "Please enter the file path for the Excel file..."
  Text 10, 50, 205, 10, "Please enter the column containing the MAXIS case numbers..."
  Text 10, 70, 50, 10, "Row to start..."
  Text 135, 70, 50, 10, "Row to end..."
  Text 10, 90, 165, 10, "Please enter your case note header..."
  Text 10, 130, 145, 10, "Please enter the body of your case note..."
  Text 15, 170, 100, 10, "Please sign your case note..."
EndDialog


'----------THE SCRIPT----------
EMConnect ""

maxis_check_function

DO
	Dialog bulk_case_note_dialog
		IF ButtonPressed = 0 THEN stopscript
		IF isnumeric(excel_col) = FALSE AND len(excel_col) > 2 THEN 
			MsgBox "Please do not use such a large column. The script cannot handle it."
		ELSE
			IF (isnumeric(right(excel_col, 1)) = TRUE AND isnumeric(left(excel_col, 1)) = FALSE) OR (isnumeric(right(excel_col, 1)) = FALSE AND isnumeric(left(excel_col, 1)) = TRUE) THEN
				MsgBox "Please use a valid Column indicator. " & excel_col & " contains BOTH a letter and a number."
			ELSE
				IF isnumeric(excel_col) = FALSE THEN call convert_letter_to_number(excel_col)
				IF isnumeric(excel_row) = false or isnumeric(end_row) = false THEN MSGBox "Please enter the Excel rows as numeric characters."
				IF end_row = "" THEN MSGBox "Please enter an end to the search. The script needs to know when to stop searching."
				IF worker_sig = "" THEN MSGBox "Please sign your case note."
			END IF
		END IF
LOOP UNTIL (isnumeric(excel_col) = true) and excel_col <> "" and isnumeric(excel_row) = true and isnumeric(end_row) = true and end_row <> "" AND worker_sig <> ""

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open(excel_file_path)
objExcel.DisplayAlerts = True

end_row = end_row * 1
excel_col = excel_col * 1
excel_row = excel_row - 1
DO
	excel_row = excel_row + 1
	case_number = ObjExcel.Cells(excel_row, excel_col).Value
	case_number_array = case_number_array & case_number & " "
LOOP UNTIL excel_row = end_row

case_number_array = trim(case_number_array)
case_number_array = split(case_number_array)

privileged_count = 0
out_of_county_count = 0
invalid_case_count = 0
FOR EACH case_number IN case_number_array
	back_to_SELF
	call navigate_to_screen("CASE", "NOTE")
	EMReadScreen invalid_case, 7, 24, 2
	EMReadScreen primary_county, 4, 21, 14
	EMReadScreen privileged_case, 10, 24, 14
	IF invalid_case <> "INVALID" THEN
		IF privileged_case = "PRIVILEGED" THEN
			privileged_array = privileged_array & case_number & " "
			privileged_count = privileged_count + 1
		ELSE
			IF ucase(primary_county) = ucase(worker_county_code) THEN
				PF9
				call write_new_line_in_case_note(case_note_header)
				call write_new_line_in_case_note(case_note_body)
				call write_new_line_in_case_note("---")
				call write_new_line_in_case_note(worker_sig)
				call write_new_line_in_case_note("---")
				call write_new_line_in_case_note("**Processed in bulk script**")
			ELSE
				out_of_county_array = out_of_county_array & case_number & " "
				out_of_county_count = out_of_county_count + 1
			END IF
		END IF
	ELSE
		invalid_array = invalid_array & case_number & " "
		invalid_case_count = invalid_case_count + 1
	END IF
NEXT

privileged_array = trim(privileged_array)
privileged_array = split(privileged_array)

out_of_county_array = trim(out_of_county_array)
out_of_county_array = split(out_of_county_array)

IF privileged_count > 0 or out_of_county_count > 0 or invalid_case_count > 0 THEN
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = True
	
	objExcel.Cells(1, 1).Value = "PRIVILEGED CASES"
	objExcel.Cells(1, 2).Value = "OUT OF COUNTY CASES"
	objExcel.Cells(1, 3).Value = "INVALID CASES"	
	
	privileged_row = 2
	FOR EACH priv_case IN privileged_array
		objExcel.Cells(privileged_row, 1).Value = priv_case
		privileged_row = privileged_row + 1
	NEXT
	
	out_of_county_row = 2
	FOR EACH out_of_county_case IN out_of_county_array
		objExcel.Cells(out_of_county_row, 2).Value = out_of_county_case
		out_of_county_row = out_of_county_row + 1
	NEXT
	
	invalid_case_row = 2
	FOR EACH invalid_maxis_case IN invalid_array
		objExcel.Cells(invalid_case_row, 3).Value = invalid_maxis_case
		invalid_case_row = invalid_case_row + 1
	NEXT
END IF

script_end_procedure("The script is now finished running. If the script did not appear to do anything, it is likely that the column you selected is empty.")