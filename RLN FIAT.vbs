'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - RLN FIAT.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 269                	'manual run time in seconds
STATS_denomination = "C"       			' is for case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/29/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Custom functions============================================================================================================
FUNCTION create_dynamic_dialog_for_FIAT(Total_members, test_array)
ReDIM FIAT_array(Total_members, 4)					'redefining 2d array with total members and 4(5) spots for variables to be stored.
FOR each member in test_array
	FOR i = 0 to Total_members								'setting initial variable values for the dialog to display
		FIAT_array(i, 0) = test_array(i)
		FIAT_array(i, 1) = "None"
		FIAT_array(i, 2) = "0"
		FIAT_array(i, 3) = MAXIS_footer_month
		FIAT_array(i, 4) = MAXIS_footer_year
		NEXT
NEXT

BeginDialog Dialog1, 0, 0, 336, (50 + (20 * Total_members)), "Dialog"
Text 30, 10, 60, 10, "Member"
Text 125, 10, 60, 10, "Sanction Reason"
Text 220, 10, 40, 10, "Occurence"
Text 275, 10, 50, 10, "Sanction Begin"
FOR i = 0 TO (Total_members -1)				'protecting against the blank last space on the array
	Text 15, (30 + (20 * i)), 85, 10, FIAT_array(i, 0)
	DropListBox 105, (30 + (20 * i)), 100, 15, "No Sanction"+chr(9)+"Employment Services"+chr(9)+"Child Support Ongoing"+chr(9)+"School Attendance"+chr(9)+"ES & CS", FIAT_array(i, 1)
	DropListBox 225, (30 + (20 * i)), 25, 15, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5", FIAT_array(i, 2)  'occurences
	EditBox 275, (30 + (20 * i)), 20, 15, FIAT_array(i, 3)	'begin month
	Text 300, (30 + (20 * i)), 5, 10, "/"
	EditBox 305, (30 + (20 * i)), 20, 15, FIAT_array(i, 4)	'begin year
NEXT
ButtonGroup ButtonPressed
	OkButton 220, (30 + (20 * Total_members)), 50, 15
	CancelButton 275, (30 + (20 * Total_members)), 50, 15
EndDialog

DO
	err_msg = ""										'clearing error message variable
	at_least_one_sanction = FALSE		'redefining sanction status at begining of loop so we can re-check it later
	dialog dialog1				'calling final dialog
	cancel_confirmation
	FOR i = 0 TO (Total_members - 1)	'making it mandatory to pick at least one sanction
		IF FIAT_array(i, 1) <> "No Sanction" THEN at_least_one_sanction = TRUE
	NEXT
	FOR i = 0 TO (Total_members - 1)  'making it mandatory to enter occurence whens selecting ES
		IF (FIAT_array(i, 1) = "ES & CS" OR FIAT_array(i, 1) = "Employment Services") AND FIAT_array(i, 2) = 0 THEN err_msg = err_msg & vbCr & "Please enter number of occurence when imposing ES sanctions."
	NEXT
	IF at_least_one_sanction = FALSE THEN err_msg = err_msg & vbCr & "Please select at least one sanction to use this script."  'making it mandatory to select at least one sanction
	IF err_msg <> "" THEN msgbox err_msg														'if the error message variable isn't blank we have error and must display/loop again.
LOOP UNTIl err_msg = ""
END FUNCTION
'End Custom functions============================================================================================================

'DIALOG==============================================================================================================================
BeginDialog case_number_dialog, 0, 0, 251, 155, "Red Lake Nation Sanction FIATER"
  EditBox 105, 10, 60, 15, MAXIS_case_number
  EditBox 105, 30, 25, 15, MAXIS_footer_month
  EditBox 140, 30, 25, 15, MAXIS_footer_year
  EditBox 105, 50, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 70, 50, 15
    CancelButton 130, 70, 50, 15
  Text 50, 15, 50, 10, "Case Number:"
  Text 10, 35, 90, 10, "Footer month/year to FIAT:"
  Text 35, 55, 65, 10, "Worker Signature: "
  Text 30, 110, 200, 25, "* This script will FIAT eligibility Sanctions for MFIP. It does not handle child support sanctions for initial applications. "
  Text 30, 135, 200, 10, "* All STAT panels must be updated before using this script."
  GroupBox 20, 95, 215, 55, "Before you begin:"
EndDialog

'END DIALOG==============================================================================================================================

'The Script=============================================================================================================================='
EMConnect ""

check_for_MAXIS(TRUE)						'checking to make sure we are in MAXIS and not passworded out.
CALL MAXIS_case_number_finder(MAXIS_case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

DO
	err_msg = ""																								'clearing error message handling
	dialog case_number_dialog																		'displaying dialog
	cancel_confirmation																					'custom function that allowed for better cancelling options'
	IF isnumeric(MAXIS_case_number)	= FALSE THEN err_msg = err_msg & vbCr & "Please enter a valid case number"	'checking to see if the case number is numeric
	IF len(MAXIS_footer_month) <> 2 THEN  err_msg = err_msg & vbCr & "Enter a 2 digit Footer month"         'making sure footer month is 2 digits
	IF len(MAXIS_footer_year) <> 2 THEN  err_msg = err_msg & vbCr & "Enter a 2 digit Footer year"						'making sure footer year is 2 digits
	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a worker signature"							'making sure worker signature is not blank
	IF err_msg <> "" THEN msgbox err_msg																																		'if we hit any issue that generate an error display error message and loop again
LOOP UNTIL err_msg = ""

check_for_MAXIS(TRUE)						'checking to make sure we are in MAXIS and not passworded out.

'Entering requested footer month
back_to_self
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
Transmit

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

DO								'reads the reference number, last name, first name, and then puts it into a single string to later convert into array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadscreen last_name, 5, 6, 30
	EMReadscreen first_name, 7, 6, 63
	EMReadscreen Mid_intial, 1, 6, 79
	last_name = replace(last_name, "_", "") & " "
	first_name = replace(first_name, "_", "") & " "
	mid_initial = replace(mid_initial, "_", "")
	client_string = ref_nbr & last_name & first_name & mid_intial
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)					'trimming spaces off of the end of string
test_array = split(client_array, "|")			'splitting the string by the | that was entered between each entry
Total_members = Ubound(test_array)				'determining the total amount of entries in the array. NOTICE: this does include the blank entry after the last |

DIM FIAT_array()								'dimming for our 2D array and the roadmap for each spot on the array
' 0 = Member number/name
' 1 = Sanction type
' 2 = Occurence
' 3 = Begin month
' 4 = Begin year

Call create_dynamic_dialog_for_FIAT(Total_members, test_array)

check_for_MAXIS(False)

'commented out as this isn't the active policy 1/6/17
'Checking stat panels first
'For i=0 to Total_members
''	IF FIAT_array(i, 0) <> "" THEN
''		IF FIAT_array(i, 1) = "School Attendance" THEN
''			Call navigate_to_MAXIS_screen("STAT","SCHL")
''			EMWriteScreen left(FIAT_array(i, 0), 2), 20, 76
''			Transmit
''			EMReadScreen School_status, 1, 6, 40				'If school sanction is selected check SCHL for Not Attending
''			IF School_status <> "N" THEN script_end_procedure("This member is not coded as not attending school. You cannot place a school sanction on an attending student.")
''		END IF
		'IF (FIAT_array(i, 1) = "Employment Services" OR FIAT_array(i, 1) = "ES & CS") AND FIAT_array(i, 2) >= 1 THEN    'if ES sanction and any occurance 1 or over is selected
		''	Call navigate_to_MAXIS_screen("STAT", "EMPS")
		''	EMWriteScreen left(FIAT_array(i, 0), 2), 20, 76
		''	Transmit
		''	EMReadScreen EMPS_reason_code, 2, 18, 40						'per instructions must be coded 03 no employment plan
		''	IF EMPS_reason_code <> "03" THEN script_end_procedure("EMPS is not coded as 03 for sanction reason. Please review.")
	''	END IF
''	END IF
'NEXT

'Determining sanction percent and if we need to FIAT sanction limit for 5 ES sanctions
sanction_percent_to_write = 0																'defining sanction % to write as 0 so we can compare it
fail_for_ES_limit = FALSE
FOR i = 0 TO Total_members
	IF FIAT_array(i, 0) <> "" THEN														'If the entry at spot 0 isn't blank then we can continue.
	IF FIAT_array(i, 1) = "Employment Services" OR FIAT_array(i, 1) = "ES & CS" THEN
		IF FIAT_array(i, 2) = "5" THEN fail_for_ES_limit = TRUE	'IF sanctioning for 5th ES violation we set up a dummy variable to be used later to save time.
	END IF
		IF FIAT_array(i, 1) = "Employment Services" THEN				'If we are using ES sanction then it could be 10 or 30 depending on occurence
			IF FIAT_array(i, 2) = 1 THEN													'If the occurence is 1 then code a 10 to the comparing number
				TEMP_sanction_percent = 10
			ELSE
				TEMP_sanction_percent = 30													'otherside code a 30 to the comparing number
			END IF
			IF abs(sanction_percent_to_write) < ABS(TEMP_sanction_percent) THEN sanction_percent_to_write = TEMP_sanction_percent				'If the % we are comparing is over the previously defined number replace the previously defined number with the higher one.
		ELSE
			IF FIAT_array(i, 1) <> "No Sanction" THEN sanction_percent_to_write = 30		'If it isn't a ES sanction then it has to be 30%, even if it's ES/CS it will be 30 because of CS
		END IF
	END IF
NEXT

'Nav to FIAT and check for unapproved version
Call navigate_to_MAXIS_screen ("FIAT", "")
EMReadScreen unapproved_check, 8, 9, 46
IF TRIM(unapproved_check) = "" THEN script_end_procedure("Unapproved MFIP results do not exist. Please review for changes and run case through background.")
'unapproved_check = "11/30/16"
IF unapproved_check <> " " OR datediff("d", unapproved_check, date) = 0 THEN  'must have unapproved version generated today else script will close
	EMWriteScreen "21", 4, 34   'Reason is 21
	EMWriteScreen "x", 9, 22   'X on MFIP
	Transmit
ELSE
	script_end_procedure("Unapproved MFIP results were not generated today. Please review for changes and run case through background.")
END IF

FOR i = 0 TO Total_members																	'For every number starting with 0 to the total number of members in the array
	IF FIAT_array(i, 0) <> "" THEN														'If the entry at spot 0 isn't blank then we can continue.
		FMSL_row = 9																					'setting default for row to read this is dynamic so we can iterate in the loop

		IF FIAT_array(i, 1) <> "No Sanction" THEN
			FMBF_row = 9																						'setting default for row to read this is dynamic so we can iterate in the loop
			EMWriteScreen "X", 17, 4																 'writing x on view budget factors
			Transmit
			EMWriteScreen sanction_percent_to_write, 15, 29														'writing in the sanction percent

			DO																												'hunting for right reference number on FMBF
				EMReadScreen FMBF_ref_num, 2, FMBF_row, 4								'reading the current row's member number
				IF FMBF_ref_num = left(FIAT_array(i, 0), 2) THEN				'If we find that they match
					EMWriteScreen "X", FMBF_row, 65													'writing x and transmitting to that person's sanction screen
					Transmit
					IF FIAT_array(i, 1) = "Child Support Ongoing" THEN EMWriteScreen "F", 7, 14    'if it is a CS sanction we fail for child support
					IF FIAT_array(i, 1) = "ES & CS" THEN																						'if it is ES AND CS sanction we fail for child support AND ES
						EMWriteScreen "F", 9, 14
						EMWriteScreen "F", 7, 14
					END IF
					IF FIAT_array(i, 1) = "School Attendance" THEN EMWriteScreen "F", 9, 14					'If it is School Attendance we fail for ES
					IF FIAT_array(i, 1) = "Employment Services" THEN EMWriteScreen "F", 9, 14					'If it is ES we fail for ES
					IF FIAT_array(i, 2) = "1" THEN									'FIAT demands 1 be written is 10% is used.
						IF FIAT_array(i, 1) = "Employment Services" THEN
							EMWriteScreen "1", 13, 24
						ELSE
							EMWriteScreen "2", 13, 24									'writing the occurence as 2 for all other instances as they are all 30% sanctions.
						END IF
					ELSE
						EMWriteScreen "2", 13, 24									'writing the occurence as 2 for all other instances as they are all 30% sanctions.
					END IF
					EMWriteScreen FIAT_array(i, 3), 13, 42									'writing the begin month
					EMWriteScreen FIAT_array(i, 4), 13, 45									'writing the begin year
					Transmit																								'transmitting twice to exit
					Transmit
				ELSE																											'if we don't find the current member number on the current row we must get to next one
					FMBF_row = FMBF_row + 1																	'adding 1 to our row counter to check the next row'
					IF FMBF_row = 15 THEN 																	'checking for additional members if we hit the bottom of the list
						PF8																										'pf8ing to the next page'
						FMBF_row = 9																					'resetting row to read to top of list
						EMReadScreen FMSL_edit_check, 2, 24, 15								'checking to make sure we have more pages to read, if not we quit because member isn't on case
						IF FMSL_edit_check = "NO" THEN script_end_procedure(FIAT_array(i, 0) & "was not found on MFIP FIAT. Please review.")
					END IF
				END IF
			LOOP UNTIL FMBF_ref_num = left(FIAT_array(i, 0), 2)					'looping until we find the HH member we are FIATing
			PF3																									'backing out of person test will cause the ELIG status to become ELIG instead of UNKN
			DO
				EMReadScreen Back_on_FMSL, 4, 3, 52								'This is a do loop to make sure we don't get hung up on a skipable edit after pf3ing
				IF FMPR_edit_check <> "FMSL" THEN PF3
			LOOP UNTIL Back_on_FMSL = "FMSL"
		END IF
	END IF
	Back_on_FMSL = ""									'clearing out read variables for the next run through.
	FMSL_edit_check = ""
	FMPR_edit_check = ""
NEXT

'checking case tests to make sure it isn't marked as UNKNOWN
EMWriteScreen "x", 16, 4
Transmit
IF fail_for_ES_limit = TRUE THEN 											'if this is the 5th ES sanction we fail for sanction limit
	EMWriteScreen "F", 9, 44
	Transmit
END IF
PF3																								'backing out to FMSL
DO
	EMReadScreen Back_on_FMSL, 4, 3, 52								'This is a do loop to make sure we don't get hung up on a skipable edit after pf3ing
	IF FMPR_edit_check <> "FMSL" THEN PF3
LOOP UNTIL Back_on_FMSL = "FMSL"

'This is where it will push through BUDGET for entire case and save fiat version
EMWriteScreen "x", 18, 4
Transmit																'transmit into budget/FMB1
Transmit																'transmit into FMB2
Transmit																'transmit into FMSM
EMWriteScreen "         ", 10, 31				'writing 9 spaces to clear out line
EMWriteScreen "MONTHLY", 10, 31					'writing MONTHLY on HRF reporting
Transmit																'transmitting to make monthly stick
PF3																			'pf3ing out of budget
PF3																			'pf3ing out of FIAT
EMSendKey "y"														'retaining FIAT version
Transmit																'confirming Y to retain FIAT version

continue_to_case_note = msgbox("Selected FIATs have been added for ELIG/MFIP " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review ELIG and approve if correct." & vbCr & vbCr & "Click Yes to continue to case note this FIAT. Click No to end script.", vbYesNo)
IF continue_to_case_note = vbNo THEN script_end_procedure("FIAT has been added, please remember to case note manually and add WCOMs to any FIATed sanction approval.")

check_for_MAXIS(FALSE)																						'checking to make sure MAXIS isn't passworded out and user is still in MAXIS.
back_to_self																											'function to go back to self
start_a_blank_CASE_NOTE																						'function to check for edit ability and open blank case note.
CALL write_variable_in_CASE_NOTE("FIAT info for Red Lake Nation TANF Sanction " & MAXIS_footer_month & "/" & MAXIS_footer_year)
CALL write_variable_in_CASE_NOTE("FIAT entered for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " by script.")
CALL write_variable_in_CASE_NOTE("---")
IF fail_for_ES_limit = FALSE THEN CALL write_variable_in_CASE_NOTE("Sanction Percent Applied: " & sanction_percent_to_write & "%")
IF fail_for_ES_limit = TRUE THEN CALL write_variable_in_CASE_NOTE("Case marked as ineligible due to 5th employment services sanction.")
CALL write_variable_in_CASE_NOTE("---")
FOR i = 0 to Total_members - 1																		'case noting every member in the array who has a sanction, and if it's ES related case note the occurence.
	IF FIAT_array(i, 1) <> "No Sanction" THEN	CALL write_variable_in_CASE_NOTE(FIAT_array(i, 0) & " - " & FIAT_array(i, 1))
	IF FIAT_array(i, 1) = "Employment Services" OR FIAT_array(i, 1) = "ES & CS" THEN CALL write_variable_in_CASE_NOTE("      Occurence: " & FIAT_array(i, 2))
NEXT
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
