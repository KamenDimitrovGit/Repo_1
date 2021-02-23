section info

'1. Place this script where your DB and dms are.
'2. Set the mandatory settings.
'3. Use Just_Run with 1 at the beginning.
'4. Press F5
'First run will be slower than the next ones. The dms can be edited between the runs. 
'If the Powershell window(the blue window) freeze or appears as doing nothing press F5 to refresh it.
'If you want to stop earlier the checking you need to close all the cmd windows(the black background windows)


'----------------- Requirements:
'Run it locally on your PC - still not tested on S: / rofs: / robuc: ....
'This script will make a 'BackUp' folder where the original mdd, ddf/dzf and dms will be stored.



'----------------- Known errors:
'1. Pre-processor error: (row): error: could not find include file: \\robuc-fp01.... in C:\Users.... - you don't have internet/VPN connection turned on.
'2. 'Method call failed: The system cannot find the file specified.' - restart Base Professional and try again.
'3. (0x800A003E) - "SyntaxError.txt" - restart Base Professional and try again.
'4. (0x800A0035) - set db_bckp = .GetFile("BackUp\"+database_name+DZForDDF).DateLastModified - Execute Error(210): Method call failed: Unknown error 0x800A0035 - you have ddf and dzf in the main folder and only dzf in the 'BackUp' folder. Copy manually the ddf from the main folder to the 'BackUp' folder and run again the script.
'5. "Error : Error creating input cursor Input - Error Creating Rowset" - check the syntax and variable names in 'SelectQuery_for_DMS'
'6. "Merge_Outputs: Error : Event(OnBeforeJobStart,Delete existing files) mrScriptEngine execution error: Execute Error(61): Method call failed: Unknown error 0x800A0046" - Close the 'Merged_Output.mdd/ddf/dzf' in dmquery/Professional and run manually 'Merge_Outputs.dms' which can be found in the folder where the qa.dms is.
'7. In case of other undefined errors like (0x800Axxxx) --> https://ss64.com/vb/syntax-errors.html



'----------------- Features to be implement:
' print dots or records processed.
'clean the cleaning - some repeated lines
'ask to merge if errors
'err handling 0x800A003E and 0x800A0035

'download locally scripts
'err handling for open/delete files
'put some output in the while in order to print smth when the script is slow
'log_file_name to var
'PS to object and catch when it finishes. refresh?
'err std output
'memmory when big metadata
'if .FileExists("mrs_log.txt") then .DeleteFile("mrs_log.txt")  - Execute Error(156): The system cannot find the file specified. - the file is open in powershell - i.e. PS should be closed.

end section


'================= Mandatory Settings
'database_name - do not add '.mmd' or '.ddf'. just the snumber/name of the db without any extensions.
#Define database_name "S20035219_ddf"
#Define DMS_Name "QA"

'Standard_IIS_DMS: 0 or 1. Set to '1' if your are going to run the standard IIS's QA.dms for checking studies. Set to 0 if you'll running a different dms.
#Define Standard_IIS_DMS 1

'Just_Run: 0 or 1. 0 = advanced settings will be available for changing. 1 = you just need to set the database/dms names and type of dms(standard iis or not), and press F5.
#Define Just_Run 1

'SelectQuery: default: "". Here you can add the part of the query after the 'WHERE' clause.
dim SelectQuery_for_DMS
SelectQuery_for_DMS = "" 'Example: returncode*{C} and IDtype*{real}. Splitted db's will contain all the data but only what's in the query will be checked.

'additional_includes: If your dms uses additional incldes like dms, xlsx, etc. files.
dim additional_includes
additional_includes = "" ' Ex.: "MCRT.mrs, blah_blah.xlsx, 123.txt". Use comma as separator. If no additional files used leave it empty - ""

'merge_cond: The db is splitted to several parts and after the check you'll have several outputs('Output.mdd/ddf'). If you want the clean data to be merged after the check set this to true.
'if there are errors(issues) found the merge will be canceled automatically. If you still want to merge the clean data you can use 'Merge_Outputs.dms' which could be found in the main folder where the dms and db's are.
dim merge_cond
merge_cond = true 'true or false.




'================= Advanced Settings(Non-Mandatory)
'Join_Cleanings/PopUpCleanings: setting specific for the standard IIS QA shell (qa.dms) for checking studies. If your dms is not that type these are automatically set to false/0.
dim Join_Cleanings, PopUpCleanings
Join_Cleanings = true 'true/false. TRUE = All cleanings will be merged in one named 'Cleaning Summary.txt'. FALSE = Cleanings will not be merged. in both cases(0 and 1) the separate cleanings will be available in 'DMS_Name'+_1/Output, 'DMS_Name'+_2/Output, ... folders. 
PopUpCleanings = false 'true/false. FALSE = Cleaning.txt and Quotatable.html will not pop-up after the end of qa.dms. TRUE = Both files will pop up as in normal run of the dms. Due to inconvenience of poping up too many cleanings and quota tables when running qa.dms you can turn them off.


'================= Settings that Require 'Just_Run' to be set to 0. If it's 1 there's no need to edit them.(Non-Mandatory)
dim core_count
core_count = 10 'the number of logical processors to be used. If set to a bigger number than available cores it wil automatically switch to the max number of cores. If you don't know how many you have set 'TellMeHowManyPrcessorsIHave' to 1 and run the script.

'Use this if you want to manually set in the advanced settings the number of cores to be used.
'TellMeHowManyPrcessorsIHave: 0 or 1. 0 = will not do anything. 1 = will tell you the number of processros in your pc and exit the script.
#Define TellMeHowManyPrcessorsIHave 0

				'------- NOT USED ANYMORE
					'--- cmd_popup: determine the way how cmd windows will appear. 7(minimized) is the default. 5 is somehow useful but interupts whatever you are doing. 0 is for absolutelly hidden.
					#Define cmd_popup 7
					
					'--- Affinity: 0 or 1. Not recommended to be set to '1'. 0 = all dms scripts will run without dedicated core(5-10% faster). 1 = every dms will have dedicated core. In case of instabillity(dms ends in the middle of its execution) try with 1 otherwise keep that to 0.
					#Define Affinity 0
				'------- END OF NOT USED ANYMORE


'================= End of settings





'----- Timers
dim StartTime
StartTime = Now()
'debug.Log(ctext(StartTime.TimeOnly()) + ": Job starts.")
delog("Job starts.")

#if Just_Run 
	#ifdef Affinity
		#Undef Affinity
		#Define Affinity 0
	#endif
#endif



if not(Standard_IIS_DMS) then
	Join_Cleanings = false
	PopUpCleanings = false
end if

dim i, k, ii, points

'------ objects
dim fso, txt_file, txt_file_text
dim mdm
dim adoRecordSet, adoConnection, adoRS, adoCommand
dim wshell
dim shellapp
dim ScriptEngine
dim LogObj
Set fso = CreateObject("Scripting.FileSystemObject")
set mdm = createObject("MDM.Document")
Set adoConnection = CreateObject("ADODB.Connection")
Set wshell = CreateObject("WScript.Shell")
'Set shellapp = CreateObject("Shell.Application")
'Set ScriptEngine = CreateObject("mrScript.ScriptEngine")
'Set LogObj = CreateObject("LogFront.Logger.2")








section ManualCheckProcessors

dim num_core
if TellMeHowManyPrcessorsIHave then
	debug.Log(ctext(StartTime.TimeOnly()) + ": Checking the system processors.")
	num_core = wshell.ExpandEnvironmentStrings("%number_of_processors%").clong()
	debug.Log(ctext(StartTime.TimeOnly()) + ": ------ Your PC have: "+ctext(num_core)+" cores. ----------")	
	debug.MsgBox("Your PC have: "+ctext(num_core)+" cores.")
	debug.Log(ctext(StartTime.TimeOnly()) + ": The script will exit now. ")
	debug.MsgBox("Set 'TellMeHowManyPrcessorsIHave'(row ~57) to '0'")
	debug.Log(ctext(StartTime.TimeOnly()) + ": Set 'tellmehowmanyprcessorsihave' to '0' in order to use the main functionality(split db) of this script.")
	exit
end if
end section

with fso

'-------- Creating log file of this mrs for reading in PS
section mrs_log
	
if .FileExists("mrs_log.txt") then .DeleteFile("mrs_log.txt")

dim run_PS, mrs_log_file
run_PS = true

'creating the file
set mrs_log_file = fso.OpenTextFile("mrs_log.txt",8,true,0)
mrs_log_file.close()

'start the PS
if run_PS then
	wshell.run("powershell.exe Get-Content mrs_log.txt -Wait -Tail 500") '-wait works only in file system drives.
	run_PS = false
end if

end section

'-------- Check if all files exist / Create TEMP and DDF
section FilesExist

dim DZForDDF

if not(.FileExists(DMS_Name+".dms") and .FileExists(database_name+".mdd") and (.FileExists(database_name+".ddf") or .FileExists(database_name+".dzf"))) then
	delog("------ You need to have '"+DMS_Name+".dms', '"+database_name+".mdd' and '"+database_name+".ddf/dzf' files in the folder. Script will exit now")
	delog("NOW CLOSE THIS POWERSHELL WINDOW.")
	exit
end if

'temp
if .FolderExists("TEMP") then .DeleteFolder("TEMP",true)
.CreateFolder("TEMP")

'mdd
.CopyFile(database_name+".mdd","TEMP\")

'check ddf
if .FileExists(database_name+".ddf") then
	DZForDDF = ".ddf"
	.CopyFile(database_name+".ddf","TEMP\")
elseif .FileExists(database_name+".dzf") then
	DZForDDF = ".dzf"
	.CopyFile(database_name+".dzf","TEMP\")
	'create ddf / delete dzf
	adoConnection.Open("Provider=mrOleDB.Provider.2; _
				Data Source=mrDataFileDsc; _
				Location=.\TEMP\"+database_name+".ddf; _
				Initial Catalog=TEMP\"+database_name+".mdd"+"; _
				MR Init MDM Access=1")
	adoconnection.close()
	.DeleteFile("TEMP\"+database_name+".dzf")
else
	delog("------ You havo to have '"+database_name+".ddf/dzf' file in the folder. Script will exit now")
	delog("NOW CLOSE THIS POWERSHELL WINDOW.")
	exit
end if
end section

'-------- Number of processors
section CoreCount

delog("Checking the system processors.")

num_core = wshell.ExpandEnvironmentStrings("%number_of_processors%").clong()
delog("Your PC have "+ctext(num_core)+" cores.")

if Just_Run then
	if num_core = 1 then
		core_count = 1 'this just doesn't make sense
	elseif num_core = 2 then
		core_count = 2
	elseif num_core = 3 then
		core_count = 3
	elseif num_core = 4 then
		core_count = 3
	else
		core_count = num_core - 2 '~75-80% of the available processors
	end if
else
	if core_count > num_core then
		core_count = num_core 'all cores!!!
	end if
end if
end section

'-------- DB New/Split needed
section DBDate

dim db, db_bckp, db_is_new, QA_Folder, make_split
db_is_new = false
make_split = false

if .FolderExists("BackUp") then
	if .FileExists("BackUp\"+database_name+".mdd") and (.FileExists("BackUp\"+database_name+".dzf") or .FileExists("BackUp\"+database_name+".ddf")) then
		set db = .GetFile(database_name+DZForDDF).DateLastModified
		set db_bckp = .GetFile("BackUp\"+database_name+DZForDDF).DateLastModified
		if datediff(db_bckp,db,"s")=0 then
			for i = 1 to 16
				QA_Folder = DMS_Name+"_"+ctext(i)
				if i<=core_count then
					if not(.FolderExists(QA_Folder) and .FileExists(QA_Folder+"\"+database_name+".mdd") and .FileExists(QA_Folder+"\"+database_name+".ddf")) then
						make_split = true
					end if
				else
					if .FolderExists(QA_Folder) then
						make_split = true
					end if
				end if
			next
		else
			db_is_new = true
		end if
	else
		db_is_new = true
	end if
else
	db_is_new = true
end if
if db_is_new then make_split = true
end section

'-------- Backup
section Backup
	
'Make BackUp folder and delete old QA folders
if db_is_new then
	if .FolderExists("BackUp") then .DeleteFolder("BackUp")
	.CreateFolder("BackUp")
	.CopyFile(database_name+".mdd","BackUp\",true)
	.CopyFile(database_name+DZForDDF,"BackUp\",true)
	.CopyFile(DMS_Name+".dms","BackUp\",true)
	for i = 1 to 16
		if .FolderExists("QA_"+ctext(i)) then .DeleteFolder("QA_"+ctext(i))
	next
	delog("The dms, mdd and ddf files were backed up in the 'BackUp' folder.")
end if
end section

'------- Checking DMS for syntax errors / Expand DMS and replace
dim DMS_Name2, err_check_obj
if Standard_IIS_DMS then
	section ExpandDMS
	
	dim ErrorTxt, ErrName
	if .FileExists("SyntaxError.txt") then .DeleteFile("SyntaxError.txt")
	
	'Check and save expanded DMS
	.OpenTextFile("cmd_err_check.cmd",2,false,-2).write("dmsrun "+DMS_Name+".dms /norun /a:"+DMS_Name+"_expanded.dms 1>SyntaxError.txt")
	set err_check_obj = wshell.exec("cmd_err_check.cmd")
	
	
	'ExitCode = 0 --> Program successfully executed.
	'Status 0 = running, 1 finished, 2 failed
	do while err_check_obj.Status = 0
		sleep(1000)
		delog("Checking for syntax errors.")
	loop

	
	if .FileExists("SyntaxError.txt") then
		ErrorTxt = fso.OpenTextFile("SyntaxError.txt",1,true,-2).readall().Trim()
		
		if find(ErrorTxt, "0 Error(s) and 0 Warning(s)") = -1 then 'If errors found
		
			ErrName = "SyntaxError_" + ctext(now()).replace(":",".").replace("/",".").replace("\",".") + ".txt" 'creating unique name for the err file
			.CopyFile("SyntaxError.txt",ErrName,false) 'make file with date in the name
			.DeleteFile("SyntaxError.txt") 'delete the old err file
			delog("DMS syntax errors found. The script will exit and error log will be opened now.")
			delog("NOW CLOSE THIS POWERSHELL WINDOW.")
			debug.Log("DMS syntax errors found. Error log saved in '" + ErrName + "'.") 'cyrilic symbols like 'г.' crash the delog func.
			shellexecute(ErrName,,,,,5)
			exit
		else
			.DeleteFile("SyntaxError.txt")
			DMS_Name2 = DMS_Name+"_expanded"
			
			'Replace
			dim DMS_Obj, txtDMS
			
			set DMS_Obj = .OpenTextFile(DMS_Name2+".dms",1,false,-2)
			txtDMS = DMS_Obj.readall()
			DMS_Obj.Close()
			.DeleteFile(DMS_Name2+".dms")
			
			'Checking if the db name is the same as the one in this mrs.
			if txtDMS.find("Location = .\" + database_name + ".ddf") = -1 or txtDMS.find("Initial Catalog = .\" + database_name + ".mdd") = -1 then
				delog("'database_name' variable and the inputs('#define input_mdd/#define input_ddf') in the dms have different values. They should be the same. The script will exit now.")
				delog("NOW CLOSE THIS POWERSHELL WINDOW.")
				exit
			end if
			
			
			'Update the query
			if SelectQuery_for_DMS.len()>0 then
				txtDMS = txtDMS.replace("SelectQuery = ""SELECT * FROM VDATA""", "SelectQuery = ""SELECT * FROM VDATA WHERE " + SelectQuery_for_DMS + """" )
			end if
			
	
			'Adding func for output
			delog("DMS succesfully expanded. Replacing functions.")
			dim add_func
			add_func = mr.CrLf + mr.CrLf
			add_func = add_func + mr.Tab + "dim fso, out_file, err_txt, prev_err_txt" + mr.CrLf
			add_func = add_func + mr.Tab + "set fso = createobject(" + """scripting.filesystemobject""" + ")" + mr.CrLf
			add_func = add_func + mr.Tab + "set out_file = fso.OpenTextFile(" + """..\mrs_log.txt""" + ",8,false,0)" + mr.CrLf
			add_func = add_func + mr.Tab + "err_txt = " + """"+DMS_Name+"_"+"""+fso.GetParentFolderName(fso.GetAbsolutePathName("""+DMS_Name2+".dms"")).Replace("+"""\"""+","+""""""+").right(1) + """+": "+""" + note + mr.CrLf" + mr.CrLf
			add_func = add_func + mr.Tab + "out_file.Write(err_txt.Replace(prev_err_txt," + """" + """" + "))" + mr.CrLf
			add_func = add_func + mr.Tab + "prev_err_txt = err_txt" + mr.CrLf
			add_func = add_func + mr.Tab + "out_file.Close()" + mr.CrLf
			add_func = add_func + mr.Tab + "set fso = null" + mr.CrLf

			txtDMS = txtDMS.replace("debug.Log(""------>>>"" + note + mr.CrLf)", "debug.Log(""------>>>"" + note + mr.CrLf)" + add_func + mr.CrLf)
			
			'Removing popups
			if not(PopUpCleanings) then
				txtDMS = txtDMS.replace("objTableDoc.Exports[""mrHtmlExport""].Properties[""LaunchApplication""] = True", "'objTableDoc.Exports[""mrHtmlExport""].Properties[""LaunchApplication""] = True")
				txtDMS = txtDMS.replace("ShellExecute(oFso.GetAbsolutePathName(""."") + ""\Output\Cleaning.txt"", , , , ,4)","'ShellExecute(oFso.GetAbsolutePathName(""."") + ""\Output\Cleaning.txt"", , , , ,4)")
			end if
			
			'Save the changes
			dim UpdateFile
			set UpdateFile = .OpenTextFile(DMS_Name2+".dms",2,true,-1)
			UpdateFile.write(txtDMS)
			UpdateFile.close()		
		end if
	else 'continue w/o syntax check
		delog("Unknown error - SyntaxError.txt cannot be created.")
		delog("PopUpCleanings will be set to true and you will see no runtime output")
		PopUpCleanings = true
		DMS_Name2 = DMS_Name
	end if
	end section
else 'Standard_IIS_DMS
	DMS_Name2 = DMS_Name
end if 'Standard_IIS_DMS


end with 'fso

'-------- Split / distribute files / run DMS
section SplitDB

if make_split then
	'------ ADO open base mdd/ddf
	delog("Starting ado connection.")
	adoConnection.Open("Provider = mrOleDB.Provider.2; _
					Data Source = mrDataFileDsc; _
					Location = TEMP\"+database_name+".ddf; _
					Initial Catalog = TEMP\"+database_name+".mdd"+"; _
					MR Init MDM Access = 1")
	
	
	'---- ADO add metadata
	delog("Adding metadata.")
	set adoCommand = CreateObject("ADODB.command")
	set adoCommand.ActiveConnection = adoConnection
	adoCommand.CommandText = "ALTER TABLE vdata ADD COLUMN AssigntoCore long null"
	adoCommand.execute()
	Set adoRS = adoConnection.Execute("select AssigntoCore from vdata")
	
	'------ ADO AssigntoCore and close Ado
	delog("Start distribution of cores.")
	Do Until adoRS.EOF
		ii = ii + 1
		if ii > clong(core_count) then ii = 1
		adoRS.Fields["AssigntoCore"] = ii
		adoRS.MoveNext() 
	Loop
	adoConnection.Close()
	delog("End distribution of cores.")
		
	
	'---------- Distributing and starting files
	delog("Distributing and starting files.")
	for i = 1 to 16

		'----- mdm/fso for each output
		with fso
			if .FolderExists(DMS_Name+"_"+ctext(i)) then .DeleteFolder(DMS_Name+"_"+ctext(i))
			if not(i>core_count) then
				'------- Creating folders
				delog("Creating folder "+DMS_Name+"_"+ctext(i)+" and copying files.")
				.CreateFolder(DMS_Name+"_"+ctext(i))
				.CopyFile(DMS_Name2+".dms",".\"+DMS_Name+"_"+ctext(i)+"\",true)
				.CopyFile("TEMP\"+database_name+".mdd",".\"+DMS_Name+"_"+ctext(i)+"\",true)
				.CopyFile("TEMP\"+database_name+".ddf",".\"+DMS_Name+"_"+ctext(i)+"\",true)
				if additional_includes.len()>0 then
					wshell.run("robocopy . "+DMS_Name+"_"+ctext(i)+" "+additional_includes.replace(","," "),0,true)
				end if
				
				'----- ado open splitted and deleting unnecessary respondents
				delog("Re-open ado connection.")
				adoConnection.Open("Provider=mrOleDB.Provider.2; _
					Data Source = mrDataFileDsc; _
					Location = .\"+DMS_Name+"_"+ctext(i)+"\"+database_name+".ddf; _
					Initial Catalog = .\"+DMS_Name+"_"+ctext(i)+"\"+database_name+".mdd; _
					MR Init MDM Access = 1")
				
				delog("Deleting unnecessary respondents.")
	'			set adoCommand = CreateObject("ADODB.command")
				set adoCommand.ActiveConnection = adoConnection
				adoCommand.CommandText = "DELETE FROM vdata WHERE AssigntoCore<>"+ctext(i)
				adoCommand.execute()
				adoConnection.Close()
			
			
				'----- deleting added metadata
				delog("Deleting added metadata.")
				mdm.Open(".\"+DMS_Name+"_"+ctext(i)+"\"+database_name+".mdd") ',,2-rw
				mdm.Fields.Remove("AssigntoCore")
				mdm.Save(".\"+DMS_Name+"_"+ctext(i)+"\"+database_name+".mdd")
				mdm.Close()

				
				'------ Set Affinity / Making the cmd command / Starting the dms / make log file
'				StartDMS(DMS_Name2, i, Affinity)
			end if
		end with
	next
else
	'---------- Distributing and starting files
	delog("Distributing files:")
	for i = 1 to 16
		if not(i>core_count) then
			'------ Update the dms
			fso.CopyFile(DMS_Name2+".dms",".\"+DMS_Name+"_"+ctext(i)+"\",true)
			if additional_includes.len()>0 then
				wshell.run("robocopy . "+DMS_Name+"_"+ctext(i)+" "+additional_includes.replace(","," "),0,true)
			end if
			
			'------ Set Affinity / Making the cmd command / Starting the dms / make log file
'			StartDMS(DMS_Name2, i, Affinity)			
		end if
	next
end if

'--------- Starting the dms files / make log file
dim Wrun1,Wrun2,Wrun3,Wrun4,Wrun5,Wrun6,Wrun7,Wrun8,Wrun9,Wrun10,Wrun11,Wrun12,Wrun13,Wrun14,Wrun15,Wrun16
dim prev_out1,prev_out2,prev_out3,prev_out4,prev_out5,prev_out6,prev_out7,prev_out8,prev_out9,prev_out10,prev_out11,prev_out12,prev_out13,prev_out14,prev_out15,prev_out16
dim curr_out1,curr_out2,curr_out3,curr_out4,curr_out5,curr_out6,curr_out7,curr_out8,curr_out9,curr_out10,curr_out11,curr_out12,curr_out13,curr_out14,curr_out15,curr_out16
dim err_txt, prev_err_txt
dim script_running, out_file
script_running = true

'---- start each dms
delog("Starting the scripts:")
for i = 1 to 16
	if not(i>core_count) then
		execute("set Wrun"+ctext(i)+" = wshell.exec("+"""dmsrun "+DMS_Name+"_"+ctext(i)+"\"+DMS_Name+"_expanded.dms"""+")")
	end if
next

'---- check the status and logging
while script_running
	dim status_str, exit_code, exit_code_arr
	status_str = ""
	exit_code = ""
	exit_code_arr = ""
	
	'------ check the status
	'ExitCode = 0 --> Program successfully executed.
	'Status 0 = running, 1 finished
	
	for i = 1 to core_count
		status_str = status_str + eval("Wrun"+ctext(i)+".Status.ctext()") + ","
		exit_code = exit_code + eval("Wrun"+ctext(i)+".ExitCode.ctext()") + ","
	next
	exit_code_arr = exit_code.split(",")
	
	if find(status_str,"0") <> -1 then 'if the dms has started
		script_running = true
				
		for i = 1 to core_count
			if exit_code_arr[i-1] = "0" then
				execute("prev_out"+ctext(i)+" = "+"curr_out"+ctext(i))
				execute("curr_out"+ctext(i)+" = "+""""+DMS_Name+"_"+ctext(i)+": """+" + Wrun"+ctext(i)+".stdout.readline().rtrim().Replace(prev_out"+ctext(i)+","+""""+""""+")")
				execute("err_txt = err_txt + curr_out"+ctext(i)+" + mr.CrLf")
			end if
		next
		

		set out_file = fso.OpenTextFile("mrs_log.txt",8,true,0)
		out_file.Write(err_txt.Replace(prev_err_txt,""))		
		out_file.close()
		prev_err_txt = err_txt
		
	else
		script_running = false
		delog("All scripts are ready...")
	end if
end while

end section

'-------- Merge Cleanings
if Standard_IIS_DMS then
section MergeCleaning

if Join_Cleanings then
	delog("Prepare for merging cleanings.")
	
	dim cleaning_sum, cl_curr_open, unsuccessfully_finished
	if fso.FileExists("Cleaning Summary.txt") then fso.DeleteFile("Cleaning Summary.txt")
	set cleaning_sum = fso.OpenTextFile("Cleaning Summary.txt",8,true,-1)
	unsuccessfully_finished = 0
	
	'sleep(2000) 'wait to write the files
	
	delog("Merging all cleaning files.")
	for i = 1 to core_count
		if exit_code_arr[i-1] = "0" then
			set cl_curr_open = fso.OpenTextFile(DMS_Name+"_"+ctext(i)+"\Output\Cleaning.txt",1,false,-1)
			cleaning_sum.writeline("====================== Cleaning from "+DMS_Name+"_"+ctext(i)+" =====================")
			cleaning_sum.write(mr.crlf+cl_curr_open.readall()+mr.crlf+mr.crlf+mr.crlf)
			set cl_curr_open = null
		else
			unsuccessfully_finished = unsuccessfully_finished + 1
			cleaning_sum.writeline("====================== Cleaning from "+DMS_Name+"_"+ctext(i)+" =====================")
			cleaning_sum.writeline(mr.crlf+"----Cleaning was NOT produced in this run. The DMS was stopped or failed.")
			
			if fso.fileexists("mrs_log.txt") then
			
				dim curr_line
				curr_line = ""
				
				cleaning_sum.writeline("----Check 'mrs_log.txt' for partial cleaning and the error which caused the stopping of the dms.")
				cleaning_sum.writeline("----The bellow text is extracted from it.")
				
				set cl_curr_open = fso.OpenTextFile("mrs_log.txt",1,false,0)
				do while cl_curr_open.AtEndOfStream <> True				    
				    curr_line = cl_curr_open.readline()
				    if find(curr_line, DMS_Name+"_"+ctext(i)+":")<>-1 then
				    	cleaning_sum.writeline(curr_line)
				    end if
				Loop
				set cl_curr_open = null
				cleaning_sum.writeline(mr.CrLf+mr.CrLf+mr.CrLf)
			end if
		end if
	next
	
	
	set cleaning_sum = null
	wshell.run("""Cleaning Summary.txt""",5,false)
	
end if 'Join_Cleanings(main 'if' in the begining of the section)
end section


'--------- Merge Output DDF's
section MergeOutputs

if merge_cond then delog("Merging process starts.")

'canceling merge if a script(s) failed to make output.
if merge_cond then
	if unsuccessfully_finished > 0 then
'		if debug.MsgBox(ctext(unsuccessfully_finished) + " of " + ctext(core_count) + iif(unsuccessfully_finished = 1, "was", "were") + " terminated prematurely. Do want to continue with merging?",4,"Some scripts failed!") = 7 then
			delog(ctext(unsuccessfully_finished) + " OF " + ctext(core_count) + " SCRIPTS " + iif(unsuccessfully_finished = 1, "WAS", "WERE") + " TERMINATED PREMATURELY. MERGE POSTPONED!")
			merge_cond = false
'		end if
	end if
end if

'canceling merge if a there's errors in the cleaning
dim flag_bad_cases
flag_bad_cases = false

if Standard_IIS_DMS and merge_cond then
	for i = 1 to core_count
		if fso.FileExists(DMS_Name+"_"+ctext(i)+"\Output\BadCases.csv") then
			if datediff(fso.GetFile(DMS_Name+"_"+ctext(i)+"\Output\BadCases.csv").DateLastModified, StartTime, "s") < 0 then 'file is new
				delog("There are bad cases in " + DMS_Name+"_"+ctext(i))
				flag_bad_cases = true
			end if
		end if
	next
	
	if flag_bad_cases then
		merge_cond = false
		delog("MERGE POSTPONED BECAUSE OF INVALIDATED RESPONDENTS!")
	end if
end if


if merge_cond then
	
	delog("Start merging the outputs.")
	
	if fso.FileExists("Merged_Output.ddf") then fso.DeleteFile("Merged_Output.ddf")
	if fso.FileExists("Merged_Output.dzf") then fso.DeleteFile("Merged_Output.dzf")
	if fso.FileExists("Merged_Output.mdd") then fso.DeleteFile("Merged_Output.mdd")
	
	dim merge_dms_txt, file_count
	merge_dms_txt = ""
	file_count = 0
	
	for i = 1 to 16
		if i <= core_count then
			QA_Folder = DMS_Name + "_" + ctext(i) + "\Output\"
			if fso.FileExists(QA_Folder + "Output.mdd") and fso.FileExists(QA_Folder + "Output.ddf") then 'check with missing folder
				file_count = file_count + 1
				merge_dms_txt = merge_dms_txt + "Inputdatasource(input" + ctext(i) + ", """ + "Input Data Source Number " + ctext(i) + """)" + mr.CrLf
				merge_dms_txt = merge_dms_txt + mr.Tab + "ConnectionString = ""Provider = mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location = " + QA_Folder + "Output.ddf; Initial Catalog = " + QA_Folder + "Output.mdd; MR Init Category Names = 1""" + mr.crlf
				merge_dms_txt = merge_dms_txt + mr.Tab + "SelectQuery = """ + "Select * from vdata" + """" + mr.crlf
				merge_dms_txt = merge_dms_txt + "End Inputdatasource"
				merge_dms_txt = merge_dms_txt + mr.CrLf + mr.CrLf
			else
				delog(QA_Folder + "Output.mdd/" + QA_Folder + "Output.mdd/ddf doesn't exists.")
			end if
		end if
	next
	
'	if file_count = 0 then merge_cond = false '???????
	if file_count < core_count then merge_cond = false
	
	if merge_cond then
		'adding output
		merge_dms_txt = merge_dms_txt + "Outputdatasource(output)" + mr.CrLf
		merge_dms_txt = merge_dms_txt + mr.Tab + "ConnectionString = ""Provider = mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location = Merged_Output.ddf; Initial Catalog = Merged_Output.mdd""" + mr.CrLf
		merge_dms_txt = merge_dms_txt + mr.Tab + "metadataOutputName = """+"Merged_Output.mdd"+"""" + mr.CrLf
		merge_dms_txt = merge_dms_txt + "End Outputdatasource" + mr.CrLf
		merge_dms_txt = merge_dms_txt + mr.CrLf + mr.CrLf
		
		'adding OnBeforeJobStart
		merge_dms_txt = merge_dms_txt + "Event(OnBeforeJobStart, ""Delete existing files"")" + mr.CrLf
		merge_dms_txt = merge_dms_txt + mr.Tab + "dim fso" + mr.CrLf
		merge_dms_txt = merge_dms_txt + mr.Tab + "set fso = CreateObject(""Scripting.FileSystemObject"")" + mr.CrLf
		merge_dms_txt = merge_dms_txt + mr.crlf
		merge_dms_txt = merge_dms_txt + mr.Tab + "if fso.FileExists(""Merged_Output.ddf"") Then fso.DeleteFile(""Merged_Output.ddf"")" + mr.crlf
		merge_dms_txt = merge_dms_txt + mr.Tab + "if fso.FileExists(""Merged_Output.dzf"") Then fso.DeleteFile(""Merged_Output.dzf"")" + mr.crlf
		merge_dms_txt = merge_dms_txt + mr.Tab + "if fso.FileExists(""Merged_Output.mdd"") Then fso.DeleteFile(""Merged_Output.mdd"")" + mr.crlf
		merge_dms_txt = merge_dms_txt + "End Event" + mr.crlf
	
		'write Merge_Outputs.dms
		if fso.FileExists("Merge_Outputs.dms") then fso.DeleteFile("Merge_Outputs.dms")
		
		dim merge_dms_file
		set merge_dms_file = fso.OpenTextFile("Merge_Outputs.dms",8,true,0)
		merge_dms_file.Write(merge_dms_txt)
		merge_dms_file.close()
		
		'run Merge_Outputs.dms
		dim prev_out, curr_out, out_txt, prev_out_txt
		dim Wrun_merge
		set Wrun_merge = wshell.exec("dmsrun ""Merge_Outputs.dms""")
		
		script_running = true
		while script_running
			'status
			status_str = Wrun_merge.Status.ctext()
			exit_code = Wrun_merge.ExitCode.ctext()
			
			'ExitCode = 0 --> Program successfully executed. -1 --> program terminated prematurely.
			'Status 0 = running, 1 finished, 2 failed
			
			if find(status_str,"0")<>-1 then
				script_running = true
				
				if exit_code = "0" then					
					prev_out = curr_out
					curr_out = Timenow().CText() + ": Merge_Outputs: " + Wrun_merge.stdout.readline().rtrim().Replace(prev_out,"")
					out_txt = out_txt + curr_out + mr.CrLf
				end if
				
				set out_file = fso.OpenTextFile("mrs_log.txt",8,true,0)
				out_file.Write(out_txt.Replace(prev_out_txt,""))		
				out_file.close()
				prev_out_txt = out_txt
				
			else
				script_running = false
				delog("Outputs should be merged now. DB name = 'Merged_Output.mdd/dzf'.")
			end if
		end while
	end if
end if
end section
end if 'Standard_IIS_DMS



set fso = null
set mdm = null
set wshell = null
set adoCommand = null
'set ScriptEngine = null
'Set LogObj = null



'-------- Finish Time
dim FinishTime, TimeTaken, TimeTaken_temp, TimeTaken_hour, TimeTaken_min, TimeTaken_sec
FinishTime = TimeNow()

TimeTaken_temp = DateDiff(StartTime.TimeOnly(),FinishTime,"s")

TimeTaken_hour = (TimeTaken_temp/60)/60
TimeTaken_min = (TimeTaken_temp - TimeTaken_hour*3600)/60
TimeTaken_sec = TimeTaken_temp - TimeTaken_hour*3600 - TimeTaken_min*60

TimeTaken = iif(TimeTaken_hour<10,"0"+TimeTaken_hour.CText(),TimeTaken_hour.CText())+":"+iif(TimeTaken_min<10,"0"+TimeTaken_min.ctext(),TimeTaken_min.ctext())+":"+iif(TimeTaken_sec<10,"0"+TimeTaken_sec.CText(),TimeTaken_sec.CText())

delog("Job ends.")
delog("Job done for "+ctext(TimeTaken)+".")
delog("NOW CLOSE THIS POWERSHELL WINDOW.")


'----------------- Functions
function WritePoints(premsg,iterator,postmsg)
	dim i, points, StartTime
	StartTime = timenow()
	for i =1 to iterator
		points = points + "."
	next
	delog(premsg + points + ctext(iterator) + postmsg)
end function

function StartDMS(dmsname, iterator, aff)
	
	
	dim k, txt_file_text, txt_file, fso, test_txt, wshell
	set fso = createobject("Scripting.FileSystemObject")
	Set wshell = CreateObject("WScript.Shell")
	k = hex(pow(2,(iterator-1)))
	
	'------ Starting the dms
	if aff then 'DEPRECATED
		'----- old not updated section
			txt_file_text = "start /Affinity "+ctext(k)+" DMSRun "+DMS_Name+"_"+ctext(iterator)+"\"+dmsname+".dms"+mr.crlf
			txt_file_text = txt_file_text + "if NOT ["+"""%errorlevel%"""+"]==["+"""0"""+"] pause"+mr.crlf
			
			'------- Making the cmd command
			set txt_file = fso.OpenTextFile("run_dms.cmd",2,true,0)
			txt_file.Write(txt_file_text)
			txt_file.close()
			delog(DMS_Name+"_"+ctext(iterator)+" is ready. Running ...")
			shellexecute("run_dms.cmd",,,,,5)
		'----
	else
		delog(DMS_Name+"_"+ctext(iterator)+" is ready. Running ...")
		wshell.run("dmsrun.exe "+DMS_Name+"_"+ctext(iterator)+"\"+dmsname+".dms",cmd_popup,false)
	end if
	
	set fso = null
	set wshell = null
	sleep(1000)

end function

function Delog(msg)
	dim msg_txt, fso, out_file, prev_msg_txt, file_name
	msg_txt = ctext(timenow()) + ": " + msg
	
	if msg <> "NOW CLOSE THIS POWERSHELL WINDOW." then debug.Log(msg_txt)
	
	file_name = "mrs_log.txt"
	
	set fso = createobject("Scripting.FileSystemObject")
	set out_file = fso.OpenTextFile(file_name,8,true,0)
		out_file.Write(msg_txt.Replace(prev_msg_txt,"")+ mr.CrLf)
		out_file.close()
		prev_msg_txt = msg_txt
end function



