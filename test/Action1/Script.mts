﻿Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username,preperation, runCase

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKAKT015-001 - Akuntansi -  Dep.CairJT - Cek Jurnal Pencairan Deposito BNI Gambir KCP MDS DTDEP202200075 PAB-1367586 PRI01-Deposito Pasar Uang.xlsx", "DPLKAKT015")
Call spGetDatatable()
'Call fnRunningIterator()
Call spReportInitiate()

Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, "Cek")
Iteration = Environment.Value("ActionIteration")
REM ------- DPLK

b = 1
For i = 1 To DataTable.GetSheet(dtLocalSheet).GetRowCount-1 STEP 1
If runCase = "RUN" Then
If i <> 1 Then
msgbox "1"
	Call spInitiateData("DPLKLib_Report.xlsx", "DPLKAKT015-001 - Akuntansi -  Dep.CairJT - Cek Jurnal Pencairan Deposito BNI Gambir KCP MDS DTDEP202200075 PAB-1367586 PRI01-Deposito Pasar Uang.xlsx", "DPLKAKT015")
	msgbox "2"
	Call spGetDatatable()
	msgbox "3"
	Call fnRunningIterator()
	msgbox "3.5"
	Call spReportInitiate()
	msgbox "4"
	preperation = Split(DataTable.Value("PREPERATION",dtlocalsheet),",")
	
	Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, preperation)
End If

msgbox i
b = b + 1

Call spReportForceSave()
DataTable.GetSheet(dtLocalSheet).SetNextRow
Call spGetDatatable()
Else
ExitTest
End If
Next

'If Iteration mod 3 = 1 Then
''	Call DA_Login()
'
'	Call spReportSave()
'	Call spReportInitiate()
'ElseIf Iteration mod 3 = 2 Then	
''	Call DA_Login()
''	Call Ambil_Kode_Jurnal_Standar()
''	Call Ambil_No_Akun_Dari_Table()
''	Call AC_GoTo_Menu()
''	Call Lihat_Setup_Akuntansi_Setup_Jurnal_Standar()
''	Call Bandingkan_No_Akun_Setup_Akuntansi_Setup_Jurnal_Standar()
''	Call DA_Logout("0")
'	Call spReportSave()
'	Call spReportInitiate()
'ElseIf Iteration mod 3 = 0 Then
''	Call DA_Login()
''	Call Ambil_Nilai_Debit(2,6)
''	Call Ambil_Dokumen_Induk()
''	Call AC_GoTo_Menu()
''	Call Lihat_Investasi_Pasar_Uang_Dealing_Ticket_Deposito()
''	Call Compare_Debit("Dealing Ticket Deposito")
''	Call DA_Logout("0")
'	Call spReportSave()
'	Call spReportInitiate()
'End If
'


	
Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDPLKPath 	= Environment.Value("TestDir")
	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	
	LibPathDPLK	= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport			= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDPLK & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	rem ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Akuntansi_Laporan.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Akuntansi_Setup.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Akuntansi_Transaksi.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Investasi_FixedIncome.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Investasi_PasarUang.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_LogMenu.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Dashboard.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Log.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Function.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Akuntansi_Laporan.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Akuntansi_Transaksi.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Akuntansi_Setup.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Investasi_FixedIncome.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Investasi_PasarUang.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Username					= DataTable.Value("USERID",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	runCase						= DataTable.Value("RUN", dtLocalSheet)
End Sub
