﻿Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult
Dim preparation, iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKKPS129-001 until DPLKKPS129-002 - Kepesertaan - Transaksi - Register Penarikan Iuran 10 Persen.xlsx", "DPLKKPS129-001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()

dtPreparation = Split(preparation, ";")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, dtPreparation)
iteration = Environment.Value("ActionIteration")

REM ------- DPLK
Call DA_Login()
Call GoTo_SidebarMenu2()
Call GoTo_SidebarSubMenu()

Call AddPenarikanIuran10Persen()

Call DA_Logout("0")
Call spReportSave()

Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 	= Createobject("Wscript.Network")	
	
	tempDPLKPath 	= Environment.Value("TestDir")
	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	
	LibPathDPLK		= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport		= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo			= PathDPLK & "Lib_Repo_Excel\Repo\"
	
	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")

	REM ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Kepesertaan_Transaksi.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Administration_Dashboard.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Kepesertaan_Transaksi.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")

End Sub

Sub spGetDatatable()
	REM --------- Data
	preparation		 		= DataTable.Value("PREPARATION",dtlocalsheet)
	
	REM --------- Reporting
	dt_TCID					= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc		= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc			= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult		= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)

End Sub
