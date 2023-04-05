Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username,preperation

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKAKT155-001 - Akuntansi - Transaksi - IDTR - RKS Trx Reksadana - Cek Jurnal Reksadana Redemption.xlsx", "DPLKAKT155-001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()	
preperation = Split(DataTable.Value("PREPERATION",dtlocalsheet),",")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, preperation)
Iteration = Environment.Value("ActionIteration")
REM ------- DPLK

Call DA_Login()
Call AC_Direct_GoTo_Menu("Entry Jurnal Transaksi",1)
Call Open_Entry_Jurnal_Transaksi()
Call Ambil_Total_Debit_Credit()
Call AC_Direct_GoTo_Menu(DataTable.Value("SIDEBAR_SUBMENU_SUBMENU",dtlocalsheet),1)
Call Lihat_Investasi_Reksadana_Dealing_Ticket_Reksadana()
Call Compare_Investasi_Reksadana_Dealing_Ticket_Reksadana()
Call DA_Logout("0")
Call Reset_Global_Var()
Call spReportForceSave()

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
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_LogMenu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Investasi_Reksadana.qfl")
	
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
	Call RepositoriesCollection.Add(LibRepo & "RP_Investasi_Reksadana.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Username					= DataTable.Value("USERID",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
End Sub
