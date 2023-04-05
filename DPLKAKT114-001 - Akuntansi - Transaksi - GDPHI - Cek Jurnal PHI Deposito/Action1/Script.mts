Dim dt_Username,preperation
Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKAKT114-001 - Akuntansi - Transaksi - GDPHI - Cek Jurnal PHI Deposito.xlsx", "DPLKAKT114-001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
preperation = Split(DataTable.Value("PREPERATION",dtlocalsheet),",")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, preperation)
Iteration = Environment.Value("ActionIteration")
User = DataTable.Value("USERID",dtlocalsheet)
REM ------- DPLK
	User = Split(DataTable.Value("USERID",dtlocalsheet),",")
	Call DA_Login_Batch_No_SS(User(0),DataTable.Value("PASSWORD",dtlocalsheet))
	Call Ambil_No_Akun_Dari_Table()
	Call Ambil_Total_Debit_Credit()
	Call AC_Logout_No_SS()
	
	Call DA_Login_Batch(User(1),DataTable.Value("PASSWORD",dtlocalsheet))
	Call AC_GoTo_Menu()
	Call Lihat_Investasi_Investasi_Umum_Generate_Dayend()
'	Call Bandingkan_No_Akun_Investasi_Investasi_Umum_Setup_Jurnal_Standar()
	Call Compare_Debit_And_Credit_Global_Antara_Entry_Journal_dan_Investasi_Investasi_Umum()
	Call AC_Logout()

Call Reset_Global_Var()
Call spReportSave()
	
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
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_LogMenu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Investasi_PasarUang.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Kepesertaan_Inquiry.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Investasi_InvestasiUmum.qfl")
	
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
	Call RepositoriesCollection.Add(LibRepo & "RP_Kepesertaan_Inquiry.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Investasi_InvestasiUmum.tsr")
	
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
