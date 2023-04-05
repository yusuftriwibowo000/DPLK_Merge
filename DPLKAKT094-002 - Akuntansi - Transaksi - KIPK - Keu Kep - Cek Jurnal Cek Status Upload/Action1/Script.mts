Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username,preperation

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKAKT094-002 - Akuntansi - Transaksi - KIPK - Keu Kep - Cek Jurnal Cek Status Upload.xlsx", "DPLKAKT094-002")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
preperation = Split(DataTable.Value("PREPERATION",dtlocalsheet),",")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, preperation)
Iteration = Environment.Value("ActionIteration")
REM ------- DPLK

User 			= Split(DataTable.value("USERID",dtlocalsheet),",")
Password		= DataTable.Value("PASSWORD",dtlocalsheet)
Keyword_Search	= Split(DataTable.value("KEYWORD_SEARCH",dtlocalsheet),",")
Direct_Menu		= Split(DataTable.value("DIRECT_MENU",dtlocalsheet),",")
Nilai_PPH_Or_Akum	= ""

Call DA_Login_Batch_No_SS(User(0),Password)
Call Ambil_Jumlah_Baris_Search_Entry_Jurnal_Transaksi
Iterator_Max = Ambil_Jumlah_Baris_Search_Entry_Jurnal_Transaksi
For Iterator = 1+2 To Iterator_Max Step 1
	Global_Total_Debit = 0
	Global_Total_Credit = 0	
	Call Ambil_Kode_Jurnal_Standar_Using_Row_Nmuber(Iterator)
	Call Ambil_No_Akun_Dari_Table_Using_Row_Number(Iterator)
	Call Ambil_Total_Debit_Credit	
	If Nilai_PPH_Or_Akum = "" Then
		Nilai_PPH_Or_Akum = Global_Total_Credit
	Else 
		Nilai_PPH_Or_Akum = Nilai_PPH_Or_Akum & "," & Global_Total_Credit
	End If
Next
Nilai = Split(Nilai_PPH_Or_Akum,",")
Call DA_Logout_No_SS("0")
Call DA_Login_Batch(User(1),Password)
For Iteratorr = 1+2 To Iterator_Max Step 1
	Call AC_Direct_GoTo_Menu(DataTable.Value("SIDEBAR_SUBMENU_SUBMENU",dtlocalsheet) ,1)
	Call Lihat_Inquiry_Pembayaran_Kepesertaan_With_Row_Number(Iteratorr-1)
	Call Bandingkan_Inquiry_Pembayaran_Kepesertaan_Global_With_Array(Nilai)
Next
Call DA_Logout("0")	

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
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Investasi_PasarUang.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_LogMenu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Keuangan_Transaksi.qfl")
	
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
	Call RepositoriesCollection.Add(LibRepo & "RP_Keuangan_Transaksi.tsr")
	
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
