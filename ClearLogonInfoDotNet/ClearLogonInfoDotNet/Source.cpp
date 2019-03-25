#include <iostream>
#include <iomanip>
#include <string>
#include <algorithm>
#include <fstream>
#include <atlstr.h>
#include <Shlobj.h>
#include <Windows.h>
#include <stdlib.h>
#include <dos.h>
#include <VersionHelpers.h>
#include <conio.h>
#include <direct.h>

#include <odbcinst.h>
#include <sql.h>
#include <sqlext.h>
#include <sqltypes.h>
#include <sqlucode.h>
#include <msdasql.h>
#include <msdadc.h>

using namespace std;

string runningVersion = "v0.0.1";
bool debugOn = false;

int getOmate32();
void header();
void cls();
void pause();
void clearLogonInfo();
void exit();

#ifndef UNICODE  
typedef std::string String;
#else
typedef std::wstring String;
#endif

//System Variables
TCHAR wWinDir[80]; String WinDir;
TCHAR wDataDir[80]; String DataDir;
TCHAR wPgmsDir[80]; String PgmsDir;

//ADOConnection Variables
TCHAR wConnectThru[80]; String ConnectThru;
TCHAR wDatabaseName[80]; String DatabaseName;
TCHAR wDataSource[80]; String DataSource;

//INSTALL Variables
TCHAR wSQLbuild[80]; String SQLbuild;
TCHAR wInstalledVersion[80]; String InstalledVersion;
TCHAR wBuild[80]; String Build;
TCHAR wServicePack[80]; String ServicePack;

//Windows Specific Variables
String rUsername; //USERNAME
String rUserprofile; //USERPROFILE
String rLocalAppData; //LOCALAPPDATA
String rHostname; //COMPUTERNAME
String rSystemRoot; //SYSTEMROOT 

int main() {
	getOmate32();
	clearLogonInfo();
	exit();
}

void header() {
	if (debugOn == true) {
		cout << "Eyefinity Officemate - ClearLogonInfoDotNet (Users Logged In) " << runningVersion << " (Debug ON)" << endl;
		cout << "--------------------------------------------------------------------------------" << string(1, '\n');
	}
	else {
		cout << "Eyefinity Officemate - ClearLogonInfoDotNet (Users Logged In) " << runningVersion << endl;
		cout << "--------------------------------------------------------------------------------" << string(1, '\n');
	}
}

void cls() {

	CHAR fill = ' ';
	COORD tl = { 0,0 };
	CONSOLE_SCREEN_BUFFER_INFO s;
	HANDLE console = GetStdHandle(STD_OUTPUT_HANDLE);

	GetConsoleScreenBufferInfo(console, &s);
	DWORD written, cells = s.dwSize.X * s.dwSize.Y;
	FillConsoleOutputCharacter(console, fill, cells, tl, &written);
	FillConsoleOutputAttribute(console, s.wAttributes, cells, tl, &written);
	SetConsoleCursorPosition(console, tl);
}

void pause() {
	char a;
	cout << string(2, '\n') << "Press any Key to Continue... ";
	a = _getch();


}

void exit() {
	cout << string(1, '\n') << " Press any Key to Exit... ";
	char a;
	a = _getch();
	exit(EXIT_SUCCESS);
}

int getOmate32() {
	//cout << "Gathering Information from C:\\Windows\\Omate32.ini..." << string(2, '\n');

	//System Section
	GetPrivateProfileString(TEXT("System"), TEXT("WinDir"), TEXT("INI NOT FOUND"), wWinDir, 255, TEXT("Omate32.ini"));
	WinDir = wWinDir;
	GetPrivateProfileString(TEXT("System"), TEXT("DataDir"), TEXT("INI NOT FOUND"), wDataDir, 255, TEXT("Omate32.ini"));
	DataDir = wDataDir;
	GetPrivateProfileString(TEXT("System"), TEXT("PgmsDir"), TEXT("INI NOT FOUND"), wPgmsDir, 255, TEXT("Omate32.ini"));
	PgmsDir = wPgmsDir;

	//ADOConnection Section
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("ConnectThru"), TEXT("INI NOT FOUND"), wConnectThru, 255, TEXT("Omate32.ini"));
	ConnectThru = wConnectThru;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DatabaseName"), TEXT("INI NOT FOUND"), wDatabaseName, 255, TEXT("Omate32.ini"));
	DatabaseName = wDatabaseName;
	GetPrivateProfileString(TEXT("ADOConnection"), TEXT("DataSource"), TEXT("INI NOT FOUND"), wDataSource, 255, TEXT("Omate32.ini"));
	DataSource = wDataSource;

	//Install Section
	GetPrivateProfileString(TEXT("Install"), TEXT("SQL_Build"), TEXT("INI NOT FOUND"), wSQLbuild, 255, TEXT("Omate32.ini"));
	SQLbuild = wSQLbuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("InstalledVersion"), TEXT("INI NOT FOUND"), wInstalledVersion, 255, TEXT("Omate32.ini"));
	InstalledVersion = wInstalledVersion;
	GetPrivateProfileString(TEXT("Install"), TEXT("Build"), TEXT("INI NOT FOUND"), wBuild, 255, TEXT("Omate32.ini"));
	Build = wBuild;
	GetPrivateProfileString(TEXT("Install"), TEXT("Service Pack"), TEXT("INI NOT FOUND"), wServicePack, 255, TEXT("Omate32.ini"));
	ServicePack = wServicePack;

	transform(WinDir.begin(), WinDir.end(), WinDir.begin(), (int(*)(int))toupper);
	transform(DataDir.begin(), DataDir.end(), DataDir.begin(), (int(*)(int))toupper);
	transform(PgmsDir.begin(), PgmsDir.end(), PgmsDir.begin(), (int(*)(int))toupper);

	transform(ConnectThru.begin(), ConnectThru.end(), ConnectThru.begin(), (int(*)(int))toupper);
	transform(DatabaseName.begin(), DatabaseName.end(), DatabaseName.begin(), (int(*)(int))toupper);
	transform(DataSource.begin(), DataSource.end(), DataSource.begin(), (int(*)(int))toupper);

	transform(SQLbuild.begin(), SQLbuild.end(), SQLbuild.begin(), (int(*)(int))toupper);
	transform(InstalledVersion.begin(), InstalledVersion.end(), InstalledVersion.begin(), (int(*)(int))toupper);
	transform(ServicePack.begin(), ServicePack.end(), ServicePack.begin(), (int(*)(int))toupper);

	if ((DataDir == "INI NOT FOUND") || (PgmsDir == "INI NOT FOUND")) {
		cout << setw(10) << left << "Omate32.ini not found in C:\\Windows. Please correct Omate32.ini to proceed" << endl;
		cout << setw(10) << left << "Is Officemate\\ExamWriter installed on this PC? Are you cloud hosted?";
		pause();
		exit(1);
	}

	return 0;
}

void showSQLErrorMsg(unsigned int handleType, const SQLHANDLE& handle) {
	//Function shows error message and state derived from preliminary function return message
	SQLCHAR SQLState[1024];
	SQLCHAR message[1024];

	if (SQL_SUCCESS == SQLGetDiagRec(handleType, handle, 1, SQLState, NULL, message, 1024, NULL)) {
		cout << left << setw(10) << "ODBC Driver Message: " << message << string(2, '\n');;
		cout << left << setw(10) << "ODBC SQL State: " << SQLState << string(2, '\n');;
	}
}

void clearLogonInfo() {
	cls();
	header();
	
	cout << "Datasource: " << DataSource << "           Database: " << DatabaseName << string(2, '\n');
	cout << "Displaying table dbo.user_logon_info: ";

	SQLHANDLE SQLEnvHandle = NULL;
	SQLHANDLE SQLConnectionHandle = NULL;
	SQLHANDLE SQLStatementHandle = NULL;
	SQLCHAR retConString[1024];

	string ConnectionString = "DRIVER={SQL Server}; SERVER=" + DataSource + "; DATABASE=" + DatabaseName + ";Uid=OM_USER;Pwd=OMSQL@2004;";

	if (SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &SQLEnvHandle) == SQL_ERROR) {
	}
	if (SQLSetEnvAttr(SQLEnvHandle, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0) == SQL_ERROR) {
	}
	if (SQLAllocHandle(SQL_HANDLE_DBC, SQLEnvHandle, &SQLConnectionHandle) == SQL_ERROR) {
	}
	if (SQLSetConnectAttr(SQLConnectionHandle, SQL_LOGIN_TIMEOUT, (SQLPOINTER)5, 0) == SQL_ERROR) {
	}

	switch (SQLDriverConnect(SQLConnectionHandle, NULL, (SQLCHAR*)ConnectionString.c_str(), SQL_NTS, retConString, 1024, NULL, SQL_DRIVER_NOPROMPT)) {
	case SQL_SUCCESS:
		cout << "Success - SQL_SUCESS" << string(2, '\n');;
		break;
	case SQL_SUCCESS_WITH_INFO:
		cout << "Success - SQL_SUCCESS_WITH_INFO" << string(2, '\n');;
		break;
	case SQL_NO_DATA_FOUND:
		cout << "Error - SQL_NO_DATA_FOUND" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_INVALID_HANDLE:
		cout << "Error - SQL_INVALID_HANDLE" << string(2, '\n');;
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	case SQL_ERROR:
		cout << "Error - SQL_ERROR" << string(2, '\n');
		showSQLErrorMsg(SQL_HANDLE_DBC, SQLConnectionHandle);
		break;
	}

	if (SQLAllocHandle(SQL_HANDLE_STMT, SQLConnectionHandle, &SQLStatementHandle) == SQL_ERROR) {
		cout << left << setw(10) << "Error allocating SQL Statement Handle" << endl;
	}

	char SQLQuery1[] = "SELECT logon_index, logon_machinename, convert(nvarchar(25),logon_datetime), App_Type FROM dbo.user_logon_info";

	// Display table and logins
	if (SQLExecDirect(SQLStatementHandle, (SQLCHAR*)SQLQuery1, SQL_NTS) == SQL_ERROR) {
		cout << left << setw(10) << "Error executing SQL Statement Handle" << endl;
		showSQLErrorMsg(SQL_HANDLE_STMT, SQLStatementHandle);
	}
	else
	{
		int logon_index;
		char logon_machinename[25];
		char logon_datetime[25];
		char App_Type[25];

		cout << setw(15) << left << "logon_index" << setw(25) << "logon_machinename" << setw(25) << left << "logon_datetime" << setw(6) << left << "App_Type" << string(2, '\n');
		while (SQLFetch(SQLStatementHandle) == SQL_SUCCESS) {
			SQLGetData(SQLStatementHandle, 1, SQL_C_DEFAULT, &logon_index, sizeof(logon_index) + 1, NULL);
			SQLGetData(SQLStatementHandle, 2, SQL_C_DEFAULT, &logon_machinename, size(logon_machinename), NULL);
			SQLGetData(SQLStatementHandle, 3, SQL_C_DEFAULT, &logon_datetime, size(logon_datetime), NULL);
			SQLGetData(SQLStatementHandle, 4, SQL_C_DEFAULT, &App_Type, sizeof(App_Type), NULL);
			cout << setw(15) << left << logon_index << setw(25) << logon_machinename << setw(25) << left << logon_datetime << setw(6) << left << App_Type << endl;
		}
	}
	SQLFreeHandle(SQL_HANDLE_STMT, SQLStatementHandle);

	char clearOpt = 'N';
	cout << string(2, '\n');
	cout << " Enter Y to clear Users Logged In, Z to exit: ";
	cin >> clearOpt;
	clearOpt = toupper(clearOpt);

	if (clearOpt == 'Y') {
		// Truncate table
		char SQLQuery2[] = "TRUNCATE TABLE dbo.user_logon_info";

		cout << endl;
		cout << left << setw(10) << " Truncating table dbo.user_logon_info: ";

		if (SQLAllocHandle(SQL_HANDLE_STMT, SQLConnectionHandle, &SQLStatementHandle) == SQL_ERROR) {
			cout << left << setw(10) << "Error allocating SQL Statement Handle" << endl;
		}

		if (SQLExecDirect(SQLStatementHandle, (SQLCHAR*)SQLQuery2, SQL_NTS) == SQL_ERROR) {
			cout << left << setw(10) << "Error executing SQL Statement Handle" << endl;
			showSQLErrorMsg(SQL_HANDLE_STMT, SQLStatementHandle);
		}
		else {
			cout << "Success - SQL_SUCESS" << endl;
			clearLogonInfo();
		}
	}

	if (clearOpt != 'Z' && clearOpt != 'Y') {
		cout << endl << " Invalid character entered. Try again." << endl;
		pause();
		clearLogonInfo();
	}
	else if (clearOpt != 'Z') {
		exit();
	}

	SQLFreeHandle(SQL_HANDLE_STMT, SQLStatementHandle);
	SQLDisconnect(SQLConnectionHandle);
	SQLFreeHandle(SQL_HANDLE_DBC, SQLConnectionHandle);
	SQLFreeHandle(SQL_HANDLE_ENV, SQLEnvHandle);
}

