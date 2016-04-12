#include <Windows.h>
#include <Commctrl.h>
#include <tlhelp32.h>
#include <string>
#include <vector>
#include <algorithm>

#include "resource.h"

using namespace std;

#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif

// Global variables
const wchar_t* AppTitle = L"Mail Sync Monitor";
const size_t MAX_PATH_EX = 1024;

HINSTANCE hAppInstance;         // Application instance
HICON hIcon32 = nullptr;        // Application icon 32x32
HICON hIcon16 = nullptr;        // Application icon 16x16

// Application options stored in the INI file
struct Options
{
	wchar_t* DataDir;       // Directory with data to backup
	wchar_t* SyncDir;       // Directory to backup into
	wchar_t* SyncList;		// Path to the list file with files to backup (for WinRAR)
	wchar_t* Password;		// Password to encrypt the archive with
	wchar_t* BackupSuffix;	// Suffix to add to the archive name
	size_t   MaxBackupsNum;	// Maximum number of latest backup archives to keep

	Options() {
		DataDir       = new wchar_t[MAX_PATH_EX];
		SyncDir       = new wchar_t[MAX_PATH_EX];
		SyncList      = new wchar_t[MAX_PATH_EX];
		Password      = new wchar_t[MAX_PATH_EX];
		BackupSuffix  = new wchar_t[MAX_PATH_EX];
		MaxBackupsNum = 10;
	}

	~Options() {
		if (DataDir)
			delete[] DataDir;
		if (SyncDir)
			delete[] SyncDir;
		if (SyncList)
			delete[] SyncList;
		if (Password)
			delete[] Password;
		if (BackupSuffix)
			delete[] BackupSuffix;
	}
};

DWORD ShowMsgBox(const wstring& msg, DWORD Flags, HWND parent = nullptr)
{
	return MessageBox(parent, msg.c_str(), AppTitle, Flags | MB_SYSTEMMODAL);
}

// Checks whether the process with the specified EXE name is running
bool CheckProcess(const wstring& exeName)
{
	bool res = false;

	PROCESSENTRY32 entry;
	entry.dwSize = sizeof(PROCESSENTRY32);

	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

	if (Process32First(snapshot, &entry) == TRUE)
	{
		while (Process32Next(snapshot, &entry) == TRUE)
		{
			if (_wcsicmp(entry.szExeFile, exeName.c_str()) == 0)
			{
				res = true;
				break;
			}
		}
	}
	CloseHandle(snapshot);
	return res;
}

// Checks whether the specified path is available (e.g. USB flash disk inserted)
bool CheckSyncDisk(const wstring& Path)
{
	return (GetFileAttributes(Path.c_str()) != INVALID_FILE_ATTRIBUTES);
}

// Gets the latest backup archive name and, optionally, removes old archives
wstring GetLatestBackup(const wstring& SyncDir, bool Cleanup, size_t MaxBackupsNum = 0)
{
	vector<wstring> BackupFiles;
	WIN32_FIND_DATA FindFileData;

	// List all available archives
	HANDLE hFF = FindFirstFile((SyncDir + L"\\*.rar").c_str(), &FindFileData);
	if (hFF == INVALID_HANDLE_VALUE)
		return L"";
	BackupFiles.push_back(FindFileData.cFileName);
	while (FindNextFile(hFF, &FindFileData) != 0)
	{
		BackupFiles.push_back(FindFileData.cFileName);
	}
	FindClose(hFF);

	// Sort the list
	sort(BackupFiles.begin(), BackupFiles.end());

	// If necessary, delete the old archives
	if (Cleanup && (BackupFiles.size() > MaxBackupsNum))
	{
		for (size_t i = 0; i < BackupFiles.size() - MaxBackupsNum; ++i)
			DeleteFile((SyncDir + L"\\" + BackupFiles[i]).c_str());
	}

	// Return the name of the latest archive
	return BackupFiles[BackupFiles.size() - 1];
}

// Constructs backup archive name so that it was higher than LatestArchName.
// Current date is used as the starting point.
wstring ConstructArchName(const wstring& LatestArchName, const wstring& Suffix)
{
	SYSTEMTIME CurrentTime;
	GetLocalTime(&CurrentTime);
	wchar_t archName[128];
	// Start from the current date. If there is already archive with same name, increase the day and try again.
	// We can go over the month boundaries - no need to keep the exact date in the archive name.
	do {
		swprintf_s(archName, L"Mail-%04d-%02d-%02d%s.rar", CurrentTime.wYear, CurrentTime.wMonth, CurrentTime.wDay, Suffix.c_str());
	} while ((LatestArchName.compare(archName) > 0) && (++CurrentTime.wDay < 100));
	if (CurrentTime.wDay >= 100)
	{
		// Failed to construct the name for reasonable amount of steps - return error
		return L"";
	}
	return archName;
}

// Runs a command line, optionally requesting user consent for it
bool Run(const wstring& cmdline, const wstring& msg = L"")
{
	bool res = true;
	if ((msg == L"") || (ShowMsgBox(msg.c_str(), MB_OKCANCEL | MB_ICONQUESTION) == IDOK))
	{
		STARTUPINFO si = { 0 };
		PROCESS_INFORMATION pi = { 0 };
		si.cb = sizeof(si);
		// Duplicate the command line as read-write - CreateProcess requires it for some reason
		wchar_t* cmdline_rw = _wcsdup(cmdline.c_str());

		if (CreateProcess(nullptr, cmdline_rw, nullptr, nullptr, FALSE, CREATE_UNICODE_ENVIRONMENT, nullptr, nullptr, &si, &pi))
		{
			// Waiting till the process finishes and check the exit code
			WaitForSingleObject(pi.hProcess, INFINITE);
			DWORD ExitCode = 0;
			BOOL GetCodeStatus = GetExitCodeProcess(pi.hProcess, &ExitCode);
			if (!GetCodeStatus || (ExitCode != 0))
			{
				// Something went wrong, report it
				wchar_t StatusMsg[128];
				if (GetCodeStatus)
					swprintf_s(StatusMsg, L"Command failed with code 0x%08x!", ExitCode);
				else
					swprintf_s(StatusMsg, L"Failed to obtain exit code of the process (error: 0x%08x)!", GetLastError());
				ShowMsgBox(StatusMsg, MB_OK | MB_ICONERROR);
				res = false;
			}
			CloseHandle(pi.hProcess);
			CloseHandle(pi.hThread);
		}
		else
		{
			// Something went wrong real bad!
			ShowMsgBox(L"Failed to start command!\nPlease, start it manually.", MB_OK | MB_ICONERROR);
			res = false;
		}
		free(cmdline_rw);
		return res;
	}
	// User cancelled the action
	return false;
}

// Thread synchronization data
enum ThreadStatus
{
	NORMAL,
	PAUSE,
	ABORT
};
volatile ThreadStatus statusRequested = NORMAL;

CRITICAL_SECTION threadStatusCS;
void SetThreadStatus(ThreadStatus st)
{
	EnterCriticalSection(&threadStatusCS);
	statusRequested = st;
	LeaveCriticalSection(&threadStatusCS);
}
ThreadStatus GetThreadStatus()
{
	ThreadStatus res;
	EnterCriticalSection(&threadStatusCS);
	res = statusRequested;
	LeaveCriticalSection(&threadStatusCS);
	return res;
}

// Dialog function to display the progress dialog when moving file
INT_PTR CALLBACK ProgressDlgFunc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (uMsg)
	{
	case WM_INITDIALOG:
		hIcon16 = (HICON)LoadImage(hAppInstance, MAKEINTRESOURCE(IDI_MAINICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
		hIcon32 = (HICON)LoadImage(hAppInstance, MAKEINTRESOURCE(IDI_MAINICON), IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR);
		SendMessage(hwndDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon16);
		SendMessage(hwndDlg, WM_SETICON, ICON_BIG, (LPARAM)hIcon32);
		SetWindowText(hwndDlg, AppTitle);
		return TRUE;
	case WM_DESTROY:
		DestroyIcon(hIcon32);
		DestroyIcon(hIcon16);
		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDCANCEL:
			// User clicked the [X] close title button
			SetThreadStatus(PAUSE);
			if (ShowMsgBox(L"Abort?", MB_OKCANCEL | MB_ICONQUESTION, hwndDlg) == IDOK)
			{
				SetThreadStatus(ABORT);
				DestroyWindow(hwndDlg);
				PostQuitMessage(0);
			}
			else
				statusRequested = NORMAL;
			return TRUE;
		case IDOK:
			// Signal about finished operation
			DestroyWindow(hwndDlg);
			PostQuitMessage(0);
			return TRUE;
		}
	}
	return FALSE;
}

// Data to pass into the thread
struct CopyFuncData
{
	HANDLE src;     // Source file
	HANDLE dst;     // Target file
	HWND wnd;       // Progress dialog
};

// Thread function for copying file
DWORD WINAPI CopyFunc(LPVOID lpParameter)
{
	DWORD res = 0;
	CopyFuncData* th_data = (CopyFuncData*)lpParameter;
	HWND hp = GetDlgItem(th_data->wnd, IDC_PROGRESS_BAR);
	const size_t BUF_SIZE = 8 * 1024 * 1024;
	BYTE* buf = new BYTE[BUF_SIZE];
	DWORD br, bw;
	size_t total = 0;

	// Starting read-write cycle
	while (ReadFile(th_data->src, buf, BUF_SIZE, &br, nullptr) && (br > 0))
	{
		total += br;
		while (!WriteFile(th_data->dst, buf, br, &bw, nullptr) || (bw != br))
		{
			if (ShowMsgBox(L"Write operation failed, retry?", MB_OKCANCEL | MB_ICONQUESTION, th_data->wnd) == IDCANCEL)
			{
				PostQuitMessage(0);
				return 1;
			}
		}
		SendMessage(hp, PBM_SETPOS, total / (1024 * 1024), 0);

		// Check whether pause or abort was requested
		ThreadStatus st;
		while ((st = GetThreadStatus()) == PAUSE)
		{
			// Wait for unpause
			Sleep(100);
		}
		if (st == ABORT)
		{
			res = 1;
			break;
		}
	}
	SendMessage(th_data->wnd, WM_COMMAND, IDOK, 0);
	return res;
}

// Perform all the gore details with moving the file:
// open files, show the progress dialog, start copying thread, perform message loop
// until the dialog closes, and if copy was successful, delete the source file.
bool MoveBackup(wstring SrcPath, wstring DstPath)
{
	bool res = true;
	CopyFuncData th_data = { nullptr, nullptr, nullptr };
	HANDLE thread = nullptr;

	try
	{
		wchar_t msg[1024];

		// Open source file for reading
		wstring FileName = SrcPath.substr(SrcPath.rfind(L'\\') + 1);
		th_data.src = CreateFile(
			SrcPath.c_str(),
			GENERIC_READ,
			FILE_SHARE_READ,
			nullptr,
			OPEN_EXISTING,
			FILE_FLAG_SEQUENTIAL_SCAN,
			nullptr);
		if (th_data.src == INVALID_HANDLE_VALUE)
		{
			swprintf_s(msg, L"Failed to open file «%s» for reading: %d!", SrcPath.c_str(), GetLastError());
			throw msg;
		}

		// Open target file for writing
		th_data.dst = CreateFile(
			DstPath.c_str(),
			GENERIC_WRITE,
			FILE_SHARE_READ,
			nullptr,
			OPEN_ALWAYS,
			FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN | FILE_FLAG_WRITE_THROUGH,
			nullptr);
		if (th_data.dst == INVALID_HANDLE_VALUE)
		{
			swprintf_s(msg, L"Failed to open file «%s» for writing: %d!", DstPath.c_str(), GetLastError());
			throw msg;
		}

		// Open, initialise and show the progress dialog
		th_data.wnd = CreateDialog(nullptr, MAKEINTRESOURCE(IDD_PROGRESS_DLG), nullptr, ProgressDlgFunc);
		if (th_data.wnd == nullptr)
		{
			swprintf_s(msg, L"Failed to create progress dialog, code: %d!", GetLastError());
			throw msg;
		}

		SetDlgItemText(th_data.wnd, IDC_MESSAGE, (L"Moving «" + FileName + L"»...").c_str());
		LARGE_INTEGER srcSize;
		if (!GetFileSizeEx(th_data.src, &srcSize))
		{
			swprintf_s(msg, L"Failed to obtain archive file size, code: %d!", GetLastError());
			throw msg;
		}
		HWND hp = GetDlgItem(th_data.wnd, IDC_PROGRESS_BAR);
		SendMessage(hp, PBM_SETRANGE32, 0, (srcSize.QuadPart / (1024 * 1024)));
		ShowWindow(th_data.wnd, SW_SHOWNORMAL);

		// Start copying thread
		thread = CreateThread(nullptr, 0, CopyFunc, &th_data, 0, nullptr);
		if (thread == nullptr)
		{
			swprintf_s(msg, L"Failed to create thread, code: %d!", GetLastError());
			throw msg;
		}

		// Dialog message loop
		BOOL bRet;
		MSG umsg;
		while ((bRet = GetMessage(&umsg, NULL, 0, 0)) != 0)
		{
			if (bRet == -1)
			{
				// Handle the error and possibly exit
			}
			else if (!IsWindow(th_data.wnd) || !IsDialogMessage(th_data.wnd, &umsg))
			{
				TranslateMessage(&umsg);
				DispatchMessage(&umsg);
			}
		}

		// Copy procedure finished or was aborted: check the status
		CloseHandle(th_data.src);
		CloseHandle(th_data.dst);
		th_data.src = th_data.dst = nullptr;

		DWORD th_res;
		if (!GetExitCodeThread(thread, &th_res))
		{
			swprintf_s(msg, L"Failed to obtain thread status, code: %d!", GetLastError());
			throw msg;
		}
		else if (th_res != 0)
			res = false;

		if (res)
		{
			// Successful completion, delete the source file
			if (!DeleteFile(SrcPath.c_str()))
			{
				swprintf_s(msg, L"Failed to delete file «%s», code: %d!", SrcPath.c_str(), GetLastError());
				throw msg;
			}
		}
	}
	catch (wstring err)
	{
		ShowMsgBox(err, MB_OK | MB_ICONERROR);
		res = false;
	}

	if (th_data.src)
		CloseHandle(th_data.src);
	if (th_data.dst)
		CloseHandle(th_data.dst);
	if (thread)
		CloseHandle(thread);
	return res;
}

int CALLBACK wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	UNREFERENCED_PARAMETER(nCmdShow);
	hAppInstance = hInstance;
	InitializeCriticalSection(&threadStatusCS);

	// Construct path to the configuration file
	wchar_t* IniPath = new wchar_t[MAX_PATH_EX];
	size_t len;
	if ((len = GetModuleFileName(nullptr, IniPath, MAX_PATH_EX)) == 0)
	{
		ShowMsgBox(L"Failed to get INI path.", MB_OK | MB_ICONERROR);
		return 1;
	}
	wcscpy_s(IniPath + len - 3, 4, L"ini");

	// Read options
	Options opts;
	GetPrivateProfileString(L"General", L"DataDir", L"T:\\AppData", opts.DataDir, MAX_PATH_EX, IniPath);
	GetPrivateProfileString(L"General", L"SyncDir", L"Q:\\MailSync", opts.SyncDir, MAX_PATH_EX, IniPath);
	GetPrivateProfileString(L"General", L"SyncList", L"T:\\AppData\\bat_sync_list.txt", opts.SyncList, MAX_PATH_EX, IniPath);
	GetPrivateProfileString(L"General", L"Password", L"", opts.Password, MAX_PATH_EX, IniPath);
	GetPrivateProfileString(L"General", L"BackupSuffix", L"", opts.BackupSuffix, MAX_PATH_EX, IniPath);

	// Set current path to the DataDir so that WinRAR used relative paths
	SetCurrentDirectory(opts.DataDir);

	// Start monitoring cycle
	bool TheBatRunning = false;
	bool SyncDiskPresent = CheckSyncDisk(opts.SyncDir);
	while (true)
	{
		bool test;
		test = CheckProcess(L"thebat.exe");
		if (TheBatRunning && !test)
		{
			// The Bat! was running and just exited - start synchronization
			if (ShowMsgBox(L"Start The Bat! backup?", MB_OKCANCEL) == IDOK)
			{
				wstring LatestArchName = GetLatestBackup(opts.SyncDir, true, opts.MaxBackupsNum - 1);
				wstring ArchName = ConstructArchName(LatestArchName, opts.BackupSuffix);
				if (ArchName == L"")
					ShowMsgBox(L"Failed to construct archive name higher than «" + LatestArchName + L"»!\nPlease, start backup manually.", MB_OK | MB_ICONERROR);
				else
				{
					// Try to create archive in temp directory and then move it,
					// but if tempdir fails for some reason just pack directly into the target sync dir
					wchar_t* TempDir = new wchar_t[MAX_PATH_EX];
					if (GetTempPath(MAX_PATH_EX, TempDir) == 0)
					{
						delete[] TempDir;
						TempDir = nullptr;
					}
					wstring cmdline = wstring(L"\"C:\\Program Files\\WinRAR\\WinRAR.exe\" a -cfg- -m1 -rr3p -s -p") + opts.Password + L" -r -x\"The Bat!\\IspRas\\\" \"" + (TempDir ? TempDir : wstring(opts.SyncDir) + L"\\") + ArchName + L"\" @\"" + opts.SyncList + L"\"";
					if (Run(cmdline))
					{
						// If necessary, move the archive from temp directory to the sync directory
						if (!TempDir || MoveBackup(wstring(TempDir) + L"\\" + ArchName, wstring(opts.SyncDir) + L"\\" + ArchName))
							ShowMsgBox(L"Mail archive backup complete:\n" + ArchName, MB_OK | MB_ICONINFORMATION);
					}
					if (TempDir)
						delete[] TempDir;
				}
			}
		}
		TheBatRunning = test;

		test = CheckSyncDisk(opts.SyncDir);
		if (!SyncDiskPresent && test)
		{
			// USB Flash with sync data was inserted - start restoring
			wstring LatestArchName = GetLatestBackup(opts.SyncDir, false);
			wstring ending = wstring(opts.BackupSuffix) + L".rar";
			wstring msg;
			if (LatestArchName.substr(LatestArchName.size() - ending.size()) == ending)
				msg = L"WARNING!\n\nThe latest archive is from the same computer (" + LatestArchName + L")!\nRestore anyway?";
			else
				msg = L"Archive «" + LatestArchName + L"» is going to be restored. Continue?";
			wstring cmdline = wstring(L"\"C:\\Program Files\\WinRAR\\WinRAR.exe\" x -cfg- -o+ -p") + opts.Password + L" \"" + opts.SyncDir + L"\\" + LatestArchName + L"\" \"" + opts.DataDir + L"\\\"";
			if (Run(cmdline, msg))
				ShowMsgBox(L"Mail archive restore complete:\n" + LatestArchName, MB_OK | MB_ICONINFORMATION);
		}
		SyncDiskPresent = test;
		Sleep(1000);
	}
	delete[] IniPath;
	DeleteCriticalSection(&threadStatusCS);
	return 0;
}