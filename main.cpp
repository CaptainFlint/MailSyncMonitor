#include <Windows.h>
#include <tlhelp32.h>

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

bool CheckProcess(LPCWSTR exeName)
{
	bool res = false;

	PROCESSENTRY32 entry;
	entry.dwSize = sizeof(PROCESSENTRY32);

	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);

	if (Process32First(snapshot, &entry) == TRUE)
	{
		while (Process32Next(snapshot, &entry) == TRUE)
		{
			if (_wcsicmp(entry.szExeFile, exeName) == 0)
			{
				res = true;
				break;
			}
		}
	}
	CloseHandle(snapshot);
	return res;
}

bool CheckSyncDisk()
{
	return (GetFileAttributes(L"Q:\\MailSync") != INVALID_FILE_ATTRIBUTES);
}

void Run(LPCWSTR cmdline, LPCWSTR msg)
{
	if ((msg == NULL) || (MessageBox(NULL, msg, L"Mail Sync Monitor", MB_OKCANCEL | MB_ICONQUESTION | MB_SYSTEMMODAL) == IDOK))
	{
		STARTUPINFO si = { 0 };
		PROCESS_INFORMATION pi = { 0 };
		si.cb = sizeof(si);
		wchar_t* cmdline_rw = _wcsdup(cmdline);

		if (CreateProcess(NULL, cmdline_rw, NULL, NULL, FALSE, CREATE_UNICODE_ENVIRONMENT, NULL, NULL, &si, &pi))
		{
			CloseHandle(pi.hProcess);
			CloseHandle(pi.hThread);
		}
		else
			MessageBox(NULL, L"Failed to start command!\nPlease, start it manually.", L"Mail Sync Monitor", MB_OK | MB_ICONERROR | MB_SYSTEMMODAL);
		free(cmdline_rw);
	}
}

int CALLBACK wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR /*lpCmdLine*/, int nCmdShow)
{
	UNREFERENCED_PARAMETER(hInstance);
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(nCmdShow);

	bool TheBatRunning = false;
	bool SyncDiskPresent = CheckSyncDisk();
	while (true)
	{
		bool test;
		test = CheckProcess(L"thebat.exe");
		if (TheBatRunning && !test)
		{
			// The Bat! was running and just exited - start synchronization
			Run(L"C:\\Windows\\System32\\WScript.exe T:\\AppData\\mail_sync.js backup", L"Start The bat! backup?");
		}
		TheBatRunning = test;

		test = CheckSyncDisk();
		if (!SyncDiskPresent && test)
		{
			// USB Flash with sync data was inserted - start restoring
			Run(L"C:\\Windows\\System32\\WScript.exe T:\\AppData\\mail_sync.js restore", L"Start The bat! restore?");
		}
		SyncDiskPresent = test;
		Sleep(1000);
	}
	return 0;
}