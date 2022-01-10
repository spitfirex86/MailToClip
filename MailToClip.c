#include "framework.h"
#include "MailToClip.h"
#include "shellapi.h"
#include "ShlObj.h"



#define RegCreateWriteKey(hKey, lpSubKey, phkResult) RegCreateKeyEx((hKey), (lpSubKey), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, (phkResult), NULL)
#define RegSetString(hKey, lpValueName, lpData, cbData) RegSetValueEx((hKey), (lpValueName), 0, REG_SZ, (BYTE const *)(lpData), (cbData))

#define Error(x) MessageBox(NULL, (x), szAppName, MB_OK | MB_ICONERROR)
#define Warn(x) MessageBox(NULL, (x), szAppName, MB_OK | MB_ICONWARNING)
#define Info(x) MessageBox(NULL, (x), szAppName, MB_OK | MB_ICONINFORMATION)


WCHAR const szAppName[] = L"MailToClip";

WCHAR const szMailto[] = L"mailto:";
size_t const cbMailto = sizeof(szMailto);
size_t const cchMailto = ARRAYSIZE(szMailto) - 1;

BOOL RegisterMailtoHandler( void );
void NotifyAssociationChange( void );


BOOL RegisterMailtoHandler( void )
{
	HKEY hRegisteredApps;
	if ( RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\RegisteredApplications", 0, KEY_WRITE, &hRegisteredApps) == ERROR_SUCCESS )
	{
		WCHAR szCapabilitiesPath[] = L"SOFTWARE\\Clients\\Mail\\MailToClip\\Capabilities";
		RegSetString(hRegisteredApps, szAppName, szCapabilitiesPath, sizeof(szCapabilitiesPath));
		RegCloseKey(hRegisteredApps);
	}
	else
	{
		Warn(L"Please run as admin to register the Mailto handler.");
		return FALSE;
	}

	HKEY hClient;
	if ( RegCreateWriteKey(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Clients\\Mail\\MailToClip", &hClient) == ERROR_SUCCESS )
	{
		RegSetString(hClient, NULL, szAppName, sizeof(szAppName));

		HKEY hCapabilities;
		if ( RegCreateWriteKey(hClient, L"Capabilities", &hCapabilities) == ERROR_SUCCESS )
		{
			WCHAR const szAppDesc[] = L"Copy email address to clipboard";
			RegSetString(hCapabilities, L"ApplicationName", szAppName, sizeof(szAppName));
			RegSetString(hCapabilities, L"ApplicationDescription", szAppDesc, sizeof(szAppDesc));

			HKEY hCapsSubKey;
			if ( RegCreateWriteKey(hCapabilities, L"UrlAssociations", &hCapsSubKey) == ERROR_SUCCESS )
			{
				WCHAR szValue[] = L"MailToClip.mailto";
				RegSetString(hCapsSubKey, L"mailto", szValue, sizeof(szValue));
				RegCloseKey(hCapsSubKey);
			}
			if ( RegCreateWriteKey(hCapabilities, L"StartMenu", &hCapsSubKey) == ERROR_SUCCESS )
			{
				RegSetString(hCapsSubKey, L"Mail", szAppName, sizeof(szAppName));
				RegCloseKey(hCapsSubKey);
			}

			RegCloseKey(hCapabilities);
		}

		RegCloseKey(hClient);
	}

	HKEY hClass;
	if ( RegCreateWriteKey(HKEY_CLASSES_ROOT, L"MailToClip.mailto", &hClass) == ERROR_SUCCESS )
	{
		WCHAR szValue[] = L"MailToClip Mailto Handler";
		RegSetString(hClass, NULL, szValue, sizeof(szValue));

		HKEY hCommand;
		if ( RegCreateWriteKey(hClass, L"shell\\open\\command", &hCommand) == ERROR_SUCCESS )
		{
			WCHAR szExePath[MAX_PATH];
			GetModuleFileName(NULL, szExePath, MAX_PATH);

			WCHAR szBuffer[MAX_PATH + 16];
			swprintf(szBuffer, MAX_PATH + 15, L"\"%s\" %%1", szExePath);

			RegSetString(hCommand, NULL, szBuffer, sizeof(szBuffer));
			RegCloseKey(hCommand);
		}

		RegCloseKey(hClass);
	}

	return TRUE;
}

void NotifyAssociationChange( void )
{
	SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_DWORD | SHCNF_FLUSH, NULL, NULL);
	Sleep(1000);

	WCHAR *szMsg = L"The Mailto handler was registered.\n\n"
		L"Set MailToClip as the default email client in Settings -> Default Apps to enable.\n\n"
		L"To keep your current email client for everything except Mailto, go to Settings -> Default Apps -> Choose default apps by protocol -> MAILTO and choose MailToClip.\n\n"
		L"(Unfortunately, this cannot be automated on Windows 8 or newer. Thanks, MSFT)";
	Info(szMsg);
	ShellExecute(NULL, NULL, L"ms-settings:defaultapps", NULL, NULL, SW_SHOW);
}

int APIENTRY wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow )
{
	if ( wcslen(lpCmdLine) == 0 )
	{
		// No args - register mailto handler if necessary
		if ( RegisterMailtoHandler() )
			NotifyAssociationChange();

		return 0;
	}

	LPWSTR szBuffer = calloc(wcslen(lpCmdLine) + 1, sizeof(WCHAR));

	if ( !szBuffer )
		return 0;

	wcscpy(szBuffer, lpCmdLine);

	LPWSTR szEmail = szBuffer;
	size_t szLen = wcslen(szEmail);

	// Remove quotes if exist
	if ( szEmail[0] == L'"' && szEmail[szLen - 1] == L'"' )
	{
		szEmail[szLen - 1] = L'\0';
		szEmail++;
		szLen = wcslen(szEmail);
	}

	// Remove "mailto:" if exists
	if ( !_wcsnicmp(szEmail, szMailto, cchMailto) )
	{
		szEmail += cchMailto;
		szLen = wcslen(szEmail);
	}

	// Check the length first, we don't want to clear the clipboard if the email is empty
	if ( szLen > 0 )
	{
		size_t cbEmail = (szLen + 1) * sizeof(WCHAR);
		HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, cbEmail);

		if ( hMem )
		{
			memcpy(GlobalLock(hMem), szEmail, cbEmail);
			GlobalUnlock(hMem);

			if ( OpenClipboard(NULL) )
			{
				EmptyClipboard();
				SetClipboardData(CF_UNICODETEXT, hMem);
				CloseClipboard();
			}

			GlobalFree(hMem);
		}
	}

	free(szBuffer);
	return 0;
}
