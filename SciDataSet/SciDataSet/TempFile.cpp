#include "stdafx.h"
#include "TempFile.h"

CTempFile::CTempFile() :
	m_grating(-1),
	m_detector(-1)
{
	this->m_szWorkingDirectory[0] = '\0';
	this->m_szFilter[0] = '\0';
}

CTempFile::~CTempFile()
{
}

BOOL CTempFile::InitTempFile(
	LPCTSTR			szWorkingDirectory)
{
	StringCchCopy(this->m_szWorkingDirectory, MAX_PATH, szWorkingDirectory);
	this->m_szFilter[0] = '\0';
	this->m_detector = -1;
	this->m_grating = -1;
	// create new temp file
	HANDLE			hLogFile;
	TCHAR			szFileName[MAX_PATH];

	this->FormFileName(szFileName, MAX_PATH);
	hLogFile = CreateFile(szFileName, GENERIC_READ | GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hLogFile)
	{
		CloseHandle(hLogFile);
	}
	return TRUE;
}

void CTempFile::SetCurrentExp(
	LPCTSTR			filter,
	short			grating,
	short			detector)
{
	TCHAR			szLine[MAX_PATH];
	if (0 != lstrcmpi(filter, this->m_szFilter) || grating != this->m_grating || detector != this->m_detector)
	{
		StringCchCopy(this->m_szFilter, MAX_PATH, filter);
		this->m_grating = grating;
		this->m_detector = detector;
		StringCchPrintf(szLine, MAX_PATH, L"Filter = %s  grating = %1d  detector = %1d", 
			this->m_szFilter, this->m_grating, this->m_detector);
		this->AppendLineToFile(szLine);
	}
}

void CTempFile::AddValue(
	double			NM,
	double			Signal)
{
	TCHAR			szLine[MAX_PATH];
	_stprintf_s(szLine, L"%8.5f\t%10.5g", NM, Signal);
	this->AppendLineToFile(szLine);
}

void CTempFile::FormFileName(
	LPTSTR			szFileName,
	UINT			nBufferSize)
{
	StringCchCopy(szFileName, nBufferSize, this->m_szWorkingDirectory);
	PathAppend(szFileName, L"SciDataSet.temp");
}

void CTempFile::AppendLineToFile(
	LPCTSTR			szString)
{
	HANDLE		hLogFile;
	TCHAR		szFileName[MAX_PATH];
	TCHAR		szText[MAX_PATH];
	size_t		slen;
	DWORD		dwPos;
	DWORD		dwNWritten;

	this->FormFileName(szFileName, MAX_PATH);

	hLogFile = CreateFile(szFileName, GENERIC_READ | GENERIC_WRITE, 0,
		NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hLogFile)
	{
		// point to the end of the file
		dwPos = SetFilePointer(hLogFile, 0, NULL, FILE_END);
		// write the string
		StringCchPrintf(szText, MAX_PATH, TEXT("%s\r\n"), szString);
		StringCchLength(szText, MAX_PATH, &slen);
		WriteFile(hLogFile, (LPCVOID)szText, slen * sizeof(TCHAR), &dwNWritten, NULL);
		CloseHandle(hLogFile);
	}
}
