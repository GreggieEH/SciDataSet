#pragma once
class CTempFile
{
public:
	CTempFile();
	~CTempFile();
	BOOL			InitTempFile(
						LPCTSTR			szWorkingDirectory);
	void			SetCurrentExp(
						LPCTSTR			filter,
						short			grating,
						short			detector);
	void			AddValue(
						double			NM,
						double			Signal);
protected:
	void			FormFileName(
						LPTSTR			szFileName,
						UINT			nBufferSize);
	void			AppendLineToFile(
						LPCTSTR			szString);
private:
	TCHAR			m_szWorkingDirectory[MAX_PATH];
	TCHAR			m_szFilter[MAX_PATH];
	short			m_grating;
	short			m_detector;
};

