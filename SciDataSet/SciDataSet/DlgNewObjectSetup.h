#pragma once

class CMyObject;

class CDlgNewObjectSetup
{
public:
	CDlgNewObjectSetup(CMyObject * pMyObject);
	~CDlgNewObjectSetup();
	BOOL			DoOpenDialog(
						HINSTANCE		hInst,
						HWND			hwndParent);
protected:
	BOOL			DialogProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	void			DisplayScanType();
	void			SetScanType(
						LPCTSTR			ScanType);
//	// idle message processing
//	void			OnEnterIdle();
	// display the comment
	void			DisplayComment();
	// set the comment
	void			OnSetComment();
	BOOL			OnGetDefID();
	// working directory
	void			DisplayWorkingDirectory();
	void			BrowseForWorkingDirectory();
	void			SetWorkingDirectory();

private:
	CMyObject	*	m_pMyObject;
	HWND			m_hwndDlg;
	BOOL			m_fAllowClose;

	friend LRESULT CALLBACK DlgProcNewObjectSetup(HWND, UINT, WPARAM, LPARAM);
};

