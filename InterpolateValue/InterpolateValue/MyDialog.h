// CMyDialog

#pragma once

class CMyDialog
{
public:
	CMyDialog(UINT nID);
	virtual ~CMyDialog();
	BOOL                DoOpenDialog(
							HWND		 hwndParent,
							HINSTANCE    hInst);
	HWND				DoCreateDialog(
							HWND		hwndParent,
							HINSTANCE	hInst);
	HWND				GetMyDialog();
protected:
	virtual BOOL		DlgProc(
							UINT		uMsg,
							WPARAM		wParam,
							LPARAM		lParam);
	virtual BOOL		OnInitDialog();
	virtual BOOL		OnCommand(
							WORD		wmID,
							WORD		wmEvent);
	virtual BOOL		OnNotify(
							LPNMHDR		pnmh);
	virtual BOOL		OnGetDefID();
	virtual BOOL		OnReturnClicked(
							UINT		nID);
	BOOL				GetAllowClose();
	void				SetAllowClose(
							BOOL		fAllowClose);
	virtual BOOL		OnOK();
	virtual void		OnCloseup();
private:
	UINT				m_nID;
	HWND				m_hwndDlg;
	BOOL				m_fAllowClose;

	friend LRESULT CALLBACK	MyDialogProc(HWND, UINT, WPARAM, LPARAM);
};
