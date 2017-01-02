// CMyDialog  - virtual base class

#include "stdafx.h"
#include "MyDialog.h"

CMyDialog::CMyDialog(UINT nID) :
	m_nID(nID),
	m_hwndDlg(NULL),
	m_fAllowClose(TRUE)
{

}

CMyDialog::~CMyDialog()
{

}
BOOL CMyDialog::DoOpenDialog(
	HWND		 hwndParent,
	HINSTANCE    hInst)
{
	return IDOK == DialogBoxParam(hInst, MAKEINTRESOURCE(this->m_nID), hwndParent, (DLGPROC)MyDialogProc, (LPARAM)this);
}
HWND CMyDialog::DoCreateDialog(
	HWND		hwndParent,
	HINSTANCE	hInst)
{
	return CreateDialogParam(hInst, MAKEINTRESOURCE(this->m_nID), hwndParent, (DLGPROC)MyDialogProc, (LPARAM)this);
}

HWND CMyDialog::GetMyDialog()
{
	return this->m_hwndDlg;
}

LRESULT CALLBACK MyDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyDialog*	pDlg = NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CMyDialog*)lParam;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		pDlg->m_hwndDlg = hwndDlg;
	}
	else
	{
		pDlg = (CMyDialog*)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
	{
		return pDlg->DlgProc(uMsg, wParam, lParam);
	}
	else
	{
		return FALSE;
	}
}

BOOL CMyDialog::DlgProc(
	UINT		uMsg,
	WPARAM		wParam,
	LPARAM		lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
	case WM_NOTIFY:
		return this->OnNotify((LPNMHDR)lParam);
	case DM_GETDEFID:
		return this->OnGetDefID();
	default:
		break;
	}
	return FALSE;
}

BOOL CMyDialog::OnInitDialog()
{
	return TRUE;
}

BOOL CMyDialog::OnCommand(
	WORD		wmID,
	WORD		wmEvent)
{
	UNREFERENCED_PARAMETER(wmEvent);
	switch (wmID)
	{
	case IDOK:
		if (!this->OnOK())
		{
			return TRUE;
		}
		// fall through
	case IDCANCEL:
		this->OnCloseup();
		EndDialog(this->m_hwndDlg, wmID);
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

BOOL CMyDialog::OnOK()
{
	if (!this->m_fAllowClose)
	{
		this->m_fAllowClose = TRUE;
		return FALSE;
	}
	return TRUE;
}

void CMyDialog::OnCloseup()
{

}


BOOL CMyDialog::OnNotify(
	LPNMHDR		pnmh)
{
	UNREFERENCED_PARAMETER(pnmh);
	return FALSE;
}

BOOL CMyDialog::OnGetDefID()
{
	SHORT		shKeyState;
	HWND		hwndFocus;
	UINT		nID;
	shKeyState = GetKeyState(VK_RETURN);
	if (0 != (0x8000 & shKeyState))
	{
		hwndFocus = GetFocus();
		nID = GetDlgCtrlID(hwndFocus);
		if (this->OnReturnClicked(nID))
		{
			this->SetAllowClose(FALSE);
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CMyDialog::OnReturnClicked(
	UINT		nID)
{
	UNREFERENCED_PARAMETER(nID);
	return FALSE;
}
	
BOOL CMyDialog::GetAllowClose()
{
	return this->m_fAllowClose;
}
	
void CMyDialog::SetAllowClose(
	BOOL		fAllowClose)
{
	this->m_fAllowClose = fAllowClose;
}