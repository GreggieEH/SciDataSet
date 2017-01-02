#include "stdafx.h"
#include "DlgNewObjectSetup.h"
#include "MyObject.h"
#include "dispids.h"

CDlgNewObjectSetup::CDlgNewObjectSetup(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_hwndDlg(NULL),
	m_fAllowClose(TRUE)

{
}

CDlgNewObjectSetup::~CDlgNewObjectSetup()
{
}

BOOL CDlgNewObjectSetup::DoOpenDialog(
	HINSTANCE		hInst,
	HWND			hwndParent)
{
	return IDOK == DialogBoxParam(
		hInst, MAKEINTRESOURCE(IDD_DIALOGNewObjectSetup), hwndParent, (DLGPROC)DlgProcNewObjectSetup, (LPARAM) this);
}

LRESULT CALLBACK DlgProcNewObjectSetup(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgNewObjectSetup*	pDlg = NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgNewObjectSetup*)lParam;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		pDlg->m_hwndDlg = hwndDlg;
	}
	else
	{
		pDlg = (CDlgNewObjectSetup*)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
	{
		return pDlg->DialogProc(uMsg, wParam, lParam);
	}
	else
	{
		return FALSE;
	}
}

BOOL CDlgNewObjectSetup::DialogProc(
	UINT			uMsg,
	WPARAM			wParam,
	LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
//	case WM_ENTERIDLE:
//		this->OnEnterIdle();
//		break;
	case DM_GETDEFID:
		return this->OnGetDefID();
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgNewObjectSetup::OnInitDialog()
{
	Utils_CenterWindow(this->m_hwndDlg);
	this->m_pMyObject->SetADInfoString();
	this->DisplayScanType();
	this->DisplayComment();
	this->DisplayWorkingDirectory();
	return TRUE;
}

BOOL CDlgNewObjectSetup::OnCommand(
	WORD			wmID,
	WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
		// check allow close flag
		if (!this->m_fAllowClose)
		{
			this->m_fAllowClose = TRUE;
			return TRUE;
		}
		// set the comment
		this->OnSetComment();
		// fail through
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		return TRUE;
	case IDC_SCANPARAMETERS:
		this->m_pMyObject->FireRequestSetupDisplay(L"Scan Parameters");
		return TRUE;
	case IDC_MONOCHROMATOR:
		this->m_pMyObject->FireRequestSetupDisplay(L"Monochromator");
		return TRUE;
	case IDC_ADBOARD:
		this->m_pMyObject->FireRequestSetupDisplay(L"AD Board");
		this->m_pMyObject->SetADInfoString();
		return TRUE;
	case IDC_SLIT:
		this->m_pMyObject->FireRequestSetupDisplay(L"Slit");
		return TRUE;
	case IDC_FILTERWHEEL:
		this->m_pMyObject->FireRequestSetupDisplay(L"FilterWheel");
		return TRUE;
	case IDC_DATASETUP:
		{
			IDispatch	*	pdisp;
			VARIANTARG		varg;
			HRESULT			hr;
			hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			if (SUCCEEDED(hr))
			{
				InitVariantFromInt32((long)GetParent(this->m_hwndDlg), &varg);
				Utils_InvokeMethod(pdisp, DISPID_ViewSetup, &varg, 1, NULL);
				pdisp->Release();
			}
		}
		return TRUE;
	case IDC_CHECKSAMPLE:
		if (BN_CLICKED == wmEvent)
		{
			this->SetScanType(L"Sample");
			return TRUE;
		}
		break;
	case IDC_CHECKINTENSITYCALIBRATION:
		if (BN_CLICKED == wmEvent)
		{
			this->SetScanType(L"Intensity Calibration");
			return TRUE;
		}
		break;
	case IDC_CHECKBACKGROUND:
		if (BN_CLICKED == wmEvent)
		{
			this->SetScanType(L"Background");
			return TRUE;
		}
		break;
	case IDC_SETCOMMENT:
		this->OnSetComment();
		return TRUE;
	case IDC_BROWSEWORKINGDIRECTORY:
		this->BrowseForWorkingDirectory();
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

void CDlgNewObjectSetup::DisplayScanType()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	TCHAR			ScanType[MAX_PATH];

	ScanType[0] = '\0';
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_ScanType, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToString(varResult, ScanType, MAX_PATH);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKSAMPLE), 0 == lstrcmpi(ScanType, L"Sample") ? BST_CHECKED : BST_UNCHECKED);
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKINTENSITYCALIBRATION), 0 == lstrcmpi(ScanType, L"Intensity Calibration") ? BST_CHECKED : BST_UNCHECKED);
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHECKBACKGROUND), 0 == lstrcmpi(ScanType, L"Background") ? BST_CHECKED : BST_UNCHECKED);

}

void CDlgNewObjectSetup::SetScanType(
	LPCTSTR			ScanType)
{
	HRESULT			hr;
	IDispatch	*	pdisp;

	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_SetStringProperty(pdisp, DISPID_ScanType, ScanType);
		pdisp->Release();
	}
	this->DisplayScanType();
}

/*
// idle message processing
void CDlgNewObjectSetup::OnEnterIdle()
{
	MSG					msg;			// message structure
	while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}
*/

BOOL CDlgNewObjectSetup::OnGetDefID()
{
	SHORT		shKeyState;
	HWND		hwndFocus;
	UINT		nID;
	shKeyState = GetKeyState(VK_RETURN);
	if (0 != (0x8000 & shKeyState))
	{
		hwndFocus = GetFocus();
		nID = GetDlgCtrlID(hwndFocus);
		if (IDC_EDITCOMMENT == nID)
		{
			this->m_fAllowClose = FALSE;
			return TRUE;
		}
	}
	return FALSE;
}

// display the comment
void CDlgNewObjectSetup::DisplayComment()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	TCHAR			Comment[MAX_PATH];
	BOOL			fSuccess = FALSE;

	Comment[0] = '\0';
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_Comment, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToString(varResult, Comment, MAX_PATH);
			fSuccess = SUCCEEDED(hr);
			if (fSuccess)
				SetDlgItemText(this->m_hwndDlg, IDC_EDITCOMMENT, Comment);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	if (!fSuccess)
	{
		SetDlgItemText(this->m_hwndDlg, IDC_EDITCOMMENT, L"");
	}
}

// set the comment
void CDlgNewObjectSetup::OnSetComment()
{
	HRESULT			hr;
	TCHAR			szString[MAX_PATH];
	IDispatch	*	pdisp;
	VARIANTARG		varg;

	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITCOMMENT, szString, MAX_PATH) > 0)
	{
		hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
		if (SUCCEEDED(hr))
		{
			InitVariantFromString(szString, &varg);
			Utils_InvokePropertyPut(pdisp, DISPID_Comment, &varg, 1);
			VariantClear(&varg);
			pdisp->Release();
		}
	}
	this->DisplayComment();
}

// working directory
void CDlgNewObjectSetup::DisplayWorkingDirectory()
{
	TCHAR			szWorkingDirectory[MAX_PATH];
	HWND			hwnd;
	HDC				hdc;
	RECT			rc;
	if (this->m_pMyObject->FirerequestWorkingDirectory(szWorkingDirectory, MAX_PATH))
	{
		hwnd = GetDlgItem(this->m_hwndDlg, IDC_EDITWORKINGDIRECTORY);
		hdc = GetDC(hwnd);
		GetClientRect(hwnd, &rc);
		PathCompactPath(hdc, szWorkingDirectory, rc.right - rc.left);
		ReleaseDC(hwnd, hdc);
		SetDlgItemText(this->m_hwndDlg, 
			IDC_EDITWORKINGDIRECTORY, szWorkingDirectory);
	}
}

void CDlgNewObjectSetup::BrowseForWorkingDirectory()
{
	this->m_pMyObject->FirebrowseForWorkingDirectory(this->m_hwndDlg);
	this->DisplayWorkingDirectory();
}

void CDlgNewObjectSetup::SetWorkingDirectory()
{
	this->DisplayWorkingDirectory();
}
