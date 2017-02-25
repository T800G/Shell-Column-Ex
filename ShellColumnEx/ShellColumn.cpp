#include "stdafx.h"
#include "ShColEx.h"
#include "ShellColumn.h"

#include <strsafe.h>
#pragma comment(lib,"strsafe.lib")
#include <Shlwapi.h>
#pragma comment(lib,"Shlwapi.lib")

#include "helpers.h"


HRESULT CShellColumnEx::FinalConstruct()
{
	//ATLTRACE("CShellColumnEx::FinalConstruct\n");
	HRESULT hRes=::CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, __uuidof(IShellLink), (void**)&m_pShellLink);
	ATLASSERT(SUCCEEDED(hRes));
return hRes;
}

void CShellColumnEx::FinalRelease()
{
	//ATLTRACE("CShellColumnEx::FinalRelease\n");
	ULONG lret=m_pShellLink->Release();
	ATLASSERT(lret==0);
	m_pShellLink=NULL;
}


HRESULT CShellColumnEx::Initialize(LPCSHCOLUMNINIT psci)
{
	//ATLTRACE("CShellColumnEx::Initialize\n");
return S_OK;
}

HRESULT CShellColumnEx::GetColumnInfo(DWORD dwIndex, SHCOLUMNINFO *psci)
{
	//ATLTRACE("CShellColumnEx::GetColumnInfo\n");
	if (dwIndex > 2) return S_FALSE;//3 columns, 0-based index

	psci->scid.fmtid = CLSID_ShellColumnEx;//format identifier
	psci->vt = VT_LPSTR; //column data is string
	psci->csFlags = SHCOLSTATE_TYPE_STR;
	psci->fmt = LVCFMT_LEFT; //left alignment

	switch (dwIndex)
	{
		case 0:
			psci->scid.pid = 0;//first column of fmtid
			psci->cChars = 14; //default column width in chars
			//Caption and description
			if FAILED(StringCchCopy(psci->wszTitle, MAX_COLUMN_NAME_LEN, _T("File Extension"))) return E_FAIL;
			if FAILED(StringCchCopy(psci->wszDescription, MAX_COLUMN_DESC_LEN, _T("Displays File Extensions"))) return E_FAIL;
		break;
		case 1:
			psci->scid.pid = 1;
			psci->cChars = 32;  //average path length?
			if FAILED(StringCchCopy(psci->wszTitle, MAX_COLUMN_NAME_LEN, _T("Shortcut target"))) return E_FAIL;
			if FAILED(StringCchCopy(psci->wszDescription, MAX_COLUMN_DESC_LEN, _T("Displays shortcut target location"))) return E_FAIL;
		break;
		case 2:
			psci->scid.pid = 2;
			psci->cChars = 48;	//average URL length?
			if FAILED(StringCchCopy(psci->wszTitle, MAX_COLUMN_NAME_LEN, _T("URL"))) return E_FAIL;
			if FAILED(StringCchCopy(psci->wszDescription, MAX_COLUMN_DESC_LEN, _T("Displays URL address"))) return E_FAIL;
		break;		
		default:return S_FALSE;
	}
return S_OK;
}

HRESULT CShellColumnEx::GetItemData(LPCSHCOLUMNID pscid, LPCSHCOLUMNDATA pscd, VARIANT *pvarData)
{
	//ATLTRACE("CShellColumnEx::GetItemData\n");
	if (!IsEqualGUID(pscid->fmtid, CLSID_ShellColumnEx)) return S_FALSE;//not our columns

	if (PathIsDirectory(pscd->wszFile)) return S_FALSE;//skip dirs / use this inside switch statement if needed

	LPTSTR ps=_T("");
	switch (pscid->pid)
	{
		case 0://column for file extensions
			if (StrStrI(_T(".lnk.pif.shb.wsh.url.appref-ms.cnf"), pscd->pwszExt)) return S_FALSE;//skip IsShortcut/NeverShowExt
			ps=CharNext(pscd->pwszExt);
		break;
		case 1://column for lnk data
			if (StrCmpI(pscd->pwszExt,_T(".lnk"))==0)
				if (ResolveShellShortcut(m_pShellLink, pscd->wszFile, m_szPath, _MAX_URL)==S_OK) ps=m_szPath;
		break;
		case 2://column for url data
			if (StrCmpI(pscd->pwszExt,_T(".url"))==0)
				if (ResolveURLFile(pscd->wszFile, m_szPath, _MAX_URL)) ps=m_szPath;
		break;
		default:return S_FALSE;
	}

	CComVariant cv(ps);
	cv.Detach(pvarData);
return S_OK;
}
