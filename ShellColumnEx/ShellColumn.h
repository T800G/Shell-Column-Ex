#ifndef _SHELLCOLUMN_F2A47379_452A_48B7_B039_430472BCF4FE_
#define _SHELLCOLUMN_F2A47379_452A_48B7_B039_430472BCF4FE_
#pragma once

#include "resource.h"       // main symbols
#include "helpers.h"

/////////////////////////////////////////////////////////////////////////////
// CShellColumn

#include <shlobj.h>

class CShellColumnEx : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CShellColumnEx,&CLSID_ShellColumnEx>,
	public IColumnProvider
{
public:
	CShellColumnEx(): m_pShellLink(NULL) {}

BEGIN_COM_MAP(CShellColumnEx)
	COM_INTERFACE_ENTRY(IColumnProvider)
END_COM_MAP()

	HRESULT FinalConstruct();
	void FinalRelease();

DECLARE_REGISTRY_RESOURCEID(IDR_ShellColumnEx)
DECLARE_PROTECT_FINAL_CONSTRUCT()


public:
	//IColumnProvider
	STDMETHOD(Initialize)(LPCSHCOLUMNINIT);
	STDMETHOD(GetColumnInfo)(DWORD, SHCOLUMNINFO*);
	STDMETHOD(GetItemData)(LPCSHCOLUMNID, LPCSHCOLUMNDATA, VARIANT*);

private:
	IShellLink* m_pShellLink;
	TCHAR m_szPath[_MAX_NTFS_PATH];//allocated along with COM object instance
};

#endif//_SHELLCOLUMN_F2A47379_452A_48B7_B039_430472BCF4FE_
