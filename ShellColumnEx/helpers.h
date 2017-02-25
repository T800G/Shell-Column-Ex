#ifndef _HELPERS_4469818A_9E07_4F0C_BDE9_E756E6C46A56_
#define _HELPERS_4469818A_9E07_4F0C_BDE9_E756E6C46A56_

#include <sal.h>

// Replace color in DIB palette
#include <windows.h>
inline BOOL ReplaceDIBColor(__inout HBITMAP* pDIB, __in COLORREF oldColor, __in COLORREF newColor)
{
	//_ASSERTE(pDIB!=NULL);
	if (pDIB==NULL) return FALSE;//?
	if (oldColor==newColor) return TRUE;//nothing to do
	BOOL bRet=FALSE;

	//get color information
	DIBSECTION ds;
	if (!::GetObject(*pDIB, sizeof(DIBSECTION), &ds)) return FALSE;
	UINT nColors = ds.dsBmih.biClrUsed ? ds.dsBmih.biClrUsed : 1<<ds.dsBmih.biBitCount;//bpp to UINT
	if (ds.dsBmih.biBitCount>8 || nColors>256) return FALSE;//8bpp is max

	HDC hDC=::CreateCompatibleDC(NULL);
	if (!hDC) return FALSE;
	HBITMAP hbmpOld=(HBITMAP)::SelectObject(hDC, *pDIB);
	if (hbmpOld!=NULL)
	{	//replace color table entries
		RGBQUAD clrtbl[256*sizeof(RGBQUAD)];//move to global scope if using _ATL_MIN_CRT
		if (::GetDIBColorTable(hDC, 0, nColors, clrtbl))
		{
			UINT i;
			for (i=0; i<nColors; i++)
			{
				if (oldColor==RGB(clrtbl[i].rgbRed, clrtbl[i].rgbGreen, clrtbl[i].rgbBlue))
				{
					clrtbl[i].rgbRed=GetRValue(newColor);
					clrtbl[i].rgbGreen=GetGValue(newColor);
					clrtbl[i].rgbBlue=GetBValue(newColor);
					bRet=TRUE;
				}
			}
			if (bRet)//set new table
				if (!::SetDIBColorTable(hDC, 0, nColors, clrtbl)) bRet=FALSE;
		}
	*pDIB=(HBITMAP)::SelectObject(hDC, hbmpOld);
	}
	::DeleteDC(hDC);
return bRet;
}


#ifndef _MAX_URL
#define _MAX_URL  2048 //IE7 limit (better than MAX_PATH)
#endif

#ifndef _MAX_NTFS_PATH
#define _MAX_NTFS_PATH   32768
#endif


#define INI_KBLIMIT  32768 // 64kb ini file limit?


#ifndef _UNICODE
	#error IPersistFile requires UNICODE
#endif
#include <shlguid.h>
#include <shobjidl.h>
inline HRESULT ResolveShellShortcut(__in LPCTSTR lpszShortcutPath, __out LPTSTR lpszFilePath, __in size_t cchFilePath)
{
	_ASSERTE(lpszShortcutPath);
	_ASSERTE(lpszFilePath);
	_ASSERTE(cchFilePath>0);
	IShellLink* ipShellLink;
	HRESULT hRes=::CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void**)&ipShellLink);
	if (hRes==S_OK)
	{
		IPersistFile* ipPersistFile;
		hRes=ipShellLink->QueryInterface(IID_IPersistFile, (void**)&ipPersistFile);
		if (hRes==S_OK)
		{
			hRes=ipPersistFile->Load(lpszShortcutPath, STGM_READ);
			if (hRes==S_OK)
				hRes = ipShellLink->GetPath(lpszFilePath, (int)cchFilePath, NULL, SLGP_RAWPATH | SLGP_UNCPRIORITY);
		ipPersistFile->Release();
		ipPersistFile=NULL;
		}
	ipShellLink->Release();
	ipShellLink=NULL;
	}
return hRes;
}
inline HRESULT ResolveShellShortcut(__in IShellLink* pShellLink, __in LPCTSTR lpszShortcutPath, __out LPTSTR lpszFilePath, __in size_t cchFilePath)
{
	_ASSERTE(pShellLink);
	_ASSERTE(lpszShortcutPath);
	_ASSERTE(lpszFilePath);
	_ASSERTE(cchFilePath>0);

	IPersistFile* ipPersistFile;
	HRESULT hRes=pShellLink->QueryInterface(IID_IPersistFile, (void**)&ipPersistFile);
	if (hRes==S_OK)
	{
		hRes=ipPersistFile->Load(lpszShortcutPath, STGM_READ);
		if (hRes==S_OK)
			hRes = pShellLink->GetPath(lpszFilePath, (int)cchFilePath, NULL, SLGP_RAWPATH | SLGP_UNCPRIORITY);
	ipPersistFile->Release();
	ipPersistFile=NULL;
	}
return hRes;
}




//extract path from shortcut
inline BOOL ResolveURLFile(LPCTSTR szUrl, LPTSTR pUrl, SIZE_T psLen)
{
return (BOOL)GetPrivateProfileString(_T("InternetShortcut"), _T("URL"), NULL, pUrl, (DWORD)psLen, szUrl);
}


//inline HRESULT HresultFromLastError() {DWORD dwErr=::GetLastError();return HRESULT_FROM_WIN32(dwErr);}
//inline HRESULT HresultFromWin32(DWORD nError) {return( HRESULT_FROM_WIN32(nError));}


// must use TCHAR else clipboard calls fail
#ifdef _UNICODE
#define CF_tTEXT CF_UNICODETEXT
#else
#define CF_tTEXT CF_TEXT
#endif
inline BOOL SetClipboardData(HWND hOwner, UINT uFormat, LPVOID pBuf, SIZE_T len)
{
	BOOL bret=FALSE;
	if ((hOwner==NULL) || (pBuf==NULL) || (len==0)) return FALSE;
	if (OpenClipboard(hOwner))
	{
		HGLOBAL hg=GlobalAlloc(GMEM_MOVEABLE, len);
		if (hg)
		{
			LPVOID pDest=GlobalLock(hg);
			if (pDest)
			{
				CopyMemory(pDest, pBuf, len);
				bret=(BOOL)SetClipboardData(uFormat, hg);
			bret=((GlobalUnlock(hg)==0) && (GetLastError()==NO_ERROR));
			}
		bret=(GlobalFree(hg)==NULL);//fix algo? release on err?
		}
	bret=CloseClipboard();
	}
return bret;
}


#include <strsafe.h>
inline BOOL CopyStringToClipboard(LPCTSTR psz, HWND hWndOwner)
{
	BOOL bret=FALSE;
	size_t cb;
	if SUCCEEDED(StringCbLength(psz, STRSAFE_MAX_CCH*sizeof(TCHAR), &cb))
	{
		cb+=sizeof(TCHAR);//StringCbLength doesn't count terminating null
		bret=(BOOL)SetClipboardData(hWndOwner, CF_tTEXT, (LPVOID)psz, cb);
		ATLTRACE(_T("string copied to clipboard: <<%s>>\n"), psz);
	}
return bret;
}


//#define INSERTEDMENUITEMS(x)  MAKE_HRESULT(SEVERITY_SUCCESS,FACILITY_NULL,(x)+1)
//inline BOOL InsertMenuSeparator(HMENU hMenu, UINT uPosition, UINT_PTR uIDNewItem)
//{return ::InsertMenu(hMenu, uPosition, MF_SEPARATOR | MF_BYPOSITION, uIDNewItem, NULL);}



template <class T> class CBuffer
{
public:
	CBuffer(): m_buf(NULL) {}
	virtual ~CBuffer(){Free();}

public:
	T* Allocate(SIZE_T s, BOOL bAutozero=FALSE)
	{
		ATLTRACE("CBuffer::Allocate size=%d , autozero %d\n", s, bAutozero);
		if (m_buf!=NULL) Free();
		m_buf=::CoTaskMemAlloc(s*sizeof(T));//COM methods compatible, 'new' throws
		if (bAutozero && m_buf) SecureZeroMemory(m_buf, s*sizeof(T));
	return (T*)m_buf;
	}
	void Free(){ ATLTRACE("CBuffer::Free\n"); ::CoTaskMemFree(m_buf); m_buf=NULL; }
	T* Attach(T* pT)
	{
		ATLTRACE("CBuffer::Attach\n");
		LPVOID p=m_buf;
		m_buf=(LPVOID)pT;
	return (T*)p;
	}
	T* Detach()
	{
		ATLTRACE("CBuffer::Detach\n");
		LPVOID p=m_buf;
		m_buf=NULL;
	return (T*)p;
	}
	operator LPVOID () {return m_buf;}
	operator T* () {return (T*)m_buf;}
	operator BOOL () {return (m_buf!=NULL);}
public:
	LPVOID m_buf;
};



#ifdef _DEBUG
#define dbgTraceCMFlags(x) _dbgTraceCMFlags((x))
static void _dbgTraceCMFlags(UINT uFlags)
{
	if (uFlags & CMF_NORMAL) ATLTRACE("CMF_NORMAL");
	if (uFlags & CMF_DEFAULTONLY) ATLTRACE("|CMF_DEFAULTONLY");
	if (uFlags & CMF_VERBSONLY) ATLTRACE("|CMF_VERBSONLY");
	if (uFlags & CMF_EXPLORE) ATLTRACE("|CMF_EXPLORE");
	if (uFlags & CMF_NOVERBS) ATLTRACE("|CMF_NOVERBS");
	if (uFlags & CMF_CANRENAME) ATLTRACE("|CMF_CANRENAME");
	if (uFlags & CMF_NODEFAULT) ATLTRACE("|CMF_NODEFAULT");
#if (NTDDI_VERSION < NTDDI_VISTA)
	if (uFlags & CMF_INCLUDESTATIC) ATLTRACE("|CMF_INCLUDESTATIC");
#endif
	if (uFlags & 0x00000080) ATLTRACE("|CMF_ITEMMENU");
	if (uFlags & CMF_EXTENDEDVERBS) ATLTRACE("|CMF_EXTENDEDVERBS");
	if (uFlags & 0x00000200) ATLTRACE("|CMF_DISABLEDVERBS");
	if (uFlags & CMF_RESERVED) ATLTRACE("|CMF_RESERVED");
	ATLTRACE("\n");
}
#else
#define dbgTraceCMFlags(x)
#endif//_DEBUG


#endif//_HELPERS_4469818A_9E07_4F0C_BDE9_E756E6C46A56_
