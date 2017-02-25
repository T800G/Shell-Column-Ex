#ifndef _STDAFX_28F0EC39_C030_4907_AD4F_26BE9CB61E08_
#define _STDAFX_28F0EC39_C030_4907_AD4F_26BE9CB61E08_
#pragma once

#define STRICT

//minimum requirements
#define _WIN32_WINNT  _WIN32_WINNT_WIN2K
//#define _WIN32_IE    _WIN32_IE_IE50 //autoselect from NT version

#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>


#endif//_STDAFX_28F0EC39_C030_4907_AD4F_26BE9CB61E08_

