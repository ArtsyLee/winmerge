
//++[artsylee]
#include "pch.h"
#include "IniOptionsMgr.h"
#include <windows.h>
#include <tchar.h>
#include <vector>
#include "varprop.h"
#include "OptionsMgr.h"

#define MAX_PATH_FULL 32767
#define BUFSIZE 20480

struct AsyncWriterThreadIniParams
{
	String name;
	varprop::VariantValue value;
};

CIniOptionsMgr::CIniOptionsMgr()
	: m_serializing(true)
	, m_dwThreadId(0)
	, m_hThread(nullptr)
	, m_dwQueueCount(0)
{
	InitializeCriticalSection(&m_cs);
	m_hThread = CreateThread(nullptr, 0, AsyncWriterThreadIniProc, this, 0, &m_dwThreadId);
}

CIniOptionsMgr::~CIniOptionsMgr()
{
	for (;;) {
		PostThreadMessage(m_dwThreadId, WM_QUIT, 0, 0);
		if (WaitForSingleObject(m_hThread, 1) != WAIT_TIMEOUT)
			break;
	}
	DeleteCriticalSection(&m_cs);
}

void CIniOptionsMgr::SplitName(const String& strName, String& strPath,
	String& strValue) const
{
	size_t pos = strName.rfind('/');
	if (pos != String::npos)
	{
		size_t len = strName.length();
		strValue = strName.substr(pos + 1, len - pos - 1);
		strPath = strName.substr(0, pos);
	}
	else
	{
		strValue = strName;
		strPath.erase();
	}
}

DWORD WINAPI CIniOptionsMgr::AsyncWriterThreadIniProc(void* pvThis)
{
	CIniOptionsMgr* pThis = reinterpret_cast<CIniOptionsMgr*>(pvThis);
	MSG msg;
	BOOL bRet;
	while ((bRet = GetMessage(&msg, 0, 0, 0)) != 0)
	{
		AsyncWriterThreadIniParams* pParam = reinterpret_cast<AsyncWriterThreadIniParams*>(msg.wParam);
		EnterCriticalSection(&pThis->m_cs);
		pThis->SaveValueToIni(pParam->name, pParam->value);
		LeaveCriticalSection(&pThis->m_cs);
		delete pParam;
		InterlockedDecrement(&pThis->m_dwQueueCount);
	}
	return 0;
}

static TCHAR* s_pAppName = _T("WinMerge");

int CIniOptionsMgr::LoadValueFromIni(const String& strKeyName, varprop::VariantValue& value)
{
	int valType = value.GetType();

	int retVal = COption::OPT_OK;
	const int BufSize = BUFSIZE;
	TCHAR buf[BufSize] = { 0 };
	DWORD len = GetPrivateProfileString(s_pAppName, strKeyName.c_str(), NULL, buf, BufSize, m_fileName.c_str());
	if (len == 0)
		return COption::OPT_NOTFOUND;

	if (valType == varprop::VT_BOOL)
	{
		UINT intVal = GetPrivateProfileInt(s_pAppName, strKeyName.c_str(), 0, m_fileName.c_str());
		value.SetBool(intVal != 0);
	}
	else if (valType == varprop::VT_INT)
	{
		int intVal = (int)GetPrivateProfileInt(s_pAppName, strKeyName.c_str(), 0, m_fileName.c_str());
		value.SetInt(intVal);
	}
	else if (valType == varprop::VT_STRING)
	{
		TCHAR strVal[MAX_PATH] = { 0 };
		GetPrivateProfileString(s_pAppName, strKeyName.c_str(), _T(""), strVal, MAX_PATH, m_fileName.c_str());
		value.SetString(strVal);
	}
	else
	{
		retVal = COption::OPT_WRONG_TYPE;
	}

	if (retVal == COption::OPT_OK)
	{
		retVal = Set(strKeyName, value);
	}

	return retVal;
}

int CIniOptionsMgr::SaveValueToIni(const String& strKeyName, const varprop::VariantValue& value)
{
	int valType = value.GetType();
	int retVal = COption::OPT_OK;

	String strVal;
	if (valType == varprop::VT_BOOL)
	{
		strVal = value.GetBool() ? _T("1") : _T("0");
	}
	else if (value.GetType() == varprop::VT_INT)
	{
		TCHAR num[20] = { 0 };
		_itot_s(value.GetInt(), num, 10);
		strVal = num;
	}
	else if (value.GetType() == varprop::VT_STRING)
	{
		strVal = value.GetString();
	}
	else
	{
		retVal = COption::OPT_UNKNOWN_TYPE;
	}

	if (retVal == COption::OPT_OK)
	{
		BOOL bRet = WritePrivateProfileString(s_pAppName, strKeyName.c_str(), strVal.c_str(), m_fileName.c_str());
		if (!bRet)
		{
			retVal = COption::OPT_ERR;
		}
	}

	return retVal;
}

//-------------------------------------------------------------------------------
//-------------------------------------------------------------------------------
int CIniOptionsMgr::InitOption(const String& name, const varprop::VariantValue& defaultValue)
{
	int valType = defaultValue.GetType();
	if (valType == varprop::VT_NULL)
		return COption::OPT_ERR;

	if (!m_serializing)
		return AddOption(name, defaultValue);

	EnterCriticalSection(&m_cs);

	int retVal = AddOption(name, defaultValue);
	if (retVal == COption::OPT_OK)
	{
		const int BufSize = BUFSIZE;
		TCHAR buf[BufSize] = { 0 };
		DWORD len = GetPrivateProfileString(s_pAppName, name.c_str(), NULL, buf, BufSize, m_fileName.c_str());
		if (len == 0)
		{
			retVal = SaveValueToIni(name, defaultValue);
		}
		else
		{
			varprop::VariantValue value(defaultValue);
			retVal = LoadValueFromIni(name, value);
		}
	}

	LeaveCriticalSection(&m_cs);

	return retVal;
}

//-------------------------------------------------------------------------------
//-------------------------------------------------------------------------------
int CIniOptionsMgr::InitOption(const String& name, const String& defaultValue)
{
	varprop::VariantValue defValue;
	defValue.SetString(defaultValue);
	return InitOption(name, defValue);
}

int CIniOptionsMgr::InitOption(const String& name, const TCHAR* defaultValue)
{
	return InitOption(name, String(defaultValue));
}

//-------------------------------------------------------------------------------
//-------------------------------------------------------------------------------
int CIniOptionsMgr::InitOption(const String& name, int defaultValue, bool serializable)
{
	varprop::VariantValue defValue;
	int retVal = COption::OPT_OK;

	defValue.SetInt(defaultValue);
	if (serializable)
		retVal = InitOption(name, defValue);
	else
		AddOption(name, defValue);
	return retVal;
}

//-------------------------------------------------------------------------------
//-------------------------------------------------------------------------------
int CIniOptionsMgr::InitOption(const String& name, bool defaultValue)
{
	varprop::VariantValue defValue;
	defValue.SetBool(defaultValue);
	return InitOption(name, defValue);
}

int CIniOptionsMgr::SaveOption(const String& name)
{
	if (!m_serializing)
		return COption::OPT_OK;

	varprop::VariantValue value;
	int retVal = COption::OPT_OK;

	value = Get(name);
	int valType = value.GetType();
	if (valType == varprop::VT_NULL)
		retVal = COption::OPT_NOTFOUND;

	if (retVal == COption::OPT_OK)
	{
		AsyncWriterThreadIniParams* pParam = new AsyncWriterThreadIniParams();
		pParam->name = name;
		pParam->value = value;
		InterlockedIncrement(&m_dwQueueCount);
		PostThreadMessage(m_dwThreadId, WM_USER, (WPARAM)pParam, 0);
	}
	return retVal;
}

int CIniOptionsMgr::SaveOption(const String& name, const varprop::VariantValue& value)
{
	int retVal = Set(name, value);
	if (retVal == COption::OPT_OK)
		retVal = SaveOption(name);
	return retVal;
}

int CIniOptionsMgr::SaveOption(const String& name, const String& value)
{
	varprop::VariantValue val;
	val.SetString(value);
	int retVal = Set(name, val);
	if (retVal == COption::OPT_OK)
		retVal = SaveOption(name);
	return retVal;
}

int CIniOptionsMgr::SaveOption(const String& name, const TCHAR* value)
{
	return SaveOption(name, String(value));
}

int CIniOptionsMgr::SaveOption(const String& name, int value)
{
	varprop::VariantValue val;
	val.SetInt(value);
	int retVal = Set(name, val);
	if (retVal == COption::OPT_OK)
		retVal = SaveOption(name);
	return retVal;
}

int CIniOptionsMgr::SaveOption(const String& name, bool value)
{
	varprop::VariantValue val;
	val.SetBool(value);
	int retVal = Set(name, val);
	if (retVal == COption::OPT_OK)
		retVal = SaveOption(name);
	return retVal;
}

int CIniOptionsMgr::RemoveOption(const String& name)
{
	int retVal = COption::OPT_OK;

	String strPath;
	String strValueName;

	SplitName(name, strPath, strValueName);

	while (InterlockedCompareExchange(&m_dwQueueCount, 0, 0) != 0)
		Sleep(0);

	EnterCriticalSection(&m_cs);

	if (!strValueName.empty())
	{
		retVal = COptionsMgr::RemoveOption(name);

		if (retVal == COption::OPT_OK)
		{
			BOOL bSuccess = WritePrivateProfileString(s_pAppName, name.c_str(), NULL, m_fileName.c_str());
			retVal = bSuccess ? COption::OPT_OK : COption::OPT_ERR;
		}
	}
	else
	{
		for (auto it = m_optionsMap.begin(); it != m_optionsMap.end(); )
		{
			if (it->first.find(strPath) == 0)
			{ 
				WritePrivateProfileString(s_pAppName, it->first.c_str(), NULL, m_fileName.c_str());
				it = m_optionsMap.erase(it);
			}
			else
				++it;
		}
		retVal = COption::OPT_OK;
	}

	LeaveCriticalSection(&m_cs);

	return retVal;

}

int CIniOptionsMgr::SetIniFileName(const String& filename)
{
	m_fileName = filename;
	return COption::OPT_OK;
}

int CIniOptionsMgr::ExportOptions(const String& filename, const bool bHexColor /*= false*/) const
{
	int retVal = COption::OPT_OK;
	OptionsMap::const_iterator optIter = m_optionsMap.begin();
	while (optIter != m_optionsMap.end() && retVal == COption::OPT_OK)
	{
		const String name(optIter->first);
		String strVal;
		varprop::VariantValue value = optIter->second.Get();
		if (value.GetType() == varprop::VT_BOOL)
		{
			if (value.GetBool())
				strVal = _T("1");
			else
				strVal = _T("0");
		}
		else if (value.GetType() == varprop::VT_INT)
		{
			if (bHexColor && (strutils::makelower(name).find(String(_T("color"))) != std::string::npos))
				strVal = strutils::format(_T("0x%06x"), value.GetInt());
			else
				strVal = strutils::to_str(value.GetInt());
		}
		else if (value.GetType() == varprop::VT_STRING)
		{
			strVal = value.GetString();
		}

		bool bRet = !!WritePrivateProfileString(_T("WinMerge"), name.c_str(),
			strVal.c_str(), filename.c_str());
		if (!bRet)
			retVal = COption::OPT_ERR;
		++optIter;
	}
	return retVal;
}

int CIniOptionsMgr::ImportOptions(const String& filename)
{
	int retVal = COption::OPT_OK;
	const int BufSize = BUFSIZE;
	TCHAR buf[BufSize] = { 0 };
	auto oleTranslateColor = [](unsigned color) -> unsigned { return ((color & 0xffffff00) == 0x80000000) ? GetSysColor(color & 0x000000ff) : color; };

	DWORD len = GetPrivateProfileString(_T("WinMerge"), nullptr, _T(""), buf, BufSize, filename.c_str());
	if (len == 0)
		return COption::OPT_NOTFOUND;

	TCHAR* pKey = buf;
	while (*pKey != '\0')
	{
		varprop::VariantValue value = Get(pKey);
		if (value.GetType() == varprop::VT_BOOL)
		{
			bool boolVal = GetPrivateProfileInt(_T("WinMerge"), pKey, 0, filename.c_str()) == 1;
			value.SetBool(boolVal);
			SaveOption(pKey, boolVal);
		}
		else if (value.GetType() == varprop::VT_INT)
		{
			int intVal = GetPrivateProfileInt(_T("WinMerge"), pKey, 0, filename.c_str());
			if (strutils::makelower(pKey).find(String(_T("color"))) != std::string::npos)
				intVal = static_cast<int>(oleTranslateColor(static_cast<unsigned>(intVal)));
			value.SetInt(intVal);
			SaveOption(pKey, intVal);
		}
		else if (value.GetType() == varprop::VT_STRING)
		{
			TCHAR strVal[MAX_PATH_FULL] = { 0 };
			GetPrivateProfileString(_T("WinMerge"), pKey, _T(""), strVal, MAX_PATH_FULL, filename.c_str());
			value.SetString(strVal);
			SaveOption(pKey, strVal);
		}
		Set(pKey, value);

		pKey += _tcslen(pKey);

		if ((pKey < buf + len) && (*(pKey + 1) != '\0'))
			pKey++;
	}
	return retVal;
}
//--[artsylee]