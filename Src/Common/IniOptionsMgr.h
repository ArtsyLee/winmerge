//++[artsylee]
#pragma once

#include <Windows.h>
#include "OptionsMgr.h"

class COptionsMgr;

/**
 * @brief IniConfig-based implementation of OptionsMgr interface
 */
class CIniOptionsMgr : public COptionsMgr
{
public:
	CIniOptionsMgr();
	virtual ~CIniOptionsMgr();
	int SetIniFileName(const String& filename);

	virtual int InitOption(const String& name, const varprop::VariantValue& defaultValue) override;
	virtual int InitOption(const String& name, const String& defaultValue) override;
	virtual int InitOption(const String& name, const TCHAR* defaultValue) override;
	virtual int InitOption(const String& name, int defaultValue, bool serializable = true) override;
	virtual int InitOption(const String& name, bool defaultValue) override;

	virtual int SaveOption(const String& name) override;
	virtual int SaveOption(const String& name, const varprop::VariantValue& value) override;
	virtual int SaveOption(const String& name, const String& value) override;
	virtual int SaveOption(const String& name, const TCHAR* value) override;
	virtual int SaveOption(const String& name, int value) override;
	virtual int SaveOption(const String& name, bool value) override;

	virtual int RemoveOption(const String& name) override;

	virtual void SetSerializing(bool serializing = true) override { m_serializing = serializing; }

	virtual int ExportOptions(const String& filename, const bool bHexColor = false) const override;
	virtual int ImportOptions(const String& filename) override;

protected:
	void SplitName(const String& strName, String& strPath, String& strValue) const;

	static DWORD WINAPI AsyncWriterThreadIniProc(void* pParam);

	int LoadValueFromIni(const String& strKeyName, varprop::VariantValue& value);
	int SaveValueToIni(const String& strKeyName, const varprop::VariantValue& value);

private:
	String m_fileName;
	bool m_serializing;
	DWORD m_dwThreadId;
	HANDLE m_hThread;
	CRITICAL_SECTION m_cs;
	DWORD m_dwQueueCount;
};
//--[artsylee]