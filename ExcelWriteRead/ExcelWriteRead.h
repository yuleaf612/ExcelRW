
// ExcelWriteRead.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelWriteReadApp: 
// �йش����ʵ�֣������ ExcelWriteRead.cpp
//

class CExcelWriteReadApp : public CWinApp
{
public:
	CExcelWriteReadApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelWriteReadApp theApp;