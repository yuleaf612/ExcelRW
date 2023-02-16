#include "stdafx.h"
#include "ExcelRW.h"


ExcelRW::ExcelRW()
{
	COleVariant covTrue((short)TRUE);
	COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);	
}

ExcelRW::~ExcelRW()
{
}

void ExcelRW::OpenTable(CString OpenPath)
{
	if (m_bServeStart)
		CloseTable();
	m_openFilePath = OpenPath;
	//创建Excel 服务器(启动Excel)
	if (!m_ExcelApp.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("启动Excel服务失败"), MB_OK | MB_ICONWARNING);
		return;
	}
	m_bServeStart = TRUE;
	/*判断当前Excel的版本*/
	CString strExcelVersion = m_ExcelApp.get_Version();//获取版本信息
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
	if ("11" == strExcelVersion)
	{
		//AfxMessageBox("当前Excel的版本是2003。");
	}
	else if ("12" == strExcelVersion)
	{
		//AfxMessageBox("当前Excel的版本是2007。");
	}
	else
	{
		//AfxMessageBox("当前Excel的版本是其他版本。");
	}
	m_ExcelApp.put_Visible(FALSE);
	m_ExcelApp.put_UserControl(FALSE);

	m_ExcelBooks.AttachDispatch(m_ExcelApp.get_Workbooks());//得到工作簿容器	
	try
	{
		m_ExcelBook = m_ExcelBooks.Open(m_openFilePath,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing);
	}
	catch(...)
	{
		m_ExcelBook = m_ExcelBooks.Add(vtMissing);
		m_bNewTable = TRUE;
	}

	/*得到工作簿中的Sheet的容器*/
	m_ExcelSheets.AttachDispatch(m_ExcelBook.get_Sheets());

	/*打开一个Sheet，如不存在，就新增一个Sheet*/
	CString strSheetName = _T("Sheet1");
	try
	{
		/*打开一个已有的Sheet*/
		m_ExcelSheet = m_ExcelSheets.get_Item(_variant_t(strSheetName));
	}
	catch (...)
	{
		/*创建一个新的Sheet*/
		m_ExcelSheet = m_ExcelSheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		m_ExcelSheet.put_Name(strSheetName);
	}             
}

void ExcelRW::WriteTable(CString cSheet, CString clocow,CString strWrite)
{
	try
	{
		/*打开一个已有的Sheet*/
		m_ExcelSheet = m_ExcelSheets.get_Item(_variant_t(cSheet));
	}
	catch (...)
	{
		/*创建一个新的Sheet*/
		m_ExcelSheet = m_ExcelSheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		m_ExcelSheet.put_Name(cSheet);
	}
	m_ExcelRange = m_ExcelSheet.get_Range(COleVariant(clocow), COleVariant(clocow));
	m_ExcelRange.put_Value2(COleVariant(strWrite));
}

void ExcelRW::CloseTable()
{
	m_ExcelBook.put_Saved(TRUE);	
	if(m_bNewTable)
		m_ExcelBook.SaveCopyAs(COleVariant(m_openFilePath));//另存为
	else
		m_ExcelBook.Save();
	// 释放对象
	m_ExcelBooks.ReleaseDispatch();
	m_ExcelBook.ReleaseDispatch();
	m_ExcelSheets.ReleaseDispatch();
	m_ExcelSheet.ReleaseDispatch();
	m_ExcelRange.ReleaseDispatch();
	m_ExcelApp.Quit();
	m_ExcelApp.ReleaseDispatch();

	m_bServeStart = FALSE;
}

void ExcelRW::OpenSheet()
{

}

//获取单元格内容
void ExcelRW::ReadTable(CString clocow, CString &strRead)
{
	variant_t rValue;
	m_ExcelRange = m_ExcelSheet.get_Range(COleVariant(clocow), COleVariant(clocow));
	rValue = m_ExcelRange.get_Value2();
	switch (rValue.vt)
	{
	case VT_R8:
		strRead.Format(_T("%f"), (float)rValue.dblVal);
		break;
	case VT_BSTR:
		strRead = rValue.bstrVal;
		break;
	case VT_I4:
		strRead.Format(_T("%ld"), (int)rValue.dblVal);
		break;
	default:
		break;
	}
}