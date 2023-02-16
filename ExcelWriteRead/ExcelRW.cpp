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
	//����Excel ������(����Excel)
	if (!m_ExcelApp.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("����Excel����ʧ��"), MB_OK | MB_ICONWARNING);
		return;
	}
	m_bServeStart = TRUE;
	/*�жϵ�ǰExcel�İ汾*/
	CString strExcelVersion = m_ExcelApp.get_Version();//��ȡ�汾��Ϣ
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
	if ("11" == strExcelVersion)
	{
		//AfxMessageBox("��ǰExcel�İ汾��2003��");
	}
	else if ("12" == strExcelVersion)
	{
		//AfxMessageBox("��ǰExcel�İ汾��2007��");
	}
	else
	{
		//AfxMessageBox("��ǰExcel�İ汾�������汾��");
	}
	m_ExcelApp.put_Visible(FALSE);
	m_ExcelApp.put_UserControl(FALSE);

	m_ExcelBooks.AttachDispatch(m_ExcelApp.get_Workbooks());//�õ�����������	
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

	/*�õ��������е�Sheet������*/
	m_ExcelSheets.AttachDispatch(m_ExcelBook.get_Sheets());

	/*��һ��Sheet���粻���ڣ�������һ��Sheet*/
	CString strSheetName = _T("Sheet1");
	try
	{
		/*��һ�����е�Sheet*/
		m_ExcelSheet = m_ExcelSheets.get_Item(_variant_t(strSheetName));
	}
	catch (...)
	{
		/*����һ���µ�Sheet*/
		m_ExcelSheet = m_ExcelSheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		m_ExcelSheet.put_Name(strSheetName);
	}             
}

void ExcelRW::WriteTable(CString cSheet, CString clocow,CString strWrite)
{
	try
	{
		/*��һ�����е�Sheet*/
		m_ExcelSheet = m_ExcelSheets.get_Item(_variant_t(cSheet));
	}
	catch (...)
	{
		/*����һ���µ�Sheet*/
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
		m_ExcelBook.SaveCopyAs(COleVariant(m_openFilePath));//���Ϊ
	else
		m_ExcelBook.Save();
	// �ͷŶ���
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

//��ȡ��Ԫ������
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