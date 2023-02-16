#pragma once
#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CRange.h"

class ExcelRW
{
public:
	ExcelRW();
	~ExcelRW();
	CApplication m_ExcelApp;            
	CWorkbook m_ExcelBook;         
	CWorkbooks m_ExcelBooks;       
	CWorksheet m_ExcelSheet;          
	CWorksheets m_ExcelSheets;        
	CRange m_ExcelRange;                        
	BOOL m_bNewTable = FALSE;
	BOOL m_bServeStart=FALSE;
	CString m_openFilePath;              

public:
	void OpenTable(CString OpenPath);//�򿪱��OpenPathΪҪ���·��
	void WriteTable(CString cSheet, CString clocow,CString strWrite);//д�������ݣ�clocowΪ���λ�ã����硰A5����,strWriteΪҪд����ַ�
	void ReadTable(CString clocow, CString &strRead);//��ȡ����
	void CloseTable();//���沢�رձ��

public:
	void OpenSheet();

};

