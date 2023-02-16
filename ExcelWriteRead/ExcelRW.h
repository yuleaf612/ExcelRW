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
	void OpenTable(CString OpenPath);//打开表格，OpenPath为要表格路径
	void WriteTable(CString cSheet, CString clocow,CString strWrite);//写入表格数据，clocow为表格位置（比如“A5”）,strWrite为要写入的字符
	void ReadTable(CString clocow, CString &strRead);//读取数据
	void CloseTable();//保存并关闭表格

public:
	void OpenSheet();

};

