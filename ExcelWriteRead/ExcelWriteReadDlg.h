
// ExcelWriteReadDlg.h : 头文件
//

#pragma once
#include "ExcelRW.h"
// CExcelWriteReadDlg 对话框
class CExcelWriteReadDlg : public CDialogEx
{
// 构造
public:
	CExcelWriteReadDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCELWRITEREAD_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	ExcelRW m_ExcelRW;
	CString m_edit_strFileName;
	CString m_edit_strReadPose;
	CString m_edit_strReadData;
	CString m_edit_strWriteSheetPose;
	CString m_eidt_strWritePose;
	CString m_edit_strWriteData;
	afx_msg void OnBnClickedButtonCreate();
	afx_msg void OnBnClickedButtonWrite();
	afx_msg void OnBnClickedButtonRead();
	afx_msg void OnBnClickedButtonClose();
	afx_msg void OnClose();
};
