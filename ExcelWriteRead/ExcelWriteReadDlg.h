
// ExcelWriteReadDlg.h : ͷ�ļ�
//

#pragma once
#include "ExcelRW.h"
// CExcelWriteReadDlg �Ի���
class CExcelWriteReadDlg : public CDialogEx
{
// ����
public:
	CExcelWriteReadDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCELWRITEREAD_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
