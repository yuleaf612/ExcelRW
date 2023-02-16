
// ExcelWriteReadDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ExcelWriteRead.h"
#include "ExcelWriteReadDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
public:
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExcelWriteReadDlg �Ի���



CExcelWriteReadDlg::CExcelWriteReadDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_EXCELWRITEREAD_DIALOG, pParent)
	, m_edit_strFileName(_T("D:\\test.xlsx"))
	, m_edit_strReadPose(_T("A1"))
	, m_edit_strReadData(_T(""))
	, m_eidt_strWritePose(_T("A1"))
	, m_edit_strWriteData(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelWriteReadDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_FILENAME, m_edit_strFileName);
	DDX_Text(pDX, IDC_EDIT_READPOSE, m_edit_strReadPose);
	DDX_Text(pDX, IDC_EDIT_READDATA, m_edit_strReadData);
	DDX_Text(pDX, IDC_EDIT_WRITESHEETPOSE, m_edit_strWriteSheetPose);
	DDX_Text(pDX, IDC_EDIT_WRITEPOSE, m_eidt_strWritePose);
	DDX_Text(pDX, IDC_EDIT_WRITEDATA, m_edit_strWriteData);

}

BEGIN_MESSAGE_MAP(CExcelWriteReadDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_Create, &CExcelWriteReadDlg::OnBnClickedButtonCreate)
	ON_BN_CLICKED(IDC_BUTTON_WRITE, &CExcelWriteReadDlg::OnBnClickedButtonWrite)
	ON_BN_CLICKED(IDC_BUTTON_READ, &CExcelWriteReadDlg::OnBnClickedButtonRead)
	ON_BN_CLICKED(IDC_BUTTON_CLOSE, &CExcelWriteReadDlg::OnBnClickedButtonClose)
	ON_WM_CLOSE()
END_MESSAGE_MAP()


// CExcelWriteReadDlg ��Ϣ�������

BOOL CExcelWriteReadDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CExcelWriteReadDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelWriteReadDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CExcelWriteReadDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CExcelWriteReadDlg::OnBnClickedButtonCreate()
{
	UpdateData(TRUE);
	m_ExcelRW.OpenTable(m_edit_strFileName);
	GetDlgItem(IDC_BUTTON_Create)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_CLOSE)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_READ)->EnableWindow(TRUE);
	GetDlgItem(IDC_BUTTON_WRITE)->EnableWindow(TRUE);
}

void CExcelWriteReadDlg::OnBnClickedButtonWrite()
{
	UpdateData(TRUE);
	m_ExcelRW.WriteTable(m_edit_strWriteSheetPose, m_eidt_strWritePose,m_edit_strWriteData);
}

void CExcelWriteReadDlg::OnBnClickedButtonRead()
{
	UpdateData(TRUE);
	m_ExcelRW.ReadTable(m_edit_strReadPose, m_edit_strReadData);
	UpdateData(FALSE);
}

void CExcelWriteReadDlg::OnBnClickedButtonClose()
{
	m_ExcelRW.CloseTable();
	GetDlgItem(IDC_BUTTON_CLOSE)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_READ)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_WRITE)->EnableWindow(FALSE);
	GetDlgItem(IDC_BUTTON_Create)->EnableWindow(TRUE);
}

void CExcelWriteReadDlg::OnClose()
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	if(m_ExcelRW.m_bServeStart)
		m_ExcelRW.CloseTable();
	CDialogEx::OnClose();
}
