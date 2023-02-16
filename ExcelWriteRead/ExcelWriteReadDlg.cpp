
// ExcelWriteReadDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelWriteRead.h"
#include "ExcelWriteReadDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CExcelWriteReadDlg 对话框



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


// CExcelWriteReadDlg 消息处理程序

BOOL CExcelWriteReadDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelWriteReadDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
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
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	if(m_ExcelRW.m_bServeStart)
		m_ExcelRW.CloseTable();
	CDialogEx::OnClose();
}
