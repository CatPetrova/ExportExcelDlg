// export_excel_dlg.cpp : 实现文件
//

#include "stdafx.h"
#include "Resource.h"
#include "export_excel_dlg.h"
#include "afxdialogex.h"
#include "ext_dll_state.h"
#include <cassert>
#include <string>
#include <thread>
#include <atomic>
#include "routinue_l.h"


// ExportExcelDlg 对话框

IMPLEMENT_DYNAMIC(ExportExcelDlg, CDialogEx)

ExportExcelDlg::ExportExcelDlg(ExportCB *export_cb, CWnd* pParent /*=NULL*/)
  : CDialogEx(IDD_EXPORT_EXCEL, pParent)
  , export_cb_(export_cb)
  , def_export_folder_(_T("\0"))
  , export_status_(kStopped)
  , export_thread_(nullptr)
  , terminate_flag_(false)
  , progress_lower_(0)
  , progress_upper_(0)
{

}

ExportExcelDlg::ExportExcelDlg(ExportCB *export_cb, const std::basic_string<TCHAR> &def_export_folder, CWnd *pParent /*= NULL*/)
  : CDialogEx(IDD_EXPORT_EXCEL, pParent)
  , export_cb_(export_cb)
  , def_export_folder_(def_export_folder)
  , export_status_(kStopped)
  , export_thread_(nullptr)
  , terminate_flag_(false)
{

}

ExportExcelDlg::~ExportExcelDlg()
{
  if (export_thread_ != nullptr) {
    terminate_flag_ = true;
    if (export_thread_->joinable())
      export_thread_->join();
    delete export_thread_;
    export_thread_ = nullptr;
  }
}

void ExportExcelDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialogEx::DoDataExchange(pDX);
  DDX_Control(pDX, IDC_EDIT_EXPORT_FOLDER, export_folder_edit_);
  DDX_Control(pDX, IDC_BTN_EXPORT, export_btn_);
  DDX_Control(pDX, IDC_PROGRESS, progress_);
}


BEGIN_MESSAGE_MAP(ExportExcelDlg, CDialogEx)
  ON_BN_CLICKED(IDC_BTN_OPEN, &ExportExcelDlg::OnBnClickedBtnOpen)
  ON_BN_CLICKED(IDC_BTN_EXPORT, &ExportExcelDlg::OnBnClickedBtnExport)
  ON_WM_CLOSE()
  ON_MESSAGE(WM_EXPORT_COMPLETE, &ExportExcelDlg::OnExportComplete)
  ON_MESSAGE(WM_EXPORT_TERMINATE, &ExportExcelDlg::OnExportTerminate)
END_MESSAGE_MAP()


// ExportExcelDlg 消息处理程序

void ExportExcelDlg::Show() {
  ExtDllState state;
  DoModal();
}

void ExportExcelDlg::SetRange(const int &lower, const int &upper) {
  progress_lower_ = lower;
  progress_upper_ = upper;
}

void ExportExcelDlg::OnBnClickedBtnOpen()
{
  CFolderPickerDialog folder_dlg(NULL, OFN_EXPLORER | OFN_HIDEREADONLY, this, 0);
  if (folder_dlg.DoModal() != IDOK)
    return;

  export_folder_edit_.SetWindowText(
      static_cast<LPCTSTR>(folder_dlg.GetPathName()));
}

void ExportExcelDlg::OnBnClickedBtnExport()
{
  CString export_folder(_T("\0"));
  assert(export_cb_ != nullptr);

  if (export_status_ == kExporting) {
    export_status_ = kStopping;
    terminate_flag_ = true;
    SetExportBtnStatus();
    return;
  }
  else if (export_status_ == kStopping) {
    return;
  }

  export_folder_edit_.GetWindowText(export_folder);
  if (export_folder.IsEmpty()) {
    MessageBox(_T("导出路径不能为空！"));
    return;
  }

  export_btn_.EnableWindow(FALSE);
  if (!routinue_l::DirectoryExists(static_cast<LPCTSTR>(export_folder))) {
    routinue_l::MakeDirectory(static_cast<LPCTSTR>(export_folder));
  }

  StopExport();
  export_thread_ = new std::thread(*export_cb_, this, static_cast<LPCTSTR>(export_folder));
  export_btn_.EnableWindow(TRUE);
  export_status_ = kExporting;
  SetExportBtnStatus();
}

void ExportExcelDlg::StopExport() {
  if (export_thread_ != nullptr) {
    terminate_flag_ = true;
    if (export_thread_->joinable())
      export_thread_->join();
    delete export_thread_;
    export_thread_ = nullptr;
  }
  terminate_flag_ = false;
  export_status_ = kStopped;
  SetExportBtnStatus();
}

void ExportExcelDlg::OnClose()
{
  StopExport();
  CDialogEx::OnClose();
}

BOOL ExportExcelDlg::OnInitDialog()
{
  CDialogEx::OnInitDialog();

  if (def_export_folder_.empty()) {
    routinue_l::GetExecuteFilePath(&def_export_folder_);
  }
  export_folder_edit_.SetWindowText(def_export_folder_.c_str());

  progress_.SetRange(progress_lower_, progress_upper_);
  ResetStatus();

  return TRUE;  // return TRUE unless you set the focus to a control
                // 异常: OCX 属性页应返回 FALSE
}

void ExportExcelDlg::SetExportBtnStatus() {
  switch (export_status_) {
  case kExporting:
    export_btn_.SetWindowText(_T("终止"));
    export_btn_.EnableWindow(TRUE);
    break;
  case kStopping:
    export_btn_.SetWindowText(_T("正在终止"));
    export_btn_.EnableWindow(FALSE);
    break;
  case kStopped:
    export_btn_.SetWindowText(_T("导出"));
    export_btn_.EnableWindow(TRUE);
    break;
  default:
    export_btn_.SetWindowText(_T("导出"));
    export_btn_.EnableWindow(TRUE);
    break;
  }
}

afx_msg LRESULT ExportExcelDlg::OnExportComplete(WPARAM wParam, LPARAM lParam)
{
  StopExport();
  CString tip(_T("\0"));
  tip.Format(_T("导出完成，共导出%d条信息。"), static_cast<int>(wParam));
  MessageBox(tip);

  ResetStatus();
  return 0;
}

afx_msg LRESULT ExportExcelDlg::OnExportTerminate(WPARAM wParam, LPARAM lParam)
{
  StopExport();
  CString tip(_T("\0"));
  tip.Format(_T("导出终止，共导出%d条信息。"), static_cast<int>(wParam));
  MessageBox(tip);

  ResetStatus();
  return 0;
}

const std::atomic_bool &ExportExcelDlg::IsTerminate() const {
  return terminate_flag_;
}

bool ExportExcelDlg::IsReachEnd() {
  int lower = 0, upper = 0;
  progress_.GetRange(lower, upper);
  if (progress_.GetPos() == upper) {
    return true;
  }
  else {
    return false;
  }
}

void ExportExcelDlg::OffsetPos(const int &pos) {
  progress_.OffsetPos(pos);
}

int ExportExcelDlg::GetPos() const{
  return progress_.GetPos();
}

void ExportExcelDlg::ResetStatus() {
  int lower = 0, upper = 0;
  progress_.GetRange(lower, upper);
  progress_.SetPos(lower);
  export_status_ = kStopped;
  SetExportBtnStatus();
}

CString ExportExcelDlg::GetExportFolder() const{
  CString folder(_T("\0"));
  export_folder_edit_.GetWindowText(folder);

  return folder;
}