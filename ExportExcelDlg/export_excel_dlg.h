#pragma once

// ExportExcelDlg 对话框

#include <string>
#include <functional>
#include <thread>
#include <atomic>
#include "afxwin.h"
#include "afxcmn.h"

#define WM_EXPORT_COMPLETE    (WM_APP + 1) //导出完成，WPARAM表示共导出多少条信息
#define WM_EXPORT_TERMINATE   (WM_APP + 2) //导出终止，WPARAM表示共导出多少条信息

class AFX_EXT_CLASS ExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ExportExcelDlg)

public:
  using ExportCB = std::function<void(ExportExcelDlg *, std::basic_string<TCHAR>)>;
	ExportExcelDlg(ExportCB *export_cb, CWnd* pParent = NULL);   // 标准构造函数
  ExportExcelDlg(ExportCB *export_cb, const std::basic_string<TCHAR> &def_export_folder, CWnd *pParent = NULL);
	virtual ~ExportExcelDlg();
  void Show();
  void SetRange(const int &lower, const int &upper);
  bool IsReachEnd();
  void OffsetPos(const int &pos);
  int GetPos() const;
  void SetPos(const int &pos);
  const std::atomic_bool &IsTerminate() const;
  CString GetExportFolder() const;

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXPORT_EXCEL };
#endif

protected:
  afx_msg void OnBnClickedBtnOpen();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
  afx_msg void OnBnClickedBtnExport();
  afx_msg void OnClose();
  virtual BOOL OnInitDialog();
  afx_msg LRESULT OnExportComplete(WPARAM wParam, LPARAM lParam);
  afx_msg LRESULT OnExportTerminate(WPARAM wParam, LPARAM lParam);

	DECLARE_MESSAGE_MAP()
private:
  enum ExportStatus {
    kExporting = 0, //正在导出
    kStopping = 1,  //正在停止
    kStopped = 2    //已经停止
  };
  void SetExportBtnStatus();
  void StopExport();
  void ResetStatus();
  std::basic_string<TCHAR> def_export_folder_;
  ExportStatus export_status_;
  CEdit export_folder_edit_;
  CButton export_btn_;
  ExportCB *export_cb_;
  std::atomic_bool terminate_flag_;
  std::thread *export_thread_;
  CProgressCtrl progress_;
  int progress_lower_;
  int progress_upper_;
};
