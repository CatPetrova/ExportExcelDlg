#pragma once

// ExportExcelDlg �Ի���

#include <string>
#include <functional>
#include <thread>
#include <atomic>
#include "afxwin.h"
#include "afxcmn.h"

#define WM_EXPORT_COMPLETE    (WM_APP + 1)
#define WM_EXPORT_TERMINATE   (WM_APP + 2)

class AFX_EXT_CLASS ExportExcelDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ExportExcelDlg)

public:
  using ExportCB = std::function<void(ExportExcelDlg *,
      const std::basic_string<TCHAR> &) >;
	ExportExcelDlg(ExportCB *export_cb, CWnd* pParent = NULL);   // ��׼���캯��
  ExportExcelDlg(ExportCB *export_cb, const std::basic_string<TCHAR> &def_export_folder, CWnd *pParent = NULL);
	virtual ~ExportExcelDlg();
  void Show();
  void SetRange(const int &lower, const int &upper);
  bool IsReachEnd();
  void OffsetPos(const int &pos);
  int GetPos() const;
  void SetPos(const int &pos);
  const std::atomic_bool &IsTerminate() const;

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXPORT_EXCEL };
#endif

protected:
  afx_msg void OnBnClickedBtnOpen();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��
  afx_msg void OnBnClickedBtnExport();
  afx_msg void OnClose();
  virtual BOOL OnInitDialog();
  afx_msg LRESULT OnExportComplete(WPARAM wParam, LPARAM lParam);
  afx_msg LRESULT OnExportTerminate(WPARAM wParam, LPARAM lParam);

	DECLARE_MESSAGE_MAP()
private:
  enum ExportStatus {
    kExporting = 0, //���ڵ���
    kStopping = 1,  //����ֹͣ
    kStopped = 2    //�Ѿ�ֹͣ
  };
  void SetExportBtnStatus();
  void StopExport();
  std::basic_string<TCHAR> def_export_folder_;
  ExportStatus export_status_;
  CEdit export_folder_edit_;
  CButton export_btn_;
  ExportCB *export_cb_;
  std::atomic_bool terminate_flag_;
  std::thread *export_thread_;
  CProgressCtrl progress_;
};
