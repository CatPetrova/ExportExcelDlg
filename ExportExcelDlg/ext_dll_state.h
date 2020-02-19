#ifndef EXPORT_EXCEL_DLG_EXPORT_EXCEL_DLG_EXT_DLL_STATE_H_
#define EXPORT_EXCEL_DLG_EXPORT_EXCEL_DLG_EXT_DLL_STATE_H_

class ExtDllState
{
public:
  ExtDllState();
  ~ExtDllState();
protected:
  HINSTANCE m_hInstOld;
};

#endif