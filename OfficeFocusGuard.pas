unit OfficeFocusGuard;

{ ============================================================================
  Unit: OfficeFocusGuard
  Muc dich: Binding Delphi cho OfficeFocusGuard.dll - chan keyboard leak
  2 chieu giua Form/control cua Add-in va Office (Excel/Word/PowerPoint).

  Day CHI la khai bao external, khong chua logic - toan bo logic nam
  trong OfficeFocusGuard.dll (viet bang C++ Builder, dung chung cho
  nhieu ngon ngu: Delphi, Visual C++, C#...). Xem chi tiet co che trong
  OfficeFocusGuard.cpp.

  Cach dung (giong het ban unit .pas cu, chi doi ten ham va truyen
  .Handle thay vi object):
  1. uses OfficeFocusGuard;
  2. Trong FormCreate (hoac Initialize cua TActiveForm):
  OfficeFocusGuard.OFG_Install(Self.Handle);
  OfficeFocusGuard.OFG_InstallControlGuard(Edit1.Handle);
  OfficeFocusGuard.OFG_InstallControlGuard(Memo1.Handle);
  3. Trong FormDestroy:
  OfficeFocusGuard.OFG_UninstallControlGuard(Edit1.Handle);
  OfficeFocusGuard.OFG_UninstallControlGuard(Memo1.Handle);
  OfficeFocusGuard.OFG_Uninstall(Self.Handle);

  Luu y: OfficeFocusGuard.dll phai nam cung thu muc voi file .dll/.xll
  cua Add-in (hoac trong PATH), neu khong se bao loi "khong tim thay
  entry point" luc load Add-in.

  Delphi 13.1 / Windows x64
  ============================================================================ }

interface

uses
  Winapi.Windows;

const
  OFG_DLL_NAME = 'OfficeFocusGuard.dll';

  // Cai guard cho 1 Form (top-level window). Goi 1 lan luc Form duoc tao.
  // Tra ve 1 = thanh cong, 0 = that bai.
function OFG_Install(AFormHandle: HWND): Integer; stdcall;
  external OFG_DLL_NAME;

// Go guard cho AFormHandle - goi luc Form bi huy, TRUOC khi Handle mat
// hieu luc.
function OFG_Uninstall(AFormHandle: HWND): Integer; stdcall;
  external OFG_DLL_NAME;

// Subclass WndProc cua 1 control con (Edit/Memo...). Goi SAU KHI control
// da co Handle that su.
function OFG_InstallControlGuard(AControlHandle: HWND): Integer; stdcall;
  external OFG_DLL_NAME;

// Go subclass cho AControlHandle - goi TRUOC khi Handle mat hieu luc.
function OFG_UninstallControlGuard(AControlHandle: HWND): Integer; stdcall;
  external OFG_DLL_NAME;

implementation

{
  procedure TForm1.FormCreate(Sender: TObject);
  begin
  OfficeFocusGuard.OFG_Install(Self.Handle);
  FormStyle := fsStayOnTop;
  end;

  procedure TForm1.FormDestroy(Sender: TObject);
  begin
  OfficeFocusGuard.OFG_Uninstall(Self.Handle);
  end;
}
end.
