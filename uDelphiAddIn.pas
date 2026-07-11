unit uDelphiAddIn;

{$WARN SYMBOL_PLATFORM OFF}
{ ============================================================================
  Unit: uDelphiAddIn
  Muc dich: COM Add-in cho Excel (IDTExtensibility2), quan ly vong doi
  Form1 (dinh nghia trong unit Form) theo cac su kien cua Excel.

  Luu y quan trong ve vong doi Form:
  - OnStartupComplete: Excel da load xong -> hien Form1 (ShowForm).
  - OnBeginShutdown: Excel BAT DAU qua trinh thoat, luc nay window
  handle va COM object cua Excel VAN CON hop le -> day la thoi diem
  AN TOAN NHAT de giai phong Form1 (FreeForm). Neu doi den luc process
  that su ket thuc moi don don, thu tu unload giua VCL / Add-in DLL /
  OfficeFocusGuard.dll co the khong dong bo va gay Access Violation
  luc thoat Excel.
  - OnDisconnection: Add-in bi go/disable thu cong (khong nhat thiet
  Excel dang thoat) -> cung phai giai phong Form1 vi FExcelApp sap
  thanh nil, Form khong con COM Application hop le de thao tac.

  Ghi chu: KHONG dung "external 'InputInExcel.dll'" de goi ShowForm nhu
  ban cu, vi Form.pas va uDelphiAddIn.pas cung nam trong MOT DLL duy
  nhat (InputInExcel.dll) - tu import lai chinh minh la khong can thiet
  va tiem an rui ro (phu thuoc export table da resolve xong hay chua
  tai thoi diem goi, de vor bug kho tai hien luc Excel dang khoi tao
  COM). Goi thang qua tham chieu unit Form la du va an toan hon.
  ============================================================================ }

interface

uses
  Excel2010,
  Winapi.Messages,
  System.SysUtils,
  ComObj, ActiveX, InputInExcel_TLB,
  StdVcl, AddInDesignerObjects_TLB;

type
  TXLComAddinFactory = class(TAutoObjectFactory)
    procedure UpdateRegistry(Register: Boolean); override;
  end;

  TDelphiAddIn1 = class(TAutoObject, IDelphiAddIn1, IDTExtensibility2)
  private
    procedure OnConnection(const Application: IDispatch;
      ConnectMode: ext_ConnectMode; const AddInInst: IDispatch;
      var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode;
      var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  protected

  end;

implementation

uses
  ComServ,
  Windows,
  Registry,
  Form;

{ TDelphiAddIn1 }
procedure TDelphiAddIn1.OnAddInsUpdate(var custom: PSafeArray);
begin
  //
end;

procedure TDelphiAddIn1.OnBeginShutdown(var custom: PSafeArray);
begin
  // Excel bat dau thoat - giai phong Form1 ngay tai day, TRUOC khi
  // handle/COM cua Excel mat hieu luc, de tranh crash luc thoat.
  Form.FreeForm;
end;

procedure TDelphiAddIn1.OnConnection(const Application: IDispatch;
  ConnectMode: ext_ConnectMode; const AddInInst: IDispatch;
  var custom: PSafeArray);
begin
  // Form1 := TForm1.Create(nil);
  FExcelApp := ExcelApplication(Application);
  // supports(Application, ExcelApplication, Form1.FExcelApp);
end;

procedure TDelphiAddIn1.OnDisconnection(RemoveMode: ext_DisconnectMode;
  var custom: PSafeArray);
begin
  // Add-in bi go/disable - giai phong Form1 truoc, roi moi bo tham
  // chieu Excel Application (FExcelApp co the dang duoc Form1 su dung).
  Form.FreeForm;
  FExcelApp := nil;
end;

procedure TDelphiAddIn1.OnStartupComplete(var custom: PSafeArray);
begin
  Form.ShowForm;
end;

{ TXLComAddinFactory }
procedure TXLComAddinFactory.UpdateRegistry(Register: Boolean);
var
  RootKey: HKEY;
  AddInKey: string;
  r: TRegistry;
begin
  RootKey := HKEY_CURRENT_USER;
  AddInKey := 'Software\Microsoft\Office\Excel\Addins\' + ProgID;
  r := TRegistry.Create;
  r.RootKey := RootKey;
  try
    if Register then
      if r.OpenKey(AddInKey, True) then
      begin
        r.WriteInteger('LoadBehavior', 3);
        r.WriteInteger('CommandLineSafe', 0);
        r.WriteString('FriendlyName', 'Delphi Excel Add-In');
        r.WriteString('Description', 'There is a bug in VCL');
        r.CloseKey;
      end
      else
        raise EOleError.Create('Can''t register Add-In ' + ProgID)
    else if r.KeyExists(AddInKey) then
      r.DeleteKey(AddInKey);
  finally
    r.Free;
  end;
  inherited;
end;

initialization

TXLComAddinFactory.Create(ComServer, TDelphiAddIn1, Class_DelphiAddIn1,
  ciMultiInstance, tmApartment);

end.
