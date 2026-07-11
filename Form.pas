unit Form;
{$WARN SYMBOL_PLATFORM OFF}
{ ============================================================================
  Unit: Form
  Muc dich: Form chinh (TForm1) hien thi tu COM Add-in (uDelphiAddIn.pas)
  khi Excel khoi dong xong (OnStartupComplete).

  Nguyen tac thiet ke:
  - Form1 la SINGLETON: ShowForm goi nhieu lan van chi co 1 the hien Form1
  (tranh tao trung form neu OnStartupComplete bi goi lai hoac user goi
  lai ham nhieu lan).
  - FreeForm la ham don gan huy Form1 mot cach chu dong, dung ben trong
  OnDisconnection / OnBeginShutdown cua Add-in de dam bao:
  + OfficeFocusGuard duoc go (OFG_Uninstall) TRUOC khi Excel that su
  pha huy window / thoat tien trinh.
  + Form1 khong con ton tai treo (dangling) khi Excel unload DLL,
  tranh Access Violation luc thoat Excel.
  ============================================================================ }

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Excel2010, Vcl.OleServer,
  Vcl.Controls, Vcl.Forms,
  Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls;

type
  TForm1 = class(TForm)
    Edit1: TEdit;
    GetRange: TButton;
    Memo1: TMemo;
    RichEdit1: TRichEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GetRangeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  FExcelApp: ExcelApplication; // TExcelApplication; khai bao su dung ExcelApp

  // Hien thi Form1. Neu Form1 da ton tai (da Show truoc do) thi chi dua
  // len truoc (BringToFront) thay vi tao moi - tranh tao trung instance.
Procedure ShowForm; stdcall;

// Giai phong Form1 mot cach chu dong va an toan. Goi ham nay TRUOC khi
// Excel thoat (OnBeginShutdown) hoac khi Add-in bi disconnect
// (OnDisconnection). Goi nhieu lan khong sao (co kiem tra Assigned).
Procedure FreeForm; stdcall;

implementation

{$R *.dfm}

uses OfficeFocusGuard;

procedure TForm1.GetRangeClick(Sender: TObject);
var
  sh: _WorksheetDisp;
begin
  if Assigned(FExcelApp) then
  begin
    sh := _WorksheetDisp(FExcelApp.ActiveSheet);
    sh.Cells._Default[1, 1].Value := 'Kieu Manh';
    ShowMessage(sh.name);
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // caFree: form tu huy sau khi dong, FormDestroy se duoc goi ke tiep
  Action := caFree;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  OfficeFocusGuard.OFG_Install(Self.Handle);
  // OfficeFocusGuard.OFG_InstallControlGuard(Edit1.Handle);
  // OfficeFocusGuard.OFG_InstallControlGuard(Memo1.Handle);
  FormStyle := fsStayOnTop;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  // OfficeFocusGuard.OFG_UninstallControlGuard(Edit1.Handle);
  // OfficeFocusGuard.OFG_UninstallControlGuard(Memo1.Handle);
  OfficeFocusGuard.OFG_Uninstall(Self.Handle);

  // Quan trong: xoa tham chieu bien global ngay khi Handle sap mat hieu
  // luc, de ShowForm/FreeForm goi sau do khong thao tac tren doi tuong
  // da bi huy (tranh Access Violation).
  Form1 := nil;
end;

Procedure ShowForm;
begin
  if not Assigned(Form1) then
  begin
    Form1 := TForm1.Create(nil);
    Form1.Show;
  end
  else
  begin
    // Form da mo san - chi dua len truoc, khong tao instance thu hai
    if Form1.WindowState = wsMinimized then
      Form1.WindowState := wsNormal;
    Form1.Show;
    Form1.BringToFront;
  end;
end;

Procedure FreeForm;
begin
  if Assigned(Form1) then
  begin
    // Goi Close (khong phai Free truc tiep) de FormClose/FormDestroy
    // chay dung trinh tu, dam bao OFG_Uninstall duoc goi truoc khi
    // Handle mat hieu luc.
    Form1.Close;
    Form1 := nil; // phong khi FormClose khong duoc trigger (VD form an)
  end;
end;

Exports ShowForm, FreeForm;

end.
