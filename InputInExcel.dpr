library InputInExcel;

uses
  ComServ,
  InputInExcel_TLB in 'InputInExcel_TLB.pas',
  uDelphiAddIn in 'uDelphiAddIn.pas' {DelphiAddIn1: CoClass},
  Form in 'Form\Form.pas' {Form1},
  OfficeFocusGuard in 'OfficeFocusGuard.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

{$R *.TLB}

{$R *.RES}

begin
end.
