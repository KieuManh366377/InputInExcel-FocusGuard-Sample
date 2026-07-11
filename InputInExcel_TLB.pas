unit InputInExcel_TLB;

// ************************************************************************ //
// WARNING
// -------
// The types declared in this file were generated from data read from a
// Type Library. If this type library is explicitly or indirectly (via
// another type library referring to this type library) re-imported, or the
// 'Refresh' command of the Type Library Editor activated while editing the
// Type Library, the contents of this file will be regenerated and all
// manual modifications will be lost.
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 01/07/2020 1:56:40 PM from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Users\ADMIN\Downloads\InputInExcel_Add-Ins\InputInExcel_Add-Ins\InputInExcel_Add-Ins\InputInExcel (1)
// LIBID: {727C4E3E-0B03-41EE-ABE9-69B3C1493A74}
// LCID: 0
// Helpfile:
// HelpString:
// DepndLst:
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers.
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  InputInExcelMajorVersion = 1;
  InputInExcelMinorVersion = 0;

  LIBID_InputInExcel: TGUID = '{727C4E3E-0B03-41EE-ABE9-69B3C1493A74}';

  IID_IDelphiAddIn1: TGUID = '{2A94A400-21E9-40A5-B4F1-09E2CA827CBE}';
  CLASS_DelphiAddIn1: TGUID = '{EC20E6BA-F5F1-4560-8372-6EB01A7ECFA5}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  IDelphiAddIn1 = interface;
  IDelphiAddIn1Disp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// (NOTE: Here we map each CoClass to its Default Interface)
// *********************************************************************//
  DelphiAddIn1 = IDelphiAddIn1;


// *********************************************************************//
// Interface: IDelphiAddIn1
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2A94A400-21E9-40A5-B4F1-09E2CA827CBE}
// *********************************************************************//
  IDelphiAddIn1 = interface(IDispatch)
    ['{2A94A400-21E9-40A5-B4F1-09E2CA827CBE}']
  end;

// *********************************************************************//
// DispIntf:  IDelphiAddIn1Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2A94A400-21E9-40A5-B4F1-09E2CA827CBE}
// *********************************************************************//
  IDelphiAddIn1Disp = dispinterface
    ['{2A94A400-21E9-40A5-B4F1-09E2CA827CBE}']
  end;

// *********************************************************************//
// The Class CoDelphiAddIn1 provides a Create and CreateRemote method to
// create instances of the default interface IDelphiAddIn1 exposed by
// the CoClass DelphiAddIn1. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoDelphiAddIn1 = class
    class function Create: IDelphiAddIn1;
    class function CreateRemote(const MachineName: string): IDelphiAddIn1;
  end;

implementation

uses System.Win.ComObj;

class function CoDelphiAddIn1.Create: IDelphiAddIn1;
begin
  Result := CreateComObject(CLASS_DelphiAddIn1) as IDelphiAddIn1;
end;

class function CoDelphiAddIn1.CreateRemote(const MachineName: string): IDelphiAddIn1;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DelphiAddIn1) as IDelphiAddIn1;
end;

end.

