unit Data2LisSvr_TLB;

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

// PASTLWTR : 1.2
// File generated on 2008-01-17 10:54:08 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\Data2Lis\Data2LisSvr.tlb (1)
// LIBID: {50040B52-E510-4681-B301-D11EAA7B5404}
// LCID: 0
// Helpfile: 
// HelpString: Data2LisSvr Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  Data2LisSvrMajorVersion = 1;
  Data2LisSvrMinorVersion = 0;

  LIBID_Data2LisSvr: TGUID = '{50040B52-E510-4681-B301-D11EAA7B5404}';

  IID_IData2Lis: TGUID = '{354EF7BD-3F0D-4E06-97EC-D15CBE16E415}';
  CLASS_Data2Lis: TGUID = '{F80B0C16-4BA5-4F4B-8B60-8CBABEFEDBE9}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IData2Lis = interface;
  IData2LisDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  Data2Lis = IData2Lis;


// *********************************************************************//
// Interface: IData2Lis
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {354EF7BD-3F0D-4E06-97EC-D15CBE16E415}
// *********************************************************************//
  IData2Lis = interface(IDispatch)
    ['{354EF7BD-3F0D-4E06-97EC-D15CBE16E415}']
    function fData2Lis(pReceiveItemInfo: OleVariant; const pSpecNo: WideString; 
                       const pCheckDate: WideString; const pGroupName: WideString; 
                       const pSpecType: WideString; const pSpecStatus: WideString; 
                       const pEquipChar: WideString; const pCombinID: WideString; 
                       const pLisClassName: WideString; const pLisFormCaption: WideString; 
                       const pConnectString: WideString; const pQuaContSpecNoG: WideString; 
                       const pQuaContSpecNo: WideString; const pQuaContSpecNoD: WideString; 
                       const pXmlPath: WideString; pIsSure: WordBool; pHasCalaItem: WordBool; 
                       const pDiagnosetype: WideString): WordBool; stdcall;
  end;

// *********************************************************************//
// DispIntf:  IData2LisDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {354EF7BD-3F0D-4E06-97EC-D15CBE16E415}
// *********************************************************************//
  IData2LisDisp = dispinterface
    ['{354EF7BD-3F0D-4E06-97EC-D15CBE16E415}']
    function fData2Lis(pReceiveItemInfo: OleVariant; const pSpecNo: WideString; 
                       const pCheckDate: WideString; const pGroupName: WideString; 
                       const pSpecType: WideString; const pSpecStatus: WideString; 
                       const pEquipChar: WideString; const pCombinID: WideString; 
                       const pLisClassName: WideString; const pLisFormCaption: WideString; 
                       const pConnectString: WideString; const pQuaContSpecNoG: WideString; 
                       const pQuaContSpecNo: WideString; const pQuaContSpecNoD: WideString; 
                       const pXmlPath: WideString; pIsSure: WordBool; pHasCalaItem: WordBool; 
                       const pDiagnosetype: WideString): WordBool; dispid 201;
  end;

// *********************************************************************//
// The Class CoData2Lis provides a Create and CreateRemote method to          
// create instances of the default interface IData2Lis exposed by              
// the CoClass Data2Lis. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoData2Lis = class
    class function Create: IData2Lis;
    class function CreateRemote(const MachineName: string): IData2Lis;
  end;

implementation

uses ComObj;

class function CoData2Lis.Create: IData2Lis;
begin
  Result := CreateComObject(CLASS_Data2Lis) as IData2Lis;
end;

class function CoData2Lis.CreateRemote(const MachineName: string): IData2Lis;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Data2Lis) as IData2Lis;
end;

end.
