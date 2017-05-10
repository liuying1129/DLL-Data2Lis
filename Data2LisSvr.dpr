library Data2LisSvr;

uses
  ComServ,
  Data2LisSvr_TLB in 'Data2LisSvr_TLB.pas',
  UData2Lis in 'UData2Lis.pas' {Data2Lis: CoClass};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
