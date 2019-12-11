unit UData2Lis;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  ComObj, ActiveX, Data2LisSvr_TLB, StdVcl,Classes{TList},Windows{TCopyDataStruct}, 
  XMLIntf{IXMLDocument},XMLDoc{TXMLDocument},AdoDB,Variants{VarArrayLock},
  SysUtils{StrPas},Messages {WM_COPYDATA},Jpeg{TJPEGImage},
  ExtCtrls{TImage},DB{ftGraphic},StrUtils{ifThen},DateUtils;

type
  TData2Lis = class(TAutoObject, IData2Lis)
  protected
    function fData2Lis(pReceiveItemInfo: OleVariant; const pSpecNo, pCheckDate,
      pGroupName, pSpecType, pSpecStatus, pEquipChar, pCombinID,
      pLisClassName, pLisFormCaption, pConnectString, pQuaContSpecNoG,
      pQuaContSpecNo, pQuaContSpecNoD, pXmlPath: WideString; pIsSure,
      pHasCalaItem: WordBool; const pDiagnosetype: WideString;
      const pBarCode: WideString;
      pEquipUnid: SYSINT;
      const pReserve1: WideString;const pReserve2: WideString;const pReserve3: WideString;const pReserve4: WideString;
      pReserve5: SYSINT;pReserve6: SYSINT;pReserve7: SYSINT;pReserve8: SYSINT;
      pReserve9: Double;pReserve10: Double;pReserve11: Double;pReserve12: Double;
      pReserve13: WordBool;pReserve14: WordBool;pReserve15: WordBool;pReserve16: WordBool
      ): WordBool;
      stdcall;
  public
    procedure Initialize;override;
    destructor destroy;override;
  end;

TYPE
  TMachineItemInfo=record
    Machine_itemid:String;
    Machine_dlttype:String;
    Machine_ItemValu:String;
    Machine_Histogram:String;
    Machine_ImagePath:String;
    //该小项目传输过来后应显示的组合项目
    //如仅属于唯一的组合项目A，值为A
    //如属于多个组合项目，判断该小项目是否属于小蝴蝶设置的组合项目代码B,如是，则值为B,否则值为空
    //如不属于任何组合项目，值为空
    Machine_CombId:String;
  end;

implementation

uses ComServ;

var
  ADOConn:TADOConnection;
  ServerDateTime:TDateTime;
  CheckDate:TDateTime;
  ReceiveItemInfo:OleVariant;//从仪器接收到的项目信息(值、联机标识等)
  MachineItemInfo:TList;//从数据库取得的机器项目信息
  SpecNo:string;//
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  EquipChar:string;
  Diagnosetype:string;//优先级别
  ConnectString:string;//CalcItemPro.dll要用到
  TransItemidString:string;//CalcItemPro.dll要用到
  IfRecLog:boolean;//是否记录日志

  //病人附加信息
  PatientName:string;
  Sex:string;
  sDateOfBirth:string;
  DateOfBirth:TDateTime;
  Age:string;
  CaseNo:string;
  DeptName:string;//送检科室
  Check_Doctor:string;//送检医生
  BedNo:string;
  Diagnose:string;//临床诊断
  Issure:string;//备注
  Operator:string;//检验操作者
  GermName:string;//细菌
  His_Unid:string;//chk_con_his.Unid
  EquipUnid:integer;//设备唯一编号


//将计算项目增加或编辑到检验结果表中
procedure addOrEditCalcItem(const Aadoconnstr:Pchar;const ComboItemID:Pchar;const checkunid: integer);stdcall;external 'CalcItemPro.dll';

//将计算数据增加或编辑到检验结果表中
procedure addOrEditCalcValu(const Aadoconnstr:Pchar;const checkunid: integer;const AifInterface:boolean;const ATransItemidString:pchar);stdcall;external 'CalcItemPro.dll';

//找到表达式中小数点位数的最大值.如56.5*100+23.01的值为2
function MaxDotLen(const ACalaExp:PChar):integer;stdcall;external 'LYFunction.dll';
function Gif2Bmp(const AGifFile,ABmpFile:Pchar):boolean;stdcall;external 'LYFunction.dll';
function Png2Bmp(const APngFile,ABmpFile:Pchar):boolean;stdcall;external 'LYFunction.dll';
function CalParserValue(const CalExpress:Pchar;var ReturnValue:single):boolean;stdcall;external 'CalParser.dll';
procedure WriteLog(const ALogStr: Pchar);stdcall;external 'LYFunction.dll';


function StrToList(const SourStr:string;const Separator:string):TStrings;
//根据指定的分隔字符串(Separator)将字符串(SourStr)导入到字符串列表中
var
  vSourStr,s:string;
  ll,lll:integer;
begin
  vSourStr:=SourStr;
  Result := TStringList.Create;
  lll:=length(Separator);

  while pos(Separator,vSourStr)<>0 do
  begin
    ll:=pos(Separator,vSourStr);
    Result.Add(copy(vSourStr,1,ll-1));
    delete(vSourStr,1,ll+lll-1);
  end;  //}
  Result.Add(vSourStr);
  s:=vSourStr;
end;

function GetServerDate: TDateTime;
var
  adotempDate:tadoquery;
begin
  adotempDate:=tadoquery.Create(NIL);
  ADOTEMPDATE.Connection:=ADOConn;
  ADOTEMPDATE.Close;
  ADOTEMPDATE.SQL.Clear;
  ADOTEMPDATE.SQL.Text:='SELECT GETDATE() as ServerDate ';
  ADOTEMPDATE.Open;
  result:=ADOTEMPDATE.fieldbyname('ServerDate').AsDateTime;
  ADOTEMPDATE.Free;
end;

procedure ClearList(const List:TList);
Var
  I : Integer;
begin
  if not Assigned(List) then exit;
  for i :=0  to List.Count-1 do
  begin
    Dispose(List.Items[I]);
  end;

  List.clear;
end;

procedure ReadMachineItem;//读取指定机器项目的所有信息
Var
    adotemp11,adotemp22:tadoquery;
    PItem : ^TMachineItemInfo;
begin
    adotemp11:=tadoquery.Create(nil);
    adotemp11.Connection:=ADOConn;
    adotemp11.Close;
    adotemp11.SQL.Clear;
    adotemp11.SQL.Text:='select * from clinicchkitem where COMMWORD='''+EquipChar+''' ';   //机器项目
    Try
      adotemp11.Open;
    except
      on E:Exception do
      begin
        WriteLog(pchar('获取仪器字母'+EquipChar+'的机器项目信息失败:'+E.Message));//有此日志，一般来说是连错数据库了
        adotemp11.Free;
        exit;
      end;
    end;

    ClearList(MachineItemInfo);
    adotemp11.First;
    while not adotemp11.Eof do
    begin
      New(PItem);
      PItem^.Machine_itemid:=TRIM(adotemp11.fieldbyname('itemid').AsString);
      PItem^.Machine_dlttype:=uppercase(adotemp11.fieldbyname('dlttype').AsString); //大写

      //20160818该小项目传输过来后应显示的组合项目START
      //如仅属于唯一的组合项目A，值为A
      //如属于多个组合项目，判断该小项目是否属于小蝴蝶设置的组合项目代码B,如是，则值为B,否则值为空
      //如不属于任何组合项目，值为空
      adotemp22:=tadoquery.Create(nil);
      adotemp22.Connection:=ADOConn;
      adotemp22.Close;
      adotemp22.SQL.Clear;
      adotemp22.SQL.Text:='select distinct B.id from CombSChkItem A,combinitem B where B.Unid=A.CombUnid and A.ItemUnid='+adotemp11.fieldbyname('unid').AsString;
      adotemp22.Open;
      if adotemp22.RecordCount=1 then PItem^.Machine_CombId:=adotemp22.fieldbyname('id').AsString
        else if adotemp22.RecordCount>1 then
             begin
               if adotemp22.Locate('id',CombinID,[loCaseInsensitive]) then PItem^.Machine_CombId:=CombinID else PItem^.Machine_CombId:='';
             end else PItem^.Machine_CombId:='';
      adotemp22.Free;
      //该小项目传输过来后应显示的组合项目STOP

      MachineItemInfo.Add(PItem);

      adotemp11.Next;
    end;
    adotemp11.Close;
    adotemp11.Free;
end;

procedure ScoutIIGetItemValue;
var
  i,j:integer;
  PItem : ^TMachineItemInfo;
begin
  TransItemidString:='';
  for  i:=0  to MachineItemInfo.Count-1 do
  begin
    PItem:=MachineItemInfo.Items[i];
    //表示变体数组ReceiveItemInfo的第1维的下边界索引、上边界索引
    for j :=VarArrayLowBound(ReceiveItemInfo,1) to VarArrayHighBound(ReceiveItemInfo,1) do
    begin
      //if trim(uppercase(PItem^.Machine_dlttype))=trim(uppercase(ReceiveItemInfo[j][0])) then
      //如果ReceiveItemInfo[j][0]为浮点数，则VarToStr转为string的结果最多保留4位小数，并且四舍五入
      if SameText(trim(PItem^.Machine_dlttype),trim(VarToStr(ReceiveItemInfo[j][0]))) then
      begin
        TransItemidString:=TransItemidString+'['+PItem^.Machine_itemid+']';

        //一个结果集中可能针对同一个联机标识有多条记录,而且可能一条是结果值，另一条是图片文件，所以要判断是否为空
        if VarToStr(ReceiveItemInfo[j][1])<>'' then PItem^.Machine_ItemValu:=ReceiveItemInfo[j][1];//结果值
        if ReceiveItemInfo[j][2]<>'' then PItem^.Machine_Histogram:=ReceiveItemInfo[j][2];//直方图数据
        if ReceiveItemInfo[j][3]<>'' then PItem^.Machine_ImagePath:=ReceiveItemInfo[j][3];//图片文件及路径
        //Break;//一个结果集中可能针对同一个联机标识有多条记录
      end;
    end;
  end;
end;

function ScalarSQLCmd(AConnectionString:string;ASQL:string):string;
var
  Conn:TADOConnection;
  Qry:TAdoQuery;
begin
  Result:='';
  Conn:=TADOConnection.Create(nil);
  Conn.LoginPrompt:=false;
  Conn.ConnectionString:=AConnectionString;
  Qry:=TAdoQuery.Create(nil);
  Qry.Connection:=Conn;
  Qry.Close;
  Qry.SQL.Clear;
  Qry.SQL.Text:=ASQL;
  Try
    Qry.Open;
  except
    on E:Exception do
    begin
      WriteLog(pchar('函数ScalarSQLCmd失败:'+E.Message+'。错误的SQL:'+ASQL));
      Qry.Free;
      Conn.Free;
      exit;
    end;
  end;
  Result:=Qry.Fields[0].AsString;
  Qry.Free;
  Conn.Free;
end;

procedure addrecord(var checkunid:integer); //增加病人信息表中记录
var                
  lsh:string;
  sqlstr:string;
  adotemp11:tadoquery;
begin
    lsh:=ScalarSQLCmd(ADOConn.ConnectionString,' select dbo.uf_GetNextSerialNum('''+GroupName+''','''+FormatDateTime('YYYY-MM-DD',CheckDate)+''','''+Diagnosetype+''') ');

    sqlstr:='Insert into chk_con (checkid,check_date,combin_id,'+
    'report_date,Diagnosetype,flagetype,typeflagcase,LSH,'+
    'patientname,sex,age,Caseno,deptname,check_doctor,bedno,diagnose,Issure,Operator,GermName,His_Unid)'+
    ' values (:P_checkid,:P_check_date,:p_combin_id,'+
    ':P_report_date,:P_Diagnosetype,:P_flagetype,:P_typeflagcase,:p_LSH,'+
    ':patientname,:sex,:age,:Caseno,:deptname,:check_doctor,:bedno,:diagnose,:Issure,:Operator,:GermName,:His_Unid ) ';
    adotemp11:=tadoquery.Create(nil);
    adotemp11.Connection:=ADOConn;
    adotemp11.Close;
    adotemp11.SQL.Clear;
    adotemp11.SQL.Add(sqlstr);
    adotemp11.SQL.Add(' SELECT SCOPE_IDENTITY() AS Insert_Identity ');
    adotemp11.Parameters.ParamByName('P_checkid').Value:=EquipChar+SpecNo ;
    adotemp11.Parameters.ParamByName('P_check_date').Value:=CheckDate ;
    adotemp11.Parameters.ParamByName('p_combin_id').Value:=GroupName ;//组别
    adotemp11.Parameters.ParamByName('P_report_date').Value:=CheckDate ;
    adotemp11.Parameters.ParamByName('P_Diagnosetype').Value:=Diagnosetype ;//edit by ly 20070629 CGYXJB->Diagnosetype
    adotemp11.Parameters.ParamByName('P_flagetype').Value:=SpecType ;
    adotemp11.Parameters.ParamByName('P_typeflagcase').Value:=SpecStatus ;
    adotemp11.Parameters.ParamByName('P_lsh').Value:=lsh ;
    adotemp11.Parameters.ParamByName('patientname').Value:=patientname ;
    adotemp11.Parameters.ParamByName('sex').Value:=sex ;
    adotemp11.Parameters.ParamByName('age').Value:=age ;
    adotemp11.Parameters.ParamByName('Caseno').Value:=Caseno ;
    adotemp11.Parameters.ParamByName('deptname').Value:=deptname ;
    adotemp11.Parameters.ParamByName('check_doctor').Value:=check_doctor ;
    adotemp11.Parameters.ParamByName('bedno').Value:=bedno ;
    adotemp11.Parameters.ParamByName('diagnose').Value:=diagnose ;
    adotemp11.Parameters.ParamByName('Issure').Value:=Issure ;
    adotemp11.Parameters.ParamByName('Operator').Value:=Operator ;
    adotemp11.Parameters.ParamByName('GermName').Value:=GermName ;
    adotemp11.Parameters.ParamByName('His_Unid').Value:=His_Unid ;
    try
      adotemp11.Open;
      checkunid:=adotemp11.fieldbyname('Insert_Identity').AsInteger;
    except
      on E:Exception do
      begin
        WriteLog(pchar('Data2Lis。方法addrecord失败:'+E.Message));
      end;
    end;
    adotemp11.Free;
end;

procedure addoreditvalueRecord(const checkunid:integer); //将仪器数据增加或编辑到检验结果表中
Var
  i:integer;
  adotemp11:tadoquery;

  PItem : ^TMachineItemInfo;

  MS:TMemoryStream;
  J1:TJPEGImage;
  ti:TImage;
  adotemp22:tadoquery;//修改图片用
  adotemp55:tadoquery;

  buf: array[0..MAX_PATH] of Char;
  hinst: HMODULE;

  chk_valu_his_valueid:string;
begin
  //取得COM自身的路径
  hinst:=GetModuleHandle('Data2LisSvr.dll');
  GetModuleFileName(hinst,buf,MAX_PATH);
  //=================

  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=ADOConn;

  for i :=0  to MachineItemInfo.Count -1 do
  begin
    PItem:=MachineItemInfo.Items[i];

    if (trim(PItem^.Machine_ItemValu)='') and (trim(PItem^.Machine_Histogram)='') then//既无检验结果，也无直方图数据
    begin
      if not FileExists(PItem^.Machine_ImagePath) then continue //也无文件
      else if (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.BMP')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.JPG')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.JPEG')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.GIF')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.PNG')
          then continue;
    end;

    //检验结果、直方图数据、图像文件，至少存在其中一个，才会往下执行

    adotemp11.Close;
    adotemp11.SQL.Clear;
    adotemp11.SQL.Text:='select itemvalue,histogram,valueid from chk_valu where itemid=:p_itemid '+
                        ' and pkunid='+inttostr(checkunid);
    adotemp11.Parameters.ParamByName('p_itemid').Value:=PItem^.machine_itemid;
    adotemp11.Open;
        
    if adotemp11.RecordCount>0 then   //检验结果表中有该检验项目的情况则修改
    begin
        while not adotemp11.Eof do
        begin
          adotemp11.Edit;
          if PItem^.Machine_ItemValu<>'' then
            adotemp11.FieldByName('itemvalue').AsString:=PItem^.Machine_ItemValu;//修改结果
          if trim(PItem^.Machine_Histogram)<>'' then//修改直方图数据
            adotemp11.FieldByName('histogram').AsString:=PItem^.Machine_Histogram;
          if (FileExists(PItem^.Machine_ImagePath))and
            ((uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.BMP')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPEG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.GIF')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.PNG'))
          then//修改图片
          begin
            IF Gif2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
              PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//图片文件及路径

            IF Png2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
              PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//图片文件及路径

            MS:=TMemoryStream.Create;
            J1:=TJPEGImage.Create;

            ti:=TImage.Create(nil);
            ti.Picture.LoadFromFile(PItem^.Machine_ImagePath);J1.Assign(ti.Picture.Graphic);
            ti.Free;

            J1.SaveToStream(MS);
            J1.Free;

            adotemp22:=tadoquery.Create(nil);
            adotemp22.Connection:=ADOConn;
            adotemp22.Close;
            adotemp22.sql.clear;
            adotemp22.sql.Text:='update chk_valu set Photo=:Photo where valueid=:valueid';
            adotemp22.Parameters.ParamByName('valueid').Value:=adotemp11.fieldbyname('valueid').AsInteger;
            adotemp22.Parameters.ParamByName('Photo').LoadFromStream(MS,ftGraphic);
            adotemp22.ExecSQL;
            adotemp22.Free;
            MS.Free;
          end;
          try
            adotemp11.Post;
          except
            adotemp11.Free;
            exit;
          end;
          adotemp11.Next;
        end;
    end else                          //检验结果表中没有该检验值的情况则插入
    begin
        //全部插入(包括待计算的计算项目及手工项目)
        adotemp11.Close;
        adotemp11.Sql.Clear;
        adotemp11.Sql.text:=
        'Insert into chk_valu ('+
        ' pkunid,pkcombin_id,itemid,itemvalue,issure,Histogram,Photo,Surem2,EquipUnid) values ('+
        ':P_pkunid,:P_pkcombin_id,:P_itemid,:P_itemvalue,:P_issure,:p_Histogram,:Photo,:Surem2,:EquipUnid) ';
        adotemp11.Parameters.ParamByName('P_pkunid').Value:=checkunid ;

        if His_Unid<>'' then
        begin
          adotemp55:=tadoquery.Create(nil);
          adotemp55.Connection:=ADOConn;
          adotemp55.Close;
          adotemp55.SQL.Clear;
          adotemp55.SQL.Text:='select valueid from chk_valu_his cvh where cast(cvh.pkunid as varchar)=:pkunid and cvh.pkcombin_id=:pkcombin_id ';
          adotemp55.Parameters.ParamByName('pkunid').Value:=His_Unid;
          adotemp55.Parameters.ParamByName('pkcombin_id').Value:=PItem^.Machine_CombId;//sCombinID;
          adotemp55.Open;
          chk_valu_his_valueid:=adotemp55.fieldbyname('valueid').AsString;
          adotemp55.Free;
        end;

        adotemp11.Parameters.ParamByName('P_pkcombin_id').Value:=PItem^.Machine_CombId;//sCombinID ;
        adotemp11.Parameters.ParamByName('P_itemid').Value:=PItem^.Machine_itemid ;
        adotemp11.Parameters.ParamByName('P_itemvalue').Value:=PItem^.Machine_ItemValu ;
        adotemp11.Parameters.ParamByName('P_issure').Value:=ifThen(PItem^.Machine_CombId='','0','1') ;//如果没有组合项目就不要显示了,让操作人员自己勾选组合项目吧//trim(sCombinID)
        adotemp11.Parameters.ParamByName('Surem2').Value:=chk_valu_his_valueid ;
        if EquipUnid>0 then
          adotemp11.Parameters.ParamByName('EquipUnid').Value:=EquipUnid
        else adotemp11.Parameters.ParamByName('EquipUnid').Value:=null;

        if trim(PItem^.Machine_Histogram)<>'' then//插入直方图数据
          adotemp11.Parameters.ParamByName('p_Histogram').Value:=PItem^.Machine_Histogram 
        ELSE adotemp11.Parameters.ParamByName('p_Histogram').Value:=Unassigned;
        if (FileExists(PItem^.Machine_ImagePath))and
          ((uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.BMP')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPEG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.GIF')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.PNG'))
        THEN//插入图片
        begin
          IF Gif2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
            PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//图片文件及路径

          IF Png2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
            PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//图片文件及路径

          MS:=TMemoryStream.Create;
          J1:=TJPEGImage.Create;

          ti:=TImage.Create(nil);
          ti.Picture.LoadFromFile(PItem^.Machine_ImagePath);J1.Assign(ti.Picture.Graphic);
          ti.Free;

          J1.SaveToStream(MS);
          J1.Free;
          adotemp11.Parameters.ParamByName('Photo').LoadFromStream(MS,ftGraphic);
          MS.Free;
        end else adotemp11.Parameters.ParamByName('Photo').Value:=Unassigned;
        try
          adotemp11.EXECSql ;
        except
          on E:Exception do
          begin
            adotemp11.Free;
            WriteLog(pchar('Data2Lis。插入明细失败:'+E.Message));
            exit;
          end;
        end;
    end;
  end;
  adotemp11.Free;
end;

procedure SaveDatatoDB(var valetudinarianInfoId:integer);
//valetudinarianInfoId为病人基本信息表中的“自动增加的唯一编号”字段值
var
  adotemp11:tadoquery;
  report_doctor:string;//审核人
Begin
  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=ADOConn;
  adotemp11.Close;
  adotemp11.SQL.Clear;
  adotemp11.SQL.Text:='select * from chk_con where checkid like :P_checkid and '+
                        'CONVERT(CHAR(10),check_date,121)=:P_check_date and Diagnosetype=:Diagnosetype ';//+PoInfoSql;
  adotemp11.Parameters.ParamByName('P_checkid').Value:='%'+EquipChar+SpecNo+'%';
  adotemp11.Parameters.ParamByName('P_check_date').Value:=FormatDateTime('YYYY-MM-DD',CheckDate);
  adotemp11.Parameters.ParamByName('Diagnosetype').Value:=Diagnosetype;
  adotemp11.Open;
  report_doctor:=adotemp11.fieldbyname('report_doctor').AsString;
  
  if adotemp11.RecordCount>0 then //有该病人基本信息的情况
  begin
    valetudinarianInfoId:=adotemp11.fieldbyname('unid').AsInteger;
  end else               //没有该病人基本信息的情况
  begin
    addrecord(valetudinarianInfoId); //增加病人信息表中记录
  end;
  adotemp11.Free;

  if report_doctor<>'' then exit;//表示已审核的检验单，不修改其结果
  
  addoreditvalueRecord(valetudinarianInfoId);   //增加或编辑检验结果表中记录

  addOrEditCalcItem(pchar(ConnectString),pchar(trim(CombinID)),valetudinarianInfoId);    //插入计算项目
  addOrEditCalcValu(pchar(ConnectString),valetudinarianInfoId,true,pchar(TransItemidString));    //将计算数据增加或编辑到检验结果表中
end;

procedure SendMsgToLIS(const valetudinarianInfoId:integer);
var
 Cds: TCopyDataStruct;
 Hwnd: THandle;
begin
 Cds.cbData := Length (inttostr(valetudinarianInfoId)) + 1;
 GetMem (Cds.lpData, Cds.cbData );
 StrCopy (Cds.lpData, PChar (inttostr(valetudinarianInfoId)));
 Hwnd := FindWindow ('tsdiappform',PCHAR(TRIM(LisFormCaption)));
 if Hwnd <> 0 then
 begin
    SendMessage (Hwnd, WM_COPYDATA, 0, Cardinal(@Cds)) ;
 end;
 FreeMem (Cds.lpData);
end;

procedure SaveDataToQuaContDB(Const QuaContType:Integer);
//QuaContType:1--高;0--常;-1:低
var
  adotemp,adotemp22:tadoquery;
  i:integer;
  head_unid:integer;
  sqlstr22,sDate:string;
  PItem : ^TMachineItemInfo;
  sItemValu,calc_express:string;
  iMaxDotLen:integer;
  l_ReturnValue:Single;
begin
  sDate:=FormatDateTime('YYYYMMDD',ServerDateTime);

  for i :=0  to MachineItemInfo.Count-1 do
  begin
    PItem:=MachineItemInfo.Items[i];
    if PItem^.Machine_ItemValu='' then Continue; 

    //自计算项目
    sItemValu:=PItem^.Machine_ItemValu;
    adotemp22:=tadoquery.Create(nil);
    adotemp22.Connection:=ADOConn;
    adotemp22.Close;
    adotemp22.SQL.Clear;
    adotemp22.SQL.Text:='select caculexpress from clinicchkitem where itemid='''+PItem^.Machine_itemid+''' and ltrim(rtrim(isnull(clinicchkitem.caculexpress,'''')))<>'''' ';
    adotemp22.Open;
    if adotemp22.RecordCount>0 then //该项目有计算公式
    begin
      calc_express:=adotemp22.fieldbyname('caculexpress').AsString;
      calc_express:=StringReplace(calc_express,'['+PItem^.Machine_itemid+']',PItem^.Machine_ItemValu,[rfReplaceAll,rfIgnoreCase]);

      iMaxDotLen:=MaxDotLen(pchar(calc_express));
      if CalParserValue(Pchar(calc_express),l_ReturnValue) then
        sItemValu:=format('%.'+inttostr(iMaxDotLen)+'f',[l_ReturnValue]);
    end;
    adotemp22.Free;
    //==========
    
    adotemp:=tadoquery.Create(nil);
    adotemp.Connection:=ADOConn;
    adotemp.Close;
    adotemp.SQL.Clear;
    adotemp.SQL.Text:='select * from qcghead where itemID=:itemID and '+
                      'qc_year=:P_qc_year and qc_month=:P_qc_month  '; 
    adotemp.Parameters.ParamByName('itemID').Value:=PItem^.Machine_itemid;
    adotemp.Parameters.ParamByName('P_qc_year').Value:=Copy(sDate,1,4);
    adotemp.Parameters.ParamByName('P_qc_month').Value:=Copy(sDate,5,2);
    adotemp.Open;
    if adotemp.RecordCount>0 then //有该项目的质控值的情况
    begin
      head_unid:=adotemp.fieldbyname('unid').AsInteger;
    end else               //没有该项目的质控值的情况
    begin
      sqlstr22:='Insert into qcghead (itemID,qc_year,qc_month)'+
      ' values (:itemID,:P_qc_year,:p_qc_month)';
      adotemp.Close;
      adotemp.SQL.Clear;
      adotemp.SQL.Add(sqlstr22);
      adotemp.SQL.Add(' SELECT SCOPE_IDENTITY() AS Insert_Identity ');
      adotemp.Parameters.ParamByName('itemID').Value:=PItem^.Machine_itemid;
      adotemp.Parameters.ParamByName('P_qc_year').Value:=Copy(sDate,1,4);
      adotemp.Parameters.ParamByName('p_qc_month').Value:=Copy(sDate,5,2);
      try
        adotemp.Open;
      except
        adotemp.Free;
        exit;
      end;
      head_unid:=adotemp.fieldbyname('Insert_Identity').AsInteger;
    end;

    adotemp.Close;
    adotemp.SQL.Clear;
    adotemp.SQL.Text:='select * from qcgdata where pkunid=:P_pkunid'+
                        ' and gettime=:P_gettime ';
    adotemp.Parameters.ParamByName('p_pkunid').Value:=head_unid;
    adotemp.Parameters.ParamByName('P_gettime').Value:=StrToIntDef(Copy(sDate,7,2),0);
    adotemp.Open;
        
    if adotemp.RecordCount>0 then   //检验结果表中有该检验值的情况则修改
    begin
      adotemp.Edit;
      case QuaContType of
        1:adotemp.FieldByName('hv_result').AsString:=sItemValu;
        0:adotemp.FieldByName('result').AsString:=sItemValu;
        -1:adotemp.FieldByName('lv_result').AsString:=sItemValu;
      end;
      try
        adotemp.Post;
      except
      end;
    end else                          //检验结果表中没有该检验值的情况则插入
    begin
      adotemp.Close;
      adotemp.Sql.Clear;
      case QuaContType of
        1:
        adotemp.Sql.text:=
        'Insert into qcgdata ('+
        ' pkunid,gettime,hv_result) values ('+
        ':P_pkunid,:P_gettime,:P_result) ';
        0:
        adotemp.Sql.text:=
        'Insert into qcgdata ('+
        ' pkunid,gettime,result) values ('+
        ':P_pkunid,:P_gettime,:P_result) ';
        -1:
        adotemp.Sql.text:=
        'Insert into qcgdata ('+
        ' pkunid,gettime,lv_result) values ('+
        ':P_pkunid,:P_gettime,:P_result) ';
      end;
      adotemp.Parameters.ParamByName('P_pkunid').Value:=head_unid ;
      adotemp.Parameters.ParamByName('P_gettime').Value:=StrToIntDef(Copy(sDate,7,2),0);
      adotemp.Parameters.ParamByName('P_result').Value:=sItemValu;
      try
        adotemp.EXECSql ;
      except
      end;
    end;
    adotemp.Free;
  end;//for
end;

{ TData2Lis }

function TData2Lis.fData2Lis(pReceiveItemInfo: OleVariant; const pSpecNo,
  pCheckDate, pGroupName, pSpecType, pSpecStatus, pEquipChar, pCombinID,
  pLisClassName, pLisFormCaption, pConnectString, pQuaContSpecNoG,
  pQuaContSpecNo, pQuaContSpecNoD, pXmlPath: WideString; pIsSure,
  pHasCalaItem: WordBool; const pDiagnosetype: WideString;
  const pBarCode: WideString;
  pEquipUnid: SYSINT;
  const pReserve1: WideString;const pReserve2: WideString;const pReserve3: WideString;const pReserve4: WideString;
  pReserve5: SYSINT;pReserve6: SYSINT;pReserve7: SYSINT;pReserve8: SYSINT;
  pReserve9: Double;pReserve10: Double;pReserve11: Double;pReserve12: Double;
  pReserve13: WordBool;pReserve14: WordBool;pReserve15: WordBool;pReserve16: WordBool
  ): WordBool;
var
  valetudinarianInfoId,i,j,k:integer;
  XMLDocument:IXMLDocument;
  ItemInfo:IXMLNode;
  lsPatientOtherInfo:TStrings;
  adotemp11,adotemp22:tadoquery;
  fs:TFormatSettings;
  LogStr:string;
  //sBarCode:string;
begin
  ADOConn.ConnectionString:=pConnectString;

  ServerDateTime:=GetServerDate();
  
  IF VarIsEmpty(pReceiveItemInfo) THEN//如果pReceiveItemInfo传空值(Unassigned)进来,则采用pXmlPath
  begin
    if not FileExists(pXmlPath) then begin result:=false;exit;end;//采用pXmlPath的方式,可以做个假文件欺骗系统
    //将项目值及联机标识导入ReceiveItemInfo中
    XMLDocument:=TXMLDocument.Create(pXmlPath);//如果文件不存在,此句报错
    ReceiveItemInfo:=VarArrayCreate([0,XMLDocument.DocumentElement.ChildNodes.Count-1],varVariant);
    for i :=0  to XMLDocument.DocumentElement.ChildNodes.Count-1 do
    begin
      ItemInfo:=XMLDocument.DocumentElement.ChildNodes[i];
      ReceiveItemInfo[i]:=VarArrayof([ItemInfo.ChildNodes['DltType'].Text,ItemInfo.ChildNodes['Value'].Text,ItemInfo.ChildNodes['Histogram'].Text,ItemInfo.ChildNodes['ImagePath'].Text]);
    end;
    //取样本号属性//暂不用了,但代码及XML字段位置保留,以便向前兼容及该位置XML字段的扩展(注释20110609)
    j:=XMLDocument.DocumentElement.AttributeNodes.count;
    for i :=0  to j-1 do
    begin
      ItemInfo:=XMLDocument.DocumentElement.AttributeNodes[i];
    end;
  end else
  begin
    ReceiveItemInfo:=pReceiveItemInfo;
  end;

  //记录调试日志start
  IfRecLog:=pIsSure;
  if IfRecLog then
  begin
    LogStr:='联机号:'+pSpecNo+';日期:'+pCheckDate+';样本类型:'+pSpecType+';样本状态:'+pSpecStatus+';优先级别:'+pDiagnosetype;
    WriteLog(pchar(LogStr));
    //表示变体数组ReceiveItemInfo的第1维的下边界索引、上边界索引
    for j :=VarArrayLowBound(ReceiveItemInfo,1) to VarArrayHighBound(ReceiveItemInfo,1) do
    begin
      LogStr:='联机标识:'+VarToStr(ReceiveItemInfo[j][0])+';'+'结果值:'+VarToStr(ReceiveItemInfo[j][1])+';'+'直方图数据:'+ReceiveItemInfo[j][2]+';'+'图片文件及路径:'+ReceiveItemInfo[j][3];
      WriteLog(pchar(LogStr));
    end;
    WriteLog('样本数据结束');
  end;
  //记录调试日志stop

  SpecNo:=pSpecNo;

  fs.DateSeparator:='-';
  fs.TimeSeparator:=':';
  fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
  CheckDate:=StrtoDateTimeDef(pCheckDate,ServerDateTime,fs);//edit by liuying 20111114 StrtoDateDef->StrtoDateTimeDef
  if  CheckDate<2 then ReplaceDate(CheckDate,ServerDateTime);//表示1899-12-30,没有给日期赋值
  if (HourOf(CheckDate)=0) and (MinuteOf(CheckDate)=0) and (SecondOf(CheckDate)=0) then ReplaceTime(CheckDate,ServerDateTime);//表示没有给时间赋值

  GroupName:=pGroupName;
  SpecType:=pSpecType ;
  SpecStatus:=pSpecStatus ;
  CombinID:=pCombinID;
  LisFormCaption:=pLisFormCaption;
  EquipChar:=pEquipChar;
  Diagnosetype:=pDiagnosetype;
  ConnectString:=pConnectString;
  EquipUnid:=pEquipUnid;

  //sBarCode:=trim(pBarCode);
  if trim(pBarCode)<>'' then
  begin
    adotemp22:=tadoquery.Create(nil);
    adotemp22.Connection:=ADOConn;
    adotemp22.Close;
    adotemp22.SQL.Clear;
    adotemp22.SQL.Text:='select cch.unid from chk_con_his cch where dbo.uf_GetExtBarcode(cch.unid) like ''%,'+trim(pBarCode)+',%'' ';
    adotemp22.Open;
    His_Unid:=adotemp22.fieldbyname('unid').AsString;
    adotemp22.Free;
  end;

  //2010-04-05 add by liuying
  lsPatientOtherInfo:=StrToList(pLisClassName,'{!@#}');
  for k:=0 to lsPatientOtherInfo.Count-1 do
  begin
    if k+1=1 then PatientName:=lsPatientOtherInfo[k];
    if k+1=2 then Sex:=lsPatientOtherInfo[k];
    if k+1=3 then sDateOfBirth:=lsPatientOtherInfo[k];
    if k+1=4 then Age:=lsPatientOtherInfo[k];
    if k+1=5 then CaseNo:=lsPatientOtherInfo[k];
    if k+1=6 then DeptName:=lsPatientOtherInfo[k];
    if k+1=7 then Check_Doctor:=lsPatientOtherInfo[k];
    if k+1=8 then BedNo:=lsPatientOtherInfo[k];
    if k+1=9 then Diagnose:=lsPatientOtherInfo[k];
    if k+1=10 then Issure:=lsPatientOtherInfo[k];
    if k+1=11 then Operator:=lsPatientOtherInfo[k];
    if k+1=12 then GermName:=lsPatientOtherInfo[k];
  end;
  lsPatientOtherInfo.Free;
  if sDateOfBirth<>'' then//根据出生日期算年龄
  begin
    DateOfBirth:=StrtoDateDef(sDateOfBirth,ServerDateTime);
    adotemp11:=tadoquery.Create(nil);
    adotemp11.Connection:=ADOConn;
    adotemp11.Close;
    adotemp11.SQL.Clear;
    adotemp11.SQL.Text:='select dbo.uf_GetAge('''+sDateOfBirth+''','''+pCheckDate+''') as sAge';
    try
      adotemp11.Open;
      Age:=adotemp11.fieldbyname('sAge').AsString;
    except
    end;
    adotemp11.Free;
  end;
  //2010-04-05 add by liuying

  ReadMachineItem;
  ScoutIIGetItemValue;
  if SpecNo=pQuaContSpecNo then
  begin     //保存到常值质控表中
      SaveDataToQuaContDB(0);
  end else
  if SpecNo=pQuaContSpecNoD then
  begin     //保存到低值质控表中
      SaveDataToQuaContDB(-1);
  end else
  if SpecNo=pQuaContSpecNoG then
  begin    //保存到高值质控表中
      SaveDataToQuaContDB(1);
  end else  //保存到病人表中
  begin
      SaveDatatoDB(valetudinarianInfoId);
      SendMsgToLIS(valetudinarianInfoId);
  end;
  //}
  result:=true;
end;

procedure TData2Lis.Initialize;
begin
  inherited;
  ADOConn:=TADOConnection.Create(nil);
  MachineItemInfo:=TList.Create;
end;

destructor TData2Lis.destroy;
begin
  ADOConn.Close;
  ADOConn.Free;
  ClearList(MachineItemInfo);MachineItemInfo.Free;
  inherited;
end;

initialization
  TAutoObjectFactory.Create(ComServer, TData2Lis, Class_Data2Lis,
    ciMultiInstance, tmApartment);
end.
