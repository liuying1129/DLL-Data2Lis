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
    //��С��Ŀ���������Ӧ��ʾ�������Ŀ
    //�������Ψһ�������ĿA��ֵΪA
    //�����ڶ�������Ŀ���жϸ�С��Ŀ�Ƿ�����С�������õ������Ŀ����B,���ǣ���ֵΪB,����ֵΪ��
    //�粻�����κ������Ŀ��ֵΪ��
    Machine_CombId:String;
  end;

implementation

uses ComServ;

var
  ADOConn:TADOConnection;
  ServerDateTime:TDateTime;
  CheckDate:TDateTime;
  ReceiveItemInfo:OleVariant;//���������յ�����Ŀ��Ϣ(ֵ��������ʶ��)
  MachineItemInfo:TList;//�����ݿ�ȡ�õĻ�����Ŀ��Ϣ
  SpecNo:string;//
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  EquipChar:string;
  Diagnosetype:string;//���ȼ���
  ConnectString:string;//CalcItemPro.dllҪ�õ�
  TransItemidString:string;//CalcItemPro.dllҪ�õ�
  IfRecLog:boolean;//�Ƿ��¼��־

  //���˸�����Ϣ
  PatientName:string;
  Sex:string;
  sDateOfBirth:string;
  DateOfBirth:TDateTime;
  Age:string;
  CaseNo:string;
  DeptName:string;//�ͼ����
  Check_Doctor:string;//�ͼ�ҽ��
  BedNo:string;
  Diagnose:string;//�ٴ����
  Issure:string;//��ע
  Operator:string;//���������
  GermName:string;//ϸ��
  His_Unid:string;//chk_con_his.Unid
  EquipUnid:integer;//�豸Ψһ���


//��������Ŀ���ӻ�༭������������
procedure addOrEditCalcItem(const Aadoconnstr:Pchar;const ComboItemID:Pchar;const checkunid: integer);stdcall;external 'CalcItemPro.dll';

//�������������ӻ�༭������������
procedure addOrEditCalcValu(const Aadoconnstr:Pchar;const checkunid: integer;const AifInterface:boolean;const ATransItemidString:pchar);stdcall;external 'CalcItemPro.dll';

//�ҵ����ʽ��С����λ�������ֵ.��56.5*100+23.01��ֵΪ2
function MaxDotLen(const ACalaExp:PChar):integer;stdcall;external 'LYFunction.dll';
function Gif2Bmp(const AGifFile,ABmpFile:Pchar):boolean;stdcall;external 'LYFunction.dll';
function Png2Bmp(const APngFile,ABmpFile:Pchar):boolean;stdcall;external 'LYFunction.dll';
function CalParserValue(const CalExpress:Pchar;var ReturnValue:single):boolean;stdcall;external 'CalParser.dll';
procedure WriteLog(const ALogStr: Pchar);stdcall;external 'LYFunction.dll';


function StrToList(const SourStr:string;const Separator:string):TStrings;
//����ָ���ķָ��ַ���(Separator)���ַ���(SourStr)���뵽�ַ����б���
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

procedure ReadMachineItem;//��ȡָ��������Ŀ��������Ϣ
Var
    adotemp11,adotemp22:tadoquery;
    PItem : ^TMachineItemInfo;
begin
    adotemp11:=tadoquery.Create(nil);
    adotemp11.Connection:=ADOConn;
    adotemp11.Close;
    adotemp11.SQL.Clear;
    adotemp11.SQL.Text:='select * from clinicchkitem where COMMWORD='''+EquipChar+''' ';   //������Ŀ
    Try
      adotemp11.Open;
    except
      on E:Exception do
      begin
        WriteLog(pchar('��ȡ������ĸ'+EquipChar+'�Ļ�����Ŀ��Ϣʧ��:'+E.Message));//�д���־��һ����˵���������ݿ���
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
      PItem^.Machine_dlttype:=uppercase(adotemp11.fieldbyname('dlttype').AsString); //��д

      //20160818��С��Ŀ���������Ӧ��ʾ�������ĿSTART
      //�������Ψһ�������ĿA��ֵΪA
      //�����ڶ�������Ŀ���жϸ�С��Ŀ�Ƿ�����С�������õ������Ŀ����B,���ǣ���ֵΪB,����ֵΪ��
      //�粻�����κ������Ŀ��ֵΪ��
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
      //��С��Ŀ���������Ӧ��ʾ�������ĿSTOP

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
    //��ʾ��������ReceiveItemInfo�ĵ�1ά���±߽��������ϱ߽�����
    for j :=VarArrayLowBound(ReceiveItemInfo,1) to VarArrayHighBound(ReceiveItemInfo,1) do
    begin
      //if trim(uppercase(PItem^.Machine_dlttype))=trim(uppercase(ReceiveItemInfo[j][0])) then
      //���ReceiveItemInfo[j][0]Ϊ����������VarToStrתΪstring�Ľ����ౣ��4λС����������������
      if SameText(trim(PItem^.Machine_dlttype),trim(VarToStr(ReceiveItemInfo[j][0]))) then
      begin
        TransItemidString:=TransItemidString+'['+PItem^.Machine_itemid+']';

        //һ��������п������ͬһ��������ʶ�ж�����¼,���ҿ���һ���ǽ��ֵ����һ����ͼƬ�ļ�������Ҫ�ж��Ƿ�Ϊ��
        if VarToStr(ReceiveItemInfo[j][1])<>'' then PItem^.Machine_ItemValu:=ReceiveItemInfo[j][1];//���ֵ
        if ReceiveItemInfo[j][2]<>'' then PItem^.Machine_Histogram:=ReceiveItemInfo[j][2];//ֱ��ͼ����
        if ReceiveItemInfo[j][3]<>'' then PItem^.Machine_ImagePath:=ReceiveItemInfo[j][3];//ͼƬ�ļ���·��
        //Break;//һ��������п������ͬһ��������ʶ�ж�����¼
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
      WriteLog(pchar('����ScalarSQLCmdʧ��:'+E.Message+'�������SQL:'+ASQL));
      Qry.Free;
      Conn.Free;
      exit;
    end;
  end;
  Result:=Qry.Fields[0].AsString;
  Qry.Free;
  Conn.Free;
end;

procedure addrecord(var checkunid:integer); //���Ӳ�����Ϣ���м�¼
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
    adotemp11.Parameters.ParamByName('p_combin_id').Value:=GroupName ;//���
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
        WriteLog(pchar('Data2Lis������addrecordʧ��:'+E.Message));
      end;
    end;
    adotemp11.Free;
end;

procedure addoreditvalueRecord(const checkunid:integer); //�������������ӻ�༭������������
Var
  i:integer;
  adotemp11:tadoquery;

  PItem : ^TMachineItemInfo;

  MS:TMemoryStream;
  J1:TJPEGImage;
  ti:TImage;
  adotemp22:tadoquery;//�޸�ͼƬ��
  adotemp55:tadoquery;

  buf: array[0..MAX_PATH] of Char;
  hinst: HMODULE;

  chk_valu_his_valueid:string;
begin
  //ȡ��COM�����·��
  hinst:=GetModuleHandle('Data2LisSvr.dll');
  GetModuleFileName(hinst,buf,MAX_PATH);
  //=================

  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=ADOConn;

  for i :=0  to MachineItemInfo.Count -1 do
  begin
    PItem:=MachineItemInfo.Items[i];

    if (trim(PItem^.Machine_ItemValu)='') and (trim(PItem^.Machine_Histogram)='') then//���޼�������Ҳ��ֱ��ͼ����
    begin
      if not FileExists(PItem^.Machine_ImagePath) then continue //Ҳ���ļ�
      else if (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.BMP')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.JPG')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.JPEG')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.GIF')
          and (uppercase(ExtractFileExt(PItem^.Machine_ImagePath))<>'.PNG')
          then continue;
    end;

    //��������ֱ��ͼ���ݡ�ͼ���ļ������ٴ�������һ�����Ż�����ִ��

    adotemp11.Close;
    adotemp11.SQL.Clear;
    adotemp11.SQL.Text:='select itemvalue,histogram,valueid from chk_valu where itemid=:p_itemid '+
                        ' and pkunid='+inttostr(checkunid);
    adotemp11.Parameters.ParamByName('p_itemid').Value:=PItem^.machine_itemid;
    adotemp11.Open;
        
    if adotemp11.RecordCount>0 then   //�����������иü�����Ŀ��������޸�
    begin
        while not adotemp11.Eof do
        begin
          adotemp11.Edit;
          if PItem^.Machine_ItemValu<>'' then
            adotemp11.FieldByName('itemvalue').AsString:=PItem^.Machine_ItemValu;//�޸Ľ��
          if trim(PItem^.Machine_Histogram)<>'' then//�޸�ֱ��ͼ����
            adotemp11.FieldByName('histogram').AsString:=PItem^.Machine_Histogram;
          if (FileExists(PItem^.Machine_ImagePath))and
            ((uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.BMP')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPEG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.GIF')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.PNG'))
          then//�޸�ͼƬ
          begin
            IF Gif2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
              PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//ͼƬ�ļ���·��

            IF Png2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
              PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//ͼƬ�ļ���·��

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
    end else                          //����������û�иü���ֵ����������
    begin
        //ȫ������(����������ļ�����Ŀ���ֹ���Ŀ)
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
        adotemp11.Parameters.ParamByName('P_issure').Value:=ifThen(PItem^.Machine_CombId='','0','1') ;//���û�������Ŀ�Ͳ�Ҫ��ʾ��,�ò�����Ա�Լ���ѡ�����Ŀ��//trim(sCombinID)
        adotemp11.Parameters.ParamByName('Surem2').Value:=chk_valu_his_valueid ;
        if EquipUnid>0 then
          adotemp11.Parameters.ParamByName('EquipUnid').Value:=EquipUnid
        else adotemp11.Parameters.ParamByName('EquipUnid').Value:=null;

        if trim(PItem^.Machine_Histogram)<>'' then//����ֱ��ͼ����
          adotemp11.Parameters.ParamByName('p_Histogram').Value:=PItem^.Machine_Histogram 
        ELSE adotemp11.Parameters.ParamByName('p_Histogram').Value:=Unassigned;
        if (FileExists(PItem^.Machine_ImagePath))and
          ((uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.BMP')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.JPEG')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.GIF')or(uppercase(ExtractFileExt(PItem^.Machine_ImagePath))='.PNG'))
        THEN//����ͼƬ
        begin
          IF Gif2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
            PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//ͼƬ�ļ���·��

          IF Png2Bmp(pchar(PItem^.Machine_ImagePath),pchar(ChangeFileExt(strpas(buf),'.bmp'))) THEN
            PItem^.Machine_ImagePath:=ChangeFileExt(strpas(buf),'.bmp');//ͼƬ�ļ���·��

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
            WriteLog(pchar('Data2Lis��������ϸʧ��:'+E.Message));
            exit;
          end;
        end;
    end;
  end;
  adotemp11.Free;
end;

procedure SaveDatatoDB(var valetudinarianInfoId:integer);
//valetudinarianInfoIdΪ���˻�����Ϣ���еġ��Զ����ӵ�Ψһ��š��ֶ�ֵ
var
  adotemp11:tadoquery;
  report_doctor:string;//�����
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
  
  if adotemp11.RecordCount>0 then //�иò��˻�����Ϣ�����
  begin
    valetudinarianInfoId:=adotemp11.fieldbyname('unid').AsInteger;
  end else               //û�иò��˻�����Ϣ�����
  begin
    addrecord(valetudinarianInfoId); //���Ӳ�����Ϣ���м�¼
  end;
  adotemp11.Free;

  if report_doctor<>'' then exit;//��ʾ����˵ļ��鵥�����޸�����
  
  addoreditvalueRecord(valetudinarianInfoId);   //���ӻ�༭���������м�¼

  addOrEditCalcItem(pchar(ConnectString),pchar(trim(CombinID)),valetudinarianInfoId);    //���������Ŀ
  addOrEditCalcValu(pchar(ConnectString),valetudinarianInfoId,true,pchar(TransItemidString));    //�������������ӻ�༭������������
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
//QuaContType:1--��;0--��;-1:��
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

    //�Լ�����Ŀ
    sItemValu:=PItem^.Machine_ItemValu;
    adotemp22:=tadoquery.Create(nil);
    adotemp22.Connection:=ADOConn;
    adotemp22.Close;
    adotemp22.SQL.Clear;
    adotemp22.SQL.Text:='select caculexpress from clinicchkitem where itemid='''+PItem^.Machine_itemid+''' and ltrim(rtrim(isnull(clinicchkitem.caculexpress,'''')))<>'''' ';
    adotemp22.Open;
    if adotemp22.RecordCount>0 then //����Ŀ�м��㹫ʽ
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
    if adotemp.RecordCount>0 then //�и���Ŀ���ʿ�ֵ�����
    begin
      head_unid:=adotemp.fieldbyname('unid').AsInteger;
    end else               //û�и���Ŀ���ʿ�ֵ�����
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
        
    if adotemp.RecordCount>0 then   //�����������иü���ֵ��������޸�
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
    end else                          //����������û�иü���ֵ����������
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
  
  IF VarIsEmpty(pReceiveItemInfo) THEN//���pReceiveItemInfo����ֵ(Unassigned)����,�����pXmlPath
  begin
    if not FileExists(pXmlPath) then begin result:=false;exit;end;//����pXmlPath�ķ�ʽ,�����������ļ���ƭϵͳ
    //����Ŀֵ��������ʶ����ReceiveItemInfo��
    XMLDocument:=TXMLDocument.Create(pXmlPath);//����ļ�������,�˾䱨��
    ReceiveItemInfo:=VarArrayCreate([0,XMLDocument.DocumentElement.ChildNodes.Count-1],varVariant);
    for i :=0  to XMLDocument.DocumentElement.ChildNodes.Count-1 do
    begin
      ItemInfo:=XMLDocument.DocumentElement.ChildNodes[i];
      ReceiveItemInfo[i]:=VarArrayof([ItemInfo.ChildNodes['DltType'].Text,ItemInfo.ChildNodes['Value'].Text,ItemInfo.ChildNodes['Histogram'].Text,ItemInfo.ChildNodes['ImagePath'].Text]);
    end;
    //ȡ����������//�ݲ�����,�����뼰XML�ֶ�λ�ñ���,�Ա���ǰ���ݼ���λ��XML�ֶε���չ(ע��20110609)
    j:=XMLDocument.DocumentElement.AttributeNodes.count;
    for i :=0  to j-1 do
    begin
      ItemInfo:=XMLDocument.DocumentElement.AttributeNodes[i];
    end;
  end else
  begin
    ReceiveItemInfo:=pReceiveItemInfo;
  end;

  //��¼������־start
  IfRecLog:=pIsSure;
  if IfRecLog then
  begin
    LogStr:='������:'+pSpecNo+';����:'+pCheckDate+';��������:'+pSpecType+';����״̬:'+pSpecStatus+';���ȼ���:'+pDiagnosetype;
    WriteLog(pchar(LogStr));
    //��ʾ��������ReceiveItemInfo�ĵ�1ά���±߽��������ϱ߽�����
    for j :=VarArrayLowBound(ReceiveItemInfo,1) to VarArrayHighBound(ReceiveItemInfo,1) do
    begin
      LogStr:='������ʶ:'+VarToStr(ReceiveItemInfo[j][0])+';'+'���ֵ:'+VarToStr(ReceiveItemInfo[j][1])+';'+'ֱ��ͼ����:'+ReceiveItemInfo[j][2]+';'+'ͼƬ�ļ���·��:'+ReceiveItemInfo[j][3];
      WriteLog(pchar(LogStr));
    end;
    WriteLog('�������ݽ���');
  end;
  //��¼������־stop

  SpecNo:=pSpecNo;

  fs.DateSeparator:='-';
  fs.TimeSeparator:=':';
  fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
  CheckDate:=StrtoDateTimeDef(pCheckDate,ServerDateTime,fs);//edit by liuying 20111114 StrtoDateDef->StrtoDateTimeDef
  if  CheckDate<2 then ReplaceDate(CheckDate,ServerDateTime);//��ʾ1899-12-30,û�и����ڸ�ֵ
  if (HourOf(CheckDate)=0) and (MinuteOf(CheckDate)=0) and (SecondOf(CheckDate)=0) then ReplaceTime(CheckDate,ServerDateTime);//��ʾû�и�ʱ�丳ֵ

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
  if sDateOfBirth<>'' then//���ݳ�������������
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
  begin     //���浽��ֵ�ʿر���
      SaveDataToQuaContDB(0);
  end else
  if SpecNo=pQuaContSpecNoD then
  begin     //���浽��ֵ�ʿر���
      SaveDataToQuaContDB(-1);
  end else
  if SpecNo=pQuaContSpecNoG then
  begin    //���浽��ֵ�ʿر���
      SaveDataToQuaContDB(1);
  end else  //���浽���˱���
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
