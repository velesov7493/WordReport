unit WordReport;

{=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-\
|  WordReport - ��������� ��� ������������� �������� ������� � MS Word.        |                                                     |
|                                                                              |
|  ������: WordReport                                                          |
|  �����: ������ ��������� ��������� (greyfox84@list.ru)                       |
|  Copyright: 2012 ������ ��������� ���������                                  |
|  ����: 10.11.2012                                                            |
|                                                                              |
|  ��������:                                                                   |
|                                                                              |
|  ��� ������� ������ Build ��������� ������ TWordReport ������� ��� �����     |
|  ���������� ��� ������, ����� ������, ����� ���� ���������� � ������ ������  |
|  �� �������, ���������� � TemplateDocFileName. ����� ���������� ����������   |
|  OnReadMaket ��� �������� ������ � ���� ������. ����� ������ ������������    |
|  � ������� �����, ������� ����������� � ��������� � ������ ResultDocFileName |
|                                                                              |
\=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-}

interface

uses
  DB, ComObj, Classes;

const
  MaxFieldsOnBand = 16;
  MaxBandsOnReport = 8;

  bvVariable = 0;
  bvField = 1;
  bvFormat = 2;

type

  PDataSet = ^TDataSet;
  TBandVariables = array [0..2,1..MaxFieldsOnBand] of string;

  PDataBand = ^TDataBand;
  TDataBand = class
  private
    fName: string;
    fRowBookmark: string;                                       // ��� �������� ��� ������ ���������
    fVariableNames: TBandVariables;                             // ������������ ����� ������� ����� Dataset`� � ������� ����������
    fVariableCount: integer;                                    // ���������� ���������� � ����������� ������
    fCounterName: string;                                       // ��� �������� � ������ ���������
    fCounterVal: integer;                                       // �������� �������� ������
    fDataSet: PDataSet;                                         // ��������� �� DataSet
    fTableNumber: integer;                                      // ����� ������� ���������, ������� ����������� ������
    fKeyName: string;                                           // ��� ����������� ���� �� ��� ������ ������
    fKeyValue: integer;                                         // �������� ����������� ���� �� ��� ���� ������. ����� ������������ ��� ������� �����������
    fNextBand: PDataBand;                                       // ��������� �� ��������� ������ � ������
    function FindVariable(VarName:string): integer;
    function VarToPattern(VarIndex:integer):string;
    function getTableNumberByBookmark: integer;
    procedure Prepare;
    procedure PokeData;
  protected
    function GetDimensionByVarName(VarName: string; Index: integer): string;
    procedure SetDimensionByVarName(VarName: string; Index: integer; Value: string);
    property TableNumber: integer read fTableNumber write fTableNumber;
    property KeyName: string read fKeyName write fKeyName;
    property KeyValue: integer read fKeyValue write fKeyValue;
    property NextBand: PDataBand read fNextBand write fNextBand;
  public
    constructor Create(aBookmark,aName,aCounterName:string);
    procedure AssignDataSet(aDataSet: PDataSet);
    procedure SetField(VariableName,FieldName,Format:string);

    property Name: string read fName write fName;
    property Field[VarName:string]: string index 1 read GetDimensionByVarName write SetDimensionByVarName;
    property Format[VarName:string]: string index 2 read GetDimensionByVarName write SetDimensionByVarName;
  end;

  TWordReport = class(TComponent)
  private
    fDataBands: array [1..MaxBandsOnReport] of TDataBand;
    fBandCount: integer;
    fTemplateDocFileName: string;
    fResultDocFileName: string;
    fShowResult: boolean;
    fOnReadMaket: TNotifyEvent;
    procedure ReadVarNames;
    procedure ReadBandNames(Pattern:string; var Band,Counter,Variable: string);
    procedure NewBandVar(aBookmark,aBandName,aCounterName,aVariableName: string);
    procedure CreateMultiBandMaket(Band1,Band2:TDataBand);
    procedure StartProgress;
  protected
    fReportVars: TParams;
    function GetReportDifficulty: integer;
    function GetBandByName(BandName:string):TDataBand;
    function GetBandNumber(BandName:string):integer;
    procedure SetTemplateFileName(Value: string);
    procedure BindVariables;
    procedure ReadMaket;
    procedure SaveToFile(FileName: string);
    procedure CloseDocument;
  public
    procedure JoinBands(BandKeyField:string; BandName1:string; KeyValue1:integer; BandName2:string; KeyValue2:integer);
    procedure SetValue(VariableName:string; Value:Variant);
    procedure Build;
    procedure Quit;
    function BandExists(BandName: string): boolean;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; reintroduce; overload;
    property Bands[Name:string]: TDataBand read GetBandByName;
    property BandCount: integer read fBandCount;
  published
    property TemplateDocFileName: string read fTemplateDocFileName write SetTemplateFileName;
    property ResultDocFileName: string read fResultDocFileName write fResultDocFileName;
    property ShowResult: boolean read fShowResult write fShowResult default true;
    property OnReadMaket: TNotifyEvent read fOnReadMaket write fOnReadMaket;
  end;

procedure Register;

implementation

uses
  SysUtils, StrUtils, Windows, Messages, Variants, Graphics,
  Controls, Forms, Dialogs, ComCtrls;

type
  TProgressForm = class(TForm)
    ProgressBar1: TProgressBar;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure DoProgress(Step: integer);
  public
    { Public declarations }
  end;

const
  wdCollapseEnd = 0;
  wdCollapseStart = 1;
  wdFindContinue = 1;

var
  wProgress: TProgressForm;     // ���������� ����� ���������
  curWordApp: Variant;          // OLE-������ Word
  curDocument: Variant;         // OLE-������ ���������

// ������� ��� ��������� ���������� �� ��� �������

function PatternToName(ParamPattern:string):string;
var
  i1,i2: integer;
begin
  i1:=Pos('#(',ParamPattern)+2;
  i2:=Length(ParamPattern);
  Result:=MidStr(ParamPattern,i1,i2-i1);
end;

// �������� ������ �� �����

function NameToPattern(ParamName:string):string;
begin
  Result:='#('+ParamName+')';
end;

// ����������� OLE-������ ������ Word
// ����:
//   FindText - ��� ������
//   RegExp - ������� FindText ���������� ���������� Word
// �����:
//   FindObj - �������������� ������ ������

procedure PrepareFindObject(FindText:string; RegExp:boolean; var FindObj:Variant);
begin
  FindObj:=curWordApp.Selection.Find;
  FindObj.ClearFormatting;
  FindObj.Text:= FindText;
  FindObj.Replacement.Text:='';
  FindObj.Forward:=true;
  FindObj.Wrap:=wdFindContinue;
  FindObj.Format:=false;
  FindObj.MatchCase:=false;
  FindObj.MatchWholeWord:=false;
  FindObj.MatchAllWordForms:=false;
  FindObj.MatchSoundsLike:=false;
  FindObj.MatchWildcards:=RegExp;
end;

//==============================================================================
// TProgress Form
//==============================================================================

{$R wrtprogress.dfm}

procedure TProgressForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:=caFree;
end;

procedure TProgressForm.DoProgress(Step: integer);
begin
  with wProgress.ProgressBar1 do Position:=Position+Step;
end;

//==============================================================================
// TDataBand
//==============================================================================

constructor TDataBand.Create(aBookmark:string; aName: string; aCounterName: string);
begin
  Name:=aName;
  fCounterName:=aCounterName;
  fRowBookmark:=aBookmark;
  fVariableCount:=0;
  fCounterVal:=1;
  fDataSet:=nil;
end;

// ����� ���������� ������, �.�. �������� �� ������
// ���������� -1, ���� ����� ���������� ���

function TDataBand.FindVariable(VarName: string):integer;
var
  i,r: integer;
begin
  r:=-1;
  for i:=1 to fVariableCount do
    if fVariableNames[bvVariable,i]=VarName then begin
      r:=i;
      Break;
    end;
  Result:=r;
end;

// �������� ����� ������� ���������, � ������� ���������
// �������� �������� ������

function TDataBand.getTableNumberByBookmark: integer;
var
  i,tableStart,tableNumber: integer;
  r,table: Variant;
begin
  tableNumber:=0;
  if not curDocument.Bookmarks.Exists(fRowBookmark) then begin
    Result:=tableNumber;
    Exit;
  end;
  r:=curDocument.Bookmarks.Item(fRowBookmark).Range;
  if r.Tables.Count>0 then begin
    table:=r.Tables.Item(1);
    if table.NestingLevel<>1 then begin
      Result:=tableNumber;
      Exit;
    end;
    tableStart:=table.Range.Start;
    for i:=1 to curDocument.Tables.Count do
      if curDocument.Tables.Item(i).Range.Start=tableStart then begin
        tableNumber:=i;
        Break;
      end;
  end else raise Exception.Create('�������� �� �������� ������ �������');
  Result:=tableNumber;
end;

// �������� �������� ���������� �� ����� ���������� � ������� ����� ���������

function TDataBand.GetDimensionByVarName(VarName: string; Index: integer): string;
var
  VarIndex: integer;
begin
  VarIndex:=FindVariable(VarName);
  if index>0 then Result:=fVariableNames[Index,VarIndex] else Result:='';
end;

// ���������� �������� ���������� �� ����� ���������� � ������� ����� ���������

procedure TDataBand.SetDimensionByVarName(VarName: string; Index: integer; Value: string);
var
  VarIndex: integer;
begin
  VarIndex:=FindVariable(VarName);
  if index<0 then Exit;
  fVariableNames[Index,VarIndex]:=Value;
end;

// �������� ������ ���������� ������ �� �������
// ������=0 � �������� ������� ������

function TDataBand.VarToPattern(VarIndex:integer):string;
var r: string;
begin
  if VarIndex=0 then r:='#('+fCounterName+')' else
    r:='#('+Name+'('+fCounterName+').'+fVariableNames[bvVariable,VarIndex]+')';
  Result:=r;
end;

// ��������� ���� ������ ������ � ��� ������
// � ���������� VariableName

procedure TDataBand.SetField(VariableName,FieldName,Format:string);
var
  VarIndex: integer;
begin
  VarIndex:=FindVariable(VariableName);
  if VarIndex<0 then begin
    raise Exception.Create('���������� "'+VariableName+'" ����������� � ������ "'+fName+'" ���������');
    Exit;
  end;
  fVariableNames[bvField,VarIndex]:=FieldName;
  fVariableNames[bvFormat,VarIndex]:=Format;
end;

// "���������" ������
// ���������� ��� ������ ������ � ��������

var
  // ������� �������� ����������� ����
  OldKeyValue: integer;

procedure TDataBand.PokeData;
label
  ValueDefinedS,ValueDefinedM;
var
  Pattern,Value: string;
  cField: TField;
  FindObj: Variant;
  i,j,c: integer;
begin
  if fDataSet=nil then Exit;
  if Length(fKeyName)=0 then begin
    // ���� ������ �� ������� � ������, �� ������������ �� �����������
    fDataSet.First;
    for i:=1 to fDataset.RecordCount do begin
      Pattern:=VarToPattern(0);
      PrepareFindObject(Pattern,false,FindObj);
      if FindObj.Execute then begin
        Value:=IntToStr(fDataset.RecNo)+'.';
        curWordApp.Selection.Text:=Value;
      end;
      for j:=1 to fVariableCount do begin
        Pattern:=VarToPattern(j);
        PrepareFindObject(Pattern,false,FindObj);
        if FindObj.Execute then begin
          cField:=fDataset.FieldByName(fVariableNames[bvField,j]);
          if cField.IsNull then begin
            Value:='';
            goto ValueDefinedS;
          end;
          if Length(fVariableNames[bvFormat,j])<>0 then begin
            Value:=Trim(SysUtils.Format(fVariableNames[bvFormat,j],[cField.AsFloat]));
            goto ValueDefinedS;
          end;
          Value:=cField.AsString;
ValueDefinedS:
          curWordApp.Selection.Text:=Value;
        end;
      end;
      fDataSet.Next;
      wProgress.DoProgress(1);
    end;
  end else begin
    // ����� ������������ ������ ������ � ����� ��������� ����������� ����
    c:=fDataset.FieldByName(fKeyName).AsInteger;
    // ���� ������� ����������� ������ ��������, �� �������� �������
    if fKeyValue>OldKeyValue then fCounterVal:=1;
    while (c=fKeyValue) and (not fDataset.Eof) do begin
      Pattern:=VarToPattern(0);
      PrepareFindObject(Pattern,false,FindObj);
      if FindObj.Execute then begin
        Value:=IntToStr(fCounterVal)+'.';
        curWordApp.Selection.Text:=Value;
      end;
      for j:=1 to fVariableCount do begin
        Pattern:=VarToPattern(j);
        PrepareFindObject(Pattern,false,FindObj);
        if FindObj.Execute then begin
          cField:=fDataset.FieldByName(fVariableNames[bvField,j]);
          if cField.IsNull then begin
            Value:='';
            goto ValueDefinedM;
          end;
          if Length(fVariableNames[bvFormat,j])<>0 then begin
            Value:=Trim(SysUtils.Format(fVariableNames[bvFormat,j],[cField.AsFloat]));
            goto ValueDefinedM;
          end;
          Value:=cField.AsString;
ValueDefinedM:
          curWordApp.Selection.Text:=Value;
        end;
      end;
      fDataset.Next;
      wProgress.DoProgress(1);
      OldKeyValue:=c;
      c:=fDataset.FieldByName(fKeyName).AsInteger;
      Inc(fCounterVal);
    end;
  end;
end;

// ����������� ������ ��� ������ �������
// ���������� ������, ����� ��� ����������� ������� ���, ������� ������� � ��

procedure TDataBand.Prepare;
begin
  if fDataSet=nil then Exit;
  if curDocument.Bookmarks.Exists(fRowBookmark) then begin
    curDocument.Bookmarks.Item(fRowBookmark).Range.Copy;
    fDataset.First;
    while not fDataset.Eof do begin
      if fTableNumber<>0 then begin
        curDocument.Tables.Item(fTableNumber).Select;
        curWordApp.Selection.Collapse(wdCollapseEnd);
      end else
        curDocument.Bookmarks.Item(fRowBookmark).Range.Select;
      curWordApp.Selection.Paste;
      fDataset.Next;
    end;
    fDataset.First;
    curDocument.Bookmarks.Item(fRowBookmark).Range.Select;
    curWordApp.Selection.Cut;
  end;
end;

// ���������� �� � ������

procedure TDataBand.AssignDataSet(aDataSet: PDataSet);
var
  tn: integer;
begin
  if not (aDataset^ is TDataset) then begin
    raise Exception.Create('������������ ��������� �� TDataset');
    Exit;
  end;
  fDataSet:=aDataSet;
  tn:=getTableNumberByBookmark;
  if tn>0 then fTableNumber:=tn;
end;

//==============================================================================
// TWordReport
//==============================================================================

constructor TWordReport.Create(AOwner:TComponent);
begin
  inherited Create(AOwner);
  fReportVars:=TParams.Create;
  fShowResult:=true;
end;

// ������� �������� ��������

procedure TWordReport.CloseDocument;
var
  i: integer;
begin
  for i:=1 to fBandCount do fDataBands[i].Free;
  fBandCount:=0;
  fReportVars.Clear;
  curDocument.Close;
  wProgress.DoProgress(10);
end;

// ��������� Word, ���� �� ������� ��� �������� - ���� �����.
// ������� � ���������� ���� ���������

procedure TWordReport.Quit;
begin
  if fShowResult then begin
    curWordApp.Selection.Collapse(wdCollapseEnd);
    curWordApp.Visible:=true;
  end else
    curWordApp.Quit;
  wProgress.DoProgress(40);
  wProgress.Close;
end;

destructor TWordReport.Destroy;
begin
  if Assigned(fReportVars) then fReportVars.Free;
  inherited Destroy;
end;

// ������ ������, ��������� � TemplateDocFileName � MS Word

procedure TWordReport.ReadMaket;
var
  Error,verS: string;
  version: real;
  cpos: integer;
begin
  Error:='';
  try
    curWordApp:=CreateOleObject('Word.Application');
    verS:=curWordApp.Version;
  except
    Error:='MS Word ����������';
  end;
  Val(verS,version,cpos);
  if (cpos=0) and (version>=9) then begin
    fBandCount:=0;
    curWordApp.Visible:=false;
  end else Error:=Error+#10#13+'��������� MS Word 2000 � ����';
  try
    curDocument:=curWordApp.Documents.Add(fTemplateDocFileName);
  except
    Error:='�� ������� ������� ������: "'+fTemplateDocFileName+'"'+#10#13#10#13+Error;
  end;
  if Length(Error)=0 then ReadVarNames else raise Exception.Create(Error);
end;

// ���������� �������� �������

procedure TWordReport.SetTemplateFileName(Value: string);
begin
  if fTemplateDocFileName<>Value then fTemplateDocFileName:=Value;
end;

// ���������� ��������� ������
// = ������������������������
// +�������������������������
// +10 �� �������� ���������
// +10 �� ���������� ���������
// +40 �� ���������� MS Word

function TWordReport.GetReportDifficulty: integer;
var
  i,s: integer;
  SkipBands: set of Byte;
begin
  s:=fReportVars.Count;
  SkipBands:=[];
  for i:=1 to fBandCount do begin
    if i in SkipBands then Continue;
    s:=s+fDataBands[i].fDataSet^.RecordCount;
    if fDataBands[i].NextBand<>nil then
      SkipBands:=SkipBands+[GetBandNumber(fDataBands[i].fNextBand^.fName)];
  end;
  s:=s+60;
  Result:=s;
end;

// ��������� ����� ������, �������� � ����������
// �� ���������� ���������� ������ � �������

procedure TWordReport.ReadBandNames(Pattern:string; var Band,Counter,Variable: string);
var
  i1,i2,c: integer;
  s: string;
begin
  s:=PatternToName(Pattern);
  i1:=1; i2:=Pos('(',s); c:=i2-i1;
  Band:=MidStr(s,i1,c);
  i1:=i2+1; i2:=Pos(')',s); c:=i2-i1;
  Counter:=MidStr(s,i1,c);
  i1:=i2+2; i2:=Length(s); c:=i2-i1+1;
  Variable:=MidStr(s,i1,c);
end;

// ������� ����� ���������� ������

procedure TWordReport.NewBandVar(aBookmark: string; aBandName: string; aCounterName: string; aVariableName: string);
var
  cBand: TDataBand;
  cParam: TParam;
begin
  cBand:=Bands[aBandName];
  if cBand<>nil then begin
    if cBand.FindVariable(aVariableName)<0 then with cBand do begin
      Inc(fVariableCount);
      fVariableNames[bvVariable,fVariableCount]:=aVariableName;
    end;
  end else begin
    Inc(fBandCount);
    fDataBands[fBandCount]:=TDataBand.Create(aBookmark,aBandName,aCounterName);
    with fDataBands[fBandCount] do begin
      Inc(fVariableCount);
      fVariableNames[bvVariable,fVariableCount]:=aVariableName;
      cParam:=fReportVars.FindParam(aCounterName);
      if cParam<>nil then fReportVars.RemoveParam(cParam);
    end;
  end;
end;

// ���������� �������� ���������� ��� ������

procedure TWordReport.SetValue(VariableName:string; Value:Variant);
begin
  if fReportVars.FindParam(VariableName)<>nil then
    fReportVars.ParamByName(VariableName).Value:=Value
  else raise Exception.Create('���������� "'+VariableName+'" ����������� � ���������');
end;

// �������� ������ �� �� �����

function TWordReport.GetBandByName(BandName:string): TDataBand;
var
  i: integer;
begin
  Result:=nil;
  for i:=1 to fBandCount do if fDataBands[i].Name=BandName then begin
    Result:=fDataBands[i];
    Break;
  end;
end;

// �������� ����� ������ �� �� �����

function TWordReport.GetBandNumber(BandName: string): integer;
var
  i:integer;
begin
  Result:=-1;
  for i:=1 to fBandCount do if fDataBands[i].Name=BandName then begin
    Result:=i;
    Break;
  end;
end;

// ������ ������ ��������� ������

procedure TWordReport.StartProgress;
begin
  with wProgress.ProgressBar1 do begin
    Min:=0;
    Max:=GetReportDifficulty;
    Position:=Min;
  end;
  wProgress.Show;
end;

// ��������� ����� �� �������

procedure TWordReport.ReadVarNames;
var
  ParamName: string;
  cParam: TParam;
  cBmName,cBandName,cCounterName,cVarName: string;
  FindObj: Variant;
begin
  PrepareFindObject('\#[\(]{1}[A-z.0-9]@[\)]{1}',true,FindObj);
  while FindObj.Execute do begin
    ParamName:=PatternToName(curWordApp.Selection.Text);
    if fReportVars.FindParam(ParamName)=nil then begin
      cParam:=fReportVars.CreateParam(ftString,ParamName,ptUnknown);
      fReportVars.AddParam(cParam);
    end;
  end;

  curDocument.Range.Select;
  curWordApp.Selection.Collapse(wdCollapseStart);

  PrepareFindObject('\#[\(]{1}[A-z0-9]@[\(]{1}[A-z0-9]@[\)]{1}.[A-z0-9]@[)]{1}',true,FindObj);
  while FindObj.Execute do begin
    if curWordApp.Selection.BookmarkID>0 then begin
      cBmName:=curDocument.Bookmarks.Item(curWordApp.Selection.BookmarkID).Name;
      if (LowerCase(LeftStr(cBmName,4))<>'data') or (Pos(cBmName[5],'12345678')=0) then
        raise Exception.Create('��� �������� ������ �� ������������� ������������');
    end else
      raise Exception.Create('���������� "'+curWordApp.Selection.Text+'" ��������� ��� ��������');
    ReadBandNames(curWordApp.Selection.Text,cBandName,cCounterName,cVarName);
    NewBandVar(cBmName,cBandName,cCounterName,cVarName);
  end;
end;

// ��������� ������, ����������� �������������, � ��������

procedure TWordReport.BindVariables;
var
  FindObj: Variant;
  ParamPattern: string;
  i: integer;
  SkipBands: set of Byte;
begin
  StartProgress;
  for i:=0 to fReportVars.Count-1 do begin
    ParamPattern:=NameToPattern(fReportVars.Items[i].Name);
    PrepareFindObject(ParamPattern,false,FindObj);
    while FindObj.Execute do
      curWordApp.Selection.Text:=fReportVars.Items[i].AsString;
    wProgress.DoProgress(1);
  end;

  SkipBands:=[];
  for i:=1 to fBandCount do
    if (fDataBands[i].fNextBand<>nil) then begin
      // ���� ������ � ������, �� �������� �� ������ �� ��������� �������
      if (i in SkipBands) then Continue;
      CreateMultiBandMaket(fDataBands[i],fDataBands[i].fNextBand^);
      while not fDataBands[i].fDataSet.Eof do begin
        fDataBands[i].PokeData;
        fDataBands[i].fNextBand^.PokeData;
      end;
      SkipBands:=SkipBands+[GetBandNumber(fDataBands[i].fNextBand^.fName)];
    end else
      // �����, ������ ����������� ...
      fDataBands[i].Prepare;

  for i:=1 to fBandCount do
    // ... � ������� �� �����������
    if fDataBands[i].fNextBand=nil then fDataBands[i].PokeData;
end;

// ��������� �������� � ����

procedure TWordReport.SaveToFile(FileName: string);
begin
  curDocument.SaveAs(FileName);
  wProgress.DoProgress(10);
end;

// ���������� ��� ������ � ������
// ��� ����� ��� ��������� �������� ������ �� ������ � ���� �� ��.

// ��������: ���������� ������ � ������� 'R' � 'A', ����������� ���� 'GRID'
// � �������� �����. ��� ����: ���� GRID=0 - ������� ������ � ������ R, � ����
// GRID=1 - � ������ A.
// JoinBands('GRID','R',0,'A',1);

procedure TWordReport.JoinBands(BandKeyField:string; BandName1:string; KeyValue1:integer; BandName2:string; KeyValue2:integer);
var
  i,j: integer;
begin
  i:=GetBandNumber(BandName1);
  j:=GetBandNumber(BandName2);
  if i<0 then raise Exception.Create('������ "'+BandName1+'" �� ������� � ��������� �������');
  if j<0 then raise Exception.Create('������ "'+BandName2+'" �� ������� � ��������� �������');
  if (i<0) or (j<0) then Exit;
  fDataBands[i].fNextBand:=@fDataBands[j];
  fDataBands[j].fNextBand:=@fDataBands[i];
  fDataBands[i].fKeyName:=BandKeyField;
  fDataBands[j].fKeyName:=BandKeyField;
  fDataBands[i].fKeyValue:=KeyValue1;
  fDataBands[j].fKeyValue:=KeyValue2;
end;

// ���������� �� � ������� ������ � ������ BandName

function TWordReport.BandExists(BandName: string): boolean;
begin
  Result:=not (GetBandNumber(BandName)<0);
end;

// ����������� ����� ������ ������ � ��������

procedure TWordReport.CreateMultiBandMaket(Band1,Band2: TDataBand);
var
  cDataset: TDataset;
  cKey,tn: integer;
begin
  if Band1.fDataSet<>Band2.fDataSet then Exit;
  cDataset:=Band1.fDataSet^;
  tn:=Band1.fTableNumber;
  cDataset.First;
  while not cDataset.Eof do begin
    curDocument.Tables.Item(tn).Select;
    curWordApp.Selection.Collapse(wdCollapseEnd);
    cKey:=cDataset.FieldByName(Band1.fKeyName).AsInteger;
    if cKey=Band1.fKeyValue then
      curDocument.Bookmarks.Item(Band1.fRowBookmark).Range.Copy;
    if cKey=Band2.fKeyValue then
      curDocument.Bookmarks.Item(Band2.fRowBookmark).Range.Copy;
    curWordApp.Selection.Paste;
    cDataset.Next;
  end;
  curDocument.Bookmarks.Item(Band1.fRowBookmark).Range.Select;
  curWordApp.Selection.Cut;
  curDocument.Bookmarks.Item(Band2.fRowBookmark).Range.Select;
  curWordApp.Selection.Cut;
  cDataset.First;
end;

// ������� �����
// ��������� ��� ������������������ ��������
// (����� �������� ������, ������� ������ ����������� ������������ � ����������� OnReadMaket)
// ��� ������������ ������ ������ �� �������
// TemplateDocFileName � ������� �������� ResultDocFileName

procedure TWordReport.Build;
begin
  ReadMaket;
  wProgress:=TProgressForm.Create(Owner);
  if Assigned(fOnReadMaket) then OnReadMaket(Self);
  BindVariables;
  SaveToFile(fResultDocFileName);
  if not fShowResult then CloseDocument;
end;

// ����������� ���������� � ������� Delphi

procedure Register;
begin
  RegisterComponents('WordReport',[TWordReport]);
end;

end.
