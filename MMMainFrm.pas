unit MMMainFrm;
{------------------------------------------------------------------------------}
(*
  Mail merging application supporting multiline field values.
  Fields are any strings inside {}.
  Supported fields:
  {NAME}, {ADDRESS1}, {ADDRESS2}, {CITY}, {ZIP}, {COUNTRY}, {FULL_ADDRESS}.
  {FULL_ADDRESS} is multiline.
  This demo generates report for all records as a single document (separated by
  page breaks).
  v1.2
*)
{------------------------------------------------------------------------------}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, RVScroll, RichView, StdCtrls, RVStyle, RVTable, CRVData, RVItem,
  RVEdit, RVUni, RVTypes;

{$I RV_Defs.inc}

type

  TPerson = class (TCollectionItem)
    public
      Name, Address1, Address2, Zip, City, Country: String;
  end;


  TForm1 = class(TForm)
    rvOutput: TRichView;
    rvsOutput: TRVStyle;
    rvsTemplate: TRVStyle;
    Label1: TLabel;
    Label2: TLabel;
    rve: TRichViewEdit;
    Button1: TButton;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    Persons: TCollection;
    Index: Integer;
    procedure LoadPersons;
    procedure CreateTemplate;
    function GetFieldValueFromDatabase(const FieldName: String): String;
    procedure FillFields(RVData: TCustomRVData; StartItemNo: Integer);
    procedure InsertMultilineText(RVData: TCustomRVData;
      var ItemNo, Pos: Integer; const Value: String);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

{------------------------------------------------------------------------------}
{ This procedure loads file CUSTOMERS.TXT into Persons.
  In real applications data can be stored in some database }
procedure TForm1.LoadPersons;
var F: TextFile;
    i, Count: Integer;
    Person: TPerson;
begin
  AssignFile(F, ExtractFilePath(Application.ExeName)+'Persons.txt');
  Reset(F);
  try
    Readln(F, Count);
    for i := 0 to Count-1 do
    begin
      Person := Persons.Add as TPerson;
      Readln(F, Person.Name);
      Readln(F, Person.Address1);
      Readln(F, Person.Address2);
      Readln(F, Person.City);
      Readln(F, Person.Zip);
      Readln(F, Person.Country);
    end;
  finally
    CloseFile(F);
  end;
  Label3.Caption := IntToStr(Persons.Count) + ' personal records in the database';
end;
{------------------------------------------------------------------------------}
{ Creating a sample template. You can edit it }
procedure TForm1.CreateTemplate;
begin
   rve.Clear;
   rve.AddNL('To {NAME}', 0, 0);
   rve.AddNL('{FULL_ADDRESS}', 0, 0);
   rve.AddNL('Letter Sample', 1, 1);
   rve.AddNL('Hello {NAME}', 0, 0);
   rve.AddNL('', 0, 0);
   rve.AddNL('You are fired.', 0, 0);
   rve.AddNL('', 0, 0);
   rve.AddNL('With best regards,', 0, 0);
   rve.AddNL('your mailing robot.', 0, 0);
   rve.AddNL('', 0, 0);
   rve.AddNL('', 0, 0);
   rve.AddNL('', 0, 0);
   rve.Format;
end;
{------------------------------------------------------------------------------}
{
  This function returns a current field value by the field name.
}
function TForm1.GetFieldValueFromDatabase(const FieldName: String): String;
var Person: TPerson;
begin
  Person := Persons.Items[Index] as TPerson;
  if FieldName='NAME' then
    Result := Person.Name
  else if FieldName='ADDRESS1' then
    Result := Person.Address1
  else if FieldName='ADDRESS2' then
    Result := Person.Address2
  else if FieldName='CITY' then
    Result := Person.City
  else if FieldName='ZIP' then
    Result := Person.ZIP
  else if FieldName='COUNTRY' then
    Result := Person.Country
  else if FieldName='FULL_ADDRESS' then
    Result := Person.Address1+#13+Person.Address2+#13+Person.City+' '+Person.ZIP+#13+Person.Country
  else
    Result := '{error: illegal field code}';
end;
{------------------------------------------------------------------------------}
procedure TForm1.FormCreate(Sender: TObject);
begin
  Persons := TCollection.Create(TPerson);
  LoadPersons;
  CreateTemplate;
end;
{------------------------------------------------------------------------------}
procedure TForm1.FormDestroy(Sender: TObject);
begin
  Persons.Free;
  Persons := nil;
end;
{------------------------------------------------------------------------------}
(*
  This function iterates through all items in RVData starting from StartItemNo,
  searches for strings inside {}, calls GetFieldValueFromDatabase()
  and replaces text between {} (including {}) with the returned value.
  You can move this function to your application unchanged.
*)
procedure TForm1.FillFields(RVData: TCustomRVData; StartItemNo: Integer);
var i,j,r,c: Integer;
    table: TRVTableItemInfo;
    FieldName, FieldValue: String;
    s: String;
    FieldStart: Integer;
begin
  i := StartItemNo;
  while i<RVData.ItemCount do begin
    if RVData.GetItemStyle(i)=rvsTable then begin // recursive processing of tables
      table := TRVTableItemInfo(RVData.GetItem(i));
      for r := 0 to table.Rows.Count-1 do
        for c := 0 to table.Rows[r].Count-1 do
          if table.Cells[r,c]<>nil then
            FillFields(table.Cells[r,c].GetRVData, 0);
      end
    else if RVData.GetItemStyle(i)>=0 then begin
      s := RVData.GetItemText(i);
      FieldStart := 0;
      j := 1;
      while j<=Length(s) do begin
        if s[j]='{' then // may be this is a beginning of field code
          FieldStart := j
        else if (s[j]='}') and (FieldStart>0) then begin // field code is found
          FieldName :=  Copy(s, FieldStart+1, j-FieldStart-1);
          FieldValue := GetFieldValueFromDatabase(FieldName);
          // deleting field code
          Delete(s, FieldStart, j-FieldStart+1);
          RVData.SetItemText(i, s);
          // inserting field value
          j := FieldStart;
          InsertMultilineText(RVData, i, j, FieldValue);
          s := RVData.GetItemText(i);
          FieldStart := 0;
          continue;
        end;
        inc(j);
      end;
    end;
    inc(i);
  end;
end;
{------------------------------------------------------------------------------}
{ Generating output }
procedure TForm1.Button1Click(Sender: TObject);
var Stream: TStream;
    ItemCount: Integer;
begin
  Index := 0; // index of the current person 
  ItemCount := 0;
  rvOutput.Clear;
  while Index<Persons.Count do begin
    // copying rve to the end of rvOutput
    Stream := TMemoryStream.Create;
    rve.SaveRVFToStream(Stream, False);
    Stream.Position := 0;
    rvOutput.InsertRVFFromStream(Stream, ItemCount);
    Stream.Free;
    // processing the last added copy of template
    if rvOutput.ItemCount>ItemCount then begin
      // starting from new page
      if ItemCount>0 then
        rvOutput.PageBreaksBeforeItems[ItemCount] := True;
      // replacing field codes
      FillFields(rvOutput.RVData, ItemCount);
    end;
    ItemCount := rvOutput.ItemCount;
    inc(Index);
  end;
  rvOutput.Format;
end;
{------------------------------------------------------------------------------}
// Converting text from String to internal representation
function ConvertStringToItemText(const Text: String;
  UnicodeItem: Boolean; CodePage: Cardinal): TRVRawByteString;
begin
  {$IFDEF RVUNICODESTR} // <-- declared in RV_Defs.inc
  if UnicodeItem then
    Result := RVU_GetRawUnicode(Text)
  else
    Result := TRVAnsiString(Text);
  {$ELSE}
  if UnicodeItem then
    Result := RVU_AnsiToUnicode(CodePage, Text)
  else
    Result := Text;
  {$ENDIF}
end;
{------------------------------------------------------------------------------}
{
  Inserts multiline text Value in the position (RVData, ItemNo, Pos).
  After inserting, (RVData, ItemNo, Pos) is a position after the inserted text.
  This method uses some undocumented RichView functions. Just copy it to your
  application.
}
procedure TForm1.InsertMultilineText(RVData: TCustomRVData;
  var ItemNo, Pos: Integer; const Value: String);
var item: TRVTextItemInfo;
    s1, s2: String;
    sr: TRVRawByteString;
    sl: TStringList;
    i, StyleNo, ParaNo: Integer;
    Unicode: Boolean;
    CodePage: TRVCodePage;
begin
  if Value='' then
    exit;
  s1 := RVData.GetItemText(ItemNo);
  s2 := Copy(s1, Pos, Length(s1));
  s1 := Copy(s1, 1, Pos-1);
  sl := TStringList.Create;
  sl.Text := Value;
  { workaround - TStringList ignores the last cr lf }
  if (Length(Value)>1) and
     (Value[Length(Value)-1]=#13) and (Value[Length(Value)]=#10) then
    sl.Add('');
  { end of workaround }
  if sl.Count=1 then begin
    RVData.SetItemText(ItemNo, s1+Value+s2);
    sl.Free;
    inc(Pos, Length(Value)-1);
    exit;
  end;
  RVData.SetItemText(ItemNo, s1+sl[0]);
  StyleNo := RVData.GetItemStyle(ItemNo);
  ParaNo := RVData.GetItemPara(ItemNo);
  Unicode := rvioUnicode in RVData.GetItem(ItemNo).ItemOptions;
  CodePage := RVData.GetItemCodePage(ItemNo);
  for i := 1 to sl.Count-1 do begin
    item := RichViewTextItemClass.Create(RVData);
    s1 := sl[i];
    if i=sl.Count-1 then
      s1 := s1+s2;
    item.StyleNo := StyleNo;
    item.ParaNo := ParaNo;
    if Unicode then
      item.ItemOptions := item.ItemOptions+[rvioUnicode];
    sr := ConvertStringToItemText(s1, Unicode, CodePage);
    item.Inserting(RVData, sr, False);
    RVData.Items.InsertObject(ItemNo+i, sr, item);
    item.Inserted(RVData, ItemNo+i);
  end;
  ItemNo := ItemNo+sl.Count-1;
  Pos := Length(sl[sl.Count-1])+1;
  sl.Free;
end;

end.
