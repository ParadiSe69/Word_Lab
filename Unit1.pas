unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Mask, Vcl.ExtCtrls,
  Vcl.ComCtrls,VBIDE_TLB,ComObj;

type
  TForm1 = class(TForm)
    LabeledEdit2: TLabeledEdit;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit3: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    Label3: TLabel;
    Edit1: TEdit;
    LabeledEdit6: TLabeledEdit;
    DateTimePicker1: TDateTimePicker;
    Label4: TLabel;
    LabeledEdit7: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    LabeledEdit8: TLabeledEdit;
    LabeledEdit9: TLabeledEdit;
    LabeledEdit10: TLabeledEdit;
    LabeledEdit11: TLabeledEdit;
    Label2: TLabel;
    LabeledEdit12: TLabeledEdit;
    LabeledEdit13: TLabeledEdit;
    LabeledEdit14: TLabeledEdit;
    LabeledEdit15: TLabeledEdit;
    LabeledEdit16: TLabeledEdit;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    Label5: TLabel;
    LabeledEdit17: TLabeledEdit;
    DateTimePicker3: TDateTimePicker;
    Label6: TLabel;
    DateTimePicker4: TDateTimePicker;
    Label7: TLabel;
    Label8: TLabel;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit18: TLabeledEdit;
    procedure Button1Click(Sender: TObject);
    procedure ReplaceField(const ADocument: OleVariant);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  WordApp, Document: Variant;
implementation

{$R *.dfm}
procedure TForm1.Button2Click(Sender: TObject);
begin
  WordApp.ActiveDocument.PrintPreview;
end;


procedure TForm1.ReplaceField(const ADocument: OleVariant);
var
  i: Integer;
  BookmarkName: string;
  Range: OleVariant;

  function CompareBm(ABmName: string; const AName: string): Boolean;
  var
    i: Integer;
  begin
    i := Pos('__', ABmName);
    if i > 0 then
      Delete(ABmName, i, Length(ABmName) - i + 1);

    Result := SameText(ABmName, AName);
  end;

begin
  for i := 23 downto 1 do
  begin
    BookmarkName := ADocument.Bookmarks.Item(i).Name;
    Range := ADocument.Bookmarks.Item(i).Range;

    if CompareBm(BookmarkName, 'chislo') then
      Range.Text := DateToStr(DateTimePicker4.Date)
    else
    if CompareBm(BookmarkName, 'FIO') then
      Range.Text := LabeledEdit2.Text
    else
    if CompareBm(BookmarkName, 'ceria') then
      Range.Text := LabeledEdit4.Text
    else
    if CompareBm(BookmarkName, 'nomer') then
      Range.Text := LabeledEdit8.Text
    else
    if CompareBm(BookmarkName, 'kemdata') then
      Range.Text := LabeledEdit3.Text
    else
    if CompareBm(BookmarkName, 'markamodel') then
      Range.Text := LabeledEdit5.Text
    else
    if CompareBm(BookmarkName, 'regnomer') then
      Range.Text := Edit1.Text
    else
    if CompareBm(BookmarkName, 'VIN') then
      Range.Text := LabeledEdit7.Text
    else
    if CompareBm(BookmarkName, 'godmashin') then
      Range.Text := DateToStr(DateTimePicker1.Date)
    else
    if CompareBm(BookmarkName, 'Ndvig') then
      Range.Text := LabeledEdit6.Text
    else
    if CompareBm(BookmarkName, 'cvet') then
      Range.Text := LabeledEdit10.Text
    else
    if CompareBm(BookmarkName, 'TCceria') then
      Range.Text := LabeledEdit11.Text
    else
    if CompareBm(BookmarkName, 'TCnomer') then
      Range.Text := LabeledEdit12.Text
    else
    if CompareBm(BookmarkName, 'vydano') then
      Range.Text := LabeledEdit13.Text
    else
    if CompareBm(BookmarkName, 'ychet1') then
      Range.Text := LabeledEdit14.Text
    else
    if CompareBm(BookmarkName, 'FIOdov') then
      Range.Text := LabeledEdit15.Text
    else
    if CompareBm(BookmarkName, 'mectodov') then
      Range.Text := LabeledEdit16.Text
    else
    if CompareBm(BookmarkName, 'crokna') then
      Range.Text := DateToStr(DateTimePicker2.Date)
    else
    if CompareBm(BookmarkName, 'mestovlad') then
      Range.Text := LabeledEdit17.Text
    else
    if CompareBm(BookmarkName, 'regceria') then
      Range.Text := LabeledEdit1.Text
    else
    if CompareBm(BookmarkName, 'NNkyzova') then
      Range.Text := LabeledEdit18.Text
    else
    if CompareBm(BookmarkName, 'Nkyzova') then
      Range.Text := LabeledEdit9.Text
    else
    if CompareBm(BookmarkName, 'datavydachi') then
      Range.Text := DateToStr(DateTimePicker3.Date);
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  TempleateFileName: string;
begin
   TempleateFileName := GetCurrentDir + '\temp.dot';
   WordApp := CreateOleObject('Word.Application');
   Document := WordApp.Documents.Add(TempleateFileName, False);

   // Заменяем закладки на данные
    ReplaceField(Document);

    // По умолчание окно Word скрыто, делаем его видимым с готовым отчетом
    WordApp.Visible := True;

end;

end.
