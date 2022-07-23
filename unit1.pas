unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Menus,
  StdCtrls, ColorBox, Spin, fpspreadsheetctrls, fpspreadsheetgrid,
  ColorSpeedButton, fpspreadsheet, LazUTF8, TAGraph, TASeries, TATools, TATypes,
  TASources, GraphUtil, fpsallformats, Grids, ComCtrls, Buttons, Clipbrd;

type

  { TForm1 }

  TForm1 = class(TForm)
    BtnAbrir: TColorSpeedButton;
    BtnAcoesG: TColorSpeedButton;
    BtnAum3: TColorSpeedButton;
    BtnHist1: TColorSpeedButton;
    BtnHist2: TColorSpeedButton;
    BtnExec: TColorSpeedButton;
    BtnNovo: TColorSpeedButton;
    BtnAum2: TColorSpeedButton;
    BtnSalvar: TColorSpeedButton;
    BtnCopiar: TColorSpeedButton;
    BtnSalvarRes: TColorSpeedButton;
    BtnSalvar2: TColorSpeedButton;
    BtnAcoesR: TColorSpeedButton;
    BtnAcoes: TColorSpeedButton;
    BtnAum1: TColorSpeedButton;
    BtnSelectAll: TColorSpeedButton;
    CB2: TComboBox;
    CB3: TComboBox;
    CB4: TComboBox;
    CBLabel: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    ColorBox1: TColorBox;
    ColorBox2: TColorBox;
    ComboBox1: TComboBox;
    grid: TsWorksheetGrid;
    gridR: TsWorksheetGrid;
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    ListBox1: TListBox;
    ListChartSource1: TListChartSource;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel9: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton4: TRadioButton;
    RadioButton5: TRadioButton;
    Separator3: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    PageControl1: TPageControl;
    PopupMenu3: TPopupMenu;
    SaveDialogF: TSaveDialog;
    Separator2: TMenuItem;
    Separator1: TMenuItem;
    OpenDialog: TOpenDialog;
    Panel1: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    PopupMenu1: TPopupMenu;
    PopupMenu2: TPopupMenu;
    RBFreq: TRadioButton;
    RBShep: TRadioButton;
    RBPej: TRadioButton;
    RBHist: TRadioButton;
    RBProb: TRadioButton;
    RBBiva: TRadioButton;
    SaveDialog: TSaveDialog;
    SaveDialogR: TSaveDialog;
    SpinEdit1: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    sWorkbookSource1: TsWorkbookSource;
    sWorkbookSource2: TsWorkbookSource;
    sWorkbookTabControl1: TsWorkbookTabControl;
    sWorkbookTabControl2: TsWorkbookTabControl;
    procedure BtnAbrirClick(Sender: TObject);
    procedure BtnAcoesGClick(Sender: TObject);
    procedure BtnAcoesRClick(Sender: TObject);
    procedure BtnAum1Click(Sender: TObject);
    procedure BtnAum2Click(Sender: TObject);
    procedure BtnAum3Click(Sender: TObject);
    procedure BtnCopiarClick(Sender: TObject);
    procedure BtnExecClick(Sender: TObject);
    procedure BtnHist1Click(Sender: TObject);
    procedure BtnHist2Click(Sender: TObject);
    procedure BtnNovoClick(Sender: TObject);
    procedure BtnAcoesClick(Sender: TObject);
    procedure BtnSalvar2Click(Sender: TObject);
    procedure BtnSalvarResClick(Sender: TObject);
    procedure BtnSelectAllClick(Sender: TObject);
    procedure BtnSalvarClick(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure CB2Change(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure gridClick(Sender: TObject);
    procedure gridEditingDone(Sender: TObject);
    procedure gridGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: string);
    procedure MenuItem10Click(Sender: TObject);
    procedure MenuItem11Click(Sender: TObject);
    procedure MenuItem12Click(Sender: TObject);
    procedure MenuItem13Click(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure MenuItem7Click(Sender: TObject);
    procedure MenuItem8Click(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure RadioButton5Click(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure RBFreqClick(Sender: TObject);
    procedure RBShepClick(Sender: TObject);
    procedure RBPejClick(Sender: TObject);
    procedure RBHistClick(Sender: TObject);
    procedure RBProbClick(Sender: TObject);
    procedure RBBivaClick(Sender: TObject);
    procedure RadioButton4Click(Sender: TObject);
    procedure PreencheListBox();
    procedure sWorkbookTabControl1Change(Sender: TObject);
  private
   FLine: TLineSeries;
   FBar: TBarSeries;
   PT03,PT05,PT10,PT15,PT16,PT20,PT25,PT30,PT35,PT45,PT50,PT55,PT65,PT70,PT75,PT84,
   PT80,PT85,PT90,PT95,PT97,Media,Sele,Ass,Curt:Extended;
   Med,Se,Assi,Cu:String;
   VeioDoAbrir,MudouGridR: Boolean;
   NoAm:Array[1..99] of ShortInt;
   MaxAm,h:Integer;
   IsMenu1Popup,IsMenu2Popup,IsMenu3Popup:Boolean;
   function Pala(q: Extended):Integer;
   procedure ConfirmSave();
   procedure Folk();
   procedure Mca();
   procedure Mcb();
   procedure Trask();
   procedure Otto();
   procedure Areia();
   procedure Selecao();
   procedure Assimetria();
   procedure Curtose();
   procedure Analisar();
   procedure Graficos();
   procedure Shep(); //E pej também
   procedure Freq();
   procedure Bivariado();
   procedure Hist(Sender: TObject);
   procedure Prob();
  public
  end;

type
 TMyTabSheet = class(TTabSheet)
 sImage: TImage;
 sChart: TChart;
end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }


procedure TForm1.ConfirmSave();
begin
if Label1.Font.Style=[fsItalic] then
 if MessageDlg('Continuar sem salvar planilha de entrada?',mtConfirmation,[mbYes,mbNo],0)=mrNo then Abort;
if MudouGridR then
 if MessageDlg('Continuar sem salvar planilha de resultados?',mtConfirmation,[mbYes,mbNo],0)=mrNo then Abort;
end;

procedure TForm1.gridGetEditText(Sender: TObject; ACol, ARow: Integer;
  var Value: string);
begin
Label1.Font.Style:=[fsItalic]; //Significa que 'mudou'
end;

procedure TForm1.PreencheListBox();
var
n,i: integer;
S: String;
begin
ListBox1.Clear;
n:=Grid.Worksheet.GetLastRowNumber;
if n=0 then Abort;
for i:=1 to n do begin
 S:=Grid.Worksheet.ReadAsText(i,0);
 if S<>'' then
  ListBox1.Items.Append(S);
end;
end;

procedure TForm1.sWorkbookTabControl1Change(Sender: TObject);
begin
PreencheListBox();
end;

procedure TForm1.gridEditingDone(Sender: TObject);
begin
if ListBox1.Count=0 then Exit;;
if (Grid.Col=1) and (not VeioDoAbrir) then
 PreencheListBox();
end;

procedure TForm1.BtnAbrirClick(Sender: TObject);
begin
IsMenu1Popup:=False;
ConfirmSave();
if not OpenDialog.Execute then Exit;
sWorkbookSource1.FileName:=OpenDialog.FileName;
Caption:='SysGran 4.0 - '+OpenDialog.FileName;
VeioDoAbrir:=True;
PreencheListBox();
ListBox1.SelectAll;
VeioDoAbrir:=False;
Label1.Font.Style:=[];
GridR.NewWorkbook(100,100);
MudouGridR:=False;
sWorkbookSource2.Worksheet.Name := 'Resultado1';
end;

procedure TForm1.BtnAcoesGClick(Sender: TObject);
var
 lowerLeft: TPoint;
 S: String;
begin
if PageControl1.PageCount=0 then Exit;
S:=PageControl1.ActivePage.Caption;
if (pos('Shep',S)>0) or (pos('Pej',S)>0) then begin
 MenuItem9.Enabled:=False;
 MenuItem10.Enabled:=False;
 MenuItem11.Enabled:=False;
 MenuItem13.Enabled:=False;
 MenuItem12.Enabled:=True;
end else begin
 MenuItem9.Enabled:=True;
 MenuItem10.Enabled:=True;
 MenuItem11.Enabled:=True;
 MenuItem13.Enabled:=True;
 MenuItem9.Checked:=(PageControl1.ActivePage as TMyTabSheet).sChart.Legend.Visible;
 MenuItem10.Checked:=(PageControl1.ActivePage as TMyTabSheet).sChart.LeftAxis.Grid.Visible;
 MenuItem11.Checked:=(PageControl1.ActivePage as TMyTabSheet).sChart.Title.Visible;
 MenuItem13.Checked:=(PageControl1.ActivePage as TMyTabSheet).sChart.BottomAxis.Title.Visible;
 MenuItem12.Enabled:=False;
end;

lowerLeft := Point(0, BtnAcoesG.Height);
lowerLeft := BtnAcoesG.ClientToScreen(lowerLeft);

if IsMenu3Popup=False then begin
 PopupMenu3.Popup(lowerLeft.X, lowerLeft.Y);
 IsMenu3Popup:=True;
end else IsMenu3Popup:=False;

end;

procedure TForm1.BtnAcoesRClick(Sender: TObject);
var lowerLeft: TPoint;
begin
lowerLeft := Point(0, BtnAcoesR.Height);
lowerLeft := BtnAcoesR.ClientToScreen(lowerLeft);
if IsMenu2Popup=False then begin
 PopupMenu2.Popup(lowerLeft.X, lowerLeft.Y);
 IsMenu2Popup:=True;
end else IsMenu2Popup:=False;
end;

procedure TForm1.BtnAum1Click(Sender: TObject);
begin
if Panel2.Width<=880 then begin
 Panel2.Width:=1500;
 Panel3.Height:=800;
end else begin
 Panel2.Width:=880;
 Panel3.Height:=491;
end;
end;

procedure TForm1.BtnAum2Click(Sender: TObject);
begin
if Panel2.Width<=880 then begin
 Panel2.Width:=1500;
 Panel3.Height:=200;
end else begin
 Panel2.Width:=880;
 Panel3.Height:=491;
end;
end;

procedure TForm1.BtnAum3Click(Sender: TObject);
begin
if Panel2.Width>=880 then begin
 Panel2.Width:=200;
 Panel7.Height:=120;
end else begin
 Panel2.Width:=880;
 Panel7.Height:=491;
end;
end;

procedure TForm1.BtnCopiarClick(Sender: TObject);
var S: String;
begin
if PageControl1.PageCount=0 then Exit;;
S:=PageControl1.ActivePage.Caption;
if (Pos('Shep',S)>0) or (Pos('Pej',S)>0) then
 Clipboard.Assign((PageControl1.ActivePage as TMyTabSheet).sImage.Picture)
else begin
 (PageControl1.ActivePage as TMyTabSheet).sChart.CopyToClipboardBitmap;
end;
end;

procedure TForm1.Selecao();
begin
If Sele<=0.35 then Se:='Muito bem selecionado';
If (Sele>0.35) and (Sele<=0.5) then Se:='Bem selecionado';
If (Sele>0.5) and (Sele<=1) then Se:='Moderadamente selecionado';
If (Sele>1) and (Sele<=2) then Se:='Pobremente selecionado';
If (Sele>2) and (Sele<=4) then Se:='Muito pobremente selecionado';
If Sele>4 then Se:='Extremamente mal selecionado';
end;
procedure TForm1.Folk();
begin
Media:=(PT16+PT50+PT84)/3;
Areia;
Sele:=((PT84-PT16)/4)+((PT95-PT05)/6.6);
Selecao;
Ass:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
Assimetria;
Curt:=(PT95-PT05)/(2.44*(PT75-PT25));
Curtose;
end;

procedure TForm1.Mca();
begin
Media:=(PT10+PT30+PT50+PT70+Pt90)/5;
Areia;
Sele:=((PT85+PT95-PT05-PT15)/5.4);
Selecao;
Ass:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
Assimetria;
Curt:=(PT95-PT05)/(2.44*(PT75-PT25));
Curtose;
end;

procedure TForm1.Mcb();
begin
Media:=(PT05+PT15+PT25+PT35+PT45+PT55+PT65+PT75+PT85+PT95)/10;
Areia;
Sele:=(PT70+PT80+PT90+PT97-PT03-PT10-PT20-PT30)/9.1;
Selecao;
Ass:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
Assimetria;
Curt:=(PT95-PT05)/(2.44*(PT75-PT25));
Curtose;
end;

procedure TForm1.Trask();
begin
Media:=PT50;
Areia;
Sele:=(PT75-PT25)/1.35;
Selecao;
Ass:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
Assimetria;
Curt:=(PT95-PT05)/(2.44*(PT75-PT25));
Curtose;
end;

procedure TForm1.Otto();
begin
Media:=(PT16+PT84)/2;
Areia;
Sele:=(PT84-PT16)/2;
Selecao;
Ass:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
Assimetria;
Curt:=(PT95-PT05)/(2.44*(PT75-PT25));
Curtose;
end;

procedure TForm1.Areia();
begin
if Media<=-8 then Med:='Matacão';
if (Media>-8) and (Media<=-6) then Med:='Calhau';
if (Media>-6) and (Media<=-2) then Med:='Seixo';
if (Media>-2) and (Media<=-1) then Med:='Granulo';
if (Media>-1) and (Media<=0) then Med:='Areia muito grossa';
if (Media>0) and (Media<=1) then Med:='Areia grossa';
if (Media>1) and (Media<=2) then Med:='Areia média';
if (Media>2) and (Media<=3) then Med:='Areia fina';
if (Media>3) and (Media<=4) then Med:='Areia muito fina';
if (Media>4) and (Media<=5) then Med:='Silte grosso';
if (Media>5) and (Media<=6) then Med:='Silte médio';
if (Media>6) and (Media<=7) then Med:='Silte fino';
if (Media>7) and (Media<=8) then Med:='Silte muito fino';
if (Media>8) and (Media<=9) then Med:='Argila grossa';
if (Media>9) then Med:='Argila';
end;

procedure TForm1.Assimetria();
begin
If (Ass>-1) and (Ass<=-0.3) then Assi:='Muito negativa';
If (Ass>-0.3) and (Ass<=-0.1) then Assi:='Negativa';
If (Ass>-0.1) and (Ass<=0.1) then Assi:='Aproximadamente simétrica';
If (Ass>0.1) and (Ass<=0.3) then Assi:='Positiva';
If (Ass>0.3) and (Ass<=1) then Assi:='Muito positiva';
end;

procedure TForm1.Curtose();
begin
If Curt<=0.67 then Cu:='Muito platicúrtica';
If (Curt>0.67) and (Curt<=0.9) then Cu:='Platicúrtica';
If (Curt>0.9) and (Curt<=1.11) then Cu:='Mesocúrtica';
If (Curt>1.11) and (Curt<=1.5) then Cu:='Leptocúrtica';
If (Curt>1.5) and (Curt<=3) then Cu:='Muito leptocúrtica';
If Curt>3 then Cu:='Extremamente leptocúrtica';
end;

procedure TForm1.Freq();
var
 k: integer;
 ATabSheet: TMyTabSheet;
 P,z,g,Amost,r,i:Integer;
 Soma:Extended;
 S:String;
 Clw:Array[0..255] of Extended;
 VFw:array[0..255] of Extended;
 Pw:array[0..255] of Extended;
 VBw:Array[0..255] of Extended;
begin
if ListBox1.SelCount>10 then begin
 MessageDlg('Não mais que 10 amostras podem ser plotadas para frequências acumuladas!', mtError, [mbOk], 0);
 Exit;
end;
try
Screen.Cursor:=crHourGlass;

k := PageControl1.PageCount + 1;
ATabSheet := TMyTabSheet.Create(PageControl1);
ATabSheet.Parent := PageControl1;
ATabSheet.Caption:='Frequência '+IntToStr(k);
ATabSheet.sChart:=TChart.Create(ATabSheet);
ATabSheet.sChart.Parent:=ATabSheet;
ATabSheet.sChart.Align:=AlClient;
ATabSheet.sChart.Legend.Visible:=True;
ATabSheet.sChart.BottomAxis.Title.Caption:='Phi';
ATabSheet.sChart.LeftAxis.Title.Caption:='%';
ATabSheet.sChart.BottomAxis.Title.Visible:=True;
ATabSheet.sChart.LeftAxis.Title.Visible:=True;

P:=0;z:=0;g:=0;Amost:=0;r:=0;i:=0;

for i:=2 to 255 do if Grid.Cells[i,1]=null then Break;
r:=i-2;
P:=0;

for z:=0 to ListBox1.Items.Count-1 do begin
 if ListBox1.Selected[z] then begin
  P:=P+1;
  S:=ListBox1.Items[z];
  Amost:=z+2
 end else continue;
 for g:=1 to r do begin
   Clw[g]:=StrToFloat(Grid.Cells[g+1,1]);
   Vbw[g]:=StrToFloat(Grid.Cells[g+1,Amost]);
 end;
 Clw[0]:=0;
 Vbw[0]:=0;
 VFw[0]:=0;
 Pw[0]:=0;
 Soma:=0;
 for i:=1 to r do Soma:=Soma + Vbw[I];
 for i:=1 to r do VFw[I]:=(Vbw[I]*100)/Soma;
 Pw[1]:=VFw[1];
 for i:=2 to r do Pw[I]:=Pw[I-1]+VFw[I];

 FLine := TLineSeries.Create(ATabSheet.sChart);
 FLine.ShowPoints := true;
 FLine.Title := S;
 ATabSheet.sChart.AddSeries(FLine);
 FLine.LinePen.Width:=2;

 if P=1 then begin
   FLine.Pointer.Style := psRectangle;
   FLine.Pointer.Brush.Color := clBlack;
   FLine.SeriesColor := clBlack;
 end;
 if P=2 then begin
   FLine.Pointer.Style := psCircle;
   FLine.Pointer.Brush.Color := clBlue;
   FLine.SeriesColor := clBlue;
 end;
 if P=3 then begin
   FLine.Pointer.Style := psTriangle;
   FLine.Pointer.Brush.Color := clRed;
   FLine.SeriesColor := clRed;
 end;
 if P=4 then begin
   FLine.Pointer.Style := psCross;
   FLine.Pointer.Brush.Color := clGreen;
   FLine.SeriesColor := clGreen;
 end;
 if P=5 then begin
   FLine.Pointer.Style := psStar;
   FLine.Pointer.Brush.Color := clYellow;
   FLine.SeriesColor := clYellow;
 end;
 if P=6 then begin
   FLine.Pointer.Style := psHexagon;
   FLine.Pointer.Brush.Color := clGray;
   FLine.SeriesColor := clGray;
 end;
 if P=7 then begin
   FLine.Pointer.Style := psDiamond;
   FLine.Pointer.Brush.Color := clSilver;
   FLine.SeriesColor := clSilver;
 end;
 if P=8 then begin
   FLine.Pointer.Style := psDiagCross;
   FLine.Pointer.Brush.Color := clNavy;
   FLine.SeriesColor := clNavy;
 end;
 if P=9 then begin
   FLine.Pointer.Style := psLowBracket;
   FLine.Pointer.Brush.Color := clCream;
   FLine.SeriesColor := clCream;
 end;
 if P=10 then begin
   FLine.Pointer.Style := psHighBracket;
   FLine.Pointer.Brush.Color := clOlive;
   FLine.SeriesColor := clOlive;
 end;

 for i:=1 to r do
  FLine.AddXY(Clw[i],Pw[i],'');

end;

if P=1 then ATabSheet.sChart.Legend.Visible:=False;

ATabSheet.sChart.Title.Text.Text:='Freqüências acumuladas';
ATabSheet.sChart.Title.Visible:=True;

PageControl1.ActivePageIndex:=PageControl1.PageCount-1;

IsMenu3Popup:=False;
Screen.Cursor:=crDefault;

except on exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
end;
end;
end;

procedure TForm1.BtnExecClick(Sender: TObject);
begin
MenuItem8Click(Self); //Testa se é planilha do SysgRan e aborta sozinho

if ListBox1.SelCount=0 then begin
 Screen.Cursor:=crDefault;
 MessageDlg('Selecione pelo menos uma amostra, por favor!', mtError, [mbOk], 0);
 Exit;
end;

if RadioButton2.Checked then Analisar else Graficos;
end;

procedure TForm1.BtnHist1Click(Sender: TObject);
begin
if h=2 then BtnHist1.Enabled:=False;
if h<=MaxAm then BtnHist2.Enabled:=True;
h:=h-1;
Hist(Self);
end;

procedure TForm1.BtnHist2Click(Sender: TObject);
begin
if h>=0 then BtnHist1.Enabled:=True;
h:=h+1;
Hist(Self);
if h=MaxAm then BtnHist2.Enabled:=False;
end;

procedure TForm1.Prob();
var k,Amost,P,I,r,z,g,C:Integer;
S:String;
Clz:Array[0..255] of Extended;
VBz:Array[0..255] of Extended;
VFz:array[0..255] of Extended;
Pz:array[0..255] of Extended;
Si: String;
Soma:Extended;
ATabSheet: TMyTabSheet;
begin
try
If ListBox1.SelCount>10 then begin
 MessageDlg('Não mais que 10 amostras podem ser plotadas em gráficos de probabilidade!', mtError, [mbOk], 0);
 Exit;
end;

Screen.Cursor := CrHourGlass;

k := PageControl1.PageCount + 1;
ATabSheet := TMyTabSheet.Create(PageControl1);
ATabSheet.Parent := PageControl1;
ATabSheet.Caption:='Probabilidades '+IntToStr(k);
ATabSheet.sChart:=TChart.Create(ATabSheet);
ATabSheet.sChart.Parent:=ATabSheet;
ATabSheet.sChart.Align:=AlClient;
ATabSheet.sChart.Legend.Visible:=True;
ATabSheet.sChart.BottomAxis.Title.Caption:='%';
ATabSheet.sChart.LeftAxis.Title.Caption:='Phi';
ATabSheet.sChart.BottomAxis.Title.Visible:=True;
ATabSheet.sChart.LeftAxis.Title.Visible:=True;
ATabSheet.sChart.LeftAxis.Inverted:=True;
ATabSheet.sChart.Title.Text.Text:='Gráfico de probabilidades';
ATabSheet.sChart.Title.Visible:=True;

Amost:=0;P:=0;I:=0;r:=0;z:=0;g:=0;C:=0;

for z:=ListBox1.Items.Count-1 downto 0 do if ListBox1.Selected[z] then Si:=ListBox1.Items[z];
for i:=2 to 255 do if Grid.Cells[i,1]=null then Break;
r:=i-2;
for z:=0 to ListBox1.Items.Count-1 do begin
 if ListBox1.Selected[z] then begin
  P:=P+1;
  S:=ListBox1.Items[z];
  Amost:=z+2
 end else continue;
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(Grid.Cells[g+1,1]);
  Vbz[g]:=StrToFloat(Grid.Cells[g+1,Amost]);
 end;

  Clz[0]:=0;Vbz[0]:=0;VFz[0]:=0;Pz[0]:=0;Soma:=0;

 for I:=1 to r do Soma:=Soma + Vbz[I];
 for I:=1 to r do VFz[I]:=(Vbz[I]*100)/Soma;

 Pz[1]:=VFz[1];
 for I:=2 to r do Pz[I]:=Pz[I-1]+VFz[I];

 FLine := TLineSeries.Create(ATabSheet.sChart);
 FLine.ShowPoints := true;
 FLine.Title := S;

 ATabSheet.sChart.AddSeries(FLine);

  if P=1 then begin
   FLine.Pointer.Style := psRectangle;
   FLine.Pointer.Brush.Color := clBlack;
   FLine.SeriesColor := clBlack;
 end;
 if P=2 then begin
   FLine.Pointer.Style := psCircle;
   FLine.Pointer.Brush.Color := clBlue;
   FLine.SeriesColor := clBlue;
 end;
 if P=3 then begin
   FLine.Pointer.Style := psTriangle;
   FLine.Pointer.Brush.Color := clRed;
   FLine.SeriesColor := clRed;
 end;
 if P=4 then begin
   FLine.Pointer.Style := psCross;
   FLine.Pointer.Brush.Color := clGreen;
   FLine.SeriesColor := clGreen;
 end;
 if P=5 then begin
   FLine.Pointer.Style := psStar;
   FLine.Pointer.Brush.Color := clYellow;
   FLine.SeriesColor := clYellow;
 end;
 if P=6 then begin
   FLine.Pointer.Style := psHexagon;
   FLine.Pointer.Brush.Color := clGray;
   FLine.SeriesColor := clGray;
 end;
 if P=7 then begin
   FLine.Pointer.Style := psDiamond;
   FLine.Pointer.Brush.Color := clSilver;
   FLine.SeriesColor := clSilver;
 end;
 if P=8 then begin
   FLine.Pointer.Style := psDiagCross;
   FLine.Pointer.Brush.Color := clNavy;
   FLine.SeriesColor := clNavy;
 end;
 if P=9 then begin
   FLine.Pointer.Style := psLowBracket;
   FLine.Pointer.Brush.Color := clCream;
   FLine.SeriesColor := clCream;
 end;
 if P=10 then begin
   FLine.Pointer.Style := psHighBracket;
   FLine.Pointer.Brush.Color := clOlive;
   FLine.SeriesColor := clOlive;
 end;


 for i:=1 to r do
  FLine.AddXY(Pz[i],Clz[i],''); //Resolve tudo

end;

PageControl1.ActivePageIndex:=PageControl1.PageCount-1;
IsMenu3Popup:=False;

Screen.Cursor := CrDefault;
except on Exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
end;
end;
end;

{procedure TForm1.Prob();
var k,Amost,P,I,r,z,g,C:Integer;
S:String;
Clz:Array[0..255] of Extended;
VBz:Array[0..255] of Extended;
VFz:array[0..255] of Extended;
Pz:array[0..255] of Extended;
Si: String;
Soma:Extended;
X,Y:array[0..99] of Extended;
ValY,ValX:array[1..99] of Integer;
MaxX,MaxY,MinX,MinY,aX,bX,aY,bY,DistX,DistY:Extended;
Bitmap:TBitmap;
ATabSheet: TMyTabSheet;
begin
try
{If ListBox1.SelCount>10 then begin
 MessageDlg('Não mais que 10 amostras podem ser plotadas em gráficos de probabilidade!', mtError, [mbOk], 0);
 Exit;
end;
}
Screen.Cursor := CrHourGlass;

Bitmap:=TBitmap.Create;
//Bitmap.Height:=376;
//Bitmap.Width:=720;
Bitmap.Height:=800;
Bitmap.Width:=1200;

k := PageControl1.PageCount + 1;
ATabSheet := TMyTabSheet.Create(PageControl1);
ATabSheet.Parent := PageControl1;

ATabSheet.Caption:='Probabilidades '+IntToStr(k);
ATabSheet.sImage := TImage.Create(ATabSheet);
ATabSheet.sImage.Parent:=ATabSheet;
ATabSheet.sImage.Align:=AlClient;
ATabSheet.sImage.Picture.Assign(Bitmap);

Amost:=0;P:=0;I:=0;r:=0;z:=0;g:=0;C:=0;
MaxX:=0;MaxY:=0;MinX:=0;MinY:=0;aX:=0;bX:=0;aY:=0;bY:=0;DistX:=0;DistY:=0;

with ATabSheet.sImage.Canvas do begin
 Brush.Color:=clWhite;
 FillRect(ATabSheet.sImage.BoundsRect);//Modo novo
 //FloodFill(1, 1, clgreen, fsBorder); //Modo velho
 Clear; //Precisa
 Pen.Color:=clBlack;

 Rectangle(61,23,480,172);
 ATabSheet.sImage.Canvas.Font.Size:=8;

 for i:=0 to 7 do begin
  MoveTo(61,171-i*21);LineTo(58,171-i*21);
 end;
 for i:=0 to 16 do begin
  MoveTo(73+i*24,171);LineTo(73+i*24,174);
 end;
 MoveTo(479,171);LineTo(479,174);MoveTo(278,171);LineTo(278,175);MoveTo(277,171);LineTo(277,175);
 for i:=0 to 16 do begin
  MoveTo(85+i*24,171);LineTo(85+i*24,173);
 end;
 //Brush.Color:=clSilver;
 //FloodFill(2, 2, clBlack, fsBorder);
 MoveTo(0,0);LineTo(488,0);MoveTo(488,0);LineTo(488,209);MoveTo(488,208);LineTo(0,208);
 MoveTo(0,208);LineTo(0,0);Pen.Color:=clSilver;MoveTo(63,24);LineTo(479,24);MoveTo(62,24);
 LineTo(62,171);Font.Name:='Arial';Font.Size:=7;TextOut(67,175,'0,1');TextOut(91,175,'0,5');
 TextOut(119,175,'1');TextOut(143,175,'2');TextOut(167,175,'5');TextOut(189,175,'10');
 TextOut(213,175,'20');TextOut(237,175,'30');TextOut(273,175,'50'); TextOut(309,175,'70');
 TextOut(333,175,'80');TextOut(357,175,'90');TextOut(381,175,'95');TextOut(405,175,'98');
 TextOut(429,175,'99');TextOut(448,175,'99,5');TextOut(470,175,'99,9'); Font.Size:=10;
 TextOut(250,190,'%');TextOut(20,80,'P');TextOut(20,95,'h');TextOut(20,110,'i');
end;

P:=0;z:=0;

for z:=ListBox1.Items.Count-1 downto 0 do if ListBox1.Selected[z] then Si:=ListBox1.Items[z];
for i:=2 to 255 do if Grid.Cells[i,1]=null then Break;
r:=i-2;
for z:=0 to ListBox1.Items.Count-1 do begin
 if ListBox1.Selected[z] then begin
  P:=P+1;
  S:=ListBox1.Items[z];
  Amost:=z+2
 end else continue;
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(Grid.Cells[g+1,1]);
  Vbz[g]:=StrToFloat(Grid.Cells[g+1,Amost]);
 end;

 Clz[0]:=0;Vbz[0]:=0;VFz[0]:=0;Pz[0]:=0;Soma:=0;

 for I:=1 to r do Soma:=Soma + Vbz[I];
 for I:=1 to r do VFz[I]:=(Vbz[I]*100)/Soma;

 Pz[1]:=VFz[1];
 for I:=2 to r do Pz[I]:=Pz[I-1]+VFz[I];

 MaxY:=50; //Era 30, botei 50
 MinY:=0;
 for i:=1 to r do if Clz[i]>MinY then MinY:=Clz[i];
 for i:=1 to r do if Clz[i]<MaxY then MaxY:=Clz[i];

 ATabSheet.sImage.Canvas.Font.Name:='Arial';
 ATabSheet.sImage.Canvas.Font.Size:=7;

 DistY:=(MinY-MaxY)/7;
 for i:=0 to 7 do
  ATabSheet.sImage.Canvas.TextOut(38,19+(i*21),FloatToStrf(MaxY+(i*DistY), ffNumber, 8, 2));

 MinX:=0.1;
 MaxX:=99.9;
 DistX:=(99.9-0.1)/17;
 DistY:=(MaxY-MinY)/7;

 bY:=((2*(((MinY+DistY)*151)+((MinY+(6*DistY))*47))) - (((MinY+DistY)+(MinY+(6*DistY)))*198))/
 ((2*(SQR((MinY+DistY))+SQR((MinY+(6*DistY))))) - (SQR((MinY+DistY)+(MinY+(6*DistY)))));
 aY:=99 - bY*(((MinY+DistY)+(MinY+(6*DistY)))/2);

for i:=1 to r do begin



 if Pz[i]<=0.1 then begin aX:=62; bX:=110; end;
 if (Pz[i]>0.1) and (Pz[i]<=0.5) then begin aX:=67; bX:=60; end;
 if (Pz[i]>0.5) and (Pz[i]<=1) then begin aX:=73; bX:=48; end;
 if (Pz[i]>1) and (Pz[i]<=2) then begin aX:=97; bX:=24; end;
 if (Pz[i]>2) and (Pz[i]<=5) then begin aX:=129; bX:=8; end;
 if (Pz[i]>5) and (Pz[i]<=10) then begin aX:=145; bX:=4.8; end;
 if (Pz[i]>10) and (Pz[i]<=30) then begin aX:=169; bX:=2.4; end;
 if (Pz[i]>30) and (Pz[i]<=50) then begin aX:=185.5; bX:=1.85; end;
 if (Pz[i]>50) and (Pz[i]<=70) then begin aX:=190.5; bX:=1.75; end;
 if (Pz[i]>70) and (Pz[i]<=90) then begin aX:=145; bX:=2.4; end;
 if (Pz[i]>90) and (Pz[i]<=95) then begin aX:=-71;bX:=4.8; end;
 if (Pz[i]>95) and (Pz[i]<=98) then begin aX:=-375; bX:=8; end;
 if (Pz[i]>98) and (Pz[i]<=99) then begin aX:=-1943;bX:=24; end;
 if (Pz[i]>99) and (Pz[i]<=99.5) then begin aX:=-4319;bX:=48; end;
 if (Pz[i]>99.5) and (Pz[i]<=99.9) then begin aX:=-3921;bX:=44; end;

 ValX[i]:=Trunc(aX+bX*Pz[i])-1;
 if Pz[i]>99.9 then ValX[i]:=478;
 ValY[i]:=Trunc(aY+bY*Clz[i])-3;

 with ATabSheet.sImage.Canvas do begin
  if P=1 then begin Brush.Color:=clBlue;
  Pen.Color:=clBlue;
  Ellipse(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
 end;
 if P=2 then begin
  Brush.Color:=clBlack;Pen.Color:=clBlack;Ellipse(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);Ellipse(500,30+17,508,38+17);
  Brush.Color:=clWhite;TextOut(513,28+17,S);Brush.Color:=clBlue;Pen.Color:=clBlue;
  Ellipse(500,30,508,38);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28,Si);
 end;
 if P=3 then begin
  Brush.Color:=clRed;Pen.Color:=clRed;Ellipse(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Ellipse(500,30+34,508,38+34);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+34,S);
 end;
 if P=4 then begin
  Brush.Color:=clYellow;Pen.Color:=clYellow;Ellipse(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Ellipse(500,30+51,508,38+51);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+51,S);
 end;
 if P=5 then begin
  Brush.Color:=clGreen;Pen.Color:=clGreen;Ellipse(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Ellipse(500,30+68,508,38+68);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+68,S);
 end;
 if P=6 then begin
  Brush.Color:=clBlue;Pen.Color:=clBlue;Rectangle(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Rectangle(500,30+85,508,38+85);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+85,S);
 end;
 if P=7 then begin
  Brush.Color:=clBlack;Pen.Color:=clBlack;Rectangle(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Rectangle(500,30+102,508,38+102);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+102,S);
 end;
 if P=8 then begin
  Brush.Color:=clRed;Pen.Color:=clRed;Rectangle(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Rectangle(500,30+119,508,38+119);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+119,S);
 end;
 if P=9 then begin
  Brush.Color:=clYellow;Pen.Color:=clYellow;Rectangle(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Rectangle(500,30+136,508,38+136);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+136,S);
 end;
 if P=10 then begin
  Brush.Color:=clGreen;Pen.Color:=clGreen;Rectangle(ValX[i],ValY[i],ValX[i]+5,ValY[i]+5);
  Rectangle(500,30+153,508,38+153);Brush.Color:=clWhite;Pen.Color:=clBlack;TextOut(513,28+153,S);
 end;
end;
end;

ATabSheet.sImage.Canvas.Pen.Color:=clBlack;
for i:=1 to r-1 do begin
 ATabSheet.sImage.Canvas.MoveTo(ValX[i],ValY[i]+2);
 ATabSheet.sImage.Canvas.LineTo(ValX[i+1]+1,ValY[i+1]+2);
end;
ATabSheet.sImage.Canvas.Brush.Color:=clWhite;
end;

with ATabSheet.sImage.Canvas do begin
 ATabSheet.sImage.Canvas.Pen.Color:=clBlack;
 ATabSheet.sImage.Canvas.Brush.Color:=clWhite;
 if P=1 then begin
  ATabSheet.sImage.Canvas.Font.Size:=9;
  ATabSheet.sImage.Canvas.TextOut(228,5,S);
 end;
 if P=2 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+38);LineTo(496,23+38);LineTo(496,23);
 end;
 if P=3 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+56);LineTo(496,23+56);LineTo(496,23);
 end;
 if P=4 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+72);LineTo(496,23+72);LineTo(496,23);
 end;
 if P=5 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+90);LineTo(496,23+90);LineTo(496,23);
 end;
 if P=6 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+108);LineTo(496,23+108);LineTo(496,23);
 end;
 if P=7 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+124);LineTo(496,23+124);LineTo(496,23);
 end;
 if P=8 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+142);LineTo(496,23+142);LineTo(496,23);
 end;
 if P=9 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+158);LineTo(496,23+158);LineTo(496,23);
 end;
 if P=10 then begin
  MoveTo(496,23);LineTo(558,23);LineTo(558,23+174);LineTo(496,23+174);LineTo(496,23);
 end;
end;

//ATabSheet.sImage.Top:=2;
//ATabSheet.sImage.Left:=2;

ATabSheet.sImage.Visible:=True;
PageControl1.ActivePageIndex:=PageControl1.PageCount-1;
IsMenu3Popup:=False;
Bitmap.Free;

Screen.Cursor := CrDefault;

except on Exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
end;
end;
end;
}

procedure TForm1.Bivariado();
var k,Amost,P,I,r,z,g,C:Integer;
Clz:Array[0..255] of Extended;
VBz:Array[0..255] of Extended;
VFz:array[0..255] of Extended;
Pz:array[0..255] of Extended;
Mid:array[0..255] of Extended;
Std:array[0..255] of Extended;
Mediaz:array[0..999] of Extended;
Selecaoz:array[0..999] of Extended;
Assimetriaz:array[0..999] of Extended;
Curtosez:array[0..999] of Extended;
X,Y:array[0..999] of Extended;
ValY,ValX:array[1..999] of Integer;
ATabSheet: TMyTabSheet;
A1,A2,A2a,A3,A4,P03,P05,P10,P15,P20,P16,P25,P30,P35,P45,P50,P55,P65,P70,P75,P80,P84,P85,P90,P95,P97,
PA03,PA05,PA10,PA15,PA20,PA16,PA25,PA30,PA35,PA45,PA50,PA55,PA65,PA70,PA75,PA80,PA84,PA85,PA90,PA95,PA97,
C03,C05,C10,C15,C16,C20,C25,C30,C35,C45,C50,C55,C65,C70,C75,C80,C84,C85,C90,C95,C97,
CA03,CA05,CA10,CA15,CA16,CA20,CA25,CA30,CA35,CA45,CA50,CA55,CA65,CA70,CA75,CA80,CA84,CA85,CA90,CA95,CA97,
A03,A05,A10,A15,A16,A20,A25,A30,A35,A45,A50,A55,A65,A70,A75,A80,A84,A85,A90,A95,A97,
B03,B05,B10,B15,B16,B20,B25,B30,B35,B45,B50,B55,B65,B70,B75,B80,B84,B85,B90,B95,B97,Soma,MaxX,MaxY,MinX,
MinY,DistX,DistY,bX,aX,bY,aY:Extended;
//PT03a,PT05a,PT10a,PT15a,PT16a,PT20a,PT25a,PT30a,PT35a,PT45a,PT50a,PT55a,PT65a,PT70a,PT75a,PT84a,
//PT80a,PT85a,PT90a,PT95a,PT97a,Mediaa,Selea,Assa,Curta:Extended;
begin
try
Screen.Cursor := CrHourGlass;

for i:=0 to 999 do begin
 Mediaz[i]:=0;
 Assimetriaz[i]:=0;
 Curtosez[i]:=0;
 Selecaoz[i]:=0;
end;
P:=0;
z:=0;
for i:=2 to 255 do if (Grid.Cells[i,1]=null) then Break;
r:=i-2;
for z:=0 to ListBox1.Items.Count-1 do begin
 if ListBox1.Selected[z] then begin
  P:=P+1;
  Amost:=z+2
 end else continue;
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(Grid.Cells[g+1,1]);
  Vbz[g]:=StrToFloat(Grid.Cells[g+1,Amost]);
 end;

 Clz[0]:=0;
 Vbz[0]:=0;
 VFz[0]:=0;
 Pz[0]:=0;
 Soma:=0;

 for I:=1 to r do Soma:=Soma + Vbz[I];
 for I:=1 to r do VFz[I]:=(Vbz[I]*100)/Soma;
 Pz[1]:=VFz[1];
 for I:=2 to r do Pz[I]:=Pz[I-1]+VFz[I];

 for I:=1 to r do begin if Pz[I]>=3 then begin P03:=Pz[I]; PA03:=Pz[I-1]; C03:=Clz[I]; CA03:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=5 then begin P05:=Pz[I]; PA05:=Pz[I-1]; C05:=Clz[I]; CA05:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=10 then begin P10:=Pz[I];PA10:=Pz[I-1]; C10:=Clz[I]; CA10:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=15 then begin P15:=Pz[I];PA15:=Pz[I-1]; C15:=Clz[I]; CA15:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=16 then begin P16:=Pz[I];PA16:=Pz[I-1]; C16:=Clz[I]; CA16:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=20 then begin P20:=Pz[I];PA20:=Pz[I-1]; C20:=Clz[I]; CA20:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=25 then begin P25:=Pz[I];PA25:=Pz[I-1]; C25:=Clz[I]; CA25:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=30 then begin P30:=Pz[I];PA30:=Pz[I-1]; C30:=Clz[I]; CA30:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=35 then begin P35:=Pz[I];PA35:=Pz[I-1]; C35:=Clz[I]; CA35:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=45 then begin P45:=Pz[I];PA45:=Pz[I-1]; C45:=Clz[I]; CA45:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=50 then begin P50:=Pz[I];PA50:=Pz[I-1]; C50:=Clz[I]; CA50:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=55 then begin P55:=Pz[I];PA55:=Pz[I-1]; C55:=Clz[I]; CA55:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=65 then begin P65:=Pz[I];PA65:=Pz[I-1]; C65:=Clz[I]; CA65:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=70 then begin P70:=Pz[I];PA70:=Pz[I-1]; C70:=Clz[I]; CA70:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=75 then begin P75:=Pz[I];PA75:=Pz[I-1]; C75:=Clz[I]; CA75:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=80 then begin P80:=Pz[I];PA80:=Pz[I-1]; C80:=Clz[I]; CA80:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=84 then begin P84:=Pz[I];PA84:=Pz[I-1]; C84:=Clz[I]; CA84:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=85 then begin P85:=Pz[I];PA85:=Pz[I-1]; C85:=Clz[I]; CA85:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=90 then begin P90:=Pz[I];PA90:=Pz[I-1]; C90:=Clz[I]; CA90:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=95 then begin P95:=Pz[I];PA95:=Pz[I-1]; C95:=Clz[I]; CA95:=Clz[I-1]; Break; end;end;
 for I:=1 to r do begin if Pz[I]>=97 then begin P97:=Pz[I];PA97:=Pz[I-1]; C97:=Clz[I]; CA97:=Clz[I-1]; Break; end;end;

B03:=((2*((P03*C03)+(PA03*CA03))) - ((P03+PA03)*(C03+CA03)))/
     (2*((C03*C03)+(CA03*CA03)) - ((C03+CA03)*(C03+CA03)));
A03:=((P03+PA03)/2)-(B03*((C03+CA03)/2));
PT03:=(3-A03)/B03;

B05:=((2*((P05*C05)+(PA05*CA05))) - ((P05+PA05)*(C05+CA05)))/
     (2*((C05*C05)+(CA05*CA05)) - ((C05+CA05)*(C05+CA05)));
A05:=((P05+PA05)/2)-(B05*((C05+CA05)/2));
PT05:=(5-A05)/B05;

B10:=((2*((P10*C10)+(PA10*CA10))) - ((P10+PA10)*(C10+CA10)))/
     (2*((C10*C10)+(CA10*CA10)) - ((C10+CA10)*(C10+CA10)));
A10:=((P10+PA10)/2)-(B10*((C10+CA10)/2));
PT10:=(10-A10)/B10;

B15:=((2*((P15*C15)+(PA15*CA15))) - ((P15+PA15)*(C15+CA15)))/
     (2*((C15*C15)+(CA15*CA15)) - ((C15+CA15)*(C15+CA15)));
A15:=((P15+PA15)/2)-(B15*((C15+CA15)/2));
PT15:=(15-A15)/B15;

B16:=((2*((P16*C16)+(PA16*CA16))) - ((P16+PA16)*(C16+CA16)))/
     (2*((C16*C16)+(CA16*CA16)) - ((C16+CA16)*(C16+CA16)));
A16:=((P16+PA16)/2)-(B16*((C16+CA16)/2));
PT16:=(16-A16)/B16;

B20:=((2*((P20*C20)+(PA20*CA20))) - ((P20+PA20)*(C20+CA20)))/
     (2*((C20*C20)+(CA20*CA20)) - ((C20+CA20)*(C20+CA20)));
A20:=((P20+PA20)/2)-(B20*((C20+CA20)/2));
PT20:=(20-A20)/B20;

B25:=((2*((P25*C25)+(PA25*CA25))) - ((P25+PA25)*(C25+CA25)))/
     (2*((C25*C25)+(CA25*CA25)) - ((C25+CA25)*(C25+CA25)));
A25:=((P25+PA25)/2)-(B25*((C25+CA25)/2));
PT25:=(25-A25)/B25;

B30:=((2*((P30*C30)+(PA30*CA30))) - ((P30+PA30)*(C30+CA30)))/
     (2*((C30*C30)+(CA30*CA30)) - ((C30+CA30)*(C30+CA30)));
A30:=((P30+PA30)/2)-(B30*((C30+CA30)/2));
PT30:=(30-A30)/B30;

B35:=((2*((P35*C35)+(PA35*CA35))) - ((P35+PA35)*(C35+CA35)))/
     (2*((C35*C35)+(CA35*CA35)) - ((C35+CA35)*(C35+CA35)));
A35:=((P35+PA35)/2)-(B35*((C35+CA35)/2));
PT35:=(35-A35)/B35;

B45:=((2*((P45*C45)+(PA45*CA45))) - ((P45+PA45)*(C45+CA45)))/
     (2*((C45*C45)+(CA45*CA45)) - ((C45+CA45)*(C45+CA45)));
A45:=((P45+PA45)/2)-(B45*((C45+CA45)/2));
PT45:=(45-A45)/B45;

B50:=((2*((P50*C50)+(PA50*CA50))) - ((P50+PA50)*(C50+CA50)))/
     (2*((C50*C50)+(CA50*CA50)) - ((C50+CA50)*(C50+CA50)));
A50:=((P50+PA50)/2)-(B50*((C50+CA50)/2));
PT50:=(50-A50)/B50;

B55:=((2*((P55*C55)+(PA55*CA55))) - ((P55+PA55)*(C55+CA55)))/
     (2*((C55*C55)+(CA55*CA55)) - ((C55+CA55)*(C55+CA55)));
A55:=((P55+PA55)/2)-(B55*((C55+CA55)/2));
PT55:=(55-A55)/B55;

B65:=((2*((P65*C65)+(PA65*CA65))) - ((P65+PA65)*(C65+CA65)))/
     (2*((C65*C65)+(CA65*CA65)) - ((C65+CA65)*(C65+CA65)));
A65:=((P65+PA65)/2)-(B65*((C65+CA65)/2));
PT65:=(65-A65)/B65;

B70:=((2*((P70*C70)+(PA70*CA70))) - ((P70+PA70)*(C70+CA70)))/
     (2*((C70*C70)+(CA70*CA70)) - ((C70+CA70)*(C70+CA70)));
A70:=((P70+PA70)/2)-(B70*((C70+CA70)/2));
PT70:=(70-A70)/B70;

B75:=((2*((P75*C75)+(PA75*CA75))) - ((P75+PA75)*(C75+CA75)))/
     (2*((C75*C75)+(CA75*CA75)) - ((C75+CA75)*(C75+CA75)));
A75:=((P75+PA75)/2)-(B75*((C75+CA75)/2));
PT75:=(75-A75)/B75;

B80:=((2*((P80*C80)+(PA80*CA80))) - ((P80+PA80)*(C80+CA80)))/
     (2*((C80*C80)+(CA80*CA80)) - ((C80+CA80)*(C80+CA80)));
A80:=((P80+PA80)/2)-(B80*((C80+CA80)/2));
PT80:=(80-A80)/B80;

B84:=((2*((P84*C84)+(PA84*CA84))) - ((P84+PA84)*(C84+CA84)))/
     (2*((C84*C84)+(CA84*CA84)) - ((C84+CA84)*(C84+CA84)));
A84:=((P84+PA84)/2)-(B84*((C84+CA84)/2));
PT84:=(84-A84)/B84;

B85:=((2*((P85*C85)+(PA85*CA85))) - ((P85+PA85)*(C85+CA85)))/
     (2*((C85*C85)+(CA85*CA85)) - ((C85+CA85)*(C85+CA85)));
A85:=((P85+PA85)/2)-(B85*((C85+CA85)/2));
PT85:=(85-A85)/B85;

B90:=((2*((P90*C90)+(PA90*CA90))) - ((P90+PA90)*(C90+CA90)))/
     (2*((C90*C90)+(CA90*CA90)) - ((C90+CA90)*(C90+CA90)));
A90:=((P90+PA90)/2)-(B90*((C90+CA90)/2));
PT90:=(90-A90)/B90;

B95:=((2*((P95*C95)+(PA95*CA95))) - ((P95+PA95)*(C95+CA95)))/
     (2*((C95*C95)+(CA95*CA95)) - ((C95+CA95)*(C95+CA95)));
A95:=((P95+PA95)/2)-(B95*((C95+CA95)/2));
PT95:=(95-A95)/B95;

B97:=((2*((P97*C97)+(PA97*CA97))) - ((P97+PA97)*(C97+CA97)))/
     (2*((C97*C97)+(CA97*CA97)) - ((C97+CA97)*(C97+CA97)));
A97:=((P97+PA97)/2)-(B97*((C97+CA97)/2));
PT97:=(97-A97)/B97;

if CB4.Text='Folk & Ward' then begin
 Mediaz[P]:=(PT16+PT50+PT84)/3;
 Selecaoz[P]:=((PT84-PT16)/4)+((PT95-PT05)/6.6);
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4.Text='McCammon (a)' then begin
 Mediaz[P]:=(PT10+PT30+PT50+PT70+Pt90)/5;
 Selecaoz[P]:=((PT85+PT95-PT05-PT15)/5.4);
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4.Text='McCammon (b)' then begin
 Mediaz[P]:=(PT05+PT15+PT25+PT35+PT45+PT55+PT65+PT75+PT85+PT95)/10;
 Selecaoz[P]:=(PT70+PT80+PT90+PT97-PT03-PT10-PT20-PT30)/9.1;
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4.Text='Trask' then begin
 Mediaz[P]:=PT50;
 Selecaoz[P]:=(PT75-PT25)/1.35;
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4.text='Otto & Inman' then begin
 Mediaz[P]:=(PT16+PT84)/2;
 Selecaoz[P]:=(PT84-PT16)/2;
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4.Text='Medida dos Momentos' then begin
 Mid[1]:=Clz[1]-((Clz[2]-Clz[1])/2);
 for I:=2 to r do Mid[i]:=Clz[i]-((Clz[i]-Clz[i-1])/2);

 A1:=0;
 for I:=1 to r do Pz[i]:=Mid[i]*VBz[i];
 for I:=1 to r do A1:=A1+Pz[i];
 A1:=A1/Soma;
 Mediaz[P]:=A1;

 for I:=1 to r do Std[i]:=Mid[i]-A1;

 A2:=0;
 for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i]));
 for I:=1 to r do A2:=A2+Pz[i];
 A2:=A2/(Soma-1);
 A2a:=SQRT(A2);
 Selecaoz[P]:=A2a;
 A4:=0;
 for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*Std[i]);
 for I:=1 to r do A4:=A4+Pz[i];
 A4:=A4/(Soma-1);
 A4:=A4/(Exp(1.5*Ln(A2)));
 Assimetriaz[P]:=A4;
 A3:=0;
 for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i]));
 for I:=1 to r do A3:=A3+Pz[i];
 A3:=A3/(Soma-1);
 A3:=A3/(Exp(2*Ln(A2)));
 Curtosez[P]:=A3;
 end;
end;

k := PageControl1.PageCount + 1;
ATabSheet := TMyTabSheet.Create(PageControl1);
ATabSheet.Parent := PageControl1;
ATabSheet.Caption:='Bivariado '+IntToStr(k);
ATabSheet.sChart:=TChart.Create(ATabSheet);
ATabSheet.sChart.Parent:=ATabSheet;
ATabSheet.sChart.Align:=AlClient;
//ATabSheet.sChart.Legend.Visible:=True;
//ATabSheet.sChart.BottomAxis.Title.Caption:='Phi';
//ATabSheet.sChart.LeftAxis.Title.Caption:='%';
ATabSheet.sChart.BottomAxis.Title.Visible:=True;
ATabSheet.sChart.LeftAxis.Title.Visible:=True;


if (CB2.text='Média') or (CB3.text='Média') then begin
 ATabSheet.sChart.BottomAxis.Title.Caption:='Média';
 for i:=0 to 999 do X[i]:=Mediaz[i];
end;
if (CB2.text='Curtose') or (CB3.Text='Curtose') then begin
 for i:=0 to 999 do Y[i]:=Curtosez[i];
 ATabSheet.sChart.LeftAxis.Title.Caption:='Curtose';
end;
if (CB2.text='Assimetria') or (CB3.text='Assimetria') then begin
 if (CB2.text='Curtose') or (CB3.text='Curtose') then begin
   for i:=0 to 999 do X[i]:=Assimetriaz[i];
   ATabSheet.sChart.BottomAxis.Title.Caption:='Assimetria';
 end else begin
  for i:=0 to 999 do Y[i]:=Assimetriaz[i];
  ATabSheet.sChart.LeftAxis.Title.Caption:='Assimetria';
 end;
end;
if (CB2.text='Seleção') or (CB3.text='Seleção') then begin
 if (CB2.text='Média') or (CB3.text='Média') then begin
  for i:=0 to 999 do Y[i]:=Selecaoz[i];
  ATabSheet.sChart.LeftAxis.Title.Caption:='Seleção';
 end else begin
  for i:=0 to 999 do X[i]:=Selecaoz[i];
  ATabSheet.sChart.BottomAxis.Title.Caption:='Seleção';
 end;
end;

MaxX:=0;
MaxY:=0;
MinX:=1000;
MinY:=1000;
for i:=1 to P do if X[i]>MaxX then MaxX:=X[i];
for i:=1 to P do if Y[i]>MaxY then MaxY:=Y[i];
for i:=1 to P do if X[i]<MinX then MinX:=X[i];
for i:=1 to P do if Y[i]<MinY then MinY:=Y[i];

FLine := TLineSeries.Create(ATabSheet.sChart);
FLine.ShowPoints := True;
FLine.ShowLines:=False;
FLine.Pointer.Style:=psCircle;
FLine.Pointer.Brush.Color:=clBlack;
ATabSheet.sChart.AddSeries(FLine);

for i:=1 to P do
 FLine.AddXY(X[i],Y[i],'');

 ATabSheet.sChart.Title.Text.Text:='Bivariado ('+CB2.Text+' X '+CB3.Text+')';
 ATabSheet.sChart.Title.Visible:=True;

 PageControl1.ActivePageIndex:=PageControl1.PageCount-1;

 IsMenu3Popup:=False;

 Screen.Cursor := CrDefault;
except
on Exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
end;
end;
end;

procedure TForm1.Hist(Sender: TObject);
var k,z,P,g,r,i:Integer;
Soma,ValorMax:Extended;
ATabSheet: TMyTabSheet;
Clw:Array[0..2550] of Extended;
VFw:array[0..255] of Extended;
VBw:Array[0..255] of Extended;
begin
try
 Screen.Cursor := CrHourGlass;
 MaxAm:=ListBox1.SelCount;
 BtnHist1.Visible:=True;
 BtnHist2.Visible:=True;
if MaxAm=1 then BtnHist2.Enabled:=False else BtnHist2.Enabled:=True;
 P:=0;
 for z:=0 to ListBox1.Items.Count-1 do begin
  if ListBox1.Selected[z] then begin
   P:=P+1;
   NoAm[P]:=z+2;
  end else continue;
 end;

 for i:=2 to 255 do if Grid.Cells[i,1]=null then Break;
 r:=i-2;
 for g:=1 to r do begin
  Clw[g]:=StrToFloat(Grid.Cells[g+1,1]);
  Vbw[g]:=StrToFloat(Grid.Cells[g+1,NoAm[h]]);
 end;
 Clw[0]:=0;
 Vbw[0]:=0;
 VFw[0]:=0;
 Soma:=0;
 ValorMax:=0;
 for i:=1 to r do Soma:=Soma + Vbw[I];
 for i:=1 to r do VFw[I]:=(Vbw[I]*100)/Soma;
 for i:=1 to r do if VFw[I]>ValorMax then ValorMax:=VFw[I];

 k := PageControl1.PageCount + 1;
 if sender=BtnExec then begin
  ATabSheet := TMyTabSheet.Create(PageControl1);
  ATabSheet.Parent := PageControl1;
  ATabSheet.Caption:='Histograma '+IntToStr(k);
  ATabSheet.sChart:=TChart.Create(ATabSheet);
  ATabSheet.sChart.Parent:=ATabSheet;
  ATabSheet.sChart.Align:=AlClient;
  ATabSheet.sChart.Legend.Visible:=False;
  ATabSheet.sChart.BottomAxis.Title.Caption:='Phi';
  ATabSheet.sChart.LeftAxis.Title.Caption:='%';
  ATabSheet.sChart.BottomAxis.Title.Visible:=True;
  ATabSheet.sChart.LeftAxis.Title.Visible:=True;
 end else
  ATabSheet :=  (PageControl1.ActivePage as TMyTabSheet);

 if sender=BtnExec then
  FBar := TBarSeries.Create(ATabSheet.sChart) else
   FBar.Clear;


 for i:=1 to r do begin
 FBar.Add(VFw[i],FloatToStr(Clw[i]));
  ListChartSource1.Add(i,Clw[i],FloatToStr(Clw[i]));
 end;

 ATabSheet.sChart.Title.Text.Text:=Grid.Cells[1,NoAm[h]];
 ATabSheet.sChart.Title.Visible:=True;

 ATabSheet.sChart.AddSeries(FBar);
 ATabSheet.sChart.BottomAxis.Marks.Source:=ListChartSource1;

 if sender=BtnExec then
  PageControl1.ActivePageIndex:=PageControl1.PageCount-1;

 IsMenu3Popup:=False;
 Screen.Cursor:=crDefault;

  except on exception do begin
   Screen.Cursor:=crDefault;
   BtnHist1.Visible:=False;
   BtnHist2.Visible:=False;
   MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
  end;
 end;
end;

procedure TForm1.Graficos();
begin
BtnHist1.Visible:=False;
BtnHist2.Visible:=False;
if (RBShep.Checked) or (RBPej.Checked) then Shep();
if RBFreq.Checked then Freq();
if RBHist.Checked then Hist(BtnExec);
if RBProb.Checked then Prob();
if RBBiva.Checked then Bivariado();
MenuItem12.Checked:=False;
end;

function TForm1.Pala(q: Extended):Integer;
var
S:String;
U:Extended;
begin
U:=Int(q);
S:=FloatToStr(U);
Pala:=StrToInt(S);
end;

procedure TForm1.Shep();
var Nc,Nl,i,z,Amost,k,r,g: integer;
 SHxa,SHya,SHxb,SHyb,ArHxa,ArHya,ArHxb,ArHyb,
 AHxa,AHya,AHxb,AHyb, ArA,ArB,SA,SB,AA,AB,
 YY,XX,YY1,XX1,YY2,XX2, Ar,Soma,AreiaG,Silte,Argila:Extended;
 ArXa,ArYa,ArXb,ArYb,SXa,SYa,SXb,SYb,AXa,AYa,AXb,AYb:LongInt;
 ATabSheet: TMyTabSheet;
 Bitmap:TBitmap;
 RB:Boolean;
 CorLabel,CorPonto:TColor;
 Clw:Array[0..255] of Extended;
 VBw:Array[0..255] of Extended;
 VFw:Array[0..255] of Extended;
begin
try
RB:=RBShep.Checked;

Bitmap:=TBitmap.Create;
Bitmap.Height:=376;
Bitmap.Width:=720;

k := PageControl1.PageCount + 1;
ATabSheet := TMyTabSheet.Create(PageControl1);
ATabSheet.Parent := PageControl1;

if RB then
 ATabSheet.Caption:='Shepard '+IntToStr(k) else
  ATabSheet.Caption:='Pejrup '+IntToStr(k);
ATabSheet.sImage := TImage.Create(ATabSheet);
ATabSheet.sImage.Parent:=ATabSheet;
ATabSheet.sImage.Align:=AlClient;
ATabSheet.sImage.Picture.Assign(Bitmap);

CorLabel:=ColorBox1.Selected;
CorPonto:=ColorBox2.Selected;
Screen.Cursor:=crHourGlass;
with ATabSheet.sImage.Canvas do begin
 Brush.Color:=clWhite;
 FillRect(ATabSheet.sImage.BoundsRect);//Modo novo
 //FloodFill(1, 1, clgreen, fsBorder); //Modo velho
 Clear; //Precisa
 Pen.Color:=clBlack;
 Pen.Width:=2;
 MoveTo(200,34);LineTo(375,337);LineTo(24,337);LineTo(200,34);Pen.Width:=1;
 if RB then begin
  Caption:='Diagrama de Shepard';
  MoveTo(155,112);LineTo(244,112);MoveTo(112,186);LineTo(200,239);LineTo(288,186);
  MoveTo(200,112);LineTo(200,186);LineTo(156,264);LineTo(245,264);LineTo(200,186);
  MoveTo(200,239);LineTo(200,337);MoveTo(68,262);LineTo(112,337);MoveTo(288,337);
  LineTo(332,262);MoveTo(90,300);LineTo(156,264);MoveTo(310,300);LineTo(245,264);
 end else begin
  Caption:='Diagrama de Pejrup';
  MoveTo(338,274);LineTo(24,336);MoveTo(288,186);LineTo(24,336);MoveTo(240,104);LineTo(24,336);
 end;
 Font.Name:='Arial';
 Font.Size:=8;
 Font.Color:=clBlack;
 Font.Style:=[fsBold];
 TextOut(180,12,'Argila');
 TextOut(378,321,'Silte');
 TextOut(3,338,'Areia');
 Font.Style:=[];
 TextOut(1,316,'100%');TextOut(355,340,'100%');TextOut(206,36,'100%');
 TextOut(286,175,'50%');TextOut(87,175,'50%');TextOut(194,340,'50%');
 if RB then begin
  TextOut(129,103,'25%');TextOut(243,103,'75%');TextOut(43,250,'75%');TextOut(330,252,'25%');
  TextOut(108,340,'25%');TextOut(280,340,'75%');TextOut(198,76,'1');TextOut(166,156,'2');
  TextOut(230,156,'3');TextOut(198,212,'4');TextOut(129,235,'5');TextOut(270,235,'8');
  TextOut(182,240,'6');TextOut(214,240,'7');TextOut(64,306,'9');TextOut(160,306,'10');
  TextOut(234,306,'11');TextOut(332,306,'12');
 end else begin
  TextOut(338,259,'20%');TextOut(242,94,'80%');MoveTo(348,337);LineTo(186,54);TextOut(318,339,'10%');
  TextOut(154,45,'10%');MoveTo(52,337);LineTo(40,309);TextOut(63,339,'90%');TextOut(18,293,'90%');
  MoveTo(200,337);LineTo(112,186);Font.Size:=12;TextOut(238,62,'I');TextOut(276,134,'II');
  TextOut(326,226,'III');TextOut(362,290,'IV');TextOut(270,344,'C');TextOut(122,344,'B');TextOut(46,344,'A');
 end;
 if RB then begin
  Rectangle(416,11,700,350);
  Font.Size:=8;Font.Color:=CorLabel;TextOut(460,15,'   CONVENÇÕES');Font.Color:=clBlack;
  TextOut(420,38,'1 - Argila ou argilito');TextOut(420,58,'2 - Argila Arenosa');
  TextOut(420,78,'3 - Argila síltica');TextOut(420,98,'4 - Argila siltico-arenosa');
  TextOut(420,118,'5 - Areia argilosa');TextOut(420,138,'6 - Areia síltico-argilosa');
  TextOut(420,158,'7 - Silte argilo-arenoso');TextOut(420,178,'8 - Silte argiloso');
  TextOut(420,198,'9 - Areia ou arenito');TextOut(420,218,'10 - Areia síltica');
  TextOut(420,238,'11 - Silte arenoso');TextOut(420,258,'12 - Silte ou siltito');
  Font.Color:=CorLabel;TextOut(460,278,'   LEGENDAS');Font.Color:=clBlack;
  TextOut(445,298,'- Fração de grânulos < 3%');TextOut(445,320,'- Fração de grânulos > 3%');
  Brush.Color := CorPonto;Pen.Color:=CorPonto;
  Ellipse(424,302,436,314);
  Polygon([Point(420, 333), Point(430, 321), Point(440, 333),Point(420, 333)]);
 end else begin
  Rectangle(360,11,620,220);Font.Size:=8;Font.Color:=CorLabel;TextOut(390,18,'   CONVENÇÕES');
  Font.Color:=clBlack;TextOut(370,42,'I - Hidrodinâmica baixa');
  TextOut(370,64,'II - Hidrodinâmica moderada');TextOut(370,86,'III - Hidrodinâmica alta');
  TextOut(370,108,'IV - Hidrodinâmica muito alta');

  Font.Color:=CorLabel;
  TextOut(390,138,'   LEGENDAS');
  Font.Color:=clBlack;
  TextOut(395,162,'- Fração de grânulos < 3%');
  TextOut(395,186,'- Fração de grânulos > 3%');
  Brush.Color := CorPonto;
  Pen.Color:=CorPonto;
  Ellipse(370,165,382,177);
  Polygon([Point(368-5, 157+48), Point(375, 150+42), Point(382+5, 157+48),Point(368-5, 157+48)]);
 end;
end;

for i:=2 to 255 do if Grid.Worksheet.ReadAsText(1,i)='' then Break;
Nc:=i-1;
for i:=2 to 16834 do if (Grid.Worksheet.ReadAsText(i,1))='' then Break;
Nl:=i-1;

Amost:=0;
for z:=0 to ListBox1.Items.Count-1 do begin
 if ListBox1.Selected[z] then Amost:=z+2 else continue;
 r:=Nc;
 for g:=1 to r do begin
  Clw[g]:=StrToFloat(Grid.Cells[g+1,1]); //Vetor de classes de phi
  Vbw[g]:=StrToFloat(Grid.Cells[g+1,Amost]); //Vetor de pesos
 end;

 Soma:=0;
 for I:=1 to r do Soma:=Soma + Vbw[I];
 for I:=1 to r do VFw[I]:=(Vbw[I]*100)/Soma;
 AreiaG:=0;
 Silte:=0;
 Argila:=0;
 for i:=1 to r do begin
  if (Clw[i]>-1) and (Clw[i]<=4) then AreiaG:=AreiaG+VFw[i];
  if (Clw[i]>4) and (Clw[i]<=8) then Silte:=Silte+VFw[i];
  if Clw[i]>8 then Argila:=Argila+VFw[i];
 end;
 with ATabSheet.sImage.canvas do begin
  ArHxa:=200-1.76*AreiaG;
  ArHya:=34+3.03*AreiaG;
  ArXa:=Pala(ArHxa);
  ArYa:=Pala(ArHya);
  ArHxb:=375-3.51*AreiaG;
  ArXb:=Pala(ArHxb);
  ArYb:=337;
  ArB:=(((2*((ArXa*ArYa)+(ArXb*ArYb))) - ((ArXa+ArXb)*(ArYa+ArYb)))+0.1)/
     ((2*((ArXa*ArXa)+(ArXb*ArXb)) - ((ArXa+ArXb)*(ArXa+ArXb)))+0.1);{Sem isso dá umas merdas}
  ArA:=((ArYa+ArYb)/2)-(ArB*((ArXa+ArXb)/2));
  SHxa:=200+1.75*Silte;
  SHya:=34+3.03*Silte;
  SXa:=Pala(SHxa);
  SYa:=Pala(SHya);
  SHxb:=24+3.51*Silte;
  SXb:=Pala(SHxb);
  SYb:=337;
  SB:=((2*((SXa*SYa)+(SXb*SYb))) - ((SXa+SXb)*(SYa+SYb)))/
     (2*((SXa*SXa)+(SXb*SXb)) - ((SXa+SXb)*(SXa+SXb)));
  SA:=((SYa+SYb)/2)-(SB*((SXa+SXb)/2));
  XX:=(SA-ArA)/(ArB-SB);
  YY:=ArA+ArB*XX;
  AHxa:=24+1.76*Argila;
  AHya:=337-3.03*Argila;
  AXa:=Pala(AHxa);
  AYa:=Pala(AHya);
  AHxb:=375-1.75*Argila;
  AHyb:=337-3.03*Argila;
  AXb:=Pala(AHxb);
  AYb:=Pala(AHyb);
  AB:=((2*((AXa*AYa)+(AXb*AYb))) - ((AXa+AXb)*(AYa+AYb)))/
     (2*((AXa*AXa)+(AXb*AXb)) - ((AXa+AXb)*(AXa+AXb)));
  AA:=((AYa+AYb)/2)-(AB*((AXa+AXb)/2));
  XX1:=(AA-ArA)/(ArB-AB);
  YY1:=ArA+ArB*XX1;
  XX2:=(SA-AA)/(AB-SB);
  YY2:=AA+AB*XX2;

  Brush.Color := CorPonto;
  Pen.Color:=CorPonto;
  Polygon([Point(Pala(XX), Pala(YY)),Point(Pala(XX1), Pala(YY1)),Point(Pala(XX2), Pala(YY2)),Point(Pala(XX), Pala(YY))]);
  Ellipse(Pala(XX)-SpinEdit1.Value,Pala(YY)+Pala(Abs(YY1-YY)/2)-SpinEdit1.Value,Pala(XX)+2*SpinEdit1.Value,Pala(YY)+Pala(Abs(YY1-YY)/2)+2*SpinEdit1.Value);

  Brush.Color := clWhite;
  Pen.Color:=clBlack;
  Font.Name:='Arial';
  Font.Size:=7;
  Font.Color:=CorLabel;
  if CBLabel.Checked then TextOut(Pala(XX)-14,Pala(YY)-14,Grid.Cells[1,Amost]);
 end;
end;

PageControl1.ActivePageIndex:=PageControl1.PageCount-1;
IsMenu3Popup:=False;
Bitmap.Free;
Screen.Cursor:=crDefault;

except
on exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar a apresentação do Diagrama de Shepard.', mtError, [mbOk], 0);
end;
end;

end;

procedure TForm1.Analisar();
var Nc,Nl,Amost,j,P,I,r,z,g,C:Integer;
S,Sc,Sd:String;
W:Boolean;
Clz:Array[0..255] of Extended;
VBz:Array[0..255] of Extended;
Vz:array[0..255] of Extended;
VFz:array[0..255] of Extended;
Pz:array[0..255] of Extended;
Mid:Array[0..255] of Extended;
Std:Array[0..255] of Extended;
Sele1,A5,A6,PAreia,PSilte,PArgila,PCascalho,P03,P05,P10,P15,P20,P16,P25,P30,P35,P45,P50,P55,P65,P70,P75,P80,P84,P85,P90,P95,P97,
PA03,PA05,PA10,PA15,PA20,PA16,PA25,PA30,PA35,PA45,PA50,PA55,PA65,PA70,PA75,PA80,PA84,PA85,PA90,PA95,PA97,
C03,C05,C10,C15,C16,C20,C25,C30,C35,C45,C50,C55,C65,C70,C75,C80,C84,C85,C90,C95,C97,
CA03,CA05,CA10,CA15,CA16,CA20,CA25,CA30,CA35,CA45,CA50,CA55,CA65,CA70,CA75,CA80,CA84,CA85,CA90,CA95,CA97,
A03,A05,A10,A15,A16,A20,A25,A30,A35,A45,A50,A55,A65,A70,A75,A80,A84,A85,A90,A95,A97,
B03,B05,B10,B15,B16,B20,B25,B30,B35,B45,B50,B55,B65,B70,B75,B80,B84,B85,B90,B95,B97,Soma:Extended;
sheetname: String;
begin
MenuItem8Click(Self); //Testa se é planilha do SysgRan e aborta sozinho

if ListBox1.SelCount=0 then begin
 Screen.Cursor:=crDefault;
 MessageDlg('Selecione pelo menos uma amostra, por favor!', mtError, [mbOk], 0);
 Exit;
end;

try
Screen.Cursor:=crHourGlass;
//Zerando o mundo
Sele1:=0;Amost:=0;
PAreia:=0;PSilte:=0;PArgila:=0;PCascalho:=0;P03:=0;P05:=0;P10:=0;P15:=0;P20:=0;P16:=0;P25:=0;P30:=0;P35:=0;P45:=0;P50:=0;P55:=0;P65:=0;P70:=0;P75:=0;
P80:=0;P84:=0;P85:=0;P90:=0;P95:=0;P97:=0;PA03:=0;PA05:=0;PA10:=0;PA15:=0;PA20:=0;PA16:=0;PA25:=0;PA30:=0;PA35:=0;PA45:=0;PA50:=0;PA55:=0;PA65:=0;PA70:=0;
PA75:=0;PA80:=0;
PA84:=0;PA85:=0;PA90:=0;PA95:=0;PA97:=0;C03:=0;C05:=0;C10:=0;C15:=0;C16:=0;C20:=0;C25:=0;C30:=0;C35:=0;C45:=0;C50:=0;C55:=0;C65:=0;C70:=0;C75:=0;C80:=0;
C84:=0;C85:=0;C90:=0;C95:=0;C97:=0;CA03:=0;CA05:=0;CA10:=0;CA15:=0;CA16:=0;CA20:=0;CA25:=0;CA30:=0;CA35:=0;CA45:=0;CA50:=0;CA55:=0;CA65:=0;CA70:=0;
CA75:=0;CA80:=0;CA84:=0;CA85:=0;CA90:=0;CA95:=0;CA97:=0;A03:=0;A05:=0;A10:=0;A15:=0;A16:=0;A20:=0;A25:=0;A30:=0;A35:=0;A45:=0;A50:=0;A55:=0;A65:=0;
A70:=0;A75:=0;A80:=0;A84:=0;A85:=0;A90:=0;A95:=0;A97:=0;B03:=0;B05:=0;B10:=0;B15:=0;B16:=0;B20:=0;B25:=0;B30:=0;B35:=0;B45:=0;B50:=0;B55:=0;B65:=0;
B70:=0;B75:=0;B80:=0;B84:=0;B85:=0;B90:=0;B95:=0;B97:=0;

for i:=0 to 255 do begin
 Clz[i]:=0;
 VBz[i]:=0;
 Vz[i]:=0;
 VFz[i]:=0;
 Pz[i]:=0;
end;
z:=0;C:=0;Nc:=0;Nl:=0;

for i:=2 to 255 do if Grid.Worksheet.ReadAsText(1,i)='' then Break;
Nc:=i-1;
for i:=2 to 16834 do if (Grid.Worksheet.ReadAsText(i,1))='' then Break;
Nl:=i-1;

P:=0;

i := sWorkbookSource2.Workbook.GetWorksheetCount;
if (i=1) and (GridR.Cells[1,2]=null) then
 sWorkbookSource2.Worksheet.Name := Format(ComboBox1.Text+' %d', [i])
else
sWorkbookSource2.Workbook.AddWorksheet(Format(ComboBox1.Text+' %d', [i+1]));

for z:=0 to ListBox1.Items.Count-1 do begin
 if ListBox1.Selected[z] then begin
  P:=P+1; //P=Número da amostra (1,2,3,4,etc)
  S:=ListBox1.Items[z]; //S=Nome da amostra (A,B,C,D,etc)
  Amost:=z+2 //Número da amostra em relação ao ListBox (14,15,16,17,etc;
 end else continue;
MudouGridR:=True;
 r:=Nc; //Número de classes de phi (17)
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(Grid.Cells[g+1,1]); //Vetor de classes de phi
  Vbz[g]:=StrToFloat(Grid.Cells[g+1,Amost]); //Vetor de pesos
 end;
 Soma:=0; //Soma do vetor Vbz
 for i:=1 to r do Soma:=Soma + Vbz[i];
 Mid[1]:=Clz[1]-((Clz[2]-Clz[1])/2); //Ponto médio das classes
 for i:=2 to r do begin Mid[i]:=Clz[i]-((Clz[i]-Clz[i-1])/2); end;
 for i:=1 to r do VFz[i]:=(Vbz[i]*100)/Soma; //Vetor das proporções do peso
 if ComboBox1.Text='Medida dos momentos' then begin
   if Checkbox2.Checked=False then begin
   GridR.Cells[1,P+1]:=S;
   Media:=0;
   for I:=1 to r do begin Pz[i]:=Mid[i]*VBz[i];end;
   for I:=1 to r do Media:=Media+Pz[i];
   Media:=Media/Soma;
   GridR.Cells[2,P+1]:=StrToFloat(FloatToStrf((Media),ffGeneral,4,18));
   for I:=1 to r do Std[i]:=Mid[i]-Media;
   Sele:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i]));
   for I:=1 to r do Sele:=Sele+Pz[i];
   Sele:=Sele/(Soma-1);
   Sele1:=SQRT(Sele);
   GridR.Cells[3,P+1]:=StrToFloat(FloatToStrf(Sele1,ffGeneral,4,18));
   Ass:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*Std[i]);
   for I:=1 to r do Ass:=Ass+Pz[i];
   Ass:=Ass/(Soma-1);
   Ass:=Ass/(Exp(1.5*Ln(Sele)));
   GridR.Cells[4,P+1]:=StrToFloat(FloatToStrf(Ass,ffGeneral,4,18));
   Curt:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do Curt:=Curt+Pz[i];
   Curt:=Curt/(Soma-1);
   Curt:=Curt/(Exp(2*Ln(Sele)));
   GridR.Cells[5,P+1]:=StrToFloat(FloatToStrf(Curt,ffGeneral,4,18));
   A5:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*Std[i]);
   for I:=1 to r do A5:=A5+Pz[i];
   A5:=A5/(Soma-1);
   GridR.Cells[6,P+1]:=StrToFloat(FloatToStrf(A5/(Exp(2.5*Ln(Sele))),ffGeneral,4,18));
   A6:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do A6:=A6+Pz[i];
   A6:=A6/(Soma-1);
   GridR.Cells[7,P+1]:=StrToFloat(FloatToStrf(A6/(Exp(3*Ln(Sele))),ffGeneral,4,18));
   GridR.Cells[2,1]:='Média';
   GridR.Cells[3,1]:='Seleção';
   GridR.Cells[4,1]:='Assimetria';
   GridR.Cells[5,1]:='Curtose';
   GridR.Cells[6,1]:='Quinto Mom.';
   GridR.Cells[7,1]:='Sexto Mom.';
  end else begin
   GridR.Cells[1,P+1]:=S;
   Media:=0;
   for I:=1 to r do Pz[i]:=Mid[i]*VBz[i];
   for I:=1 to r do Media:=Media+Pz[i];
   Media:=Media/Soma;
   GridR.Cells[2,P+1]:=StrToFloat(FloatToStrf((Media),ffGeneral,4,18));
   Areia;
   GridR.Cells[3,P+1]:=Med;
   for I:=1 to r do Std[i]:=Mid[i]-Media;
   Sele:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i]));
   for I:=1 to r do Sele:=Sele+Pz[i];
   Sele:=Sele/(Soma-1);
   Sele1:=SQRT(Sele);
   GridR.Cells[4,P+1]:=StrToFloat(FloatToStrf(Sele1,ffGeneral,4,18));
   Selecao;
   GridR.Cells[5,P+1]:=Se;
   Ass:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*Std[i]);
   for I:=1 to r do Ass:=Ass+Pz[i];
   Ass:=Ass/(Soma-1);
   Ass:=Ass/(Exp(1.5*Ln(Sele)));
   GridR.Cells[6,P+1]:=StrToFloat(FloatToStrf(Ass,ffGeneral,4,18));
   if Ass<=-1.3 then Assi:='Muito negativa';
   if (Ass>-1.3) and (Ass<=-0.43) then Assi:='Negativa';
   if (Ass>-0.43) and (Ass<=0.43) then Assi:='Aproximadamente simétrica';
   if (Ass>0.43) and (Ass<=1.3) then Assi:='Positiva';
   if Ass>1.3 then Assi:='Muito positiva';
   GridR.Cells[7,P+1]:=Assi;
   Curt:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do Curt:=Curt+Pz[i];
   Curt:=Curt/(Soma-1);
   Curt:=Curt/(Exp(2*Ln(Sele)));
   GridR.Cells[8,P+1]:=StrToFloat(FloatToStrf(Curt,ffGeneral,4,18));
   if Curt<=1.70 then Cu:='Muito platicúrtica';
   if (Curt>1.70) and (Curt<=2.55) then Cu:='Platicúrtica';
   if (Curt>2.55) and (Curt<=3.70) then Cu:='Mesocúrtica';
   if (Curt>3.70) and (Curt<=7.40) then Cu:='Leptocúrtica';
   if (Curt>7.40) and (Curt<=15) then Cu:='Muito leptocúrtica';
   if Curt>15 then Cu:='Extremamente leptocúrtica';
   GridR.Cells[9,P+1]:=Cu;
   A5:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*Std[i]);
   for I:=1 to r do A5:=A5+Pz[i];
   A5:=A5/(Soma-1);
   GridR.Cells[10,P+1]:=StrToFloat(FloatToStrf(A5/(Exp(2.5*Ln(Sele))),ffGeneral,4,18));
   A6:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do A6:=A6+Pz[i];
   A6:=A6/(Soma-1);
   GridR.Cells[11,P+1]:=StrToFloat(FloatToStrf(A6/(Exp(3*Ln(Sele))),ffGeneral,4,18));
   GridR.Cells[2,1]:='Média';
   GridR.Cells[3,1]:='Classificação';
   GridR.Cells[4,1]:='Seleção';
   GridR.Cells[5,1]:='Classificação';
   GridR.Cells[6,1]:='Assimetria';
   GridR.Cells[7,1]:='Classificação';
   GridR.Cells[8,1]:='Curtose';
   GridR.Cells[9,1]:='Classificação';
   GridR.Cells[10,1]:='Quinto Mom.';
   GridR.Cells[11,1]:='Sexto Mom.';
  end;

 end else begin //Fim das medidas de momento

  Pz[1]:=VFz[1];
  for i:=2 to r do Pz[i]:=Pz[i-1]+VFz[i];

  for I:=1 to r do begin if Pz[I]>=3 then begin P03:=Pz[I]; PA03:=Pz[I-1]; C03:=Clz[I]; CA03:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=5 then begin P05:=Pz[I]; PA05:=Pz[I-1]; C05:=Clz[I]; CA05:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=10 then begin P10:=Pz[I];PA10:=Pz[I-1]; C10:=Clz[I]; CA10:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=15 then begin P15:=Pz[I];PA15:=Pz[I-1]; C15:=Clz[I]; CA15:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=16 then begin P16:=Pz[I];PA16:=Pz[I-1]; C16:=Clz[I]; CA16:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=20 then begin P20:=Pz[I];PA20:=Pz[I-1]; C20:=Clz[I]; CA20:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=25 then begin P25:=Pz[I];PA25:=Pz[I-1]; C25:=Clz[I]; CA25:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=30 then begin P30:=Pz[I];PA30:=Pz[I-1]; C30:=Clz[I]; CA30:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=35 then begin P35:=Pz[I];PA35:=Pz[I-1]; C35:=Clz[I]; CA35:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=45 then begin P45:=Pz[I];PA45:=Pz[I-1]; C45:=Clz[I]; CA45:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=50 then begin P50:=Pz[I];PA50:=Pz[I-1]; C50:=Clz[I]; CA50:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=55 then begin P55:=Pz[I];PA55:=Pz[I-1]; C55:=Clz[I]; CA55:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=65 then begin P65:=Pz[I];PA65:=Pz[I-1]; C65:=Clz[I]; CA65:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=70 then begin P70:=Pz[I];PA70:=Pz[I-1]; C70:=Clz[I]; CA70:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=75 then begin P75:=Pz[I];PA75:=Pz[I-1]; C75:=Clz[I]; CA75:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=80 then begin P80:=Pz[I];PA80:=Pz[I-1]; C80:=Clz[I]; CA80:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=84 then begin P84:=Pz[I];PA84:=Pz[I-1]; C84:=Clz[I]; CA84:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=85 then begin P85:=Pz[I];PA85:=Pz[I-1]; C85:=Clz[I]; CA85:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=90 then begin P90:=Pz[I];PA90:=Pz[I-1]; C90:=Clz[I]; CA90:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=95 then begin P95:=Pz[I];PA95:=Pz[I-1]; C95:=Clz[I]; CA95:=Clz[I-1]; Break; end;end;
  for I:=1 to r do begin if Pz[I]>=97 then begin P97:=Pz[I];PA97:=Pz[I-1]; C97:=Clz[I]; CA97:=Clz[I-1]; Break; end;end;

  B03:=((2*((P03*C03)+(PA03*CA03))) - ((P03+PA03)*(C03+CA03)))/(2*((C03*C03)+(CA03*CA03)) - ((C03+CA03)*(C03+CA03)));
  A03:=((P03+PA03)/2)-(B03*((C03+CA03)/2));
  PT03:=(3-A03)/B03;
  B05:=((2*((P05*C05)+(PA05*CA05))) - ((P05+PA05)*(C05+CA05)))/(2*((C05*C05)+(CA05*CA05)) - ((C05+CA05)*(C05+CA05)));
  A05:=((P05+PA05)/2)-(B05*((C05+CA05)/2));
  PT05:=(5-A05)/B05;
  B10:=((2*((P10*C10)+(PA10*CA10))) - ((P10+PA10)*(C10+CA10)))/(2*((C10*C10)+(CA10*CA10)) - ((C10+CA10)*(C10+CA10)));
  A10:=((P10+PA10)/2)-(B10*((C10+CA10)/2));
  PT10:=(10-A10)/B10;
  B15:=((2*((P15*C15)+(PA15*CA15))) - ((P15+PA15)*(C15+CA15)))/(2*((C15*C15)+(CA15*CA15)) - ((C15+CA15)*(C15+CA15)));
  A15:=((P15+PA15)/2)-(B15*((C15+CA15)/2));
  PT15:=(15-A15)/B15;
  B16:=((2*((P16*C16)+(PA16*CA16))) - ((P16+PA16)*(C16+CA16)))/(2*((C16*C16)+(CA16*CA16)) - ((C16+CA16)*(C16+CA16)));
  A16:=((P16+PA16)/2)-(B16*((C16+CA16)/2));
  PT16:=(16-A16)/B16;
  B20:=((2*((P20*C20)+(PA20*CA20))) - ((P20+PA20)*(C20+CA20)))/(2*((C20*C20)+(CA20*CA20)) - ((C20+CA20)*(C20+CA20)));
  A20:=((P20+PA20)/2)-(B20*((C20+CA20)/2));
  PT20:=(20-A20)/B20;
  B25:=((2*((P25*C25)+(PA25*CA25))) - ((P25+PA25)*(C25+CA25)))/(2*((C25*C25)+(CA25*CA25)) - ((C25+CA25)*(C25+CA25)));
  A25:=((P25+PA25)/2)-(B25*((C25+CA25)/2));
  PT25:=(25-A25)/B25;
  B30:=((2*((P30*C30)+(PA30*CA30))) - ((P30+PA30)*(C30+CA30)))/(2*((C30*C30)+(CA30*CA30)) - ((C30+CA30)*(C30+CA30)));
  A30:=((P30+PA30)/2)-(B30*((C30+CA30)/2));
  PT30:=(30-A30)/B30;
  B35:=((2*((P35*C35)+(PA35*CA35))) - ((P35+PA35)*(C35+CA35)))/(2*((C35*C35)+(CA35*CA35)) - ((C35+CA35)*(C35+CA35)));
  A35:=((P35+PA35)/2)-(B35*((C35+CA35)/2));
  PT35:=(35-A35)/B35;
  B45:=((2*((P45*C45)+(PA45*CA45))) - ((P45+PA45)*(C45+CA45)))/(2*((C45*C45)+(CA45*CA45)) - ((C45+CA45)*(C45+CA45)));
  A45:=((P45+PA45)/2)-(B45*((C45+CA45)/2));
  PT45:=(45-A45)/B45;
  B50:=((2*((P50*C50)+(PA50*CA50))) - ((P50+PA50)*(C50+CA50)))/(2*((C50*C50)+(CA50*CA50)) - ((C50+CA50)*(C50+CA50)));
  A50:=((P50+PA50)/2)-(B50*((C50+CA50)/2));
  PT50:=(50-A50)/B50;
  B55:=((2*((P55*C55)+(PA55*CA55))) - ((P55+PA55)*(C55+CA55)))/(2*((C55*C55)+(CA55*CA55)) - ((C55+CA55)*(C55+CA55)));
  A55:=((P55+PA55)/2)-(B55*((C55+CA55)/2));
  PT55:=(55-A55)/B55;
  B65:=((2*((P65*C65)+(PA65*CA65))) - ((P65+PA65)*(C65+CA65)))/(2*((C65*C65)+(CA65*CA65)) - ((C65+CA65)*(C65+CA65)));
  A65:=((P65+PA65)/2)-(B65*((C65+CA65)/2));
  PT65:=(65-A65)/B65;
  B70:=((2*((P70*C70)+(PA70*CA70))) - ((P70+PA70)*(C70+CA70)))/(2*((C70*C70)+(CA70*CA70)) - ((C70+CA70)*(C70+CA70)));
  A70:=((P70+PA70)/2)-(B70*((C70+CA70)/2));
  PT70:=(70-A70)/B70;
  B75:=((2*((P75*C75)+(PA75*CA75))) - ((P75+PA75)*(C75+CA75)))/(2*((C75*C75)+(CA75*CA75)) - ((C75+CA75)*(C75+CA75)));
  A75:=((P75+PA75)/2)-(B75*((C75+CA75)/2));
  PT75:=(75-A75)/B75;
  B80:=((2*((P80*C80)+(PA80*CA80))) - ((P80+PA80)*(C80+CA80)))/(2*((C80*C80)+(CA80*CA80)) - ((C80+CA80)*(C80+CA80)));
  A80:=((P80+PA80)/2)-(B80*((C80+CA80)/2));
  PT80:=(80-A80)/B80;
  B84:=((2*((P84*C84)+(PA84*CA84))) - ((P84+PA84)*(C84+CA84)))/(2*((C84*C84)+(CA84*CA84)) - ((C84+CA84)*(C84+CA84)));
  A84:=((P84+PA84)/2)-(B84*((C84+CA84)/2));
  PT84:=(84-A84)/B84;
  B85:=((2*((P85*C85)+(PA85*CA85))) - ((P85+PA85)*(C85+CA85)))/(2*((C85*C85)+(CA85*CA85)) - ((C85+CA85)*(C85+CA85)));
  A85:=((P85+PA85)/2)-(B85*((C85+CA85)/2));
  PT85:=(85-A85)/B85;
  B90:=((2*((P90*C90)+(PA90*CA90))) - ((P90+PA90)*(C90+CA90)))/(2*((C90*C90)+(CA90*CA90)) - ((C90+CA90)*(C90+CA90)));
  A90:=((P90+PA90)/2)-(B90*((C90+CA90)/2));
  PT90:=(90-A90)/B90;
  B95:=((2*((P95*C95)+(PA95*CA95))) - ((P95+PA95)*(C95+CA95)))/(2*((C95*C95)+(CA95*CA95)) - ((C95+CA95)*(C95+CA95)));
  A95:=((P95+PA95)/2)-(B95*((C95+CA95)/2));
  PT95:=(95-A95)/B95;
  B97:=((2*((P97*C97)+(PA97*CA97))) - ((P97+PA97)*(C97+CA97)))/(2*((C97*C97)+(CA97*CA97)) - ((C97+CA97)*(C97+CA97)));
  A97:=((P97+PA97)/2)-(B97*((C97+CA97)/2));
  PT97:=(97-A97)/B97;

  if Combobox1.Text='Folk & Ward' then Folk;
  if Combobox1.Text='McCammon (a)' then Mca;
  if Combobox1.Text='McCammon (b)' then Mcb;
  if Combobox1.Text='Trask' then Trask;
  if Combobox1.Text='Otto & Inman' then Otto;

  PAreia:=0;
  PCascalho:=0;
  PSilte:=0;
  PArgila:=0;
  for i:=1 to r do begin
   if (Clz[i]>-1) and (Clz[i]<=4) then PAreia:=PAreia+VFz[i];
   if (Clz[i]>4) and (Clz[i]<=8) then PSilte:=PSilte+VFz[i];
   if Clz[i]>8 then PArgila:=PArgila+VFz[i];
  end;
  PCascalho:=100-(PAreia+PSilte+PArgila);
  if PCascalho<0.001 then PCascalho:=0;
//  if EntryRC[Amost,Nc+1]<>'' then W:=True else W:=False;


  if RadioButton5.Checked then begin
   GridR.Cells[1,P*4-2]:=S+'(Peso)';
   GridR.Cells[1,P*4-1]:=S+'(Porc.)';
   GridR.Cells[1,P*4]:=S+'(Porc.Acum.)';
   for i:=1 to r do begin
    GridR.Cells[i+1,1]:=StrToFloat(FloatToStrf(Clz[i],ffGeneral,4,18));
    GridR.Cells[i+1,P*4-2]:=StrToFloat(FloatToStrf(VBz[i],ffGeneral,4,18));
    GridR.Cells[i+1,P*4-1]:=StrToFloat(FloatToStrf(VFz[i],ffGeneral,4,18));
    GridR.Cells[i+1,P*4]:=StrToFloat(FloatToStrf(Pz[i],ffGeneral,4,18));
   end;
  end else begin
   GridR.Cells[1,P+1]:=S;
   if (CheckBox2.Checked=false) and  (CheckBox3.Checked=false) then begin
    GridR.Cells[2,1]:='Média';
    GridR.Cells[3,1]:='Mediana';
    GridR.Cells[4,1]:='Seleção';
    GridR.Cells[5,1]:='Assimetria';
    GridR.Cells[6,1]:='Curtose';
    GridR.Cells[7,1]:='% Cascalho';
    GridR.Cells[8,1]:='% Areia';
    GridR.Cells[9,1]:='% Silte';
    GridR.Cells[10,1]:='% Argila';
    GridR.Cells[2,P+1]:=StrToFloat(FloatToStrf(Media,ffGeneral,4,18));
    GridR.Cells[3,P+1]:=StrToFloat(FloatToStrf(PT50,ffGeneral,4,18));
    GridR.Cells[4,P+1]:=StrToFloat(FloatToStrf(Sele,ffGeneral,4,18));
    GridR.Cells[5,P+1]:=StrToFloat(FloatToStrf(Ass,ffGeneral,4,18));
    GridR.Cells[6,P+1]:=StrToFloat(FloatToStrf(Curt,ffGeneral,4,18));
    GridR.Cells[7,P+1]:=StrToFloat(FloatToStrf(PCascalho,ffGeneral,4,18));
    GridR.Cells[8,P+1]:=StrToFloat(FloatToStrf(PAreia,ffGeneral,4,18));
    GridR.Cells[9,P+1]:=StrToFloat(FloatToStrf(PSilte,ffGeneral,4,18));
    GridR.Cells[10,P+1]:=StrToFloat(FloatToStrf(PArgila,ffGeneral,4,18));
   end;
   if (CheckBox2.Checked=false) and (CheckBox3.Checked=true) then begin
    GridR.Cells[2,1]:='Média';
    GridR.Cells[3,1]:='Mediana';
    GridR.Cells[4,1]:='Seleção';
    GridR.Cells[5,1]:='Assimetria';
    GridR.Cells[6,1]:='Curtose';
    GridR.Cells[7,1]:='% Cascalho';
    GridR.Cells[8,1]:='% Areia';
    GridR.Cells[9,1]:='% Silte';
    GridR.Cells[10,1]:='% Argila';
    GridR.Cells[2,P+1]:=StrToFloat(FloatToStrf(Media,ffGeneral,4,18));
    GridR.Cells[3,P+1]:=StrToFloat(FloatToStrf(PT50,ffGeneral,4,18));
    GridR.Cells[4,P+1]:=StrToFloat(FloatToStrf(Sele,ffGeneral,4,18));
    GridR.Cells[5,P+1]:=StrToFloat(FloatToStrf(Ass,ffGeneral,4,18));
    GridR.Cells[6,P+1]:=StrToFloat(FloatToStrf(Curt,ffGeneral,4,18));
    GridR.Cells[7,P+1]:=StrToFloat(FloatToStrf(PCascalho,ffGeneral,4,18));
    GridR.Cells[8,P+1]:=StrToFloat(FloatToStrf(PAreia,ffGeneral,4,18));
    GridR.Cells[9,P+1]:=StrToFloat(FloatToStrf(PSilte,ffGeneral,4,18));
    GridR.Cells[10,P+1]:=StrToFloat(FloatToStrf(PArgila,ffGeneral,4,18));
    GridR.Cells[11,1]:='Phi-03';
    GridR.Cells[12,1]:='Phi-05';
    GridR.Cells[13,1]:='Phi-10';
    GridR.Cells[14,1]:='Phi-15';
    GridR.Cells[15,1]:='Phi-16';
    GridR.Cells[16,1]:='Phi-20';
    GridR.Cells[17,1]:='Phi-25';
    GridR.Cells[18,1]:='Phi-30';
    GridR.Cells[19,1]:='Phi-35';
    GridR.Cells[20,1]:='Phi-45';
    GridR.Cells[21,1]:='Phi-50';
    GridR.Cells[22,1]:='Phi-55';
    GridR.Cells[23,1]:='Phi-65';
    GridR.Cells[24,1]:='Phi-70';
    GridR.Cells[25,1]:='Phi-75';
    GridR.Cells[26,1]:='Phi-80';
    GridR.Cells[27,1]:='Phi-84';
    GridR.Cells[28,1]:='Phi-85';
    GridR.Cells[29,1]:='Phi-90';
    GridR.Cells[30,1]:='Phi-95';
    GridR.Cells[31,1]:='Phi-97';
    GridR.Cells[11,P+1]:=StrToFloat(FloatToStrf(PT03,ffGeneral,4,18));
    GridR.Cells[12,P+1]:=StrToFloat(FloatToStrf(PT05,ffGeneral,4,18));
    GridR.Cells[13,P+1]:=StrToFloat(FloatToStrf(PT10,ffGeneral,4,18));
    GridR.Cells[14,P+1]:=StrToFloat(FloatToStrf(PT15,ffGeneral,4,18));
    GridR.Cells[15,P+1]:=StrToFloat(FloatToStrf(PT16,ffGeneral,4,18));
    GridR.Cells[16,P+1]:=StrToFloat(FloatToStrf(PT20,ffGeneral,4,18));
    GridR.Cells[17,P+1]:=StrToFloat(FloatToStrf(PT25,ffGeneral,4,18));
    GridR.Cells[18,P+1]:=StrToFloat(FloatToStrf(PT30,ffGeneral,4,18));
    GridR.Cells[19,P+1]:=StrToFloat(FloatToStrf(PT35,ffGeneral,4,18));
    GridR.Cells[20,P+1]:=StrToFloat(FloatToStrf(PT45,ffGeneral,4,18));
    GridR.Cells[21,P+1]:=StrToFloat(FloatToStrf(PT50,ffGeneral,4,18));
    GridR.Cells[22,P+1]:=StrToFloat(FloatToStrf(PT55,ffGeneral,4,18));
    GridR.Cells[23,P+1]:=StrToFloat(FloatToStrf(PT65,ffGeneral,4,18));
    GridR.Cells[24,P+1]:=StrToFloat(FloatToStrf(PT70,ffGeneral,4,18));
    GridR.Cells[25,P+1]:=StrToFloat(FloatToStrf(PT75,ffGeneral,4,18));
    GridR.Cells[26,P+1]:=StrToFloat(FloatToStrf(PT80,ffGeneral,4,18));
    GridR.Cells[27,P+1]:=StrToFloat(FloatToStrf(PT84,ffGeneral,4,18));
    GridR.Cells[28,P+1]:=StrToFloat(FloatToStrf(PT85,ffGeneral,4,18));
    GridR.Cells[29,P+1]:=StrToFloat(FloatToStrf(PT90,ffGeneral,4,18));
    GridR.Cells[30,P+1]:=StrToFloat(FloatToStrf(PT95,ffGeneral,4,18));
    GridR.Cells[31,P+1]:=StrToFloat(FloatToStrf(PT97,ffGeneral,4,18));
   end;
   if (CheckBox2.Checked=true) and (CheckBox3.Checked=false) then begin
    GridR.Cells[2,1]:='Média';
    GridR.Cells[3,1]:='Classificação';
    GridR.Cells[4,1]:='Mediana';
    GridR.Cells[5,1]:='Seleção';
    GridR.Cells[6,1]:='Classificação';
    GridR.Cells[7,1]:='Assimetria';
    GridR.Cells[8,1]:='Classificação';
    GridR.Cells[9,1]:='Curtose';
    GridR.Cells[10,1]:='Classificação';
    GridR.Cells[11,1]:='% Cascalho';
    GridR.Cells[12,1]:='% Areia';
    GridR.Cells[13,1]:='% Silte';
    GridR.Cells[14,1]:='% Argila';
    GridR.Cells[2,P+1]:=StrToFloat(FloatToStrf(Media,ffGeneral,4,18));
    GridR.Cells[3,P+1]:=Med;
    GridR.Cells[4,P+1]:=StrToFloat(FloatToStrf(PT50,ffGeneral,4,18));
    GridR.Cells[5,P+1]:=StrToFloat(FloatToStrf(Sele,ffGeneral,4,18));
    GridR.Cells[6,P+1]:=Se;
    GridR.Cells[7,P+1]:=StrToFloat(FloatToStrf(Ass,ffGeneral,4,18));
    GridR.Cells[8,P+1]:=Assi;
    GridR.Cells[9,P+1]:=StrToFloat(FloatToStrf(Curt,ffGeneral,4,18));
    GridR.Cells[10,P+1]:=Cu;
    GridR.Cells[11,P+1]:=StrToFloat(FloatToStrf(PCascalho,ffGeneral,4,18));
    GridR.Cells[12,P+1]:=StrToFloat(FloatToStrf(PAreia,ffGeneral,4,18));
    GridR.Cells[13,P+1]:=StrToFloat(FloatToStrf(PSilte,ffGeneral,4,18));
    GridR.Cells[14,P+1]:=StrToFloat(FloatToStrf(PArgila,ffGeneral,4,18));
   end;
   if (CheckBox2.Checked=true) and (CheckBox3.Checked=true) then begin
    GridR.Cells[2,1]:='Média';
    GridR.Cells[3,1]:='Classificação';
    GridR.Cells[4,1]:='Mediana';
    GridR.Cells[5,1]:='Seleção';
    GridR.Cells[6,1]:='Classificação';
    GridR.Cells[7,1]:='Assimetria';
    GridR.Cells[8,1]:='Classificação';
    GridR.Cells[9,1]:='Curtose';
    GridR.Cells[10,1]:='Classificação';
    GridR.Cells[11,1]:='% Cascalho';
    GridR.Cells[12,1]:='% Areia';
    GridR.Cells[13,1]:='% Silte';
    GridR.Cells[14,1]:='% Argila';
    GridR.Cells[2,P+1]:=StrToFloat(FloatToStrf(Media,ffGeneral,4,18));
    GridR.Cells[3,P+1]:=Med;
    GridR.Cells[4,P+1]:=StrToFloat(FloatToStrf(PT50,ffGeneral,4,18));
    GridR.Cells[5,P+1]:=StrToFloat(FloatToStrf(Sele,ffGeneral,4,18));
    GridR.Cells[6,P+1]:=Se;
    GridR.Cells[7,P+1]:=StrToFloat(FloatToStrf(Ass,ffGeneral,4,18));
    GridR.Cells[8,P+1]:=Assi;
    GridR.Cells[9,P+1]:=StrToFloat(FloatToStrf(Curt,ffGeneral,4,18));
    GridR.Cells[10,P+1]:=Cu;
    GridR.Cells[11,P+1]:=StrToFloat(FloatToStrf(PCascalho,ffGeneral,4,18));
    GridR.Cells[12,P+1]:=StrToFloat(FloatToStrf(PAreia,ffGeneral,4,18));
    GridR.Cells[13,P+1]:=StrToFloat(FloatToStrf(PSilte,ffGeneral,4,18));
    GridR.Cells[14,P+1]:=StrToFloat(FloatToStrf(PArgila,ffGeneral,4,18));
    GridR.Cells[15,1]:='Phi-03';
    GridR.Cells[16,1]:='Phi-05';
    GridR.Cells[17,1]:='Phi-10';
    GridR.Cells[18,1]:='Phi-15';
    GridR.Cells[19,1]:='Phi-16';
    GridR.Cells[20,1]:='Phi-20';
    GridR.Cells[21,1]:='Phi-25';
    GridR.Cells[22,1]:='Phi-30';
    GridR.Cells[23,1]:='Phi-35';
    GridR.Cells[24,1]:='Phi-45';
    GridR.Cells[25,1]:='Phi-50';
    GridR.Cells[26,1]:='Phi-55';
    GridR.Cells[27,1]:='Phi-65';
    GridR.Cells[28,1]:='Phi-70';
    GridR.Cells[29,1]:='Phi-75';
    GridR.Cells[30,1]:='Phi-80';
    GridR.Cells[31,1]:='Phi-84';
    GridR.Cells[32,1]:='Phi-85';
    GridR.Cells[33,1]:='Phi-90';
    GridR.Cells[34,1]:='Phi-95';
    GridR.Cells[35,1]:='Phi-97';
    GridR.Cells[15,P+1]:=StrToFloat(FloatToStrf(PT03,ffGeneral,4,18));
    GridR.Cells[16,P+1]:=StrToFloat(FloatToStrf(PT05,ffGeneral,4,18));
    GridR.Cells[17,P+1]:=StrToFloat(FloatToStrf(PT10,ffGeneral,4,18));
    GridR.Cells[18,P+1]:=StrToFloat(FloatToStrf(PT15,ffGeneral,4,18));
    GridR.Cells[19,P+1]:=StrToFloat(FloatToStrf(PT16,ffGeneral,4,18));
    GridR.Cells[20,P+1]:=StrToFloat(FloatToStrf(PT20,ffGeneral,4,18));
    GridR.Cells[21,P+1]:=StrToFloat(FloatToStrf(PT25,ffGeneral,4,18));
    GridR.Cells[22,P+1]:=StrToFloat(FloatToStrf(PT30,ffGeneral,4,18));
    GridR.Cells[23,P+1]:=StrToFloat(FloatToStrf(PT35,ffGeneral,4,18));
    GridR.Cells[24,P+1]:=StrToFloat(FloatToStrf(PT45,ffGeneral,4,18));
    GridR.Cells[25,P+1]:=StrToFloat(FloatToStrf(PT50,ffGeneral,4,18));
    GridR.Cells[26,P+1]:=StrToFloat(FloatToStrf(PT55,ffGeneral,4,18));
    GridR.Cells[27,P+1]:=StrToFloat(FloatToStrf(PT65,ffGeneral,4,18));
    GridR.Cells[28,P+1]:=StrToFloat(FloatToStrf(PT70,ffGeneral,4,18));
    GridR.Cells[29,P+1]:=StrToFloat(FloatToStrf(PT75,ffGeneral,4,18));
    GridR.Cells[30,P+1]:=StrToFloat(FloatToStrf(PT80,ffGeneral,4,18));
    GridR.Cells[31,P+1]:=StrToFloat(FloatToStrf(PT84,ffGeneral,4,18));
    GridR.Cells[32,P+1]:=StrToFloat(FloatToStrf(PT85,ffGeneral,4,18));
    GridR.Cells[33,P+1]:=StrToFloat(FloatToStrf(PT90,ffGeneral,4,18));
    GridR.Cells[34,P+1]:=StrToFloat(FloatToStrf(PT95,ffGeneral,4,18));
    GridR.Cells[35,P+1]:=StrToFloat(FloatToStrf(PT97,ffGeneral,4,18));
   end; //Último Begin
  end; //end else begin
 end;
 Screen.Cursor:=crDefault;
end; //Fim -> for z:=0 to ListBix1.Count-1 do

//Aciona exceção global
except //try inical
on Exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('IMPOSSÍVEL MOSTRAR OS RESULTADOS!!!'+#13+#13+'Possíveis causas:'+#13+
 ' - Esta não é uma planilha padrão do SysGran'+#13+
 ' - O valores colocados resultam em expressões matematicamente impossíveis.'+#13+
 ' - Existem valores vazios.'+#13+
 ' - O separador decimal utilizado é diferente daquele padrão da sua versão do Windows.'+#13+#13+
 'Tente novamente após revisar a planilha.',mtError, [mbOk], 0);
 end;
end;

end;

procedure TForm1.BtnNovoClick(Sender: TObject);
begin
IsMenu1Popup:=False;
ConfirmSave();
Grid.SetFocus;
Grid.NewWorkbook(100,100);
GridR.NewWorkbook(100,100);
Caption:='SysGran 4.0';
ListBox1.Clear;
Label1.Font.Style:=[];
MudouGridR:=False;
OpenDialog.FileName:='';
sWorkbookSource1.Worksheet.Name := 'Planilha1';
sWorkbookSource2.Worksheet.Name := 'Resultado1';
end;

procedure TForm1.BtnAcoesClick(Sender: TObject);
var
  lowerLeft: TPoint;
begin
  lowerLeft := Point(0, BtnAcoes.Height);
  lowerLeft := BtnAcoes.ClientToScreen(lowerLeft);
 if IsMenu1Popup=False then begin
  PopupMenu1.Popup(lowerLeft.X, lowerLeft.Y);
  IsMenu1Popup:=True;
 end else IsMenu1Popup:=False;
end;

procedure TForm1.BtnSalvar2Click(Sender: TObject);
var S: String;
begin
if PageControl1.PageCount=0 then Exit;;
S:=PageControl1.ActivePage.Caption;
if not SaveDialogF.Execute then Exit;
if (Pos('Shep',S)>0) or (Pos('Pej',S)>0) then
 (PageControl1.ActivePage as TMyTabSheet).sImage.Picture.SaveToFile(SaveDialogF.FileName)
else begin
 if SaveDialogF.FilterIndex=1 then
  (PageControl1.ActivePage as TMyTabSheet).sChart.SaveToFile(TJpegImage,SaveDialogF.FileName) else
 if SaveDialogF.FilterIndex=2 then
  (PageControl1.ActivePage as TMyTabSheet).sChart.SaveToFile(TPortableNetworkGraphic,SaveDialogF.FileName) else
 if SaveDialogF.FilterIndex=3 then
  (PageControl1.ActivePage as TMyTabSheet).sChart.SaveToBitmapFile(SaveDialogF.FileName);
 if SaveDialogF.FilterIndex=4 then
  MessageDlg('Não é possível salvar o gráfico no formato Tiff. Por favor, escolha outro formato de gráficos',mtInformation,[mbOk],0);
end;
end;

procedure TForm1.BtnSalvarResClick(Sender: TObject);
begin
IsMenu2Popup:=False;
if SaveDialogR.FileName='' then begin
 if SaveDialogR.Execute then begin
  Screen.Cursor := crHourglass;
  GridR.SaveToSpreadsheetFile(SaveDialogR.FileName);
  MudouGridR:=False;
  Screen.Cursor := crDefault;
 end;
end else begin
 Screen.Cursor := crHourglass;
 GridR.SaveToSpreadsheetFile(SaveDialogR.FileName);
 MudouGridR:=False;
 Screen.Cursor := crDefault;
end;
end;

procedure TForm1.BtnSelectAllClick(Sender: TObject);
begin
ListBox1.SelectAll;
end;

procedure TForm1.BtnSalvarClick(Sender: TObject);
begin
IsMenu1Popup:=False;
if OpenDialog.FileName='' then
 MenuItem4Click(Self) else begin //Save as
 Screen.Cursor := crHourglass;
 Grid.SaveToSpreadsheetFile(OpenDialog.FileName);
 Label1.Font.Style:=[];
 Screen.Cursor := crDefault;
end;
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
if ComboBox1.Text='Medida dos momentos' then
 CheckBox3.Enabled:=False else
  CheckBox3.Enabled:=True;
end;

procedure TForm1.CB2Change(Sender: TObject);
begin
if CB2.Text='Média' then begin
 CB3.Items.Clear;
 CB3.Items.Add('Seleção');
 CB3.Items.Add('Assimetria');
 CB3.Items.Add('Curtose');
end;
if CB2.Text='Seleção' then begin
 CB3.Items.Clear;
 CB3.Items.Add('Média');
 CB3.Items.Add('Assimetria');
 CB3.Items.Add('Curtose');
end;
if CB2.Text='Assimetria' then begin
 CB3.Items.Clear;
 CB3.Items.Add('Média');
 CB3.Items.Add('Seleção');
 CB3.Items.Add('Curtose');
end;
if CB2.Text='Curtose' then begin
 CB3.Items.Clear;
 CB3.Items.Add('Média');
 CB3.Items.Add('Seleção');
 CB3.Items.Add('Assimetria');
end;
CB3.ItemIndex:=0;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
ConfirmSave();
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
sWorkbookSource1.Worksheet.Name := 'Planilha1';
sWorkbookSource2.Worksheet.Name := 'Resultado1';
h:=1;
end;

procedure TForm1.gridClick(Sender: TObject);
begin
IsMenu1Popup:=False;
end;

procedure TForm1.MenuItem10Click(Sender: TObject);
begin
(PageControl1.ActivePage as TMyTabSheet).sChart.LeftAxis.Grid.Visible:=not
(PageControl1.ActivePage as TMyTabSheet).sChart.LeftAxis.Grid.Visible;
(PageControl1.ActivePage as TMyTabSheet).sChart.BottomAxis.Grid.Visible:=not
(PageControl1.ActivePage as TMyTabSheet).sChart.BottomAxis.Grid.Visible;
IsMenu3Popup:=False;
end;

procedure TForm1.MenuItem11Click(Sender: TObject);
begin
(PageControl1.ActivePage as TMyTabSheet).sChart.Title.Visible:=not
(PageControl1.ActivePage as TMyTabSheet).sChart.Title.Visible;
IsMenu3Popup:=False;
end;

procedure TForm1.MenuItem12Click(Sender: TObject);
begin
(PageControl1.ActivePage as TMyTabSheet).sImage.Stretch:=MenuItem12.Checked;
IsMenu3Popup:=False;
end;

procedure TForm1.MenuItem13Click(Sender: TObject);
begin
(PageControl1.ActivePage as TMyTabSheet).sChart.BottomAxis.Title.Visible:=not
(PageControl1.ActivePage as TMyTabSheet).sChart.BottomAxis.Title.Visible;
(PageControl1.ActivePage as TMyTabSheet).sChart.LeftAxis.Title.Visible:=not
(PageControl1.ActivePage as TMyTabSheet).sChart.LeftAxis.Title.Visible;
IsMenu3Popup:=False;
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
var
  sheetname: String;
  i: Integer;
begin
  IsMenu2Popup:=False;
  i := sWorkbookSource2.Workbook.GetWorksheetCount;
  repeat
    inc(i);
    sheetName := Format('Resultado %d', [i]);
  until (sWorkbookSource2.Workbook.GetWorksheetByName(sheetname) = nil);
  sWorkbookSource2.Workbook.AddWorksheet(sheetName);
end;

procedure TForm1.MenuItem2Click(Sender: TObject);
begin
IsMenu2Popup:=False;
if sWorkbookSource2.Workbook.GetWorksheetCount = 1 then
  MessageDlg('Pelo menos 1 planilha deve existir.', mtError, [mbOK], 0)
else
if MessageDlg('Deseja apagar a planilha?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  sWorkbookSource2.Workbook.RemoveWorksheet(sWorkbookSource2.Worksheet);
end;

procedure TForm1.MenuItem3Click(Sender: TObject);
var
  s: String;
begin
  IsMenu2Popup:=False;
  s := sWorkbookSource2.Worksheet.Name;
  if InputQuery('Editar nome da planilha', 'Novo nome', s) then
  begin
    if sWorkbookSource2.Workbook.ValidWorksheetName(s) then
      sWorkbookSource2.Worksheet.Name := s
    else
      MessageDlg('Nome inválido.', mtError, [mbOK], 0);
  end;
end;

procedure TForm1.MenuItem4Click(Sender: TObject);
var err: String;
begin
IsMenu1Popup:=False;
if SaveDialog.Execute then begin
 Screen.Cursor := crHourglass;
 try
  Grid.SaveToSpreadsheetFile(SaveDialog.FileName);
  finally
   Screen.Cursor := crDefault;
   err := Grid.Workbook.ErrorMsg;
   if err <> '' then MessageDlg(err, mtError, [mbOK], 0);
   Caption:='SysGran 4.0 - '+SaveDialog.FileName;
   Label1.Font.Style:=[];
  end;
 end;
end;

procedure TForm1.MenuItem5Click(Sender: TObject);
var
  sheetname: String;
  i: Integer;
begin
IsMenu1Popup:=False;
  i := sWorkbookSource1.Workbook.GetWorksheetCount;
  repeat
    inc(i);
    sheetName := Format('Planilha%d', [i]);
  until (sWorkbookSource1.Workbook.GetWorksheetByName(sheetname) = nil);
  sWorkbookSource1.Workbook.AddWorksheet(sheetName);
end;

procedure TForm1.MenuItem6Click(Sender: TObject);
begin
IsMenu1Popup:=False;
if sWorkbookSource1.Workbook.GetWorksheetCount = 1 then
  MessageDlg('Pelo menos 1 planilha deve existir.', mtError, [mbOK], 0)
else
if MessageDlg('Deseja apagar a planilha?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  sWorkbookSource1.Workbook.RemoveWorksheet(sWorkbookSource1.Worksheet);
end;

procedure TForm1.MenuItem7Click(Sender: TObject);
var
  s: String;
begin
  IsMenu1Popup:=False;
  s := sWorkbookSource1.Worksheet.Name;
  if InputQuery('Editar nome da planilha', 'Novo nome', s) then
  begin
    if sWorkbookSource1.Workbook.ValidWorksheetName(s) then
      sWorkbookSource1.Worksheet.Name := s
    else
      MessageDlg('Nome inválido.', mtError, [mbOK], 0);
  end;
end;

procedure TForm1.MenuItem8Click(Sender: TObject);
var number: Boolean;
d: Double;
begin
IsMenu1Popup:=False;
if Grid.Cells[1,1]<>null then
 if TryStrToFloat(Grid.Cells[1,1],d) then begin
 MessageDlg('Esta não é uma planilha de entrada do SysGran.'+#13#13+'Célula A1 da planilha de entrada deveria estar vazia.',mtError,[mbOk],0);
 Abort;
end;
//Tem que repetir
if Grid.Cells[1,1]<>null then
  if not (Grid.Cells[1,1]=' ') then begin
  MessageDlg('Esta não é uma planilha de entrada do SysGran.'+#13#13+'Célula A1 da planilha de entrada deveria estar vazia.',mtError,[mbOk],0);
  Abort;
 end;
if Grid.Cells[1,2]=null then begin
 MessageDlg('Esta não é uma planilha de entrada do SysGran.'+#13#13+'Célula A2 da planilha de entrada deveria conter algum valor.',mtError,[mbOk],0);
 Abort;
end;
if Grid.Cells[2,1]=null then begin
 MessageDlg('Esta não é uma planilha de entrada do SysGran.'+#13#13+'Célula B1 da planilha de entrada deveria conter algum valor.',mtError,[mbOk],0);
 Abort;
end;
//Se tentar o try com célula vazia dá erro, mas foi testado na condição anterior
if not TryStrToFloat(Grid.Cells[2,1],d) then begin
 MessageDlg('Esta não é uma planilha de entrada do SysGran.'+#13#13+'Célula B1 da planilha de entrada deveria ser numérica.',mtError,[mbOk],0);
 Abort;
end;
if Sender=MenuItem8 then
 MessageDlg('Esta planilha é adequada para o SysGran!',mtInformation,[mbOk],0);
end;

procedure TForm1.MenuItem9Click(Sender: TObject);
begin
(PageControl1.ActivePage as TMyTabSheet).sChart.Legend.Visible:=not
(PageControl1.ActivePage as TMyTabSheet).sChart.Legend.Visible;
IsMenu3Popup:=False;
end;

procedure TForm1.PageControl1Change(Sender: TObject);
var S: String;
begin
S:=PageControl1.ActivePage.Caption;
if (pos('Shep',S)>0) or (pos('Pej',S)>0) then
 MenuItem12.Checked:=(PageControl1.ActivePage as TMyTabSheet).sImage.Stretch;
end;

procedure TForm1.RadioButton5Click(Sender: TObject);
begin
ComboBox1.Enabled:=False;
CheckBox2.Enabled:=False;
CheckBox3.Enabled:=False;
end;

procedure TForm1.RadioButton1Click(Sender: TObject);
begin
Panel12.Visible:=True;
Panel15.Visible:=False;
end;

procedure TForm1.RadioButton2Click(Sender: TObject);
begin
Panel12.Visible:=False;
Panel15.Visible:=True;
end;

procedure TForm1.RBFreqClick(Sender: TObject);
begin
Panel13.Visible:=False;
Panel14.Visible:=False;
end;

procedure TForm1.RBShepClick(Sender: TObject);
begin
Panel13.Visible:=True;
Panel14.Visible:=False;
end;

procedure TForm1.RBPejClick(Sender: TObject);
begin
Panel13.Visible:=True;
Panel14.Visible:=False;
end;

procedure TForm1.RBHistClick(Sender: TObject);
begin
Panel13.Visible:=False;
Panel14.Visible:=False;
end;

procedure TForm1.RBProbClick(Sender: TObject);
begin
Panel13.Visible:=False;
Panel14.Visible:=False;
end;

procedure TForm1.RBBivaClick(Sender: TObject);
begin
Panel13.Visible:=False;
Panel14.Visible:=True;
end;

procedure TForm1.RadioButton4Click(Sender: TObject);
begin
ComboBox1.Enabled:=True;
CheckBox2.Enabled:=True;
CheckBox3.Enabled:=True;
end;

end.

