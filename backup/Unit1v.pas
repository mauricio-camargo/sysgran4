//ATabSheet: TMyTabSheet;

{
type
  TMyTabSheet = class(TTabSheet)
  GridNew : TsWorkSheetGrid;
  end;
}
{k := PageControl1.PageCount + 1;
ATabSheet := TMyTabSheet.Create(PageControl1);
ATabSheet.Parent := PageControl1;
ATabSheet.Caption:='Resultado '+IntToStr(k);
ATabSheet.GridNew := TsWorkSheetGrid.Create(ATabSheet);
ATabSheet.GridNew.Parent := ATabSheet;
ATabSheet.GridNew.Align:=alClient;
ATabSheet.GridNew.AutoEdit:=True;
ATabSheet.GridNew.Options:=[goEditing,goFixedHorzLine,goFixedVertLine,goHorzLine,goVertLine,goRangeSelect];
PageControl1.SelectNextPage(True,True);
}

unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ImgList, ComCtrls, ToolWin, frstatus, ActnList, PlaniUni,
  GrafUni, Registry, MRUList, ExtCtrls, Gauges, StrHlder,VCF1, StdCtrls, RXCtrls,
  AxCtrls, OleCtrls, System.Actions, System.ImageList;

type
  TForm1 = class(TForm)
    MainMenu1: TMainMenu;
    Arquivo1: TMenuItem;
    N3: TMenuItem;
    Sair1: TMenuItem;
    Exibir1: TMenuItem;
    Barradeferramentas1: TMenuItem;
    Barradeestatus1: TMenuItem;
    Ajuda1: TMenuItem;
    Ajuda2: TMenuItem;
    N9: TMenuItem;
    Sobre1: TMenuItem;
    ImageList1: TImageList;
    ToolBarPlani: TToolBar;
    TBNovo: TToolButton;
    TBAbrir: TToolButton;
    TBSalvar: TToolButton;
    ToolButton4: TToolButton;
    TBImprimir: TToolButton;
    ToolButton6: TToolButton;
    TBCortar: TToolButton;
    TBCopiar: TToolButton;
    TBColar: TToolButton;
    ToolButton10: TToolButton;
    TBAnalise: TToolButton;
    TBGrafico: TToolButton;
    Novo1: TMenuItem;
    Abrir1: TMenuItem;
    ActionList1: TActionList;
    NovoAc: TAction;
    OpenFileDialog: TOpenDialog;
    Janela1: TMenuItem;
    Ladoalado1: TMenuItem;
    Cascata1: TMenuItem;
    Arranjarcones1: TMenuItem;
    CofigurarImpressora1: TMenuItem;
    Configurarpgina1: TMenuItem;
    N1: TMenuItem;
    PSTD1: TPrinterSetupDialog;
    N2: TMenuItem;
    Opes1: TMenuItem;
    Timer2: TTimer;
    procedure Sair1Click(Sender: TObject);
    procedure Novo1Click(Sender: TObject);
    procedure Ladoalado1Click(Sender: TObject);
    procedure Cascata1Click(Sender: TObject);
    procedure Arranjarcones1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Abrir1Click(Sender: TObject);
    procedure TBSalvarClick(Sender: TObject);
    procedure TBCortarClick(Sender: TObject);
    procedure TBCopiarClick(Sender: TObject);
    procedure TBColarClick(Sender: TObject);
    procedure CofigurarImpressora1Click(Sender: TObject);
    procedure Configurarpgina1Click(Sender: TObject);
    procedure Opes1Click(Sender: TObject);
    procedure MRU1Click(Sender: TObject; const RecentName, Caption: String;
      UserData: Integer);
    procedure Arquivo1Click(Sender: TObject);
    procedure Barradeferramentas1Click(Sender: TObject);
    procedure Barradeestatus1Click(Sender: TObject);
    procedure TBImprimirClick(Sender: TObject);
    procedure Sobre1Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TBAnaliseClick(Sender: TObject);
    procedure TBGraficoClick(Sender: TObject);
    procedure Ajuda2Click(Sender: TObject);
  private
   PT03,PT05,PT10,PT15,PT16,PT20,PT25,PT30,PT35,PT45,PT50,PT55,PT65,PT70,PT75,PT84,
   PT80,PT85,PT90,PT95,PT97,Media,Sele,Ass,Curt:Extended;
   Med,Se,Assi,Cu:String;
    procedure FecharVazios;
    procedure ShowHint(Sender: TObject);
    procedure Folk;
    procedure Mca;
    procedure Mcb;
    procedure Trask;
    procedure Otto;
    procedure Areia;
    procedure Selecao;
    procedure Assimetria;
    procedure Curtose;
  public
    Nc,Nl:Integer;
    procedure AjeitarEntrada;
    procedure AjeitarSaida;
    procedure FecharTodas;
    procedure Analisar;
    procedure Graficos;
  end;

var
  Form1: TForm1;

implementation
Uses PaginaUni, OpcoesUni, AboutUni, AnaUni,
  TeeProcs, TeEngine, Chart, Series, ResultUni;

{$R *.dfm}

procedure TForm1.Folk;
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

procedure TForm1.Mca;
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

procedure TForm1.Mcb;
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

procedure TForm1.Trask;
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

procedure TForm1.Otto;
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

procedure TForm1.Areia;
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

procedure TForm1.Selecao;
begin
If Sele<=0.35 then Se:='Muito bem selecionado';
If (Sele>0.35) and (Sele<=0.5) then Se:='Bem selecionado';
If (Sele>0.5) and (Sele<=1) then Se:='Moderadamente selecionado';
If (Sele>1) and (Sele<=2) then Se:='Pobremente selecionado';
If (Sele>2) and (Sele<=4) then Se:='Muito pobremente selecionado';
If Sele>4 then Se:='Extremamente mal selecionado';
end;

procedure TForm1.Assimetria;
begin
If (Ass>-1) and (Ass<=-0.3) then Assi:='Muito negativa';
If (Ass>-0.3) and (Ass<=-0.1) then Assi:='Negativa';
If (Ass>-0.1) and (Ass<=0.1) then Assi:='Aproximadamente simétrica';
If (Ass>0.1) and (Ass<=0.3) then Assi:='Positiva';
If (Ass>0.3) and (Ass<=1) then Assi:='Muito positiva';
end;

procedure TForm1.Curtose;
begin
If Curt<=0.67 then Cu:='Muito platicúrtica';
If (Curt>0.67) and (Curt<=0.9) then Cu:='Platicúrtica';
If (Curt>0.9) and (Curt<=1.11) then Cu:='Mesocúrtica';
If (Curt>1.11) and (Curt<=1.5) then Cu:='Leptocúrtica';
If (Curt>1.5) and (Curt<=3) then Cu:='Muito leptocúrtica';
If Curt>3 then Cu:='Extremamente leptocúrtica';
end;

procedure TForm1.Analisar;
var Amost,j,P,I,r,z,g,C:Integer;
S,Sc,Sd:String;
W:Boolean;
Clz:Array[0..255] of Extended;
VBz:Array[0..255] of Extended;
Vz:array[0..255] of Extended;
VFz:array[0..255] of Extended;
Pz:array[0..255] of Extended;
Mid:Array[0..255] of Extended;
Std:Array[0..255] of Extended;
Sele1,A5,A6,
PAreia,PSilte,PArgila,PCascalho,P03,P05,P10,P15,P20,P16,P25,P30,P35,P45,P50,P55,P65,P70,P75,P80,P84,P85,P90,P95,P97,
PA03,PA05,PA10,PA15,PA20,PA16,PA25,PA30,PA35,PA45,PA50,PA55,PA65,PA70,PA75,PA80,PA84,PA85,PA90,PA95,PA97,
C03,C05,C10,C15,C16,C20,C25,C30,C35,C45,C50,C55,C65,C70,C75,C80,C84,C85,C90,C95,C97,
CA03,CA05,CA10,CA15,CA16,CA20,CA25,CA30,CA35,CA45,CA50,CA55,CA65,CA70,CA75,CA80,CA84,CA85,CA90,CA95,CA97,
A03,A05,A10,A15,A16,A20,A25,A30,A35,A45,A50,A55,A65,A70,A75,A80,A84,A85,A90,A95,A97,
B03,B05,B10,B15,B16,B20,B25,B30,B35,B45,B50,B55,B65,B70,B75,B80,B84,B85,B90,B95,B97,Soma:Extended;
F1src,F1Dest:TF1Book;
sheetname: String;
begin
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

if ActiveMDIChild=nil then Exit;
Screen.Cursor:=crHourGlass;
for i:=0 to 255 do begin
 Clz[i]:=0;
 VBz[i]:=0;
 Vz[i]:=0;
 VFz[i]:=0;
 Pz[i]:=0;
end;
z:=0;C:=0;Nc:=0;Nl:=0;

try

F1src:=TPlaniForm(Form1.ActiveMDIChild).F1;

with F1src do try

with TAnaForm.Create(Application) do try
Caption:='Multi-análise';
GroupBox3.Caption:='Selecione as amostras a serem analisadas:';
ListBox1.Items.Clear;
GroupBox2.Visible:=False;
GroupBox1.Visible:=True;
for i:=2 to 255 do if (EntryRC[1,i]='')then Break;
Nc:=i-1;
for i:=2 to 16834 do if EntryRC[i,1]='' then Break;
Nl:=i-1;
for i:=2 to Nl do ListBox1.Items.Add(EntryRC[i,1]);
if ShowModal<>mrOK then begin
 Screen.Cursor:=crDefault;
 Exit;
end;

If ListBox1.SelCount=0 then begin
 Screen.Cursor:=crDefault;
 MessageDlg('Selecione pelo menos uma amostra, por favor!', mtError, [mbOk], 0);
 Analisar;
 Abort;
end;

P:=0;
PB1.Visible:=True;
PB1.Max:=Nl;

for z:=0 to ListBox1.Items.Count-1 do begin
 PB1.StepIt;
 if ListBox1.Selected[z] then begin
  P:=P+1; //P=Número da amostra (1,2,3,4,etc)
  S:=ListBox1.Items[z]; //S=Nome da amostra (A,B,C,D,etc)
  Amost:=z+2 //Número da amostra em relação ao ListBox (14,15,16,17,etc;
 end else continue;
 r:=Nc-1; //Número de classes de phi (17)
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(EntryRC[1,g+1]); //Vetor de classes de phi
  Vbz[g]:=StrToFloat(EntryRC[Amost,g+1]); //Vetor de pesos
 end;
 Soma:=0; //Soma do vetor Vbz
 for i:=1 to r do Soma:=Soma + Vbz[i];
 Mid[1]:=Clz[1]-((Clz[2]-Clz[1])/2); //Ponto médio das classes
 for i:=2 to r do begin Mid[i]:=Clz[i]-((Clz[i]-Clz[i-1])/2); end;
 for i:=1 to r do VFz[i]:=(Vbz[i]*100)/Soma; //Vetor das proporções do peso

 //Mandei este códiigo lá para baixo, pois só será usado por Folk & Ward, etc...
 // Pz[1]:=VFz[1];
 // for i:=2 to r do Pz[i]:=Pz[i-1]+VFz[i];

 if ComboBox1.Text='Medida dos Momentos' then begin
  if P=1 then begin
   Sc:=TResultForm(Form1.ActiveMDIChild).Caption;
   with TResultForm.Create(Application) do begin
    F1Dest:=F1;
    Caption:='Sem nome - Resultados (Medida dos momentos) de "'+Sc+'"';
    PathResult:='Sem nome';
   end;
  end;

  if Checkbox2.Checked=False then begin
   F1Dest.EntryRC[P+1,1]:=S;
   Media:=0;
   for I:=1 to r do begin Pz[i]:=Mid[i]*VBz[i];end;
   for I:=1 to r do Media:=Media+Pz[i];
   Media:=Media/Soma;
   F1Dest.EntryRC[P+1,2]:=FloatToStrf((Media),ffGeneral,4,18);
   for I:=1 to r do Std[i]:=Mid[i]-Media;
   Sele:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i]));
   for I:=1 to r do Sele:=Sele+Pz[i];
   Sele:=Sele/(Soma-1);
   Sele1:=SQRT(Sele);
   F1Dest.EntryRC[P+1,3]:=FloatToStrf(Sele1,ffGeneral,4,18);
   Ass:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*Std[i]);
   for I:=1 to r do Ass:=Ass+Pz[i];
   Ass:=Ass/(Soma-1);
   Ass:=Ass/(Exp(1.5*Ln(Sele)));
   F1Dest.EntryRC[P+1,4]:=FloatToStrf(Ass,ffGeneral,4,18);
   Curt:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do Curt:=Curt+Pz[i];
   Curt:=Curt/(Soma-1);
   Curt:=Curt/(Exp(2*Ln(Sele)));
   F1Dest.EntryRC[P+1,5]:=FloatToStrf(Curt,ffGeneral,4,18);
   A5:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*Std[i]);
   for I:=1 to r do A5:=A5+Pz[i];
   A5:=A5/(Soma-1);
   F1Dest.EntryRC[P+1,6]:=FloatToStrf(A5/(Exp(2.5*Ln(Sele))),ffGeneral,4,18);
   A6:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do A6:=A6+Pz[i];
   A6:=A6/(Soma-1);
   F1Dest.EntryRC[P+1,7]:=FloatToStrf(A6/(Exp(3*Ln(Sele))),ffGeneral,4,18);
   F1Dest.EntryRC[1,2]:='Média';
   F1Dest.EntryRC[1,3]:='Seleção';
   F1Dest.EntryRC[1,4]:='Assimetria';
   F1Dest.EntryRC[1,5]:='Curtose';
   F1Dest.EntryRC[1,6]:='Quinto Mom.';
   F1Dest.EntryRC[1,7]:='Sexto Mom.';
  end else begin
   F1Dest.EntryRC[P+1,1]:=S;
   Media:=0;
   for I:=1 to r do Pz[i]:=Mid[i]*VBz[i];
   for I:=1 to r do Media:=Media+Pz[i];
   Media:=Media/Soma;
   F1Dest.EntryRC[P+1,2]:=FloatToStrf((Media),ffGeneral,4,18);
   Areia;
   F1Dest.EntryRC[P+1,3]:=Med;
   for I:=1 to r do Std[i]:=Mid[i]-Media;
   Sele:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i]));
   for I:=1 to r do Sele:=Sele+Pz[i];
   Sele:=Sele/(Soma-1);
   Sele1:=SQRT(Sele);
   F1Dest.EntryRC[P+1,4]:=FloatToStrf(Sele1,ffGeneral,4,18);
   Selecao;
   F1Dest.EntryRC[P+1,5]:=Se;
   Ass:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*Std[i]);
   for I:=1 to r do Ass:=Ass+Pz[i];
   Ass:=Ass/(Soma-1);
   Ass:=Ass/(Exp(1.5*Ln(Sele)));
   F1Dest.EntryRC[P+1,6]:=FloatToStrf(Ass,ffGeneral,4,18);
   if Ass<=-1.3 then Assi:='Muito negativa';
   if (Ass>-1.3) and (Ass<=-0.43) then Assi:='Negativa';
   if (Ass>-0.43) and (Ass<=0.43) then Assi:='Aproximadamente simétrica';
   if (Ass>0.43) and (Ass<=1.3) then Assi:='Positiva';
   if Ass>1.3 then Assi:='Muito positiva';
   F1Dest.EntryRC[P+1,7]:=Assi;
   Curt:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do Curt:=Curt+Pz[i];
   Curt:=Curt/(Soma-1);
   Curt:=Curt/(Exp(2*Ln(Sele)));
   F1Dest.EntryRC[P+1,8]:=FloatToStrf(Curt,ffGeneral,4,18);
   if Curt<=1.70 then Cu:='Muito platicúrtica';
   if (Curt>1.70) and (Curt<=2.55) then Cu:='Platicúrtica';
   if (Curt>2.55) and (Curt<=3.70) then Cu:='Mesocúrtica';
   if (Curt>3.70) and (Curt<=7.40) then Cu:='Leptocúrtica';
   if (Curt>7.40) and (Curt<=15) then Cu:='Muito leptocúrtica';
   if Curt>15 then Cu:='Extremamente leptocúrtica';
   F1Dest.EntryRC[P+1,9]:=Cu;
   A5:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*Std[i]);
   for I:=1 to r do A5:=A5+Pz[i];
   A5:=A5/(Soma-1);
   F1Dest.EntryRC[P+1,10]:=FloatToStrf(A5/(Exp(2.5*Ln(Sele))),ffGeneral,4,18);
   A6:=0;
   for I:=1 to r do Pz[i]:=VBz[i]*(SQR(Std[i])*SQR(Std[i])*SQR(Std[i]));
   for I:=1 to r do A6:=A6+Pz[i];
   A6:=A6/(Soma-1);
   F1Dest.EntryRC[P+1,11]:=FloatToStrf(A6/(Exp(3*Ln(Sele))),ffGeneral,4,18);
   F1Dest.EntryRC[1,2]:='Média';
   F1Dest.EntryRC[1,3]:='Classificação';
   F1Dest.EntryRC[1,4]:='Seleção';
   F1Dest.EntryRC[1,5]:='Classificação';
   F1Dest.EntryRC[1,6]:='Assimetria';
   F1Dest.EntryRC[1,7]:='Classificação';
   F1Dest.EntryRC[1,8]:='Curtose';
   F1Dest.EntryRC[1,9]:='Classificação';
   F1Dest.EntryRC[1,10]:='Quinto Mom.';
   F1Dest.EntryRC[1,11]:='Sexto Mom.';
  end;

 end else begin //Fim das medidas de momento

 //Veio lá de cima. Tem que ver se funciona...
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
 if EntryRC[Amost,Nc+1]<>'' then W:=True else W:=False;

 if P=1 then begin
  Sd:=TResultForm(Form1.ActiveMDIChild).Caption;
  with TResultForm.Create(Application) do begin
   F1Dest:=F1;
   Caption:='Sem nome - Resultados ('+ComboBox1.Text+') de "'+Sd+'"';
   PathResult:='Sem nome';
  end;
 end;

 if RadioButton5.Checked then begin
  F1Dest.EntryRC[P*4-2,1]:=S+'(Peso)';
  F1Dest.EntryRC[P*4-1,1]:=S+'(Porc.)';
  F1Dest.EntryRC[P*4,1]:=S+'(Porc.Acum.)';
  for i:=1 to r do begin
   F1Dest.EntryRC[1,i+1]:=FloatToStrf(Clz[i],ffGeneral,4,18);
   F1Dest.EntryRC[P*4-2,i+1]:=FloatToStrf(VBz[i],ffGeneral,4,18);
   F1Dest.EntryRC[P*4-1,i+1]:=FloatToStrf(VFz[i],ffGeneral,4,18);
   F1Dest.EntryRC[P*4,i+1]:=FloatToStrf(Pz[i],ffGeneral,4,18);
  end;
 end else begin
  F1Dest.EntryRC[P+1,1]:=S;
  if (CheckBox2.Checked=false) and  (CheckBox3.Checked=false) then begin
   F1Dest.EntryRC[1,2]:='Média';
   F1Dest.EntryRC[1,3]:='Mediana';
   F1Dest.EntryRC[1,4]:='Seleção';
   F1Dest.EntryRC[1,5]:='Assimetria';
   F1Dest.EntryRC[1,6]:='Curtose';
   F1Dest.EntryRC[1,7]:='% Cascalho';
   F1Dest.EntryRC[1,8]:='% Areia';
   F1Dest.EntryRC[1,9]:='% Silte';
   F1Dest.EntryRC[1,10]:='% Argila';
   F1Dest.EntryRC[P+1,2]:=FloatToStrf(Media,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,3]:=FloatToStrf(PT50,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,4]:=FloatToStrf(Sele,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,5]:=FloatToStrf(Ass,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,6]:=FloatToStrf(Curt,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,7]:=FloatToStrf(PCascalho,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,8]:=FloatToStrf(PAreia,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,9]:=FloatToStrf(PSilte,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,10]:=FloatToStrf(PArgila,ffGeneral,4,18);
  end;
  if (CheckBox2.Checked=false) and (CheckBox3.Checked=true) then begin
   F1Dest.EntryRC[1,2]:='Média';
   F1Dest.EntryRC[1,3]:='Mediana';
   F1Dest.EntryRC[1,4]:='Seleção';
   F1Dest.EntryRC[1,5]:='Assimetria';
   F1Dest.EntryRC[1,6]:='Curtose';
   F1Dest.EntryRC[1,7]:='% Cascalho';
   F1Dest.EntryRC[1,8]:='% Areia';
   F1Dest.EntryRC[1,9]:='% Silte';
   F1Dest.EntryRC[1,10]:='% Argila';
   F1Dest.EntryRC[P+1,2]:=FloatToStrf(Media,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,3]:=FloatToStrf(PT50,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,4]:=FloatToStrf(Sele,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,5]:=FloatToStrf(Ass,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,6]:=FloatToStrf(Curt,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,7]:=FloatToStrf(PCascalho,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,8]:=FloatToStrf(PAreia,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,9]:=FloatToStrf(PSilte,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,10]:=FloatToStrf(PArgila,ffGeneral,4,18);
   F1Dest.EntryRC[1,11]:='Phi-03';
   F1Dest.EntryRC[1,12]:='Phi-05';
   F1Dest.EntryRC[1,13]:='Phi-10';
   F1Dest.EntryRC[1,14]:='Phi-15';
   F1Dest.EntryRC[1,15]:='Phi-16';
   F1Dest.EntryRC[1,16]:='Phi-20';
   F1Dest.EntryRC[1,17]:='Phi-25';
   F1Dest.EntryRC[1,18]:='Phi-30';
   F1Dest.EntryRC[1,19]:='Phi-35';
   F1Dest.EntryRC[1,20]:='Phi-45';
   F1Dest.EntryRC[1,21]:='Phi-50';
   F1Dest.EntryRC[1,22]:='Phi-55';
   F1Dest.EntryRC[1,23]:='Phi-65';
   F1Dest.EntryRC[1,24]:='Phi-70';
   F1Dest.EntryRC[1,25]:='Phi-75';
   F1Dest.EntryRC[1,26]:='Phi-80';
   F1Dest.EntryRC[1,27]:='Phi-84';
   F1Dest.EntryRC[1,28]:='Phi-85';
   F1Dest.EntryRC[1,29]:='Phi-90';
   F1Dest.EntryRC[1,30]:='Phi-95';
   F1Dest.EntryRC[1,31]:='Phi-97';
   F1Dest.EntryRC[P+1,11]:=FloatToStrf(PT03,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,12]:=FloatToStrf(PT05,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,13]:=FloatToStrf(PT10,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,14]:=FloatToStrf(PT15,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,15]:=FloatToStrf(PT16,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,16]:=FloatToStrf(PT20,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,17]:=FloatToStrf(PT25,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,18]:=FloatToStrf(PT30,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,19]:=FloatToStrf(PT35,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,20]:=FloatToStrf(PT45,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,21]:=FloatToStrf(PT50,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,22]:=FloatToStrf(PT55,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,23]:=FloatToStrf(PT65,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,24]:=FloatToStrf(PT70,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,25]:=FloatToStrf(PT75,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,26]:=FloatToStrf(PT80,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,27]:=FloatToStrf(PT84,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,28]:=FloatToStrf(PT85,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,29]:=FloatToStrf(PT90,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,30]:=FloatToStrf(PT95,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,31]:=FloatToStrf(PT97,ffGeneral,4,18);
  end;
  if (CheckBox2.Checked=true) and (CheckBox3.Checked=false) then begin
   F1Dest.EntryRC[1,2]:='Média';
   F1Dest.EntryRC[1,3]:='Classificação';
   F1Dest.EntryRC[1,4]:='Mediana';
   F1Dest.EntryRC[1,5]:='Seleção';
   F1Dest.EntryRC[1,6]:='Classificação';
   F1Dest.EntryRC[1,7]:='Assimetria';
   F1Dest.EntryRC[1,8]:='Classificação';
   F1Dest.EntryRC[1,9]:='Curtose';
   F1Dest.EntryRC[1,10]:='Classificação';
   F1Dest.EntryRC[1,11]:='% Cascalho';
   F1Dest.EntryRC[1,12]:='% Areia';
   F1Dest.EntryRC[1,13]:='% Silte';
   F1Dest.EntryRC[1,14]:='% Argila';
   F1Dest.EntryRC[P+1,2]:=FloatToStrf(Media,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,3]:=Med;
   F1Dest.EntryRC[P+1,4]:=FloatToStrf(PT50,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,5]:=FloatToStrf(Sele,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,6]:=Se;
   F1Dest.EntryRC[P+1,7]:=FloatToStrf(Ass,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,8]:=Assi;
   F1Dest.EntryRC[P+1,9]:=FloatToStrf(Curt,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,10]:=Cu;
   F1Dest.EntryRC[P+1,11]:=FloatToStrf(PCascalho,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,12]:=FloatToStrf(PAreia,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,13]:=FloatToStrf(PSilte,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,14]:=FloatToStrf(PArgila,ffGeneral,4,18);
  end;
  if (CheckBox2.Checked=true) and (CheckBox3.Checked=true) then begin
   F1Dest.EntryRC[1,2]:='Média';
   F1Dest.EntryRC[1,3]:='Classificação';
   F1Dest.EntryRC[1,4]:='Mediana';
   F1Dest.EntryRC[1,5]:='Seleção';
   F1Dest.EntryRC[1,6]:='Classificação';
   F1Dest.EntryRC[1,7]:='Assimetria';
   F1Dest.EntryRC[1,8]:='Classificação';
   F1Dest.EntryRC[1,9]:='Curtose';
   F1Dest.EntryRC[1,10]:='Classificação';
   F1Dest.EntryRC[1,11]:='% Cascalho';
   F1Dest.EntryRC[1,12]:='% Areia';
   F1Dest.EntryRC[1,13]:='% Silte';
   F1Dest.EntryRC[1,14]:='% Argila';
   F1Dest.EntryRC[P+1,2]:=FloatToStrf(Media,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,3]:=Med;
   F1Dest.EntryRC[P+1,4]:=FloatToStrf(PT50,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,5]:=FloatToStrf(Sele,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,6]:=Se;
   F1Dest.EntryRC[P+1,7]:=FloatToStrf(Ass,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,8]:=Assi;
   F1Dest.EntryRC[P+1,9]:=FloatToStrf(Curt,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,10]:=Cu;
   F1Dest.EntryRC[P+1,11]:=FloatToStrf(PCascalho,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,12]:=FloatToStrf(PAreia,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,13]:=FloatToStrf(PSilte,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,14]:=FloatToStrf(PArgila,ffGeneral,4,18);
   F1Dest.EntryRC[1,15]:='Phi-03';
   F1Dest.EntryRC[1,16]:='Phi-05';
   F1Dest.EntryRC[1,17]:='Phi-10';
   F1Dest.EntryRC[1,18]:='Phi-15';
   F1Dest.EntryRC[1,19]:='Phi-16';
   F1Dest.EntryRC[1,20]:='Phi-20';
   F1Dest.EntryRC[1,21]:='Phi-25';
   F1Dest.EntryRC[1,22]:='Phi-30';
   F1Dest.EntryRC[1,23]:='Phi-35';
   F1Dest.EntryRC[1,24]:='Phi-45';
   F1Dest.EntryRC[1,25]:='Phi-50';
   F1Dest.EntryRC[1,26]:='Phi-55';
   F1Dest.EntryRC[1,27]:='Phi-65';
   F1Dest.EntryRC[1,28]:='Phi-70';
   F1Dest.EntryRC[1,29]:='Phi-75';
   F1Dest.EntryRC[1,30]:='Phi-80';
   F1Dest.EntryRC[1,31]:='Phi-84';
   F1Dest.EntryRC[1,32]:='Phi-85';
   F1Dest.EntryRC[1,33]:='Phi-90';
   F1Dest.EntryRC[1,34]:='Phi-95';
   F1Dest.EntryRC[1,35]:='Phi-97';
   F1Dest.EntryRC[P+1,15]:=FloatToStrf(PT03,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,16]:=FloatToStrf(PT05,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,17]:=FloatToStrf(PT10,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,18]:=FloatToStrf(PT15,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,19]:=FloatToStrf(PT16,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,20]:=FloatToStrf(PT20,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,21]:=FloatToStrf(PT25,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,22]:=FloatToStrf(PT30,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,23]:=FloatToStrf(PT35,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,24]:=FloatToStrf(PT45,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,25]:=FloatToStrf(PT50,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,26]:=FloatToStrf(PT55,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,27]:=FloatToStrf(PT65,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,28]:=FloatToStrf(PT70,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,29]:=FloatToStrf(PT75,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,30]:=FloatToStrf(PT80,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,31]:=FloatToStrf(PT84,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,32]:=FloatToStrf(PT85,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,33]:=FloatToStrf(PT90,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,34]:=FloatToStrf(PT95,ffGeneral,4,18);
   F1Dest.EntryRC[P+1,35]:=FloatToStrf(PT97,ffGeneral,4,18);
  end; //Último Begin
 end; //end else begin
end;
end;

Screen.Cursor:=crDefault;
PB1.Position:=0;
PB1.Visible:=False;

finally //TAnaForm.Create
Free;
end;

finally //with F1away do try
 F1src:=Nil;
end;


//Aciona exceção global
except //try inical
on Exception do begin
 Screen.Cursor:=crDefault;
 PB1.Position:=0;
 PB1.Visible:=False;
 MessageDlg('IMPOSSÍVEL MOSTRAR OS RESULTADOS!!!'+#13+#13+'Possíveis causas:'+#13+
 ' - Esta não é uma planilha padrão do SysGran'+#13+
 ' - O valores colocados resultam em expressões matematicamente impossíveis.'+#13+
 ' - Existem valores vazios.'+#13+
 ' - O separador decimal utilizado é diferente daquele padrão da sua versão do Windows.'+#13+#13+
 'Consulte o ajuda para maiores detalhes.',mtError, [mbOk], 0);
 end;
end;

end;

procedure TForm1.ShowHint(Sender: TObject);
begin
if Length(Application.Hint) > 0 then begin
 PB1.Visible:=False;
 FST1.SimplePanel:=True;
 FST1.SimpleText:=Application.Hint;
end else begin
 PB1.Visible:=True;
 FST1.SimplePanel:=False;
end;
end;

procedure TForm1.FecharTodas;
var i:integer;
begin
for i:=0 to MDIChildCount-1 do
 TPlaniForm(MDIChildren[i]).Close;
end;

procedure TForm1.FecharVazios;
var i:integer;
begin
for i:=0 to MDIChildCount-1 do
 if Pos('Sem nome',TPlaniForm(MDIChildren[i]).Caption)<>0 then
    if (TPlaniForm(MDIChildren[i]).F1.Modified=False) then
      TPlaniForm(MDIChildren[i]).Close;
end;

procedure TForm1.AjeitarEntrada;
var i:integer;
begin
for i:=0 to MDIChildCount-1 do begin
 if (ActiveMDIChild.ClassName='TPlaniForm') or (ActiveMDIChild.ClassName='TResultForm') then begin
  Form1.TBSalvar.Visible:=True;
  Form1.TBImprimir.Visible:=True;
  Form1.TBCortar.Visible:=True;
  Form1.TBCopiar.Visible:=True;
  Form1.TBColar.Visible:=True;
 end;
 if ActiveMDIChild.ClassName='TPlaniForm' then begin
  Form1.TBAnalise.Visible:=True;
  Form1.TBGrafico.Visible:=True;
 end;
 if ActiveMDIChild.ClassName='TResultForm' then begin
  Form1.TBAnalise.Visible:=False;
  Form1.TBGrafico.Visible:=False;
 end;
 if ActiveMDIChild.ClassName='TGrafForm' then begin
  Form1.TBAnalise.Visible:=False;
  Form1.TBGrafico.Visible:=False;
  Form1.TBSalvar.Visible:=False;
  Form1.TBImprimir.Visible:=False;
  Form1.TBCortar.Visible:=False;
  Form1.TBCopiar.Visible:=False;
  Form1.TBColar.Visible:=False;
 end;

end;
end;


procedure TForm1.AjeitarSaida;
var i:integer;
begin
for i:=0 to MDIChildCount-1 do begin
 if (ActiveMDIChild.ClassName='TPlaniForm') or (ActiveMDIChild.ClassName='TResultForm') then begin
  Form1.TBAnalise.Visible:=False;
  Form1.TBGrafico.Visible:=False;
  Form1.TBSalvar.Visible:=False;
  Form1.TBImprimir.Visible:=False;
  Form1.TBCortar.Visible:=False;
  Form1.TBCopiar.Visible:=False;
  Form1.TBColar.Visible:=False;
 end;
end;

Form1.FST1.Panels[0].Text:='';
Form1.FST1.Panels[1].Text:='';
end;

procedure TForm1.Sair1Click(Sender: TObject);
begin
Close;
end;

procedure TForm1.Novo1Click(Sender: TObject);
begin
FecharVazios;
with TPlaniForm.Create(Application) do begin
 Caption:='Sem nome';
 PathPlani:='Sem nome';
end;
AjeitarEntrada;
end;

procedure TForm1.Ladoalado1Click(Sender: TObject);
begin
Tile;
end;

procedure TForm1.Cascata1Click(Sender: TObject);
begin
Cascade;
end;

procedure TForm1.Arranjarcones1Click(Sender: TObject);
begin
ArrangeIcons;
end;

procedure TForm1.FormCreate(Sender: TObject);
var Ini:TRegInifile;
begin
Application.OnHint:=ShowHint;
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
if Ini.ReadBool('Exibir','Bar1',True)=False then begin
 BarradeFerramentas1.Checked:=False;

end;
if Ini.ReadBool('Exibir','BarSta',True)=False then begin
 Barradeestatus1.Checked:=False;
 FST1.Visible:=False;
end;
Ini.Free;
end;

procedure TForm1.Abrir1Click(Sender: TObject);
var Ini: TRegIniFile;
i:integer;
si:SmallInt;
S,Tipo:String;
F_Pre_Open1:TF1Book;
F:Single;
begin
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
OpenFileDialog.InitialDir:=Ini.ReadString('WorkDir','Path','');
if OpenFileDialog.Execute then begin
 for i:=0 to MDIChildCount-1 do if TPlaniForm(MDIChildren[i]).Caption=OpenFileDialog.FileName then begin
  MessageDlg('A planilha "'+OpenFileDialog.FileName+'" já está aberta!',mtInformation,[mbOK],0);
  TPlaniForm(MDIChildren[i]).BringToFront;
  Abort;
 end;

 F_Pre_Open1:=TF1Book.Create(Application);
 F_Pre_Open1.Read(OpenFileDialog.FileName,si);
 S:=F_Pre_Open1.EntryRC[1,2];
 try
  Tipo:='Planilha';
  F:=StrToFloat(S);
 except on Exception do Tipo:='Resultado'; end;
 F_Pre_Open1.Free;

 if Tipo='Planilha' then begin
  with TPlaniForm.Create(Application) do begin
   Open(OpenFileDialog.FileName); //Lá no PlaniUni
   ClipboardChanged;
  end;
 end;
 if Tipo='Resultado' then begin
 with TResultForm.Create(Application) do begin
  Open(OpenFileDialog.FileName); //Lá no ResultUni
 end;
 end;

  MRU1.Add(OpenFileDialog.FileName,0);
  MRU1.SaveToRegistry(Ini,'MRU');
  FecharVazios;
end;
Ini.Free;
end;

procedure TForm1.TBSalvarClick(Sender: TObject);
begin
if ActiveMDIChild is TPlaniForm  then TPlaniForm(ActiveMDIChild).Salvar1Click(Self) else
 if ActiveMDIChild is TResultForm  then TresultForm(ActiveMDIChild).Salvar1Click(Self);
end;

procedure TForm1.TBCortarClick(Sender: TObject);
begin
if ActiveMDIChild is TPlaniForm  then TPlaniForm(ActiveMDIChild).Recortar1Click(Sender) else
 if ActiveMDIChild is TResultForm  then TresultForm(ActiveMDIChild).Recortar1Click(Sender);
end;

procedure TForm1.TBCopiarClick(Sender: TObject);
begin
if ActiveMDIChild is TPlaniForm  then TPlaniForm(ActiveMDIChild).Copiar1Click(Sender) else
 if ActiveMDIChild is TResultForm  then TresultForm(ActiveMDIChild).Copiar1Click(Sender);
end;

procedure TForm1.TBColarClick(Sender: TObject);
begin
if ActiveMDIChild is TPlaniForm  then TPlaniForm(ActiveMDIChild).Colar1Click(Sender) else
 if ActiveMDIChild is TResultForm  then TresultForm(ActiveMDIChild).Colar1Click(Sender);
end;

procedure TForm1.CofigurarImpressora1Click(Sender: TObject);
begin
PSTD1.Execute;
end;

procedure TForm1.Configurarpgina1Click(Sender: TObject);
var Ini: TRegIniFile;
begin
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
with TPaginaForm.Create(Self) do
 try
  case Ini.ReadInteger('ConfPag','HInt',1) of
   1: RadioButton7.Checked:=True;
   2: RadioButton3.Checked:=True;
   3: RadioButton4.Checked:=True;
   4: RadioButton5.Checked:=True;
   5: RadioButton6.Checked:=True;
  end;
  Edit1.Text:=Ini.ReadString('ConfPag','Header','');
  case Ini.ReadInteger('ConfPag','FInt',1) of
   1: RadioButton8.Checked:=True;
   2: RadioButton9.Checked:=True;
   3: RadioButton10.Checked:=True;
   4: RadioButton11.Checked:=True;
   5: RadioButton12.Checked:=True;
  end;
  Edit2.Text:=Ini.ReadString('ConfPag','Footer','');
  Edit3.Text:=Ini.ReadString('ConfPag','Topo','1,0');
  Edit4.Text:=Ini.ReadString('ConfPag','Fundo','1,0');
  Edit5.Text:=Ini.ReadString('ConfPag','Esq','0,75');
  Edit6.Text:=Ini.ReadString('ConfPag','Dir','0,75');
  Edit7.Text:=IntToStr(Ini.ReadInteger('ConfPag','Zoom',100));
  if Ini.ReadBool('ConfPag','Lands',False)=true then RadioButton1.Checked:=True else RadioButton2.Checked:=True;
  if Ini.ReadBool('ConfPag','Grid',False)=True then CheckBox1.Checked:=True;
  if Ini.ReadBool('ConfPag','Col',False)=True then CheckBox2.Checked:=True;
  if Ini.ReadBool('ConfPag','Lin',False)=True then CheckBox3.Checked:=True;
  if Ini.ReadBool('ConfPag','PB',True)=True then CheckBox4.Checked:=True;
  if ShowModal=mrOK then begin
   if RadioButton7.Checked then Ini.WriteInteger('ConfPag','HInt',1);
   if RadioButton3.Checked then Ini.WriteInteger('ConfPag','HInt',2);
   if RadioButton4.Checked then Ini.WriteInteger('ConfPag','HInt',3);
   if RadioButton5.Checked then Ini.WriteInteger('ConfPag','HInt',4);
   if RadioButton6.Checked then Ini.WriteInteger('ConfPag','HInt',5);
   Ini.WriteString('ConfPag','Header',Edit1.Text);
   if RadioButton8.Checked then Ini.WriteInteger('ConfPag','FInt',1);
   if RadioButton9.Checked then Ini.WriteInteger('ConfPag','FInt',2);
   if RadioButton10.Checked then Ini.WriteInteger('ConfPag','FInt',3);
   if RadioButton11.Checked then Ini.WriteInteger('ConfPag','FInt',4);
   if RadioButton12.Checked then Ini.WriteInteger('ConfPag','FInt',5);
   Ini.WriteString('ConfPag','Footer',Edit2.Text);
   Ini.WriteString('ConfPag','Topo',Edit3.Text);
   Ini.WriteString('ConfPag','Fundo',Edit4.Text);
   Ini.WriteString('ConfPag','Esq',Edit5.Text);
   Ini.WriteString('ConfPag','Dir',Edit6.Text);
   Ini.WriteString('ConfPag','Zoom',Edit7.Text);
   if RadioButton1.Checked=True then Ini.WriteBool('ConfPag','Lands',False) else Ini.WriteBool('ConfPag','Lands',True);
   if CheckBox1.Checked then Ini.WriteBool('ConfPag','Grid',True) else Ini.WriteBool('ConfPag','Grid',False);
   if CheckBox2.Checked then Ini.WriteBool('ConfPag','Col',True) else Ini.WriteBool('ConfPag','Col',False);
   if CheckBox3.Checked then Ini.WriteBool('ConfPag','Lin',True) else Ini.WriteBool('ConfPag','Lin',False);
   if CheckBox4.Checked then Ini.WriteBool('ConfPag','PB',True) else Ini.WriteBool('ConfPag','PB',False);
  end;
 finally
  Ini.Free;
  Free;
 end;
end;

procedure TForm1.Opes1Click(Sender: TObject);
var Ini: TRegIniFile;
begin
with TOpcoesForm.Create(Self) do try
 Ini:=TRegIniFile.Create('\Software\'+Application.Title);
 DE1.Text:=Ini.ReadString('WorkDir','Path','');
 if ShowModal=mrOK then begin
  Ini.WriteString('WorkDir','Path',DE1.Text);
 end;
 Ini.Free;
finally Free;
end;

end;

procedure TForm1.MRU1Click(Sender: TObject; const RecentName,
  Caption: String; UserData: Integer);
var i:integer;
begin
for i:=0 to MDIChildCount-1 do if MDIChildren[i].Caption=RecentName then begin
 MDIChildren[i].BringToFront;
 Abort;
end;
 with TPlaniForm.Create(Self) do begin
  FecharVazios;
  Open(RecentName); //Lá no PlaniUni
  ClipboardChanged;
end;
end;

procedure TForm1.Arquivo1Click(Sender: TObject);
var Ini:TRegIniFile;
begin
MRU1.RecentMenu:=Arquivo1;
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
MRU1.LoadFromRegistry(Ini,'MRU');
Form1.MRU1.UpdateRecentMenu;
Ini.Free;
end;

procedure TForm1.Barradeferramentas1Click(Sender: TObject);
var Ini:TRegInifile;
i:integer;
begin
for i:=0 to MDIChildCount-1 do
 TPlaniForm(MDIChildren[i]).BarradeFerramentas1.Checked:=not TPlaniForm(MDIChildren[i]).BarradeFerramentas1.Checked;
BarradeFerramentas1.Checked:= not BarradeFerramentas1.Checked;
ToolBarPlani.Visible:= not ToolBarPlani.Visible;
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
if BarradeFerramentas1.Checked then
 Ini.WriteBool('Exibir','Bar1',True) else Ini.WriteBool('Exibir','Bar1',False);
Ini.Free;
end;

procedure TForm1.Barradeestatus1Click(Sender: TObject);
var Ini:TRegInifile;
i:integer;
begin
for i:=0 to MDIChildCount-1 do
 TPlaniForm(MDIChildren[i]).BarradeEstatus1.Checked:=not TPlaniForm(MDIChildren[i]).BarradeEstatus1.Checked;
BarradeEstatus1.Checked:= not BarradeEstatus1.Checked;
FST1.Visible:= not FST1.Visible;
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
if BarradeEstatus1.Checked then
Ini.WriteBool('Exibir','BarSta',True) else Ini.WriteBool('Exibir','BarSta',False);
Ini.Free;
end;

procedure TForm1.TBImprimirClick(Sender: TObject);
begin
if ActiveMDIChild is TPlaniForm  then TPlaniForm(ActiveMDIChild).Imprimirseleoatual1Click(Self) else
 if ActiveMDIChild is TResultForm  then TresultForm(ActiveMDIChild).Imprimirseleoatual1Click(Self);
end;

procedure TForm1.Sobre1Click(Sender: TObject);
var
  MS: TMemoryStatus;
begin
  with TAboutForm.Create(Self) do
  try
    Caption:='Sobre ...';
    Height:=360;
    BorderStyle:=bsDialog;
    GlobalMemoryStatus(MS);
    PhysMem.Caption := FormatFloat('#,###" KB"', MS.dwTotalPhys / 1024);
    FreeRes.Caption := Format('%d %%', [MS.dwMemoryLoad]);
    ShowModal;
  finally
    Free;
  end;
end;

procedure TForm1.Timer2Timer(Sender: TObject);
begin
AboutForm.Destroy;
Timer2.Enabled:=False;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var Ini:TRegInifile;
begin
Ini:=TRegIniFile.Create('\Software\'+Application.Title);
Form1.MRU1.SaveToRegistry(Ini,'MRU');
Ini.Free;
end;

procedure TForm1.TBAnaliseClick(Sender: TObject);
begin
Analisar;
end;

procedure TForm1.TBGraficoClick(Sender: TObject);
begin
Graficos;
end;

procedure TForm1.Graficos;
var z,P:Integer;
F1away:TF1Book;
label Inic;
begin
try

 F1away:=TPlaniForm(Form1.ActiveMDIChild).F1;

  with TAnaForm.Create(Application) do try
   Inic:
   Caption:='Multi-gráficos';
   GroupBox3.Caption:='Selecione as amostras a serem graficadas:';
   ListBox1.Items.Clear;
   GroupBox2.Visible:=True;
   GroupBox1.Visible:=False;

   Nl:=0;
   for z:=2 to 16834 do if F1away.EntryRC[z,1]='' then Break;
   Nl:=z-1;
   for z:=2 to Nl do ListBox1.Items.Add(F1away.EntryRC[z,1]);
   if ShowModal=mrOK then begin
    if ListBox1.SelCount=0 then begin
     MessageDlg('Por favor, selecione pelo menos uma amostra!', mtError, [mbOk], 0);
     goto Inic;
    end;
    STH1.Clear;
    if (RBShep.Checked) or (RBPej.Checked) then begin
     with TGrafForm.Create(Application) do begin
      RxSpeedButton1.Visible:=False;
      RxSpeedButton2.Visible:=False;
      BTFormat.Visible:=False;
      Height:=440;
      F1Graf:=F1away;
      CorLabel:=ColorComboBox1.ColorValue;
      CorPonto:=ColorComboBox2.ColorValue;
      TamPonto:=StrToInt(FloatToStr(RXSpinEdit1.Value));
      NuShep:=ListBox1.SelCount;
      for z:= 0 to ListBox1.Items.Count-1 do begin
       if ListBox1.Selected[z] then STH1.Strings.Add(ListBox1.Items[z]) else continue;
      end;
      Shep(ListBox1,RBShep.Checked{Shep ou Pejrup},CBLabel.Checked );
      Exit;
     end;
    end;
    if RBFreq.Checked then begin
     If ListBox1.SelCount>10 then begin
      MessageDlg('Não mais que 10 amostras podem ser plotadas para frequência acumulada!', mtError, [mbOk], 0);
      goto Inic;
     end;
     with TGrafForm.Create(Application) do begin
      BTFormat.Visible:=True;
      F1Graf:=F1away;
      Freq(ListBox1);
     end;
    end;
    if RBBivar.Checked then begin
     If ListBox1.SelCount>999 then begin
      MessageDlg('Não mais que 999 amostras podem ser plotadas em gráficos bivariados!', mtError, [mbOk], 0);
      goto Inic;
     end;
     with TGrafForm.Create(Application) do begin
      BTFormat.Visible:=True;
      F1Graf:=F1away;
      Bivariado(ListBox1,ComboBox2.Text,ComboBox3.Text,ComboBox4.Text);
     end;
    end;
    if RBHist.Checked then begin
     if ListBox1.SelCount>99 then begin
      MessageDlg('Não mais que 99 amostras podem ser plotadas em histogramas!', mtError, [mbOk], 0);
      goto Inic;
     end;
     with TGrafForm.Create(Application) do begin
      BTFormat.Visible:=True;
      MaxAm:=ListBox1.SelCount;
      if MaxAm=1 then RxSB6.Enabled:=False else
       RxSB6.Enabled:=True;
      P:=0;
      for z:=0 to ListBox1.Items.Count-1 do begin
       if ListBox1.Selected[z] then begin
        P:=P+1;
        NoAm[P]:=z+2;
       end else continue;
      end;
      h:=1;
      F1Graf:=F1away;
      Hist;
     end;
    end;
    if RBProb.Checked then begin
     If ListBox1.SelCount>10 then begin
      MessageDlg('Não mais que 10 amostras podem ser plotadas em gráficos de probabilidade!', mtError, [mbOk], 0);
      goto Inic;
     end;
     with TGrafForm.Create(Application) do begin
      RxSpeedButton1.Visible:=False;
      RxSpeedButton2.Visible:=False;
      BTFormat.Visible:=False;
      F1Graf:=F1away;
      Prob(ListBox1);
     end;

    end;
   end;//ShowModal=mrOK
  finally Free; //TAnaForm.Create
  end;
  F1away:=Nil;

//Aciona exceção global
except //try inical
on Exception do begin
 Screen.Cursor:=crDefault;
 PB1.Position:=0;
 PB1.Visible:=False;
 MessageDlg('IMPOSSÍVEL MOSTRAR OS RESULTADOS!!!'+#13+#13+'Possíveis causas:'+#13+
 ' - Esta não é uma planilha padrão do SysGran'+#13+
 ' - O valores colocados resultam em expressões matematicamente impossíveis.'+#13+
 ' - Existem valores vazios.'+#13+
 ' - O separador decimal utilizado é diferente daquele padrão da sua versão do Windows.'+#13+#13+
 'Consulte o ajuda para maiores detalhes.',mtError, [mbOk], 0);
 end;
end;
end;


procedure TForm1.Ajuda2Click(Sender: TObject);
begin
Application.HelpCommand(HELP_FINDER, 0);
end;

end.



