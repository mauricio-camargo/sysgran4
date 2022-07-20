unit GrafUni;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, TeeProcs, TeEngine, Chart, Printers,
  Series, ExtDlgs, VCF1, Menus, RXCtrls, ComCtrls, ToolWin, Clipbrd;

type
  TGrafForm = class(TForm)
    Panel2: TPanel;
    Image1: TImage;
    SaveDialog1: TSaveDialog;
    PrintDialog1: TPrintDialog;
    Te1: TChart;
    Series1: TLineSeries;
    Series2: TLineSeries;
    Series3: TLineSeries;
    Series4: TLineSeries;
    Series5: TLineSeries;
    Series6: TLineSeries;
    Series7: TLineSeries;
    Series8: TLineSeries;
    Series9: TLineSeries;
    Series10: TLineSeries;
    Te2: TChart;
    Series11: TBarSeries;
    Te3: TChart;
    Series12: TPointSeries;
    MainMenu1: TMainMenu;
    Arquivo1: TMenuItem;
    Novo1: TMenuItem;
    Abrir1: TMenuItem;
    Fechar1: TMenuItem;
    Fechartodas1: TMenuItem;
    N1: TMenuItem;
    Salvar1: TMenuItem;
    Salvarcomo1: TMenuItem;
    N3: TMenuItem;
    Configurarimpressora1: TMenuItem;
    N7: TMenuItem;
    Imprimirtudo1: TMenuItem;
    N2: TMenuItem;
    Sair1: TMenuItem;
    Exibir1: TMenuItem;
    Barradeferramentas1: TMenuItem;
    Barradeestatus1: TMenuItem;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    RxSpeedButton1: TRxSpeedButton;
    ToolButton9: TToolButton;
    RxSpeedButton2: TRxSpeedButton;
    RxSB5: TRxSpeedButton;
    ToolButton6: TToolButton;
    RxSB6: TRxSpeedButton;
    BTFormat: TRxSpeedButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure RxSpeedButton1Click(Sender: TObject);
    procedure RxSpeedButton2Click(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure Zoom1Click(Sender: TObject);
    procedure Copiar2Click(Sender: TObject);
    procedure BTFormatClick(Sender: TObject);
    procedure Novo1Click(Sender: TObject);
    procedure Abrir1Click(Sender: TObject);
    procedure Fechar1Click(Sender: TObject);
    procedure Fechartodas1Click(Sender: TObject);
    procedure Salvar1Click(Sender: TObject);
    procedure Salvarcomo1Click(Sender: TObject);
    procedure Configurarimpressora1Click(Sender: TObject);
    procedure Configurarpgina1Click(Sender: TObject);
    procedure Imprimirtudo1Click(Sender: TObject);
    procedure Sair1Click(Sender: TObject);
    procedure Barradeferramentas1Click(Sender: TObject);
    procedure Barradeestatus1Click(Sender: TObject);
  private
  public
   Ser:Array[1..10] of TLineSeries;
   NoAm:Array[1..99] of ShortInt;
   MaxAm,h:ShortInt;
   F1Graf:TF1Book;
   TeAtivo:TChart;
   NuShep:integer;
   CorPonto,CorLabel:TColor;
   TamPonto:ShortInt;
   procedure Hist;
   procedure Prob(LB:TListBox);
   procedure Freq(LB:TListBox);
   procedure Bivariado(LB:TListBox;CB2,CB3,CB4:string);
   procedure Shep(LB:TListBox;RB:Boolean{Shep ou Pejrup};Lab:Boolean{Plot Label});
  end;

var
  GrafForm: TGrafForm;

implementation

uses Unit1, ConfigGraf;

{$R *.DFM}

function Pala(q: Extended):Integer;
var
S:String;
U:Extended;
begin
U:=Int(q);
S:=FloatToStr(U);
Pala:=StrToInt(S);
end;

procedure TGrafForm.Shep(LB:TListBox;RB:Boolean{Shep ou Pejrup};Lab:Boolean{Plot Label});
var z,Amost,x,C,g,r,Ops,i:Integer;
SHxa,SHya,SHxb,SHyb,
ArHxa,ArHya,ArHxb,ArHyb,
AHxa,AHya,AHxb,AHyb,
ArA,ArB,SA,SB,AA,AB,
YY,XX,YY1,XX1,YY2,XX2,
Ar,Soma,Areia,Silte,Argila:Extended;
ArXa,ArYa,ArXb,ArYb,
SXa,SYa,SXb,SYb,
AXa,AYa,AXb,AYb:LongInt;
Clw:Array[1..30] of Extended;
VFw:array[1..30] of Extended;
VBw:Array[1..30] of Extended;
Bitmap:TBitmap;
begin
try
Bitmap:=TBitmap.Create;
Bitmap.Height:=376;
Bitmap.Width:=560;
Image1.Picture.Assign(Bitmap);

Screen.Cursor:=crHourGlass;
with Image1.Canvas do begin
 Brush.Color:=clWhite;
 FloodFill(0, 0, clgreen, fsBorder);
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
 TextOut(186,20,'Argila');
 TextOut(374,321,'Silte');
 TextOut(3,338,'Areia');
 Font.Style:=[];
 TextOut(1,313,'100%');TextOut(365,340,'100%');TextOut(210,36,'100%');TextOut(87,175,'50%');
 TextOut(292,175,'50%');TextOut(87,175,'50%');TextOut(194,340,'50%');
 if RB then begin
  TextOut(129,103,'25%');TextOut(249,103,'75%');TextOut(43,250,'75%');TextOut(336,252,'25%');
  TextOut(108,340,'25%');TextOut(280,340,'75%');TextOut(198,76,'1');TextOut(166,156,'2');
  TextOut(230,156,'3');TextOut(198,212,'4');TextOut(129,235,'5');TextOut(270,235,'8');
  TextOut(182,240,'6');TextOut(214,240,'7');TextOut(64,306,'9');TextOut(160,306,'10');
  TextOut(234,306,'11');TextOut(332,306,'12');
 end else begin
  TextOut(338,259,'20%');TextOut(246,94,'80%');MoveTo(348,337);LineTo(186,54);TextOut(338,339,'10%');
  TextOut(164,42,'10%');MoveTo(52,337);LineTo(40,309);TextOut(50,339,'90%');TextOut(18,293,'90%');
  MoveTo(200,337);LineTo(112,186);Font.Size:=12;TextOut(238,62,'I');TextOut(276,134,'II');
  TextOut(326,226,'III');TextOut(362,290,'IV');TextOut(270,344,'C');TextOut(122,344,'B');TextOut(32,344,'A');
 end;
 if RB then begin
  Rectangle(360,11,550,288);
  Font.Size:=10;Font.Color:=CorLabel;TextOut(370,18,'   CONVENÇÕES');Font.Color:=clBlack;
  TextOut(370,34,'1 - Argila ou argilito');TextOut(370,50,'2 - Argila Arenosa');
  TextOut(370,66,'3 - Argila síltica');TextOut(370,82,'4 - Argila siltico-arenosa');
  TextOut(370,98,'5 - Areia argilosa');TextOut(370,114,'6 - Areia síltico-argilosa');
  TextOut(370,130,'7 - Silte argilo-arenoso');TextOut(370,146,'8 - Silte argiloso');
  TextOut(370,162,'9 - Areia ou arenito');TextOut(370,178,'10 - Areia síltica');
  TextOut(370,194,'11 - Silte arenoso');TextOut(370,210,'12 - Silte ou siltito');
  Font.Color:=CorLabel;TextOut(370,230,'   LEGENDAS');Font.Color:=clBlack;
  TextOut(387,246,'- Fração de grânulos < 3%');TextOut(387,262,'- Fração de grânulos > 3%');
  Brush.Color := CorPonto;Pen.Color:=CorPonto;Ellipse(370,250,380,260);
  Polygon([Point(368, 273), Point(375, 266), Point(382, 273),Point(368, 273)]);
 end else begin
  Rectangle(360,11,550,172);Font.Size:=10;Font.Color:=CorLabel;TextOut(370,18,'   CONVENÇÕES');
  Font.Color:=clBlack;TextOut(370,38,'I - Hidrodinâmica baixa');
  TextOut(370,54,'II - Hidrodinâmica moderada');TextOut(370,70,'III - Hidrodinâmica alta');
  TextOut(370,86,'IV - Hidrodinâmica muito alta');

  Font.Color:=CorLabel;
  TextOut(370,110,'   LEGENDAS');
  Font.Color:=clBlack;
  TextOut(387,130,'- Fração de grânulos < 3%');
  TextOut(387,146,'- Fração de grânulos > 3%');
  Brush.Color := CorPonto;
  Pen.Color:=CorPonto;
  Ellipse(370,134,380,144);
  Polygon([Point(368, 157), Point(375, 150), Point(382, 157),Point(368, 157)]);
 end;
end;

for i:=2 to 255 do if (F1Graf.EntryRC[1,i]='') then Break;
Form1.Nc:=i-1;
for i:=2 to 16834 do if F1Graf.EntryRC[i,1]='' then Break;
Form1.Nl:=i-1;

Amost:=0;
Form1.PB1.Max:=Form1.Nl;
for z:=0 to LB.Items.Count-1 do begin
 Form1.PB1.StepIt;
 if LB.Selected[z] then Amost:=z+2 else continue;
 r:=Form1.Nc-1;
 for g:=1 to r do begin
  Clw[g]:=StrToFloat(F1Graf.EntryRC[1,g+1]);
  Vbw[g]:=StrToFloat(F1Graf.EntryRC[Amost,g+1]);
 end;

 Soma:=0;
 for I:=1 to r do Soma:=Soma + Vbw[I];
 for I:=1 to r do VFw[I]:=(Vbw[I]*100)/Soma;
 Areia:=0;
 Silte:=0;
 Argila:=0;
 for i:=1 to r do begin
  if (Clw[i]>-1) and (Clw[i]<=4) then Areia:=Areia+VFw[i];
  if (Clw[i]>4) and (Clw[i]<=8) then Silte:=Silte+VFw[i];
  if Clw[i]>8 then Argila:=Argila+VFw[i];
 end;
 with Image1.canvas do begin
  ArHxa:=200-1.76*Areia;
  ArHya:=34+3.03*Areia;
  ArXa:=Pala(ArHxa);
  ArYa:=Pala(ArHya);
  ArHxb:=375-3.51*Areia;
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
  Ellipse(Pala(XX)-TamPonto,Pala(YY)+Pala(Abs(YY1-YY)/2)-TamPonto,Pala(XX)+2*TamPonto,Pala(YY)+Pala(Abs(YY1-YY)/2)+2*TamPonto);

  Brush.Color := clWhite;
  Pen.Color:=clBlack;
  Font.Name:='Arial';
  Font.Size:=7;
  Font.Color:=CorLabel;
  if Lab then TextOut(Pala(XX)-14,Pala(YY)-14,F1Graf.EntryRC[Amost,1]);
 end;
end;

Form1.PB1.Position:=0;
Form1.PB1.Visible:=False;
Image1.Top:=2;
Image1.Left:=2;
Image1.Visible:=True;
Screen.Cursor:=crDefault;

Bitmap.Free;
except
on exception do begin
 Form1.PB1.Position:=0;
 Form1.PB1.Visible:=False;
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar a apresentação do Diagrama de Shepard.', mtError, [mbOk], 0);
end;
end;
end;

procedure TGrafForm.Prob(LB:TListBox);
var Amost,P,I,r,z,g,C:Integer;
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
begin
try
Screen.Cursor := CrHourGlass;
Amost:=0;P:=0;I:=0;r:=0;z:=0;g:=0;C:=0;
MaxX:=0;MaxY:=0;MinX:=0;MinY:=0;aX:=0;bX:=0;aY:=0;bY:=0;DistX:=0;DistY:=0;

Bitmap:=TBitmap.Create;
Bitmap.Height:=209;
Bitmap.Width:=571;
Image1.Picture.Assign(Bitmap);

With Image1.Canvas do begin
 Pen.Color:=clBlack;
 Rectangle(61,23,480,172);
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

for z:=LB.Items.Count-1 downto 0 do if LB.Selected[z] then Si:=LB.Items[z];
for i:=2 to 255 do if F1Graf.EntryRC[1,i]='' then Break;
r:=i-2;
for z:=0 to LB.Items.Count-1 do begin
 if LB.Selected[z] then begin
  P:=P+1;
  S:=LB.Items[z];
  Amost:=z+2
 end else continue;
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(F1Graf.EntryRC[1,g+1]);
  Vbz[g]:=StrToFloat(F1Graf.EntryRC[Amost,g+1]);
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

 Image1.Canvas.Font.Name:='Arial';
 Image1.Canvas.Font.Size:=7;

 DistY:=(MinY-MaxY)/7;
 for i:=0 to 7 do
  Image1.Canvas.TextOut(38,19+(i*21),FloatToStrf(MaxY+(i*DistY), ffNumber, 8, 2));

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

 with Image1.Canvas do begin
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

Image1.Canvas.Pen.Color:=clBlack;
for i:=1 to r-1 do begin
 Image1.Canvas.MoveTo(ValX[i],ValY[i]+2);
 Image1.Canvas.LineTo(ValX[i+1]+1,ValY[i+1]+2);
end;
Image1.Canvas.Brush.Color:=clWhite;
end;

with Image1.Canvas do begin
 Image1.Canvas.Pen.Color:=clBlack;
 Image1.Canvas.Brush.Color:=clWhite;
 if P=1 then begin
  Image1.Canvas.Font.Size:=9;
  Image1.Canvas.TextOut(228,5,S);
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

Image1.Top:=2;
Image1.Left:=2;
Image1.Visible:=True;

RxSpeedButton1.Enabled:=False;
RxSpeedButton2.Enabled:=False;

Bitmap.Free;

Screen.Cursor := CrDefault;

except on Exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
end;
end;
end;

procedure TGrafForm.FormClose(Sender: TObject; var Action: TCloseAction);
var i:integer;
begin

//Form1.BarradeEstatus1.Checked:=BarradeEstatus1.Checked;//Primeiro passa os estados dos checks
//Form1.BarradeFerramentas1.Checked:=BarradeFerramentas1.Checked;//para o Form1 e lá no Ajeitar vão
                                                         //ser passados para os novos forms filhos
//Form1.AtualWin:=''; //Limpando o nome da janela atual

Form1.AjeitarSaida;
Action:=caFree;
end;

procedure TGrafForm.BitBtn2Click(Sender: TObject);
begin
Close;
end;

procedure TGrafForm.BitBtn3Click(Sender: TObject);
begin
if Image1.Visible then SaveDialog1.Filter:='Bitmap (*.bmp)|*.bmp';
if SaveDialog1.Execute then begin
 if Image1.Visible then Image1.Picture.SaveToFile(SaveDialog1.FileName);
 if Te1.Visible then begin
  case SaveDialog1.FilterIndex of
  1:Te1.SaveToBitmapFile(SaveDialog1.FileName);
  2:Te1.SaveToMetafile(SaveDialog1.FileName);
  3:Te1.SaveToMetafileEnh(SaveDialog1.FileName);
  end;
 end;
 if Te2.Visible then begin
  case SaveDialog1.FilterIndex of
  1:Te2.SaveToBitmapFile(SaveDialog1.FileName);
  2:Te2.SaveToMetafile(SaveDialog1.FileName);
  3:Te2.SaveToMetafileEnh(SaveDialog1.FileName);
  end;
 end;
 if Te3.Visible then begin
  case SaveDialog1.FilterIndex of
  1:Te3.SaveToBitmapFile(SaveDialog1.FileName);
  2:Te3.SaveToMetafile(SaveDialog1.FileName);
  3:Te3.SaveToMetafileEnh(SaveDialog1.FileName);
  end;
 end;
end;
end;

procedure TGrafForm.BitBtn4Click(Sender: TObject);
var Rect:TRect;
begin
if PrintDialog1.Execute then begin
  if Image1.Visible then begin
   Rect.Left:=0;
   Rect.Top:=0;
   Rect.Right:=Image1.Picture.Width*4;
   Rect.Bottom:=Image1.Picture.Height*4;
   with Printer do begin
    BeginDoc;
    Canvas.StretchDraw(Rect,Image1.Picture.Bitmap);
    EndDoc;
   end;
  end;
  if Te1.Visible then Te1.Print;
  if Te2.Visible then Te2.Print;
  if Te3.Visible then Te3.Print;
  end;
 end;

procedure TGrafForm.Hist;
var g,r,i:Integer;
Soma,ValorMax:Extended;
Clw:Array[0..2550] of Extended;
VFw:array[0..255] of Extended;
VBw:Array[0..255] of Extended;
begin
try
 Screen.Cursor := CrHourGlass;
 for i:=2 to 255 do if F1Graf.EntryRC[1,i]='' then Break;
 r:=i-2;
 for g:=1 to r do begin
  Clw[g]:=StrToFloat(F1Graf.EntryRC[1,g+1]);
  Vbw[g]:=StrToFloat(F1Graf.EntryRC[NoAm[h],g+1]);
 end;
 Clw[0]:=0;
 Vbw[0]:=0;
 VFw[0]:=0;
 Soma:=0;
 ValorMax:=0;
 for i:=1 to r do Soma:=Soma + Vbw[I];
 for i:=1 to r do VFw[I]:=(Vbw[I]*100)/Soma;
 for i:=1 to r do if VFw[I]>ValorMax then ValorMax:=VFw[I];
 Series11.Clear;
 Te2.Title.Text.Clear;
 Te2.Title.Text.Add(F1Graf.EntryRC[NoAm[h],1]);
 Te2.LeftAxis.Maximum:=ValorMax+5;
 for g:=1 to r do
 Series11.Add(VFw[g],FloatToStr(Clw[g]),ClTeeColor);
 Screen.Cursor:=crDefault;
 Te2.Align:=alClient;
 Te2.Visible:=true;
 RxSB5.Visible:=True;
 RxSB6.Visible:=True;
 TeAtivo:=Te2;
 except on exception do
  begin
   Screen.Cursor:=crDefault;
   MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
  end;
 end;
end;

procedure TGrafForm.Freq(LB:TListBox);
var P,z,g,Amost,r,i:Integer;
Soma:Extended;
S:String;
Clw:Array[0..255] of Extended;
VFw:array[0..255] of Extended;
Pw:array[0..255] of Extended;
VBw:Array[0..255] of Extended;
begin
try
 P:=0;z:=0;g:=0;Amost:=0;r:=0;i:=0;

 Screen.Cursor := CrHourGlass;
 Te1.Legend.Visible:=True;
 for i:=1 to 10 do
 Ser[i]:=TLineSeries(FindComponent('Series'+IntToStr(i)));
 for i:=2 to 255 do if F1Graf.EntryRC[1,i]='' then Break;
 r:=i-2;
 P:=0;
 for z:=0 to LB.Items.Count-1 do begin
  if LB.Selected[z] then begin
   P:=P+1;
   Ser[P].Active:=True;
   S:=LB.Items[z];
   Amost:=z+2
  end else continue;
  for g:=1 to r do begin
    Clw[g]:=StrToFloat(F1Graf.EntryRC[1,g+1]);
    Vbw[g]:=StrToFloat(F1Graf.EntryRC[Amost,g+1]);
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
  for i:=1 to r do
  Ser[P].addXY(Clw[i],Pw[i],'',clTeeColor);
  Ser[P].Title:=S;
 end;
 if P=1 then Te1.Legend.Visible:=False;
 Screen.Cursor:=crDefault;
 Te1.Align:=alClient;
 Te1.Title.Text.Add('Freqüências acumuladas');
 TeAtivo:=Te1;
 Te1.Visible:=true;
except on exception do
 begin
  Screen.Cursor:=crDefault;
  MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
 end;
end;
end;

procedure TGrafForm.BitBtn6Click(Sender: TObject);
begin
if h=MaxAm-1 then RxSB6.Enabled:=False;
if h>=0 then RxSB5.Enabled:=True;
h:=h+1;
Hist;
end;

procedure TGrafForm.BitBtn5Click(Sender: TObject);
begin
if h=2 then RxSB5.Enabled:=False;
if h<=MaxAm then RxSB6.Enabled:=True;
h:=h-1;
Hist;
end;

procedure TGrafForm.FormActivate(Sender: TObject);
begin
Form1.AjeitarEntrada;
end;

procedure TGrafForm.RxSpeedButton1Click(Sender: TObject);
begin
TeAtivo.ZoomPercent(90);
end;

procedure TGrafForm.Bivariado(LB:TListBox;CB2,CB3,CB4:string);
var Amost,P,I,r,z,g,C:Integer;
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
A1,A2,A2a,A3,A4,P03,P05,P10,P15,P20,P16,P25,P30,P35,P45,P50,P55,P65,P70,P75,P80,P84,P85,P90,P95,P97,
PA03,PA05,PA10,PA15,PA20,PA16,PA25,PA30,PA35,PA45,PA50,PA55,PA65,PA70,PA75,PA80,PA84,PA85,PA90,PA95,PA97,
C03,C05,C10,C15,C16,C20,C25,C30,C35,C45,C50,C55,C65,C70,C75,C80,C84,C85,C90,C95,C97,
CA03,CA05,CA10,CA15,CA16,CA20,CA25,CA30,CA35,CA45,CA50,CA55,CA65,CA70,CA75,CA80,CA84,CA85,CA90,CA95,CA97,
A03,A05,A10,A15,A16,A20,A25,A30,A35,A45,A50,A55,A65,A70,A75,A80,A84,A85,A90,A95,A97,
B03,B05,B10,B15,B16,B20,B25,B30,B35,B45,B50,B55,B65,B70,B75,B80,B84,B85,B90,B95,B97,Soma,MaxX,MaxY,MinX,
MinY,DistX,DistY,bX,aX,bY,aY:Extended;
PT03,PT05,PT10,PT15,PT16,PT20,PT25,PT30,PT35,PT45,PT50,PT55,PT65,PT70,PT75,PT84,
PT80,PT85,PT90,PT95,PT97,Media,Sele,Ass,Curt:Extended;
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
for i:=2 to 255 do if (F1Graf.EntryRC[1,i]='') then Break;
r:=i-2;
for z:=0 to LB.Items.Count-1 do begin
 if LB.Selected[z] then begin
  P:=P+1;
  Amost:=z+2
 end else continue;
 for g:=1 to r do begin
  Clz[g]:=StrToFloat(F1Graf.EntryRC[1,g+1]);
  Vbz[g]:=StrToFloat(F1Graf.EntryRC[Amost,g+1]);
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

if CB4='Folk & Ward' then begin
 Mediaz[P]:=(PT16+PT50+PT84)/3;
 Selecaoz[P]:=((PT84-PT16)/4)+((PT95-PT05)/6.6);
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4='McCammon (a)' then begin
 Mediaz[P]:=(PT10+PT30+PT50+PT70+Pt90)/5;
 Selecaoz[P]:=((PT85+PT95-PT05-PT15)/5.4);
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4='McCammon (b)' then begin
 Mediaz[P]:=(PT05+PT15+PT25+PT35+PT45+PT55+PT65+PT75+PT85+PT95)/10;
 Selecaoz[P]:=(PT70+PT80+PT90+PT97-PT03-PT10-PT20-PT30)/9.1;
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4='Trask' then begin
 Mediaz[P]:=PT50;
 Selecaoz[P]:=(PT75-PT25)/1.35;
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4='Otto & Inman' then begin
 Mediaz[P]:=(PT16+PT84)/2;
 Selecaoz[P]:=(PT84-PT16)/2;
 Assimetriaz[P]:=((PT16+PT84-2*PT50)/(2*(PT84-PT16)))+((PT05+PT95-2*PT50)/(2*(PT95-PT05)));
 Curtosez[P]:=(PT95-PT05)/(2.44*(PT75-PT25));
end;

if CB4='Medida dos Momentos' then begin
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

if (CB2='Média') or (CB3='Média') then begin
 Te3.BottomAxis.Title.Caption:='Média';
 for i:=0 to 999 do X[i]:=Mediaz[i];
end;
if (CB2='Curtose') or (CB3='Curtose') then begin
 for i:=0 to 999 do Y[i]:=Curtosez[i];
 Te3.LeftAxis.Title.Caption:='Curtose';
end;
if (CB2='Assimetria') or (CB3='Assimetria') then begin
 if (CB2='Curtose') or (CB3='Curtose') then begin
   for i:=0 to 999 do X[i]:=Assimetriaz[i];
   Te3.BottomAxis.Title.Caption:='Assimetria';
 end else begin
  for i:=0 to 999 do Y[i]:=Assimetriaz[i];
  Te3.LeftAxis.Title.Caption:='Assimetria';
 end;
end;
if (CB2='Seleção') or (CB3='Seleção') then begin
 if (CB2='Média') or (CB3='Média') then begin
  for i:=0 to 999 do Y[i]:=Selecaoz[i];
  Te3.LeftAxis.Title.Caption:='Seleção';
 end else begin
  for i:=0 to 999 do X[i]:=Selecaoz[i];
  Te3.BottomAxis.Title.Caption:='Seleção';
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

for i:=1 to P do
Series12.AddXY(X[i],Y[i],'',clTeeColor);

Screen.Cursor := CrDefault;
Te3.Align:=alClient;
Te3.Title.Text.Add('Bivariado ('+CB2+' X '+CB3+')');
TeAtivo:=Te3;
Te3.Visible:=True;

except
on Exception do begin
 Screen.Cursor:=crDefault;
 MessageDlg('Não foi possível preparar o gráfico para os valores atuais.', mtError, [mbOk], 0);
end;
end;
end;

procedure TGrafForm.RxSpeedButton2Click(Sender: TObject);
begin
TeAtivo.ZoomPercent(110);
end;

procedure TGrafForm.ToolButton1Click(Sender: TObject);
begin
if Image1.Visible then Clipboard.Assign(TGrafForm(Form1.ActiveMDIChild).Image1.Picture) else
 TEAtivo.CopyToClipboardBitmap;
end;

procedure TGrafForm.Zoom1Click(Sender: TObject);
begin
TeAtivo.ZoomPercent(110);
end;

procedure TGrafForm.Copiar2Click(Sender: TObject);
begin
TEAtivo.CopyToClipboardBitmap;
end;

procedure TGrafForm.BTFormatClick(Sender: TObject);
begin
with TConfigGrafForm.Create(self) do try
 CB3D.Checked:=TEAtivo.View3D;
 CBTitulo.Checked:=TEAtivo.Title.Visible;
 Edit1.Text:=TEAtivo.Title.Text[0];
 Edit1.Font:=TEAtivo.Title.Font;
 if TEAtivo.Title.Color=clDefault then
  Edit1.Color:=clBtnFace else
   Edit1.Color:=TEAtivo.Title.Color;
 Panel1.Color:=TEAtivo.Color;
 if TEAtivo.BackColor=clDefault then
  Panel2.Color:=clBtnFace else
   Panel2.Color:=TEAtivo.BackColor;
 if Te1.Visible then Panel6.Caption:='Cores múltiplas';
 if Te2.Visible then Panel6.Color:=Series11.SeriesColor;
 if Te3.Visible then Panel6.Color:=Series12.Pointer.Brush.Color;

//Eixo X
 CBX.Checked:=TEAtivo.BottomAxis.Automatic;
 SKMinX.Text:=FloatToStr(TEAtivo.BottomAxis.Minimum);
 SKMaxX.Text:=FloatToStr(TEAtivo.BottomAxis.Maximum);
 SKMaxX.Font:=TEAtivo.BottomAxis.LabelsFont;
 SKMinX.Font:=TEAtivo.BottomAxis.LabelsFont;
 EditTituloX.Text:=TEAtivo.BottomAxis.Title.Caption;
 EditTituloX.Font:=TEAtivo.BottomAxis.Title.Font;
//Eixo Y
 CBY.Checked:=TEAtivo.LeftAxis.Automatic;
 SKMinY.Text:=FloatToStr(TEAtivo.LeftAxis.Minimum);
 SKMaxY.Text:=FloatToStr(TEAtivo.LeftAxis.Maximum);
 SKMaxY.Font:=TEAtivo.LeftAxis.LabelsFont;
 SKMinY.Font:=TEAtivo.LeftAxis.LabelsFont;
 EditTituloY.Text:=TEAtivo.LeftAxis.Title.Caption;
 EditTituloY.Font:=TEAtivo.LeftAxis.Title.Font;


 //Legenda
 if TE1.Visible then begin
  CBLeg.Checked:=Te1.Legend.Visible;
  Panel3.Color:=Te1.Legend.Color;
  Panel4.Font:=Te1.Legend.Font;
 end else LegendaSheet.PageControl:=nil;

 if ShowModal<>mrOK then Exit;
 TEAtivo.View3D:=CB3D.Checked;
 TEAtivo.Title.Visible:=CBTitulo.Checked;
 TEAtivo.Title.Text[0]:= Edit1.Text;
 TEAtivo.Title.Font:=Edit1.Font;
 TEAtivo.Title.Color:=Edit1.Color;
 TEAtivo.Color:= Panel1.Color;
 TEAtivo.BackColor:=Panel2.Color;
 if Te2.Visible then Series11.SeriesColor:=Panel6.Color;
 if Te3.Visible then Series12.Pointer.Brush.Color:=Panel6.Color;
 if (Te1.Visible) and (Panel6.Caption='') then begin
  Series1.SeriesColor:=Panel6.Color;
  Series2.SeriesColor:=Panel6.Color;
  Series3.SeriesColor:=Panel6.Color;
  Series4.SeriesColor:=Panel6.Color;
  Series5.SeriesColor:=Panel6.Color;
  Series6.SeriesColor:=Panel6.Color;
  Series7.SeriesColor:=Panel6.Color;
  Series8.SeriesColor:=Panel6.Color;
  Series9.SeriesColor:=Panel6.Color;
  Series10.SeriesColor:=Panel6.Color;
 end;

//Eixo X
 TEAtivo.BottomAxis.Automatic:=CBX.Checked;
 TEAtivo.BottomAxis.AutomaticMaximum:=CBX.Checked;
 TEAtivo.BottomAxis.AutomaticMinimum:=CBX.Checked;
 TEAtivo.BottomAxis.Minimum:=StrToFloat(SKMinX.Text);
 TEAtivo.BottomAxis.Maximum:=StrToFloat(SKMaxX.Text);
 TEAtivo.BottomAxis.LabelsFont:=SKMaxX.Font;
 TEAtivo.BottomAxis.Title.Caption:=EditTituloX.Text;
 TEAtivo.BottomAxis.Title.Font:=EditTituloX.Font;
//Eixo Y
 TEAtivo.LeftAxis.Automatic:=CBY.Checked;
 TEAtivo.LeftAxis.AutomaticMaximum:=CBY.Checked;
 TEAtivo.LeftAxis.AutomaticMinimum:=CBY.Checked;
 TEAtivo.LeftAxis.Minimum:=StrToFloat(SKMinY.Text);
 TEAtivo.LeftAxis.Maximum:=StrToFloat(SKMaxY.Text);
 TEAtivo.LeftAxis.LabelsFont:=SKMaxY.Font;
 TEAtivo.LeftAxis.Title.Caption:=EditTituloY.Text;
 TEAtivo.LeftAxis.Title.Font:=EditTituloY.Font;
//Legenda
 if TE1.Visible then begin
  Te1.Legend.Visible:=CBLeg.Checked;
  Te1.Legend.Color:=Panel3.Color;
  Te1.Legend.Font:=Panel4.Font;
 end;

finally free;
end;

end;

procedure TGrafForm.Novo1Click(Sender: TObject);
begin
Form1.Novo1Click(Sender);
end;

procedure TGrafForm.Abrir1Click(Sender: TObject);
begin
Form1.Abrir1Click(Sender);
end;

procedure TGrafForm.Fechar1Click(Sender: TObject);
begin
Close;
end;

procedure TGrafForm.Fechartodas1Click(Sender: TObject);
begin
Form1.FecharTodas;
end;

procedure TGrafForm.Salvar1Click(Sender: TObject);
begin
BitBtn3Click(Self);
end;

procedure TGrafForm.Salvarcomo1Click(Sender: TObject);
begin
BitBtn3Click(Self);
end;

procedure TGrafForm.Configurarimpressora1Click(Sender: TObject);
begin
Form1.PSTD1.Execute;
end;

procedure TGrafForm.Configurarpgina1Click(Sender: TObject);
begin
Form1.Configurarpgina1Click(Sender);
end;

procedure TGrafForm.Imprimirtudo1Click(Sender: TObject);
begin
BitBtn4Click(Self);
end;

procedure TGrafForm.Sair1Click(Sender: TObject);
begin
Form1.Sair1Click(Sender);
end;

procedure TGrafForm.Barradeferramentas1Click(Sender: TObject);
begin
Form1.Barradeferramentas1Click(Self);
end;

procedure TGrafForm.Barradeestatus1Click(Sender: TObject);
begin
Form1.Barradeestatus1Click(Self);
end;

end.
