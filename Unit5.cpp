#include <vcl.h>
#pragma hdrstop

#include "Unit5.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "Word_XP_srvr"
#pragma resource "*.dfm"
TForm5 *Form5;
//---------------------------------------------------------------------------
__fastcall TForm5::TForm5(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm5::Button1Click(TObject *Sender)
{

UDPClient1->Host = EHost->Text;
UDPClient1->Send(EMessage->Text);
}
//---------------------------------------------------------------------------
void __fastcall TForm5::UDPServer1UDPRead(TIdUDPListenerThread *AThread, const TIdBytes AData,
          TIdSocketHandle *ABinding)
{
LMessages->Items->Add(BytesToString(AData));
}

//End of UDP 

//---------------------------------------------------------------------------
void __fastcall TForm5::Button2Click(TObject *Sender)
{
//aratam din document excel
if(ComboBox1->Items->Count <1) return;
ADOQuery1->SQL->Text = "Select * from [" + ComboBox1->Items->Strings[ComboBox1->ItemIndex] + "]";
ADOQuery1->Open();
}
//---------------------------------------------------------------------------
void __fastcall TForm5::FormCreate(TObject *Sender)
{
//conectam la Excel document

//UnicodeString Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
//ExtractFilePath(Application->ExeName)+ "test.xls" + ";Extended Properties=Excel 8.0";

UnicodeString Connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
ExtractFilePath(Application->ExeName)+ "test.xls" + ";Extended Properties=Excel 12.0 Xml";
ADOConnection1->ConnectionString = Connection;

//luam tabelele
TStringList *SL = new TStringList;
SL = new TStringList;
ComboBox1->Clear();
ADOConnection1->GetTableNames(SL);
for (int i = 0; i < SL->Count; i++) {
	ComboBox1->Items->Add(SL->Strings[i]);
	delete SL;
}
}



void WordLine(Word_xp::TWordDocument* WD,Word_xp::TWordFont* WF, UnicodeString Line, AnsiString Font,
int height, int bold, int italic, int underline, int shadow,
TColor Color = (TColor)RGB(0,0,0), Word_xp::WdParagraphAlignment Position = WdParagraphAlignment::wdAlignParagraphLeft)
{
	WD->GetDefaultInterface()->Paragraphs->get_Last()->set_Alignment(Position);
	WD->Range(EmptyParam(),EmptyParam())->InsertAfter(StringToOleStr(Line));
	WF->ConnectTo(WD->Sentences->get_Last()->get_Font());
	WF->GetDefaultInterface()->Size = height;
	WF->GetDefaultInterface()->Shadow = shadow;
	WF->GetDefaultInterface()->Bold = bold;
	WF->GetDefaultInterface()->Italic = italic;
	WF->GetDefaultInterface()->Underline = (Word_xp::WdUnderline)underline;
	WF->GetDefaultInterface()->Color = (Word_xp::WdColor)Color;
	WF->set_Name(StringToOleStr(Font));
}

//---------------------------------------------------------------------------
void __fastcall TForm5::Button3Click(TObject *Sender)
{
	Button3->Visible=false;
	Button4->Visible=true;

	OleVariant Template = EmptyParam();
	OleVariant NewTemplate = False;
	OleVariant ItemIndex = 1;

	//conectam la Word
	try{
		WordApplication1->Connect();
	}
	catch(...){
	Application->MessageBoxW(L"Microsoft Wrd is not installed!", L"Error", MB_OK|MB_ICONSTOP);
	return;
	}

	WordApplication1->Documents->Add(Template,NewTemplate);
	WordDocument1->ConnectTo(WordApplication1->Documents->Item(ItemIndex));

	for(int i=1; i<=15;i++)
	WordLine(WordDocument1,WordFont1,IntToStr(i)+ "\n","Calibri",i+10,i%2,i%3,0,i%4,TColor(RGB(i*25,0,250-i*10)));


	WordApplication1->GetDefaultInterface()->Visible=true;
	WordApplication1->Disconnect();
}
//---------------------------------------------------------------------------
void __fastcall TForm5::Button4Click(TObject *Sender)
{
Button4->Visible=false;
Button3->Visible=true;
}
//---------------------------------------------------------------------------
