//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
int CamcCheckDigit (ULONG CardId)
{
     short i, digit, ckdig;
     short odd_sum = 0, even_sum = 0;
     unsigned long num = CardId;

     for (i=1 ; i<12 ; i++)
      {
           digit = (short) (num % 10);
           num = num / 10L;
           if ((i % 2) == 1)   {
            digit *= 2;
            if (digit >= 10) digit -=9;
            odd_sum += digit;
           }
           else {
            even_sum += digit;
           }
      }
     ckdig = (10 - ((even_sum + odd_sum) % 10)) % 10;
     return(ckdig);
}

__fastcall TForm1::TForm1(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
    int ans;
    ans = CamcCheckDigit(1000);

    Memo1->Lines->Add('abc');
}
//---------------------------------------------------------------------------
