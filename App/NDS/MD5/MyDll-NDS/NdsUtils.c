//---------------------------------------------------------------------------

#include <windows.h>
//---------------------------------------------------------------------------
//   Important note about DLL memory management when your DLL uses the
//   static version of the RunTime Library:
//
//   If your DLL exports any functions that pass String objects (or structs/
//   classes containing nested Strings) as parameter or function results,
//   you will need to add the library MEMMGR.LIB to both the DLL project and
//   any other projects that use the DLL.  You will also need to use MEMMGR.LIB
//   if any other projects which use the DLL will be performing new or delete
//   operations on any non-TObject-derived classes which are exported from the
//   DLL. Adding MEMMGR.LIB to your project will change the DLL and its calling
//   EXE's to use the BORLNDMM.DLL as their memory manager.  In these cases,
//   the file BORLNDMM.DLL should be deployed along with your DLL.
//
//   To avoid using BORLNDMM.DLL, pass string information using "char *" or
//   ShortString parameters.
//
//   If your DLL uses the dynamic version of the RTL, you do not need to
//   explicitly add MEMMGR.LIB as this will be done implicitly for you
//---------------------------------------------------------------------------

#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void* lpReserved)
{
    return 1;
}
//---------------------------------------------------------------------------
#define PVCS_VERSION "$Name: v2_4_0_3 $ $Header: /users/ndcsrc/cvsroot/ndclib/util/ndcutil1.c,v 2.9 1998/12/07 11:41:02 edarshan Exp $"
#define  ULONG_BITS_SIZE    (sizeof(unsigned long) * 8)
#define  AND   &
#define  OR    |
#define  XOR   ^ 

#if !defined(NHP)
#define NLENDIAN
#endif


static int f(short t, unsigned long x, unsigned long y, unsigned long z, unsigned long *ret)
{
   if (t >= 0  &&  t <= 19) {
       *ret = (x AND y) OR (~x AND z);
   }
   else  if (t >= 20  &&  t <= 39) {
       *ret =  x XOR y XOR z;
   }
   else  if (t >= 40  &&  t <= 59) {
       *ret = (x AND y) OR (x AND z) OR (y AND z);
   }
   else  if (t >= 60  &&  t <= 79) {
       *ret = x XOR y XOR z;
   }
   else
       return(-1);
   return(0);
}
/***************************************************************************/
static int K(short t, unsigned long *ret)
{
   static unsigned long k[4] = { 0x5a827999L, 0x6ed9eba1L, 0x8f1bbcdcL, 0xca62c1d6L };

   if (t >= 0  &&  t <= 19) {
       *ret = k[0];
   }
   else  if (t >= 20  &&  t <= 39) {
       *ret = k[1];
   }
   else  if (t >= 40  &&  t <= 59) {
       *ret = k[2];
   }
   else  if (t >= 60  &&  t <= 79) {
       *ret = k[3];
   }
   else 
       return(-1);
   return(0);
}


static int S(short n, unsigned long X, unsigned long *ret)
{
   if (n >= 0  &&  n < 32) {
       *ret = (X << n) OR (X >> (32-n));
       return(0);
   }
   else
       return (-1);
}
static void  Swap(unsigned char *b0, unsigned char *b1)
{
#ifdef NLENDIAN
    unsigned char tmp = *b0;
    *b0 = *b1;
    *b1 = tmp;
#endif
}

static void pad(unsigned long *from,unsigned long len,unsigned long *to,unsigned long l)
{
   short          i;
   unsigned long  nn;
   unsigned long  tmp, wrk_one = 1;
   unsigned long  mask = 0xFFFFFFFF;
   unsigned long *ptr = to;
   
                                  /* 1-st whole ULONGs */
   for (nn = len / ULONG_BITS_SIZE;  nn > 0;  nn--, l -= ULONG_BITS_SIZE) {
       tmp = *(from++);
       Swap ((unsigned char *)&tmp, ((unsigned char *)&tmp)+3);    /* Machine depended */
       Swap (((unsigned char *)&tmp)+1, ((unsigned char *)&tmp)+2);
       *(to++) = tmp;
   }

                                  /* now ULONG with 1 appended */
   i = (unsigned short)(len % ULONG_BITS_SIZE);
   if (i > 0) {
       int j = (unsigned short)ULONG_BITS_SIZE - i - 1;
       tmp = *from;
       Swap ((unsigned char *)&tmp, ((unsigned char *)&tmp)+3);    /* Machine depended */
       Swap (((unsigned char *)&tmp)+1, ((unsigned char *)&tmp)+2);
       mask <<= j;
       wrk_one <<= j;
       tmp &= mask;               /* turn off the bits that do not belong */
       tmp |= wrk_one;            /* append 1 */
   }
   else  {
       tmp = 0x80000000L;
   }
   *(to++) = tmp;
   l -= ULONG_BITS_SIZE;

                                  /* Add ULONGs of 0s */
   for (tmp = 0;  l > 2 * ULONG_BITS_SIZE;  l -= ULONG_BITS_SIZE) {
      *(to++) = tmp;
   }

       /* Add 2 ULONGs of start length (the 1-st is 0 - length < 2**32) */
   *(to++) = 0;
   *to = len;
}

static int md4(unsigned char *input, unsigned long input_len, unsigned char *out)
{                                         /* input_len - number of bytes */
   unsigned long        h0, h1, h2, h3, h4;
   unsigned long        A, B, C, D, E;
   unsigned long        W[80];
   unsigned long        TEMP;
   unsigned long        *ptr;
   unsigned long        *base;
   unsigned long        f_tmp, k_tmp, s_tmp;
   short                t;
   unsigned long        bitslen;
   unsigned long        ltmp;
   unsigned long        ll;
   unsigned char       *bptr;
   int                  status;

   /* Pad msg */
   bitslen = input_len * 8;
   ll = ((bitslen + 576) / 512) * 512;     /* Future length in bits */
                                           /* 576 = 512 +(512 - 448)*/
   ltmp = (unsigned long)(ll + 7) / 8;             /* In bytes              */

   if(!(ptr=malloc(ltmp)))
     return -1;
   base = ptr;

   pad ((unsigned long *)input, bitslen, ptr, ll);

   /* main procedure */
   /* prepared message consist now of 16 * N ULONGS - main iteration N */

   h0=0x67452301L;
   h1=0xefcdab89L;
   h2=0x98badcfeL;
   h3=0x10325476L;
   h4=0xc3d2e1f0L;

   for (ll /= (16 * ULONG_BITS_SIZE);  ll > 0;  ll--)  { /*....................*/
                                           /* step (a) */
       for (t = 0;  t < 16;  t++, ptr++) {
           W[t] = *ptr;
       }
                                           /* step (b) */
       for (t = 16;  t <= 79;  t++) {      
           W[t] = W[t-3]  XOR  W[t-8]  XOR  W[t-14]  XOR  W[t-16];
       }
                                           /* step (c) */
       A = h0;
       B = h1;
       C = h2;
       D = h3;
       E = h4;
                                          /* step (d) */ 
       for (t = 0;  t <= 79;  t++) {
           if ((status = S(5,A, &s_tmp)) != 0)
               return(status);
           if ((status = f(t,B,C,D, &f_tmp)) != 0)
               return(status);
           if ((status = K(t, &k_tmp)) != 0)
               return(status);
           TEMP =  s_tmp+ f_tmp + E + W[t] + k_tmp;
           E = D;
           D = C;
           if ((status =S(30,B, &C)) != 0)
               return(status);
           B = A;
           A = TEMP;
       }
                                           /* step (e) */
       h0 += A;
       h1 += B;
       h2 += C;
       h3 += D;
       h4 += E;
   }

   free(base);
   ptr = (unsigned long *)out;
   /*******************************************************************
   This code was changed to convert 5 long numbers into two longs
   *(ptr++) = h0;
   *(ptr++) = h1;
   *(ptr++) = h2;
   *(ptr++) = h3;
   *(ptr++) = h4;
   *********************************************************************/
   h1  ^= h0;
   h2  ^= h0;
   h3  ^= h0;
   h4  ^= h0;
   *(ptr++) = h1 + h2;
   *(ptr++) = h3 + h4;
   /********************************************************************/
   /* changed t = 5 to t = 2 to only put out two longs                 */
   /********************************************************************/
   for (bptr = out, t = 2;  t > 0;  t--, bptr += sizeof(unsigned long))  {
       Swap (bptr, bptr+3);
       Swap (bptr+1, bptr+2);
   }
   
   return(0);
}


int UTIL_MD5(unsigned char *input,unsigned short input_len,unsigned char *key,
   unsigned short key_len, unsigned char *out)
{
   int            retval;
/*   int            status;*/
   unsigned char *in_buffer;
   unsigned char *out_buffer;

   in_buffer = malloc(input_len+key_len);
   if (!in_buffer)
     return -1;
   memcpy(in_buffer, input, input_len);       /* Concat of input and key  */
   memcpy(in_buffer + input_len, key, key_len);
   /* NDRqa08337 - md4 assumes that the input and output buffers are correctly
      aligned for longs, even though they are typed as char*. Violation of
      this assumption on RISC machines can cause core dumps. malloc
      guarantees that the buffer has correct alignment. We already malloced
      the input buffer because we had to concatenate the input and key; now
      we malloc the output buffer too */
   out_buffer = malloc(2*sizeof(long));
   if (!out_buffer)
      return -1;
   
   retval = md4(in_buffer, (unsigned long)input_len + key_len, out_buffer);

   memcpy(out, out_buffer, 2*sizeof(long));
   free(in_buffer);
   free(out_buffer);
   return(retval);
}







/***************************************************************************
 MD5 block update operation. Continues an MD5 message-digest
 operation, processing another message block, and updating the context.
***************************************************************************/

/***************************************************************************
 NDS_RealMD5_8 returns 8 bytes result of MD5 algorithm on input + key after xor
 between i,i+8 bytes of the result.
***************************************************************************/
int __declspec(dllexport) NDS_RealMD5_8(unsigned char *input,unsigned short input_len,unsigned char *key,unsigned short key_len, unsigned char *out)
{
   int            retval;
/*   int            status;*/
   unsigned char *in_buffer;
   unsigned char *out_buffer;

   in_buffer = malloc(input_len+key_len);
   if (!in_buffer)
     return -1;
   memcpy(in_buffer, input, input_len);       /* Concat of input and key  */
   memcpy(in_buffer + input_len, key, key_len);
   /* NDRqa08337 - md4 assumes that the input and output buffers are correctly
      aligned for longs, even though they are typed as char*. Violation of
      this assumption on RISC machines can cause core dumps. malloc
      guarantees that the buffer has correct alignment. We already malloced
      the input buffer because we had to concatenate the input and key; now
      we malloc the output buffer too */
   out_buffer = malloc(2*sizeof(long));
   if (!out_buffer)
      return -1;
   
   retval = md4(in_buffer, (unsigned long)input_len + key_len, out_buffer);

   memcpy(out, out_buffer, 2*sizeof(long));
   free(in_buffer);
   free(out_buffer);
   return(retval);

}

