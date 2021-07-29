#define PVCS_VERSION "$Name: v2_4_0_3 $ $Header: /users/ndcsrc/cvsroot/ndclib/util/ndcutil1.c,v 2.9 1998/12/07 11:41:02 edarshan Exp $"
/***************************************************************************/
/* File: ndcutil1.c                                                        */
/* Desc: General Purpose Utility routines Source Code                      */
/*-------------------------------------------------------------------------*/
/* Date        Programmer     Description                                  */
/* ----        ----------     -----------                                  */
/*04-Jun-1992  Ofra Branner   Original Version                             */
/*17-Mar-1996  Asher Sterkin  Doesn't use NDCLIB - can be sent outside NDC */
/*                            Define NLENDIAN for little-endian machines   */
/*                            (NDCLIB takes care of this automatically)    */
/***************************************************************************/
#include <stdlib.h>
#include <memory.h>

#define  ULONG_BITS_SIZE    (sizeof(unsigned long) * 8)
#define  AND   &
#define  OR    |
#define  XOR   ^ 

/* All our targets are little-endian except HP */
#if !defined(NHP)
#define NLENDIAN
#endif

/***************************************************************************/
static void  Swap(unsigned char *b0, unsigned char *b1)
{
#ifdef NLENDIAN
    unsigned char tmp = *b0;
    *b0 = *b1;
    *b1 = tmp;
#endif
}
/***************************************************************************/
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
/***************************************************************************/
static int S(short n, unsigned long X, unsigned long *ret)
{
   if (n >= 0  &&  n < 32) {
       *ret = (X << n) OR (X >> (32-n));
       return(0);
   }
   else
       return (-1);
}
/***************************************************************************/
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
/***************************************************************************/
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

/**RS***********************************************************************/
/**                                                                       **/
/** Definition : NSTATUS UTIL_MD5(NUCHAR *input, NUSHORT input_len,       **/
/**                               NUCHAR *key, NUSHORT key_len,           **/
/**                               NUCHAR *out)                            **/
/**                                                                       **/
/** Description:                                                          **/
/**      This routine used to compute the full MD4 signature on an input  **/
/**      buffer and an optional key buffer. The key buffer is padded if   **/
/**      to the input before computing the signature if needed. 'out' is  **/
/**      the result of the signature and it always contains 2 longs       **/
/**                                                                       **/
/**RE***********************************************************************/
#ifdef   NWINDOWS
 #ifdef WIN32
   #ifdef BUILD_DLL
     #define  NWEXPORT   __declspec( dllexport )
   #elif  STATIC_LIB
     #define  NWEXPORT
   #else
     #define  NWEXPORT   __declspec( dllimport )
   #endif
 #endif
#else
  #define NWEXPORT
#endif

int NWEXPORT UTIL_MD5(unsigned char *input,unsigned short input_len,unsigned char *key,
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

#ifdef TEST
#include <stdio.h>
#include <assert.h>
#include <string.h>
 
static struct
{
  unsigned char *input;
  unsigned char *output;
} test_data[] = 
{ 
  {"0001M0035010001000119940206125158ER01","C6B82780F0E1A6B6"}
 ,{"0001M0032010002000119950906114824O"   ,"A5DF5AF100F7A2A1"}
 ,{"0001M0035010002000119950906114844ES04","7C9FAB78D84E4E14"}
 ,{"0001M0035010002000119950906114904ES04","60D1A9A700366989"}
};

int main(int argc,char **argv)
{
  int i,j;
  unsigned char output[8];
  unsigned char buf[17];
  unsigned char *p;

  for(i=0;i<4;i++)
  {
    p = test_data[i].input;
    assert(!UTIL_MD5(p,strlen(p),"12345678",8,output));
    for(j=0;j<8;j++)
      sprintf(buf+2*j,"%02X",(unsigned int)output[j]);
    assert(!strcmp(buf,test_data[i].output));
  }
}

#endif
