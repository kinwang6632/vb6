/***************************************************************************/
// MD5.H - header file for MD5C.C
 

/* Copyright (C) 1991-2, RSA Data Security, Inc. Created 1991. All
   rights reserved.

   License to copy and use this software is granted provided that it
   is identified as the "RSA Data Security, Inc. MD5 Message-Digest
   Algorithm" in all material mentioning or referencing this software
   or this function.

   License is also granted to make and use derivative works provided
   that such works are identified as "derived from the RSA Data
   Security, Inc. MD5 Message-Digest Algorithm" in all material
   mentioning or referencing the derived work.  
                                                                    
   RSA Data Security, Inc. makes no representations concerning either
   the merchantability of this software or the suitability of this
   software for any particular purpose. It is provided "as is"
   without express or implied warranty of any kind.  
                                                                    
   These notices must be retained in any copies of any part of this
   documentation and/or software.  
 */
/***************************************************************************/


#ifndef _MD5_H_
#define _MD5_H_ 

#ifdef __cplusplus
extern "C" {
#endif	


/***************************************************************************
/ NDS_RealMD5_16 - a regular RSA MD5(returns 16 bytes),
/**************************************************************************/
int NDS_RealMD5_16(unsigned char *input,unsigned short input_len,unsigned char *key,
			unsigned short key_len, unsigned char *out);


/***************************************************************************
 NDS_RealMD5_8 - a special NDS version of MD5(returns 8 bytes).
***************************************************************************/
int NDS_RealMD5_8(unsigned char  *input, unsigned short input_len,unsigned char *key,
			unsigned short key_len, unsigned char *out);

#ifdef __cplusplus
}
#endif


#endif


