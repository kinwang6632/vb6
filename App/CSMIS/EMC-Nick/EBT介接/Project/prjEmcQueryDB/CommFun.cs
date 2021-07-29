
using System;
using System.Web;




namespace prjEmcQueryDB
{
	/// <summary>
	/// CommFun ���K�n�y�z�C
	/// </summary>
	public class CommFun : System.Web.UI.Page
	{

		private enum CompCode : byte
		{
			KS = 1, 
			PN = 2, 
			NT = 3,
			NCC = 5,
			FM = 6,
			CT = 7,
			UC = 8,
			YMS = 9,
			NTP = 10,
			KP = 11,
			DW = 12,
			TC = 13,
			ST = 14,
			LK = 15,
			NTY = 16 
		}
		
		private static string[] CompName = new string[16] 
     {
			"�[�@",
			"�̫n", 
		   "�n��",
		   string.Empty,
		   "�s�W�D",
		   "�׷�",
		   "���D",		  
			"���p",
			"�����s",
			"�s�x�_",
			"���W�D",
			"�j�w��s",
			"�s��",
			"�j�s��",
			"�R�a",
			"�_���",
	  };

		public CommFun()
		{
			//
			// TODO: �b���[�J�غc�禡���{���X
			//
		}

		
		public static bool getAddressValue(String sI_InputValue,  String sI_Unit, out String sI_Result)
		{
			//�P�_�ϥΪ̿�J���ѧ������Ȫ��榡�O�_���T, �æ�X�A���r��
			bool bL_IsErrorValue = false;
			String sL_TmpString="";
			int nL_Tmp=0;
			sI_Result = "";

			try
			{
					

				sL_TmpString = sI_InputValue;
				

				if ((sL_TmpString!="") && (sL_TmpString.Substring(0,1)!="-"))
				{
					nL_Tmp = System.Convert.ToInt32(sL_TmpString);
					sI_Result= sL_TmpString + sI_Unit;
				}
				else
				{
					bL_IsErrorValue = true;
					sI_Result ="�z��J��'" + sI_Unit + "'���榡���~";
					return bL_IsErrorValue;
				}

			}
			catch
			{
				//down, ��ܤ��O�¼Ʀr...
				int nL_Ndx = -1, nL_StrLength=0;
				String sL_Header="", sL_Tailer="";			

				nL_StrLength = sL_TmpString.Length;
				nL_Ndx = sL_TmpString.IndexOf("-",0,nL_StrLength);

				if (nL_Ndx>=0)
				{
					sL_Header = sL_TmpString.Substring(0, nL_Ndx);
					sL_Tailer = sL_TmpString.Substring(nL_Ndx+1, nL_StrLength - nL_Ndx - 1);

					try
					{
						
						if ((sL_Header.IndexOf("-",0,sL_Header.Length)>=0) || (sL_Tailer.IndexOf("-",0,sL_Tailer.Length)>=0))
						{
							bL_IsErrorValue = true;
							sI_Result ="�z��J��'" + sI_Unit + "'���榡���~";
							return bL_IsErrorValue;
						}
						
							
					}
					catch
					{
						bL_IsErrorValue = true;
						sI_Result ="�z��J��'" + sI_Unit + "'���榡���~";
						return bL_IsErrorValue;
					}
					try
					{
						if ((System.Convert.ToInt32(sL_Header)!=1111111) && ((System.Convert.ToInt32(sL_Tailer)!=1111111)))
						{
							sI_Result = sL_TmpString + sI_Unit;
						}
					}
					catch
					{
						sI_Result = sL_TmpString;
					}

				}
				else
				{
					try
					{
						if (System.Convert.ToInt32(sL_TmpString)!=1111111) 
						{
							sI_Result = sL_TmpString + sI_Unit;
						}
					}
					catch
					{
						sI_Result = sL_TmpString;
					}
				}	


				//up, ��ܤ��O�¼Ʀr...
			}
			return bL_IsErrorValue;
		}

		
		public static string GetCompChineseName(byte aCompCode)
		{
         string aResult = string.Empty;

			if ( aCompCode <= ( (Array)CompName ).Length ) 
			{
				aResult = CompName[aCompCode - 1];
			}
         
			return aResult;
		}
	
	
		public static string GetQueryViewName(byte aCompCode, int aViewSeq)
		{
         string aResult = string.Empty;

			if ( !IsEbtDataArea( aCompCode.ToString() ) )
			{
				aResult = string.Empty;
			}
			else if ( Enum.IsDefined( typeof( CompCode ), aCompCode ) )
			{
				aResult = string.Format( "V_{0}_EMCEBT00{1}", 
					( (CompCode)Enum.Parse( typeof( CompCode ), aCompCode.ToString() ) ).ToString(), aViewSeq );
			}
			return aResult;
		}


		public static string GetNormalDateTimeStr()
		{
			return string.Format(  "{0}/{1}/{2} {3}:{4}:{5}",
				DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 
				DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second );
		}


		public static string getOracleDateTimeStr()
		{
			return string.Format( "TO_DATE( '{0}', 'YYYY/MM/DD HH24:MI:SS' )", GetNormalDateTimeStr() );
		}

		
		public static bool IsEbtDataArea(string aCompCode)
		{
			return !( ( aCompCode == "5" ) || ( aCompCode == "6" ) ||
				( aCompCode == "7" ) || ( aCompCode == "16" ) );
		}

		
		public string GetCookie(System.Web.HttpRequest aRequest, string aKey)
		{
			string aResult = string.Empty;			
			HttpCookie MyCookie = aRequest.Cookies["USERINFO"];

			if ( MyCookie != null )
			{
				aResult = MyCookie.Values[aKey.ToUpper()];
			}

			if ( aResult == null )
				aResult = "";

			return aResult;
		}

		
		public void RedirectToLogin(System.Web.UI.Page aPage)
		{
			if ( Session.Count <= 0 )
			{
				Session.Abandon();
				Server.Transfer( "Logout.htm" );
			}
		}

	}
}
