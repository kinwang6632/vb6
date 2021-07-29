using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace IVRCaller
{
    public partial class fmMain : Form
    {
        public fmMain()
        {
            InitializeComponent();
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            string strXMLFile = @"D:\CableSoft Document\IVR_Net\XML Sample.xml";
            string strTxt = File.ReadAllText(strXMLFile);
            strTxt = txtUrl.Text + "?ID=" + strTxt;
            System.Net.HttpWebRequest aRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create( strTxt   );
            aRequest.Method = "POST";
            aRequest.Timeout = 1000;
            aRequest.Credentials = System.Net.CredentialCache.DefaultCredentials;
            System.Net.HttpWebResponse aResponse = (System.Net.HttpWebResponse)aRequest.GetResponse();
            System.IO.Stream aStream = aResponse.GetResponseStream();
            System.IO.StreamReader aReader = new System.IO.StreamReader( aStream, Encoding.UTF8 );
            PostRecver.Text = aReader.ReadToEnd();
            aStream.Close();
            aReader.Close();
            aResponse.Close();
            MessageBox.Show("完成");
        }


        private void btnSend2_Click(object sender, EventArgs e)
        {
            //string strXMLFile = @"D:\CableSoft Document\IVR_Net\XML Sample.xml";
            //string strTxt = "?Company=3&Language=1&Func=2&History=2,1&Tel=8525718&Floor=3&CMD=N21";
            //string strTxt = "?Company=5&Language=1&Func=2&History=2,1&Tel=8357828&Floor=3&CMD=N21";
            HttpWebResponse aTest;
            try
            {
                string strTxt = string.Empty;
                
                strTxt = txtUrl.Text + txtPara.Text;
                //strTxt = textBox1.Text;
                System.Net.HttpWebRequest aRequest = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(strTxt);
                aRequest.Method = "POST";
                aRequest.Timeout = 1000;
                aRequest.Credentials = System.Net.CredentialCache.DefaultCredentials;
                System.Net.HttpWebResponse aResponse = (System.Net.HttpWebResponse)aRequest.GetResponse();
                aTest = aResponse;
                System.IO.Stream aStream = aResponse.GetResponseStream();
                System.IO.StreamReader aReader = new System.IO.StreamReader(aStream, Encoding.UTF8);
                PostRecver.Text = aReader.ReadToEnd();
                aStream.Close();
                aReader.Close();
                aResponse.Close();

                MessageBox.Show("完成");
            }
            catch (WebException  ex)
            {
                if (ex.Status == WebExceptionStatus.Timeout)
                {

                    MessageBox.Show("TimeOut");

                }
                else
                {
                    MessageBox.Show(ex.Message.ToString());
                    
                }

                
                    
                
                
            }


            
        }

        private void btnSend3_Click(object sender, EventArgs e)
        {
            string strTxt = string.Empty;
            strTxt = txtUrl.Text + txtPara.Text;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strTxt);
            request.MaximumAutomaticRedirections = 4;
            request.MaximumResponseHeadersLength = 4;
            // Set credentials to use for this request.
            request.Credentials = CredentialCache.DefaultCredentials;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            PostRecver.Text = response.Headers.ToString();
            MessageBox.Show("完成");

            /*
                Console.WriteLine("Content length is {0}", response.ContentLength);
                Console.WriteLine("Content type is {0}", response.ContentType);

                // Get the stream associated with the response.
                Stream receiveStream = response.GetResponseStream();

                // Pipes the stream to a higher level stream reader with the required encoding format. 
                StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);

                Console.WriteLine("Response stream received.");
                Console.WriteLine(readStream.ReadToEnd());
                response.Close();
                readStream.Close();
            */

        }

        

    }
}