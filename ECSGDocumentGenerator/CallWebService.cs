using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Text;
using System.Threading.Tasks;

namespace ECSGDocumentGenerator
{
    public class CallWebService
    {

        public static HttpWebResponse makePostRequest(string url)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Headers.Add("X-Authorization: eyJhbGciOiJIUzI1NiJ9.eyJlbmFibGVkIjp0cnVlLCJhY2NvdW50Tm9uRXhwaXJlZCI6dHJ1ZSwiZGVzY2VuZGFudFNjb3BlcyI6WyIqIl0sImFjY291bnROb25Mb2NrZWQiOnRydWUsImNyZWRlbnRpYWxzTm9uRXhwaXJlZCI6dHJ1ZSwidXNlcm5hbWUiOm51bGwsImF1dGhvcml0aWVzIjpbeyJhdXRob3JpdHkiOiJST0xFX1BPTElORSJ9XSwicGFzc3dvcmQiOm51bGwsImlhdCI6MTU3NDI3NTExMSwiZXhwIjoxNTc5NTM0NzExfQ.zo5MNcCCnxm1U5OHWoHBpkap7-tjbYvvHqqL8xy6Cuw");
            req.Headers.Add("X-Application: themis-inf-test");
            req.Method = "POST";
            req.ContentType = "application/json";
            req.Accept = "application/json";
            req.MediaType = "application/json";
            //req.Host = "s-themis-acc.net1.cec.eu.int:8044";

            System.Text.ASCIIEncoding encoding = new ASCIIEncoding();
            string json = "{\"memberState\":\"SE\"}";
            byte[] postByteArray = encoding.GetBytes(json);
            Stream requestStream = req.GetRequestStream();
            requestStream.Write(postByteArray, 0, postByteArray.Length);
            requestStream.Close();

            WebResponse response = req.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader streamReader = new StreamReader(responseStream);
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            return resp;
        }


        public static string makePostRequestUsingWebClient(string url)
        {

            WebClient client = new WebClient();

            client.Headers[HttpRequestHeader.Authorization] = "eyJhbGciOiJIUzI1NiJ9.eyJlbmFibGVkIjp0cnVlLCJhY2NvdW50Tm9uRXhwaXJlZCI6dHJ1ZSwiZGVzY2VuZGFudFNjb3BlcyI6WyIqIl0sImFjY291bnROb25Mb2NrZWQiOnRydWUsImNyZWRlbnRpYWxzTm9uRXhwaXJlZCI6dHJ1ZSwidXNlcm5hbWUiOm51bGwsImF1dGhvcml0aWVzIjpbeyJhdXRob3JpdHkiOiJST0xFX1BPTElORSJ9XSwicGFzc3dvcmQiOm51bGwsImlhdCI6MTU3NDI3NTExMSwiZXhwIjoxNTc5NTM0NzExfQ.zo5MNcCCnxm1U5OHWoHBpkap7-tjbYvvHqqL8xy6Cuw";
            client.Headers[HttpRequestHeader.ContentType] = "application/json";
            client.Headers[HttpRequestHeader.Accept] = "application/json";
            

            Stream requestStream = client.OpenWrite(url, "POST");

            ASCIIEncoding encoding = new ASCIIEncoding();
            string json = "{\"memberState\":\"SE\"}";
            byte[] postByteArray = encoding.GetBytes(json);
            requestStream.Write(postByteArray, 0, postByteArray.Length);
            //using WebClient you cannot read the stream
            //there is no way of reading the response back.
            //does not provide the ability to read the POST back in a streaming fashion. If you want to read the response of the post, 
            //then you have to use one of the other methods which is not stream based, like uploaddata method
            //data is only sent when the req stream is closed below.

            requestStream.Close();

            // to get the data
            requestStream = client.OpenRead(url);
            StreamReader sr = new StreamReader(requestStream);
            string data = sr.ReadToEnd();
            requestStream.Close();

            return data;

        }

    }
}

//StreamWriter sw = new StreamWriter(req.GetRequestStream());
//sw.WriteLine(formData);

//sw.Write(postByteArray);
//sw.Flush();
//sw.Close();

//req.Headers.Add("Host: s-themis-acc.net1.cec.eu.int:8044");

//WebHeaderCollection headers = req.Headers;
//headers.Add("X-Authorization: eyJhbGciOiJIUzI1NiJ9.eyJlbmFibGVkIjp0cnVlLCJhY2NvdW50Tm9uRXhwaXJlZCI6dHJ1ZSwiZGVzY2VuZGFudFNjb3BlcyI6WyIqIl0sImFjY291bnROb25Mb2NrZWQiOnRydWUsImNyZWRlbnRpYWxzTm9uRXhwaXJlZCI6dHJ1ZSwidXNlcm5hbWUiOm51bGwsImF1dGhvcml0aWVzIjpbeyJhdXRob3JpdHkiOiJST0xFX1BPTElORSJ9XSwicGFzc3dvcmQiOm51bGwsImlhdCI6MTU3NDI3NTExMSwiZXhwIjoxNTc5NTM0NzExfQ.zo5MNcCCnxm1U5OHWoHBpkap7-tjbYvvHqqL8xy6Cuw");
//headers.Add("X-Application: themis-inf-test");
//headers.Add("Host: s-themis-acc.net1.cec.eu.int:8044");

//req.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
//req.AuthenticationLevel = System.Net.Security.AuthenticationLevel.None;


//"X-Application: themis-inf-test"
//"Content-Length: 20"
//"Host: s-themis-acc.net1.cec.eu.int:8044"
//"Connection: Keep-Alive"
