using System;
using System.Net;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace StormpathAPIDemo
{
    class StormpathAuth
    {
        private string _appId;					 // This is the ID found in your REST API URLs 
        private string _appKeyId;				 // Get this when you create a key
        private string _appKeySecret;			 // Get this when you create a key
        private string _accountURL;
        private string _loginAttemptsURL;
        private NetworkCredential _stormCred;
        private CredentialCache _stormCredCache; // Since posting JSON will require auth us to auth

        public StormpathAuth(string appId, string appKeyId, string appKeySecret)
        {
            _appId = appId;
            _appKeyId = appKeyId;
            _appKeySecret = appKeySecret;

            /* Maybe in the future, Stormpath might change it's URL path (I mean, there's versioning in it!) 
             * For now, I'm keeping it here, but will be updated if changes occur */
            _accountURL = "https://api.stormpath.com/v1/applications/" + _appId + "/accounts";
            _loginAttemptsURL = "https://api.stormpath.com/v1/applications/" + _appId + "/loginAttempts";

            _stormCred = new NetworkCredential(_appKeyId, _appKeySecret);
            _stormCredCache = new CredentialCache();
            _stormCredCache.Add(new Uri("https://api.stormpath.com/v1/applications/" + _appId), "Basic", _stormCred);
        }

        public bool Login(string username, string password)
        {
            /* JSON POST to login a user */

            HttpWebRequest httpReq = PrepareWebRequest(_loginAttemptsURL);
            HttpWebResponse httpResp = null;
            string combo = GetBase64String(username + ":" + password);
            var json = new JavaScriptSerializer().Serialize(
                new { type = "basic", value = combo }
            );

            if (SendWebRequest(httpReq, httpResp, json))
            {
                //FUTURE: Do something with httpResp
                return true;
            }
            else
            {
                //FUTURE: Do something with httpResp, like look into why Login failed...
                return false;
            }
        }

        public bool CreateAccount(string _username, string _email, string _password, string _givenname = "default", string _surname = "default")
        {
            /* JSON POST to create a user account */

            HttpWebRequest httpReq = PrepareWebRequest(_accountURL);
            HttpWebResponse httpResp = null;

            var json = new JavaScriptSerializer().Serialize(
                new { givenName = _givenname, surname = _surname, username = _username, email = _email, password = _password }
            );

            if (SendWebRequest(httpReq, httpResp, json))
            {
                //FUTURE: Do something with httpResp
                return true;
            }
            else
            {
                //FUTURE: Do something with httpResp, like look into why CreateAccount failed...
                return false;
            }
        }

        private string GetBase64String(string input)
        {
            return Convert.ToBase64String(Encoding.ASCII.GetBytes(input));
        }

        private HttpWebRequest PrepareWebRequest(string URL)
        {
            var httpReq = (HttpWebRequest)WebRequest.Create(URL);
            httpReq.ContentType = "application/json";
            httpReq.Method = "POST";
            httpReq.Credentials = _stormCredCache;
            return httpReq;
        }

        private bool SendWebRequest(HttpWebRequest httpReq, HttpWebResponse httpResp, string json)
        {
            using (var writer = new StreamWriter(httpReq.GetRequestStream()))
            {
                writer.Write(json); writer.Flush(); writer.Close();

                try
                {
                    httpResp = (HttpWebResponse)httpReq.GetResponse();
                    return true;

                }
                catch (Exception ex)
                {
                    return false;
                }
            }
        }

    }
}