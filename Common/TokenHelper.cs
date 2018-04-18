using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Security;

namespace Common
{
    public static class TokenHelper
    {

        public static string checkAuthorizedUser(string name, string upassword)
        {
            string _userNameAdmin = ConfigurationManager.AppSettings["DomainAdmin"];
            string _userPasswordAdmin = ConfigurationManager.AppSettings["DomainAdminPassword"];
            // bool Authorized = false;
            string UserLoggedInName = string.Empty;
            try
            {
                using (ProjectContext ctx = new ProjectContext(ConfigurationManager.AppSettings["PPMServerURL"]))
                {

                    SecureString passWord = new SecureString();
                    foreach (char c in _userPasswordAdmin) passWord.AppendChar(c);
                    ctx.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);


                    var user = ctx.Web.EnsureUser(name);
                    ctx.Load(user);
                    ctx.ExecuteQuery();

                    if (user != null)
                    {

                        // Authorized = true;
                        UserLoggedInName = user.Title;

                    }
                    //else
                    //    Authorized = false;
                }
            }
            catch (Exception ex)
            {
                //UserLoggedInName = ex.Message;
                UserLoggedInName = string.Empty;
                //Authorized = false;
            }

            //return Authorized;
            return UserLoggedInName;
        }


        public static string Datevalues(object obj, string keyNeed)
        {
            string keyval = string.Empty;
            if (typeof(IDictionary).IsAssignableFrom(obj.GetType()))
            {
                IDictionary idict = (IDictionary)obj;

                Dictionary<string, string> newDict = new Dictionary<string, string>();
                foreach (object key in idict.Keys)
                {
                    if (keyNeed == key.ToString())
                    {
                        keyval = idict[key].ToString();
                        //newDict.Add(key.ToString(), idict[key].ToString());
                        break;
                    }
                }
            }
            return keyval;

        }
    }
}