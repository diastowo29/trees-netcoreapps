using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Authenticators;
using System.Net.Mail;

using System.Net.Http;
using System.Net;

namespace newConsole {
	class Users {
		public void checkDealer(string orgId, string dealerId, string dealerName, string emailOwner, string emailPicOwner, string emailPicOutlet) {
			CallingApi callingApi = new CallingApi();
			string srcUser = "/api/v2/search.json?query=type:user";
			StringBuilder srcParameter = new StringBuilder();
			srcParameter.Append(srcUser);

            string createUser = "/api/v2/users.json";

            Dictionary<string, object> usersDict = new Dictionary<string,object>();
            Dictionary<string, string> userFields = new Dictionary<string,string>();
            Dictionary<string, object> newUser = new Dictionary<string,object>();
            if (emailOwner != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailOwner).Append("\"");

                userFields = new Dictionary<string,string>();
                newUser = new Dictionary<string,object>();
                userFields.Add("dealer_id", dealerId);
                userFields.Add("jabatan", "Owner");
                userFields.Add("kuadran", "Q1");
                userFields.Add("dealer_name", dealerName);
                MailAddress emailAddress = new MailAddress(emailOwner);
                string emailUsername = emailAddress.User;
                newUser.Add("name", emailUsername);
                newUser.Add("email", emailOwner);
                newUser.Add("user_fields", userFields);
                userList.Add(newUser);
            }
            if (emailPicOwner != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailPicOwner).Append("\"");

                userFields = new Dictionary<string,string>();
                newUser = new Dictionary<string,object>();
                userFields.Add("dealer_id", dealerId);
                userFields.Add("jabatan", "PIC Owner");
                userFields.Add("kuadran", "Q1");
                userFields.Add("dealer_name", dealerName);
                MailAddress emailAddress = new MailAddress(emailPicOwner);
                string emailUsername = emailAddress.User;
                newUser.Add("name", emailUsername);
                newUser.Add("email", emailPicOwner);
                newUser.Add("user_fields", userFields);
                userList.Add(newUser);
            }
            if (emailPicOutlet != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailPicOutlet).Append("\"");

                userFields = new Dictionary<string,string>();
                newUser = new Dictionary<string,object>();
                userFields.Add("dealer_id", dealerId);
                userFields.Add("jabatan", "PIC Outlet");
                userFields.Add("kuadran", "Q1");
                userFields.Add("dealer_name", dealerName);
                MailAddress emailAddress = new MailAddress(emailPicOutlet);
                string emailUsername = emailAddress.User;
                newUser.Add("name", emailUsername);
                newUser.Add("email", emailPicOutlet);
                newUser.Add("user_fields", userFields);
                userList.Add(newUser);
            }
            
            string srcUserResponse = callingApi.callApi(srcParameter);
            JObject srcUserJoResponse = JObject.Parse(srcUserResponse);
            JArray srcUserArray = (JArray)srcUserJoResponse["results"];
            for (int i=0; i<srcUserArray.Count; i++) {
            	
            }
            // userFields.Add("dealer_id", dealerId);
            // userFields.Add("email_owner", emailOwner);
            // userFields.Add("email_pic_owner", emailPicOwner);
            // userFields.Add("email_pic_outlet", emailPicOutlet);
            // newUser.Add("name", dealerName);
            // newUser.Add("user_fields", userFields);

            usersDict.Add("users", userList);
            Console.WriteLine(JsonConvert.SerializeObject(usersDict));

            // string createUpdateUser = callApiPost();
        }
	}
}