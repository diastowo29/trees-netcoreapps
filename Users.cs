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
		public void checkDealer(List<Dictionary<string, string>> entries,string orgId, bool doneOwner, bool donePicOwner, bool donePicOutlet, string jenisGroup, string branchCode, string kuadran, string dealerId, string dealerName, string emailOwner, string emailPicOwner, string emailPicOutlet) {
            Console.WriteLine("doneOwner: {0}, donePicOwner: {1}, donePicOutlet: {2}", doneOwner, donePicOwner, donePicOutlet);
			CallingApi callingApi = new CallingApi();
			string srcUser = "/api/v2/search.json?query=type:user";
			StringBuilder srcParameter = new StringBuilder();
			srcParameter.Append(srcUser);

            List<Dictionary<string,string>> orgMember = new List<Dictionary<string,string>>();
            Dictionary<string,string> members = new Dictionary<string,string>();

            string createUserApi = "/api/v2/users.json";

            if (emailOwner != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailOwner).Append("\"");
            }
            if (emailPicOwner != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailPicOwner).Append("\"");
            }
            if (emailPicOutlet != "#N/A") {
            	srcParameter.Append(" email:\"").Append(emailPicOutlet).Append("\"");
            }
            
            string srcUserResponse = callingApi.callApi(srcParameter.ToString());
            JObject srcUserJoResponse = JObject.Parse(srcUserResponse);
            JArray srcUserArray = (JArray)srcUserJoResponse["results"];
            string ownerFound = "0";
            string picOwnerFound = "0";
            string picOutletFound = "0";

            JObject ownerDict = new JObject();
            JObject picOwnerDict = new JObject();
            JObject picOutletDict = new JObject();

            bool tagOwnerFound = false;
            bool tagPicOwnerFound = false;
            bool tagPicOutletFound = false;

            string userId = "";

            for (int i=0; i<srcUserArray.Count; i++) {
        		if (srcUserArray[i]["email"].ToString().ToLower() == emailOwner.ToLower()) {
        			ownerFound = srcUserArray[i]["id"].ToString();
                    ownerDict = (JObject)srcUserArray[i];
                    JArray tagArray = (JArray)srcUserArray[i]["tags"];
                    for (int t=0 ; t < tagArray.Count; t++) {
                        if (tagArray[t].ToString() ==  "dealer_" + dealerId) {
                            tagOwnerFound = true;
                        }
                    }
        		}

        		if (srcUserArray[i]["email"].ToString().ToLower() == emailPicOwner.ToLower()) {
        			picOwnerFound = srcUserArray[i]["id"].ToString();
                    picOwnerDict = (JObject)srcUserArray[i];
                    JArray tagArray = (JArray)srcUserArray[i]["tags"];
                    for (int t=0 ; t < tagArray.Count; t++) {
                        if (tagArray[t].ToString() ==  "dealer_" + dealerId) {
                            tagPicOwnerFound = true;
                        }
                    }
        		}

        		if (srcUserArray[i]["email"].ToString().ToLower() == emailPicOutlet.ToLower()) {
        			picOutletFound = srcUserArray[i]["id"].ToString();
                    picOutletDict = (JObject)srcUserArray[i];
                    JArray tagArray = (JArray)srcUserArray[i]["tags"];
                    for (int t=0 ; t < tagArray.Count; t++) {
                        if (tagArray[t].ToString() ==  "dealer_" + dealerId) {
                            tagPicOutletFound = true;
                        }
                    }
        		}
            }

            if (!emailOwner.Equals("#N/A")) {
        		Dictionary<string,string> ownerUserCustomField = new Dictionary<string,string>();
                MailAddress emailAddress = new MailAddress(emailOwner);
                string emailUsername = emailAddress.User;
        		Dictionary<string,object> ownerUserField = new Dictionary<string,object>();
                ownerUserField.Add("name", emailUsername);
                ownerUserField.Add("email", emailOwner);
                // ownerUserField.Add("organization_id", orgId);
                ownerUserCustomField.Add("jabatan", "owner");
                ownerUserCustomField.Add("jenis_group", jenisGroup);
                Dictionary<string,object> ownerUser = new Dictionary<string,object>();
            	if (ownerFound != "0") {

                    if (!doneOwner) {
                        ownerUserCustomField.Add("dealer_name", dealerName);
                        ownerUserCustomField.Add("dealer_id", dealerId);
                        ownerUserCustomField.Add("branch_id", branchCode);
                        ownerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        ownerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                    } else {
                        if (!ownerDict["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                            ownerUserCustomField.Add("dealer_id", ownerDict["user_fields"]["dealer_id"] +  ";" + dealerId);
                        }
                        if (!ownerDict["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                            ownerUserCustomField.Add("branch_id", ownerDict["user_fields"]["branch_id"] +  ";" + branchCode);
                        }
                        if (!ownerDict["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                            ownerUserCustomField.Add("kuadran", ownerDict["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                        }
                        if (!ownerDict["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                            ownerUserCustomField.Add("dealer_name", ownerDict["user_fields"]["dealer_name"] +  ";" + dealerName);
                        }
                    }

                    ownerUserField.Add("user_fields", ownerUserCustomField);
                    ownerUser.Add("user", ownerUserField);

                    string updateUserApi = "/api/v2/users/" + ownerFound + ".json";
                    string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(ownerUser), updateUserApi);
                    JObject updateUserResponse = JObject.Parse(updateUser);
                    userId = updateUserResponse["user"]["id"].ToString();
                    // Console.WriteLine(userId);

                    if (doneOwner) {
                        List<string> tagsList = new List<string>();
                        string tagsInput = "" + dealerId + "~" + branchCode + "~" + kuadran;
                        tagsList.Add(tagsInput);
                        Dictionary<string, List<string>> tags = new Dictionary<string, List<string>>();
                        tags.Add("tags", tagsList);
                        string addTagsApi = "/api/v2/users/" + ownerFound + "/tags.json";
                        string addTags = callingApi.callApiPut(JsonConvert.SerializeObject(tags), addTagsApi);
                    }
            	} else {
                    ownerUserCustomField.Add("dealer_id", dealerId);
                    ownerUserCustomField.Add("branch_id", branchCode);
                    ownerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                    ownerUserCustomField.Add("dealer_name", dealerName);
                    ownerUserField.Add("user_fields", ownerUserCustomField);
                    ownerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                    ownerUser.Add("user", ownerUserField);
            		string createUser = callingApi.callApiPost(JsonConvert.SerializeObject(ownerUser), createUserApi);
                    JObject createdUser = JObject.Parse(createUser);
                    userId = createdUser["user"]["id"].ToString();
                    // Console.WriteLine(userId);
            	}
                members.Add("organization_id", orgId);
                members.Add("user_id", userId);
                orgMember.Add(members);
                members = new Dictionary<string,string>();
            }

            if (!emailPicOwner.Equals("#N/A")) {
        		Dictionary<string,string> picOwnerUserCustomField = new Dictionary<string,string>();
                MailAddress emailAddress = new MailAddress(emailPicOwner);
                string emailUsername = emailAddress.User;
                Dictionary<string,object> picOwnerUserField = new Dictionary<string,object>();
                picOwnerUserField.Add("name", emailUsername);
                picOwnerUserField.Add("email", emailPicOwner);
                // picOwnerUserField.Add("organization_id", orgId);
                picOwnerUserCustomField.Add("jabatan", "pic_owner");
                picOwnerUserCustomField.Add("jenis_group", jenisGroup);
                Dictionary<string,object> picOwnerUser = new Dictionary<string,object>();
                if (picOwnerFound != "0") {

                    if (!donePicOwner) {
                        picOwnerUserCustomField.Add("dealer_name", dealerName);
                        picOwnerUserCustomField.Add("dealer_id", dealerId);
                        picOwnerUserCustomField.Add("branch_id", branchCode);
                        picOwnerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        picOwnerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                    } else {
                        if (!picOwnerDict["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                            picOwnerUserCustomField.Add("dealer_id", picOwnerDict["user_fields"]["dealer_id"] +  ";" + dealerId);
                        }
                        if (!picOwnerDict["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                            picOwnerUserCustomField.Add("branch_id", picOwnerDict["user_fields"]["branch_id"] +  ";" + branchCode);
                        }
                        if (!picOwnerDict["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                            picOwnerUserCustomField.Add("kuadran", picOwnerDict["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                        }
                        if (!picOwnerDict["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                            picOwnerUserCustomField.Add("dealer_name", picOwnerDict["user_fields"]["dealer_name"] +  ";" + dealerName);
                        }
                    }

                    picOwnerUserField.Add("user_fields", picOwnerUserCustomField);
                    picOwnerUser.Add("user", picOwnerUserField);

                    string updateUserApi = "/api/v2/users/" + picOwnerFound + ".json";
                    string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(picOwnerUser), updateUserApi);
                    JObject updateUserResponse = JObject.Parse(updateUser);
                    userId = updateUserResponse["user"]["id"].ToString();
                    // Console.WriteLine(userId);

                    if (donePicOwner) {
                        List<string> tagsList = new List<string>();
                        string tagsInput = "" + dealerId + "~" + branchCode + "~" + kuadran;
                        tagsList.Add(tagsInput);
                        Dictionary<string, List<string>> tags = new Dictionary<string, List<string>>();
                        tags.Add("tags", tagsList);
                        string addTagsApi = "/api/v2/users/" + picOwnerFound + "/tags.json";
                        string addTags = callingApi.callApiPut(JsonConvert.SerializeObject(tags), addTagsApi);
                    }
                } else {
                    picOwnerUserCustomField.Add("dealer_id", dealerId);
                    picOwnerUserCustomField.Add("branch_id", branchCode);
                    picOwnerUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                    picOwnerUserCustomField.Add("dealer_name", dealerName);
                    picOwnerUserField.Add("user_fields", picOwnerUserCustomField);
                    picOwnerUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                    picOwnerUser.Add("user", picOwnerUserField);
                    string createUser = callingApi.callApiPost(JsonConvert.SerializeObject(picOwnerUser), createUserApi);
                    JObject createdUser = JObject.Parse(createUser);
                    userId = createdUser["user"]["id"].ToString();
                }
                members.Add("organization_id", orgId);
                members.Add("user_id", userId);
                orgMember.Add(members);
                members = new Dictionary<string,string>();
            }


            if (!emailPicOutlet.Equals("#N/A")) {
        		Dictionary<string,string> picOutletUserCustomField = new Dictionary<string,string>();
                MailAddress emailAddress = new MailAddress(emailPicOutlet);
                string emailUsername = emailAddress.User;
                Dictionary<string,object> picOutletUserField = new Dictionary<string,object>();
                picOutletUserField.Add("name", emailUsername);
                picOutletUserField.Add("email", emailPicOutlet);
                // picOutletUserField.Add("organization_id", orgId);
                picOutletUserCustomField.Add("jabatan", "pic_outlet");
                picOutletUserCustomField.Add("jenis_group", jenisGroup);
                Dictionary<string,object> picOutletUser = new Dictionary<string,object>();
                if (picOutletFound != "0") {

                    if (!donePicOutlet) {
                        picOutletUserCustomField.Add("dealer_name", dealerName);
                        picOutletUserCustomField.Add("dealer_id", dealerId);
                        picOutletUserCustomField.Add("branch_id", branchCode);
                        picOutletUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                        picOutletUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                    } else {
                        if (!picOutletDict["user_fields"]["dealer_id"].ToString().Contains(dealerId)) {
                            picOutletUserCustomField.Add("dealer_id", picOutletDict["user_fields"]["dealer_id"] +  ";" + dealerId);
                        }
                        if (!picOutletDict["user_fields"]["branch_id"].ToString().Contains(branchCode)) {
                            picOutletUserCustomField.Add("branch_id", picOutletDict["user_fields"]["branch_id"] +  ";" + branchCode);
                        }
                        if (!picOutletDict["user_fields"]["kuadran"].ToString().Contains(kuadran.ToString().ToLower())) {
                            picOutletUserCustomField.Add("kuadran", picOutletDict["user_fields"]["kuadran"] +  ";" + "kuadran_" + kuadran.ToLower());
                        }
                        if (!picOutletDict["user_fields"]["dealer_name"].ToString().Contains(dealerName)) {
                            picOutletUserCustomField.Add("dealer_name", picOutletDict["user_fields"]["dealer_name"] +  ";" + dealerName);
                        }
                    }

                    picOutletUserField.Add("user_fields", picOutletUserCustomField);
                    picOutletUser.Add("user", picOutletUserField);

                    string updateUserApi = "/api/v2/users/" + picOutletFound + ".json";
                    string updateUser = callingApi.callApiPut(JsonConvert.SerializeObject(picOutletUser), updateUserApi);
                    JObject updateUserResponse = JObject.Parse(updateUser);
                    userId = updateUserResponse["user"]["id"].ToString();
                    // Console.WriteLine(userId);
                    if (donePicOutlet) {
                        List<string> tagsList = new List<string>();
                        string tagsInput = "" + dealerId + "~" + branchCode + "~" + kuadran;
                        tagsList.Add(tagsInput);
                        Dictionary<string, List<string>> tags = new Dictionary<string, List<string>>();
                        tags.Add("tags", tagsList);
                        string addTagsApi = "/api/v2/users/" + picOutletFound + "/tags.json";
                        string addTags = callingApi.callApiPut(JsonConvert.SerializeObject(tags), addTagsApi);
                    }
                } else {
                    picOutletUserCustomField.Add("dealer_id", dealerId);
                    picOutletUserCustomField.Add("branch_id", branchCode);
                    picOutletUserCustomField.Add("kuadran", "kuadran_" + kuadran.ToLower());
                    picOutletUserCustomField.Add("dealer_name", dealerName);
                    picOutletUserField.Add("user_fields", picOutletUserCustomField);
                    picOutletUserField.Add("tags",  "" + dealerId + "~" + branchCode + "~" + kuadran);
                    picOutletUser.Add("user", picOutletUserField);
                    string createUser = callingApi.callApiPost(JsonConvert.SerializeObject(picOutletUser), createUserApi);
                    JObject createdUser = JObject.Parse(createUser);
                    userId = createdUser["user"]["id"].ToString();
                }
                members.Add("organization_id", orgId);
                members.Add("user_id", userId);
                orgMember.Add(members);
                members = new Dictionary<string,string>();
            }
            if (orgMember.Count > 0) {
                createOrgMemberships(orgMember);
                for (int i=0; i<orgMember.Count; i++) {
                    checkUserMemberships(orgMember[i], entries);
                }
            }
        }
        public void createOrgMemberships(List<Dictionary<string,string>> orgMember) {
            CallingApi callingApi = new CallingApi();
            string createManyOrgMembershipApi = "/api/v2/organization_memberships/create_many.json";
            Dictionary<string,object> orgMembership = new Dictionary<string,object>();
            orgMembership.Add("organization_memberships", orgMember);
            string createMembership = callingApi.callApiPost(JsonConvert.SerializeObject(orgMembership), createManyOrgMembershipApi);
        }

        public void checkUserMemberships(Dictionary<string,string> userMember, List<Dictionary<string, string>> entries) {
            CallingApi callingApi = new CallingApi();
            string checkMembershipApi = "/api/v2/users/" + userMember["user_id"].ToString() + "/organization_memberships.json?include=organizations,users";
            string checkMembership = callingApi.callApi(checkMembershipApi);
            JObject checkMembershipResponse = JObject.Parse(checkMembership);

            JArray membershipArray = (JArray)checkMembershipResponse["organization_memberships"];
            JArray organizationArray = (JArray)checkMembershipResponse["organizations"];
            JArray usersArray = (JArray)checkMembershipResponse["users"];
            string userEmail = usersArray[0]["email"].ToString();

            StringBuilder deleteList = new StringBuilder();

            if (membershipArray.Count > 1) {
                for (int i=0; i<membershipArray.Count; i++) {
                    bool orgFound = false;
                    if (organizationArray[i]["name"].ToString().StartsWith("D ")) {
                        string[] orgBranchCode = organizationArray[i]["name"].ToString().Split(" ");

                        for (int j=0; j<entries.Count; j++) {
                            if (userEmail.ToLower().Equals(entries[j]["Alamat email PIC Owner"].ToString().ToLower())) {
                                if (entries[j]["Branch Code"].ToString().Equals(orgBranchCode[1])) {
                                    orgFound = true;
                                }
                            }
                            if (userEmail.ToLower().Equals(entries[j]["Alamat email owner"].ToString().ToLower())) {
                                if (entries[j]["Branch Code"].ToString().Equals(orgBranchCode[1])) {
                                    orgFound = true;
                                }
                            }
                            if (userEmail.ToLower().Equals(entries[j]["Alamat email PIC Outlet"].ToString().ToLower())) {
                                if (entries[j]["Branch Code"].ToString().Equals(orgBranchCode[1])) {
                                    orgFound = true;
                                }
                            }
                        }
                    }
                    if (!orgFound) {
                        deleteList.Append(membershipArray[i]["id"].ToString()).Append(",");
                    }
                }
                // Console.WriteLine(JsonConvert.SerializeObject(deleteList));
                if (deleteList.ToString().Length > 0) {
                    deleteMembership(deleteList.ToString().Remove((deleteList.ToString().Length)-1, 1));
                }
            }
        }

        public void deleteMembership(string deleteList) {
            string deleteManyOrgMemberApi = "/api/v2/organization_memberships/destroy_many.json?ids=" + deleteList;
            CallingApi callingApi = new CallingApi();
            callingApi.callApiDelete(deleteManyOrgMemberApi);
        }
	}
}