using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Net.Mail;
using System.Net;

using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Authenticators;

// using Organizations;

namespace newConsole
{
    class Program
    {
        static string zendeskTeamLeaderRoleId = "11434967"; /*Team Leader*/
        // static string zendeskTeamLeaderRoleId = "8595408"; /*Staff*/

        static string zendeskUsername = "";
        static string zendeskPassword = "";

        static string originDirectory = "D:/WORK/Dotnet/trees-netcoreapps/ZD Dealer Survey/Upload/";
        static string destDirectory = "D:/WORK/Dotnet/trees-netcoreapps/ZD Dealer Survey/Archive/";

        // static string originDirectory = "C:/ZD Dealer Survey/Upload/";
        // static string destDirectory = "C:/ZD Dealer Survey/Archive/";

        static string emailDomain = "@example.com";
        // static string emailDomain = "@fifgroup.co.id";
        
        // static string originDirectory = "doc/";
        // static string destDirectory = "doc_done/";

        static List<string> userList = new List<string>();
        static string supportGroupId = "";
        static List<Dictionary<string,string>> doneList = new List<Dictionary<string,string>>();

        // static int excelLimit = -1;

        static List<JToken> allGroups = new List<JToken>();
        
        static void Main(string[] args)
        {
            List<Dictionary<string,string>> errorList = new List<Dictionary<string,string>>();
            Program myProgram = new Program();
            try {
                zendeskUsername = args[0];
                zendeskPassword = args[1];
            } catch {
                Console.WriteLine("Wrong Username or Password.. Program will exit");
                myProgram.createLog(errorList);
                Environment.Exit(0);
            }

            // // // // // deleteGroups();
            initiate();
        	// System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        	string[] array1 = Directory.GetFiles(originDirectory);
        	foreach (string filePath in array1) {
                File.Copy(filePath, destDirectory + Path.GetFileName(filePath), true);
        		if (filePath.Contains("xlsx")) {
        			doXlsx(filePath, errorList);
        		} else if (filePath.Contains("xls")) {
        			doXls(filePath, errorList);
        		}
                // File.Delete(filePath);
        	}
            if (errorList.Count > 0) {
                myProgram.createLog(errorList);
            }
            // Console.Read();

        }

        public static void doXlsx (string filePath, List<Dictionary<string,string>> errorList) {
            Dictionary<string, string> mappingList = new Dictionary<string, string>();
        	Console.WriteLine("===== DO XLSX =====");
            List<String> keys = new List<String>();
            List<Dictionary<string, string>> mappingArray = new List<Dictionary<string, string>>();
            int skipIndex = 0;

            try {
            	Console.WriteLine(filePath);
            	var package = new ExcelPackage(new FileInfo(filePath));
    			ExcelWorksheet sheet = package.Workbook.Worksheets[1];
    			var rowCount = sheet.Dimension.End.Row;
    			var colCount = sheet.Dimension.End.Column;
    			for (int i=1; i<=rowCount; i++) {
                    mappingList = new Dictionary<string, string>();
    				// Console.WriteLine("===== NEW ROW =====");
                    int myColCounter = 0;

    				for (int j=1; j<=colCount; j++) {
                        if (i == 1) {
                            // bool keyExist = false;
                            for (int k=0; k<keys.Count; k++) {
                                if (keys[k] == sheet.Cells[i,j].Value.ToString()) {
                                    skipIndex = j;
                                }
                            }
                            keys.Add(sheet.Cells[i,j].Value.ToString());
                        } else {
                            if (j != skipIndex) {
                                string values = "";
                                if (sheet.Cells[i,j].Value == null) {
                                    values = "#N/A";
                                } else {
                                    values = sheet.Cells[i,j].Value.ToString();
                                }
                                mappingList.Add(keys[j-1], values);
                                myColCounter++;
                            }
                        }
    				}
                    if (i != 1) {
                        mappingArray.Add(mappingList);
                    }
    			}
                doProcessMapping(mappingArray, errorList);
                // string jsonString = JsonConvert.SerializeObject(mappingArray);
                // Console.WriteLine(jsonString);
            } catch (Exception e) {
                Console.WriteLine("file error..");
                Console.WriteLine(e);
                Dictionary<string,string> newError = new Dictionary<string,string>();
                newError.Add("file_error", "file: " + filePath + " has different format or has an error");
                errorList.Add(newError);
            }
        }

        public static void doXls (string filePath, List<Dictionary<string,string>> errorList) {
            Dictionary<string, string> mappingList = new Dictionary<string, string>();
            List<Dictionary<string, string>> mappingArray = new List<Dictionary<string, string>>();
            var rowCount = 0;
            List<String> keys = new List<String>();
            int skipIndex = 0;
            try {
            	Console.WriteLine("===== DO XLS =====");
            	Console.WriteLine(filePath);
            	using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read)) {
    				using (var reader = ExcelReaderFactory.CreateReader(stream)) {
    					do {
    						while (reader.Read()) {
                                rowCount++;
                                mappingList = new Dictionary<string, string>();
    							// Console.WriteLine("===== NEW ROW =====");
    							for (int i=0; i<reader.FieldCount; i++) {
                                    if (rowCount == 1) {
                                        for (int k=0; k<keys.Count; k++) {
                                            if (keys[k] == reader.GetValue(i).ToString()) {
                                                skipIndex = i;
                                            }
                                        }
                                        keys.Add(reader.GetValue(i).ToString());
                                    } else {
                                        if (i != skipIndex) {
                                            string values = "";
                                            if (reader.GetValue(i) == null) {
                                                values = "#N/A";
                                            } else {
                                                values = reader.GetValue(i).ToString();
                                            }
                                            mappingList.Add(keys[i], values);
                                        }
                                    }
    							}
                                if (rowCount != 1) {
                                    mappingArray.Add(mappingList);
                                }
    						}
                            doProcessMapping(mappingArray, errorList);
                            // string jsonString = JsonConvert.SerializeObject(mappingArray);
                            // Console.WriteLine(jsonString);
    					} while (reader.NextResult());
    				}
    			}                
            } catch {
                Console.WriteLine("file error..");
                Dictionary<string,string> newError = new Dictionary<string,string>();
                newError.Add("file_error", "file: " + filePath + " has different format or has an error");
                errorList.Add(newError);
            }
        }

        public void createLog(List<Dictionary<string,string>> errorList) {
            DateTime dateTime = DateTime.Now;
            string fileName = dateTime.ToString().Replace("/", ".").Replace(":", ".");

            // string newFilePath = @"C:\Users\Administrator\Documents\Zendesk\Deploy Prod\Dias Prod\logs\" + fileName + ".csv";
            string newFilePath = @"D:\WORK\Dotnet\trees-netcoreapps\logs\" + fileName + ".csv";

            StringBuilder sb = new StringBuilder();
            if (errorList.Count > 0) {
                for (int i=0; i<errorList.Count; i++) {
                    if (errorList[i].ContainsKey("file_error")) {
                        sb.AppendLine(errorList[i]["file_error"]);
                    } else if (errorList[i].ContainsKey("user_error")) {
                        sb.AppendLine(errorList[i]["user_error"]);
                    }
                }
            } else {
                sb.Append("Argument Error");
            }
            File.WriteAllText(newFilePath, sb.ToString());
        }

        public static void doProcessMapping(List<Dictionary<string, string>> entries, List<Dictionary<string,string>> errorList){
            Organizations myOrg = new Organizations();
            Users users = new Users();
            Groups myGroup = new Groups();

            string groupMMid = "0";
            string groupAHid = "0";
            string groupDHid = "0";

            string userMmId = "0";
            string userDhId = "0";
            string userAhId = "0";

            string orgId = "0";

            doneList = new List<Dictionary<string,string>>();
            List<JToken> userCreated = new List<JToken>();
            Dictionary<string,string> groupList = new Dictionary<string, string>();
            Dictionary<string,string> agentList = new Dictionary<string, string>();
            List<string> doneAgent = new List<string>();
            List<Dictionary<string,string>> agentGroupList = new List<Dictionary<string, string>>();

            for (int i=0; i<entries.Count; i++) {
                Dictionary<string,string> userDone = new Dictionary<string,string>();
                List<JToken> userJustCreated = new List<JToken>();
                Console.WriteLine("=========== NEW ROW ==========");
                string groupMM = "Group D MM " + entries[i]["Branch Code"].ToString();
                string groupAH = "Group D Area Head " + entries[i]["AREA"].ToString();
                string groupDH = "Group D Dept Head " + entries[i]["Region Dept Head"].ToString();
                string orgName = "D " + entries[i]["Branch Code"].ToString();

                string namaMm = entries[i]["MM"].ToString();
                string namaDh = entries[i]["Dept. Head"].ToString();
                string namaAh = entries[i]["NAMA Area Head"].ToString();

                string emailMm = entries[i]["NPK MM"].ToString() + emailDomain;
                string emailDh = entries[i]["NPK Dept. Head"].ToString() + emailDomain;
                string emailAh = entries[i]["NPK Area Head"].ToString() + emailDomain;

                string newUserMmId = "0";
                string newUserDhId = "0";
                string newUserAhId = "0";

                string newGroupMMid = "nil";
                string newGroupDHid = "nil";
                string newGroupAHid = "nil";

                string emailPicOwner = entries[i]["Alamat email PIC Owner"].ToString();
                string emailOwner = entries[i]["Alamat email owner"].ToString();
                string emailPicOutlet = entries[i]["Alamat email PIC Outlet"].ToString();

                Dictionary<string,string> agentMembership = new Dictionary<string,string>();
                List<Dictionary<string,string>> agentMembershipList = new List<Dictionary<string,string>>();
                List<string> deleteGroupList = new List<string>();

                Program newCon = new Program();

                // if (i < excelLimit) {
                    /*NEW LOGIC*/
                    groupList = myGroup.searchGroup(groupMM, groupDH, groupAH, zendeskUsername, zendeskPassword);
                    Console.WriteLine(JsonConvert.SerializeObject(groupList));
                    newGroupMMid = groupList["groupMMid"].ToString();
                    newGroupDHid = groupList["groupDHid"].ToString();
                    newGroupAHid = groupList["groupAHid"].ToString();

                    if (newGroupMMid.Equals("0")) {
                        if (i == 0) {
                            /*DO CREATE GROUPS*/
                            newGroupMMid = myGroup.createGroup(groupMM, zendeskUsername, zendeskPassword);
                            groupMMid = newGroupMMid;
                        } else {
                            if (entries[i]["Branch Code"] == entries[i-1]["Branch Code"]) {
                                newGroupMMid = groupMMid;
                            } else {
                                /*DO CREATE GROUPS*/
                                newGroupMMid = myGroup.createGroup(groupMM, zendeskUsername, zendeskPassword);
                                groupMMid = newGroupMMid;
                            }
                        }
                    }

                    if (newGroupDHid.Equals("0")) {
                        if (i == 0) {
                            /*DO CREATE GROUPS*/
                            newGroupDHid = myGroup.createGroup(groupDH, zendeskUsername, zendeskPassword);
                            groupDHid = newGroupDHid;
                        } else {
                            if (entries[i]["Region Dept Head"] == entries[i-1]["Region Dept Head"]) {
                                newGroupDHid = groupDHid;
                            } else {
                                /*DO CREATE GROUPS*/
                                newGroupDHid = myGroup.createGroup(groupDH, zendeskUsername, zendeskPassword);
                                groupDHid = newGroupDHid;
                            }
                        }
                    }

                    if (newGroupAHid.Equals("0")) {
                        if (i == 0) {
                            /*DO CREATE GROUPS*/
                            newGroupAHid = myGroup.createGroup(groupAH, zendeskUsername, zendeskPassword);
                            groupAHid = newGroupAHid;
                        } else {
                            if (entries[i]["AREA"] == entries[i-1]["AREA"]) {
                                newGroupAHid = groupAHid;
                            } else {
                                /*DO CREATE GROUPS*/
                                newGroupAHid = myGroup.createGroup(groupAH, zendeskUsername, zendeskPassword);
                                groupAHid = newGroupAHid;
                            }
                        }
                    }

                    agentList = users.searchAgent(emailMm, emailDh, emailAh, zendeskUsername, zendeskPassword);
                    newUserMmId = agentList["userMmId"].ToString();
                    newUserDhId = agentList["userDhId"].ToString();
                    newUserAhId = agentList["userAhId"].ToString();

                    if (namaMm != "#N/A") {
                        if (newUserMmId == "0") {
                            if (i == 0) {
                                /*DO CREATE AGENT*/
                                newUserMmId = users.createAgent(namaMm, emailMm, newGroupMMid, zendeskTeamLeaderRoleId, zendeskUsername, zendeskPassword);
                                userMmId = newUserMmId;
                                if (newUserMmId == "0") {
                                    Dictionary<string,string> newError = new Dictionary<string,string>();
                                    newError.Add("user_error", "row: " + i + ", error when create agent: " + emailMm);
                                    errorList.Add(newError);
                                }
                            } else {
                                if (entries[i]["MM"] == entries[i-1]["MM"]) {
                                    newUserMmId = userMmId;
                                    agentMembership.Add("group_id", newGroupMMid);
                                    agentMembership.Add("user_id", newUserMmId);
                                    agentMembershipList.Add(agentMembership);
                                    agentGroupList.Add(agentMembership);
                                    agentMembership = new Dictionary<string,string>();
                                } else {
                                    /*DO CREATE AGENT*/
                                    newUserMmId = users.createAgent(namaMm, emailMm, newGroupMMid, zendeskTeamLeaderRoleId, zendeskUsername, zendeskPassword);
                                    userMmId = newUserMmId;
                                    if (newUserMmId == "0") {
                                        Dictionary<string,string> newError = new Dictionary<string,string>();
                                        newError.Add("user_error", "row: " + i + ", error when create agent: " + emailMm);
                                        errorList.Add(newError);
                                    }
                                }
                            }
                        } else {
                            userMmId = newUserMmId;
                            agentMembership.Add("group_id", newGroupMMid);
                            agentMembership.Add("user_id", newUserMmId);
                            agentMembershipList.Add(agentMembership);
                            agentGroupList.Add(agentMembership);
                            agentMembership = new Dictionary<string,string>();
                        }
                    } else {
                        deleteGroupList.Add(newGroupMMid);
                        // myGroup.deleteAllMember(newGroupMMid);
                        /*DELETE ALL MEMBERLIST*/
                    }

                    if (namaDh != "#N/A") {
                        if (newUserDhId == "0") {
                            if (i == 0) {
                                /*DO CREATE AGENT*/
                                newUserDhId = users.createAgent(namaDh, emailDh, newGroupDHid, zendeskTeamLeaderRoleId, zendeskUsername, zendeskPassword);
                                userDhId = newUserDhId;
                                if (newUserDhId == "0") {
                                    Dictionary<string,string> newError = new Dictionary<string,string>();
                                    newError.Add("user_error", "row: " + i + ", error when create agent: " + emailDh);
                                    errorList.Add(newError);
                                }
                            } else {
                                if (entries[i]["Dept. Head"] == entries[i-1]["Dept. Head"]) {
                                    newUserDhId = userDhId;
                                    agentMembership.Add("group_id", newGroupDHid);
                                    agentMembership.Add("user_id", newUserDhId);
                                    agentMembershipList.Add(agentMembership);
                                    agentGroupList.Add(agentMembership);
                                    agentMembership = new Dictionary<string,string>();
                                } else {
                                    /*DO CREATE AGENT*/
                                    newUserDhId = users.createAgent(namaDh, emailDh, newGroupDHid, zendeskTeamLeaderRoleId, zendeskUsername, zendeskPassword);
                                    userDhId = newUserDhId;
                                    if (newUserDhId == "0") {
                                        Dictionary<string,string> newError = new Dictionary<string,string>();
                                        newError.Add("user_error", "row: " + i + ", error when create agent: " + emailDh);
                                        errorList.Add(newError);
                                    }
                                }
                            }
                        } else {
                            agentMembership.Add("group_id", newGroupDHid);
                            agentMembership.Add("user_id", newUserDhId);
                            agentMembershipList.Add(agentMembership);
                            agentGroupList.Add(agentMembership);
                            agentMembership = new Dictionary<string,string>();
                        }
                    } else {
                        deleteGroupList.Add(newGroupDHid);
                        // myGroup.deleteAllMember(newGroupDHid);
                        /*DELETE ALL MEMBERLIST*/
                    }

                    if (namaAh != "#N/A") {
                        if (newUserAhId == "0") {
                            if (i == 0) {
                                /*DO CREATE AGENT*/
                                newUserAhId = users.createAgent(namaAh, emailAh, newGroupAHid, zendeskTeamLeaderRoleId, zendeskUsername, zendeskPassword);
                                userAhId = newUserAhId;
                                if (newUserAhId == "0") {
                                    Dictionary<string,string> newError = new Dictionary<string,string>();
                                    newError.Add("user_error", "row: " + i + ", error when create agent: " + emailAh);
                                    errorList.Add(newError);
                                }
                            } else {
                                if (entries[i]["NAMA Area Head"] == entries[i-1]["NAMA Area Head"]) {
                                    newUserAhId = userAhId;
                                    agentMembership.Add("group_id", newGroupAHid);
                                    agentMembership.Add("user_id", newUserAhId);
                                    agentMembershipList.Add(agentMembership);
                                    agentGroupList.Add(agentMembership);
                                    agentMembership = new Dictionary<string,string>();
                                } else {
                                    /*DO CREATE AGENT*/
                                    newUserAhId = users.createAgent(namaAh, emailAh, newGroupAHid, zendeskTeamLeaderRoleId, zendeskUsername, zendeskPassword);
                                    userAhId = newUserAhId;
                                    if (newUserAhId == "0") {
                                        Dictionary<string,string> newError = new Dictionary<string,string>();
                                        newError.Add("user_error", "row: " + i + ", error when create agent: " + emailAh);
                                        errorList.Add(newError);
                                    }
                                }
                            }
                        } else {
                            agentMembership.Add("group_id", newGroupAHid);
                            agentMembership.Add("user_id", newUserAhId);
                            agentMembershipList.Add(agentMembership);
                            agentGroupList.Add(agentMembership);
                            agentMembership = new Dictionary<string,string>();
                        }
                    } else {
                        deleteGroupList.Add(newGroupAHid);
                        // myGroup.deleteAllMember(newGroupAHid);
                        /*DELETE ALL MEMBERLIST*/
                    }
                    if (deleteGroupList.Count > 0) {
                        myGroup.deleteAllMember(deleteGroupList, zendeskUsername, zendeskPassword);
                    }
                    
                    myGroup.addMembership(agentMembershipList, agentGroupList, zendeskUsername, zendeskPassword);
                    if (!doneAgent.Contains(newUserMmId)) {
                        myGroup.checkAgentMembership(newUserMmId, newGroupMMid, entries, "MM", zendeskUsername, zendeskPassword);
                        doneAgent.Add(newUserMmId);
                    }
                    if (!doneAgent.Contains(newUserDhId)) {
                        myGroup.checkAgentMembership(newUserDhId, newGroupDHid, entries, "DH", zendeskUsername, zendeskPassword);
                        doneAgent.Add(newUserDhId);
                    }
                    if (!doneAgent.Contains(newUserAhId)) {
                        myGroup.checkAgentMembership(newUserAhId, newGroupAHid, entries, "AH", zendeskUsername, zendeskPassword);
                        doneAgent.Add(newUserAhId);
                    }

                    if (i > 0) {
                        if (entries[i]["Branch Code"] != entries[i-1]["Branch Code"]) {
                            orgId = myOrg.searchOrganizations(orgName, newGroupMMid, newUserMmId, newGroupDHid, newUserDhId, newGroupAHid, newUserAhId, zendeskUsername, zendeskPassword);
                        }
                    } else {
                        orgId = myOrg.searchOrganizations(orgName, newGroupMMid, newUserMmId, newGroupDHid, newUserDhId, newGroupAHid, newUserAhId, zendeskUsername, zendeskPassword);
                    }

                    if (i == 0) {
                        userJustCreated = users.checkDealer(userCreated, entries, orgId, false, false, false, entries[i], zendeskUsername, zendeskPassword);
                        for (int j=0; j<userJustCreated.Count; j++) {
                            userCreated.Add(userJustCreated[j]);
                        }
                        userList.Add(emailOwner);
                        userList.Add(emailPicOwner);
                        userList.Add(emailPicOutlet);
                    } else {
                        bool doneOwner = false;
                        bool donePicOwner = false;
                        bool donePicOutlet = false;

                        for (int u=0; u<userList.Count; u++) {
                            if (emailOwner.Equals(userList[u])) {
                                doneOwner = true;
                            }
                            if (emailPicOwner.Equals(userList[u])) {
                                donePicOwner = true;
                            }
                            if (emailPicOutlet.Equals(userList[u])) {
                                donePicOutlet = true;
                            }
                        }

                        userJustCreated = users.checkDealer(userCreated, entries ,orgId, doneOwner, donePicOwner, donePicOutlet,  entries[i], zendeskUsername, zendeskPassword);
                        for (int j=0; j<userJustCreated.Count; j++) {
                            userCreated.Add(userJustCreated[j]);
                        }
                        if (!doneOwner) {
                            userList.Add(emailOwner);
                        }
                        if (!donePicOwner) {
                            userList.Add(emailPicOwner);
                        }
                        if (!donePicOutlet) {
                            userList.Add(emailPicOutlet);
                        }
                    }

                /*END IF EXCELLIMIT*/
                // }

                /*OLD LOGIC*/
                // if (i==0) {

                //     if (isGroupExist(groupMM) == "0") {
                //         string createResponse = newCon.doCreateGroup(groupMM);
                //         JObject createObject = JObject.Parse(createResponse);
                //         groupMMid = createObject["group"]["id"].ToString();
                //     } else {
                //         groupMMid = isGroupExist(groupMM);
                //     }

                //     if (isGroupExist(groupAH) == "0") {
                //         string createResponse = newCon.doCreateGroup(groupAH);
                //         JObject createObject = JObject.Parse(createResponse);
                //         groupAHid = createObject["group"]["id"].ToString();
                //     } else {
                //         groupAHid = isGroupExist(groupAH);
                //     }

                //     if (isGroupExist(groupDH) == "0") {
                //         string createResponse = newCon.doCreateGroup(groupDH);
                //         JObject createObject = JObject.Parse(createResponse);
                //         groupDHid = createObject["group"]["id"].ToString();
                //     } else {
                //         groupDHid = isGroupExist(groupDH);
                //     }

                //     string nameParameter = "name:\"" + namaMm.Replace("#N/A", "") + "\"+name:\"" + namaDh.Replace("#N/A", "") +/* "\"+name:\"" + namaBm +*/ "\"+name:\"" + namaAh.Replace("#N/A", "") + "\"";
                //     string userResponse = isUserExist(nameParameter);
                //     JObject userJoResponse = JObject.Parse(userResponse);
                //     JArray usersList = (JArray)userJoResponse["results"];
                //     bool mmFound = false;
                //     bool dhFound = false;
                //     bool ahFound = false;
                //     for (int u=0; u<usersList.Count; u++){
                //         if (usersList[u]["name"].ToString().ToLower() == namaMm.ToLower()) {
                //             mmFound = true;
                //             userMmId = usersList[u]["id"].ToString();
                //         }
                //         if (usersList[u]["name"].ToString().ToLower() == namaDh.ToLower()) {
                //             dhFound = true;
                //             userDhId = usersList[u]["id"].ToString();
                //         }
                //         if (usersList[u]["name"].ToString().ToLower() == namaAh.ToLower()) {
                //             ahFound = true;
                //             userAhId = usersList[u]["id"].ToString();
                //         }
                //     }
                //     if (namaMm != "#N/A"){
                //         if (!mmFound) {
                //             string userCreate = newCon.createUser(namaMm);
                //             JObject userCreateJoResponse = JObject.Parse(userCreate);
                //             userMmId = userCreateJoResponse["user"]["id"].ToString();
                //         }
                //     } else {
                //         userMmId = "0";
                //         newCon.listGroupMembership(groupMMid);
                //     }

                //     if (namaDh != "#N/A") {
                //         if (!dhFound) {
                //             string userCreate = newCon.createUser(namaDh);
                //             JObject userCreateJoResponse = JObject.Parse(userCreate);
                //             userDhId = userCreateJoResponse["user"]["id"].ToString();
                //         }
                //     } else {
                //         userDhId = "0";
                //         newCon.listGroupMembership(groupDHid);
                //     }

                //     if (namaAh != "#N/A") {
                //         if (!ahFound) {
                //             string userCreate = newCon.createUser(namaAh);
                //             JObject userCreateJoResponse = JObject.Parse(userCreate);
                //             userAhId = userCreateJoResponse["user"]["id"].ToString();
                //         }
                //     } else {
                //         userAhId = "0";
                //         newCon.listGroupMembership(groupAHid);
                //     }

                //     List<Dictionary<string,string>> groupMembershipsList = createGroupMembership(groupMMid, userMmId, groupDHid, userDhId, groupAHid, userAhId);
                //     checkGroupMemberships(groupMembershipsList);
                //     orgId = org.searchOrganizations(orgName, groupMMid, userMmId, groupDHid, userDhId, groupAHid, userAhId);
                //     userJustCreated = users.checkDealer(userCreated, entries, orgId, false, false, false, entries[i]);
                //     for (int j=0; j<userJustCreated.Count; j++) {
                //         userCreated.Add(userJustCreated[j]);
                //     }
                //     userList.Add(emailOwner);
                //     userList.Add(emailPicOwner);
                //     userList.Add(emailPicOutlet);
                // } else if (i <= excelLimit) {

                //     if (entries[i]["Branch Code"].ToString() != entries[i-1]["Branch Code"].ToString()) {
                //         if (isGroupExist(groupMM) == "0") {
                //             string createResponse = newCon.doCreateGroup(groupMM);
                //             JObject createObject = JObject.Parse(createResponse);
                //             // JObject group = (JObject)createObject["group"];
                //             groupMMid = createObject["group"]["id"].ToString();
                //         } else {
                //             groupMMid = isGroupExist(groupMM);
                //         }
                //     }

                //     if (entries[i]["AREA"].ToString() != entries[i-1]["AREA"].ToString()) {
                //         if (isGroupExist(groupAH) == "0") {
                //             string createResponse = newCon.doCreateGroup(groupAH);
                //             JObject createObject = JObject.Parse(createResponse);
                //             // JObject group = (JObject)createObject["group"];
                //             groupAHid = createObject["group"]["id"].ToString();
                //         } else {
                //             groupAHid = isGroupExist(groupAH);
                //         }
                //     }

                //     if (entries[i]["Region Dept Head"].ToString() != entries[i-1]["Region Dept Head"].ToString()) {
                //         if (isGroupExist(groupDH) == "0") {
                //             string createResponse = newCon.doCreateGroup(groupDH);
                //             JObject createObject = JObject.Parse(createResponse);
                //             // JObject group = (JObject)createObject["group"];
                //             groupDHid = createObject["group"]["id"].ToString();
                //         } else {
                //             groupDHid = isGroupExist(groupDH);
                //         }
                //     }

                //     string nameParameter = "name:\"" + namaMm.Replace("#N/A", "") + "\"+name:\"" + namaDh.Replace("#N/A", "") +/* "\"+name:\"" + namaBm + */"\"+name:\"" + namaAh.Replace("#N/A", "") + "\"";
                //     string userResponse = isUserExist(nameParameter);
                //     JObject userJoResponse = JObject.Parse(userResponse);
                //     JArray usersList = (JArray)userJoResponse["results"];
                //     bool mmFound = false;
                //     bool dhFound = false;
                //     bool ahFound = false;
                //     for (int u=0; u<usersList.Count; u++){
                //         if (usersList[u]["name"].ToString().ToLower() == namaMm.ToLower()) {
                //             mmFound = true;
                //             userMmId = usersList[u]["id"].ToString();
                //         }
                //         if (usersList[u]["name"].ToString().ToLower() == namaDh.ToLower()) {
                //             dhFound = true;
                //             userDhId = usersList[u]["id"].ToString();
                //         }
                //         if (usersList[u]["name"].ToString().ToLower() == namaAh.ToLower()) {
                //             ahFound = true;
                //             userAhId = usersList[u]["id"].ToString();
                //         }
                //     }

                //     if (namaMm != "#N/A") {
                //         if (!mmFound) {
                //             if (entries[i]["MM"] != entries[i-1]["MM"]) {
                //                 string userCreate = newCon.createUser(namaMm);
                //                 JObject userCreateJoResponse = JObject.Parse(userCreate);
                //                 // JObject user = (JObject)userCreateJoResponse["user"];
                //                 userMmId = userCreateJoResponse["user"]["id"].ToString();
                //             }
                //         }
                //     } else {
                //         userMmId = "0";
                //         newCon.listGroupMembership(groupMMid);
                //     }

                //     if (namaDh != "#N/A") {
                //         if (!dhFound) {
                //             if (entries[i]["Dept. Head"] != entries[i-1]["Dept. Head"]) {
                //                 string userCreate = newCon.createUser(namaDh);
                //                 JObject userCreateJoResponse = JObject.Parse(userCreate);
                //                 // JObject user = (JObject)userCreateJoResponse["user"];
                //                 userDhId = userCreateJoResponse["user"]["id"].ToString();
                //             }
                //         }
                //     } else {
                //         userDhId = "0";
                //         newCon.listGroupMembership(groupDHid);
                //     }

                //     if (namaAh != "#N/A") {
                //         if (!ahFound) {
                //             if (entries[i]["NAMA Area Head"] != entries[i-1]["NAMA Area Head"]) {
                //                 string userCreate = newCon.createUser(namaAh);
                //                 JObject userCreateJoResponse = JObject.Parse(userCreate);
                //                 // JObject user = (JObject)userCreateJoResponse["user"];
                //                 userAhId = userCreateJoResponse["user"]["id"].ToString();
                //             }
                //         }
                //     } else {
                //         userAhId = "0";
                //         newCon.listGroupMembership(groupAHid);
                //     }

                //     List<Dictionary<string,string>> groupMembershipsList = createGroupMembership(groupMMid, userMmId, groupDHid, userDhId/*, groupBMid, userBmId*/, groupAHid, userAhId);
                //     checkGroupMemberships(groupMembershipsList);
                //     if (entries[i]["Branch Code"] != entries[i-1]["Branch Code"]) {
                //         orgId = org.searchOrganizations(orgName, groupMMid, userMmId, groupDHid, userDhId, groupAHid, userAhId);
                //     }
                //     bool doneOwner = false;
                //     bool donePicOwner = false;
                //     bool donePicOutlet = false;

                //     for (int u=0; u<userList.Count; u++) {
                //         if (emailOwner.Equals(userList[u])) {
                //             doneOwner = true;
                //         }
                //         if (emailPicOwner.Equals(userList[u])) {
                //             donePicOwner = true;
                //         }
                //         if (emailPicOutlet.Equals(userList[u])) {
                //             donePicOutlet = true;
                //         }
                //     }

                //     userJustCreated = users.checkDealer(userCreated, entries ,orgId, doneOwner, donePicOwner, donePicOutlet,  entries[i]);
                //     for (int j=0; j<userJustCreated.Count; j++) {
                //         userCreated.Add(userJustCreated[j]);
                //     }
                //     if (!doneOwner) {
                //         userList.Add(emailOwner);
                //     }
                //     if (!donePicOwner) {
                //         userList.Add(emailPicOwner);
                //     }
                //     if (!donePicOutlet) {
                //         userList.Add(emailPicOutlet);
                //     }
                // }
            }
        }

        // public static void checkGroupMemberships (List<Dictionary<string,string>> groupsIds) {
        //     /*MAKE IT IF REACH 100 ARRAY THEN EXECUTE*/
        //     List<string> willBeDelete = new List<string>();
        //     for (int i=0; i<groupsIds.Count; i++){
        //         if (groupsIds[i]["user_id"] != "0") {
        //             CallingApi callingApi = new CallingApi();
        //             string groupMembershipApi = "/api/v2/groups/" + groupsIds[i]["group_id"] + "/memberships.json";
        //             string groupMembershipRseponse = callingApi.callApi(groupMembershipApi);
        //             // Console.WriteLine(groupMembershipRseponse);
        //             JObject groupMembershipJoResponse = JObject.Parse(groupMembershipRseponse);
        //             JArray memberList = (JArray)groupMembershipJoResponse["group_memberships"];

        //             if (memberList.Count > 1) {
        //                 for (int j=0; j<memberList.Count; j++) {
        //                     if (memberList[j]["user_id"].ToString() != groupsIds[i]["user_id"]) {
        //                         willBeDelete.Add(memberList[j]["id"].ToString());
        //                     }
        //                 }
        //             }

        //             string userGroupsApi = "/api/v2/users/" + groupsIds[i]["user_id"] + "/group_memberships.json";
        //             string userGroupResponse = callingApi.callApi(userGroupsApi);
        //             JObject userGroupJoResponse = JObject.Parse(userGroupResponse);
        //             JArray groupsList = (JArray)userGroupJoResponse["group_memberships"];
        //             bool groupFound = false;
        //             if (groupsList.Count > 1) {
        //                 for (int k=0; k<groupsList.Count; k++) {
        //                     if (groupsList[k]["group_id"].ToString() != groupsIds[i]["group_id"]) {
        //                         if (groupsList[k]["group_id"].ToString() != supportGroupId) {
        //                             for (int l=0; l<doneList.Count; l++) {
        //                                 if (doneList[l]["user_id"].ToString() == groupsIds[i]["user_id"].ToString()) {
        //                                     if (doneList[l]["group_id"].ToString() == groupsList[k]["group_id"].ToString()){
        //                                         groupFound = true;
        //                                     }
        //                                 }
        //                             }
        //                             if (!groupFound) {
        //                                 willBeDelete.Add(groupsList[k]["id"].ToString());
        //                             }
        //                         }
        //                     }
        //                 }
        //             }
        //         }
        //     }
        //     doDeleteMemberships(willBeDelete);
        // }

        // public static void doDeleteMemberships(List<string> memberIds) {
        //     if (memberIds.Count > 0) {
        //         StringBuilder deleteMembershipApi = new StringBuilder();
        //         deleteMembershipApi.Append(zendeskDomain);
        //         deleteMembershipApi.Append("/api/v2/group_memberships/destroy_many.json?ids=");

        //         for (int i=0; i<memberIds.Count; i++) {
        //             deleteMembershipApi.Append(memberIds[i]);
        //             if (i != memberIds.Count-1) {
        //                 deleteMembershipApi.Append(",");
        //             }
        //         }

        //         Console.WriteLine("CALL DELETE: " + deleteMembershipApi);
        //         var client = new RestClient(deleteMembershipApi.ToString());
        //         client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

        //         var request = new RestRequest("", Method.DELETE);

        //         IRestResponse response = client.Execute(request);
        //         var content = response.Content;
        //     }
        // }

        // public static string isGroupExist (string groupName) {
        //     string groupFound = "0";
        //     for (int i=0; i<allGroups.Count; i++) {
        //         JObject group = (JObject)allGroups[i];
        //         if (group["name"].ToString() == groupName) {
        //             groupFound = group["id"].ToString();
        //         }
        //     }
        //     return groupFound;
        // }

        // public static string isUserExist (string nameParameter) {
        //     CallingApi callingApi = new CallingApi();
        //     var searchUserApi =  "/api/v2/search.json?query=type:user%20" + nameParameter;
        //     string response = callingApi.callApi(searchUserApi);
        //     return response;
        // }

        // public string createUser (string userName) {
        //     CallingApi callingApi = new CallingApi();
        //     Dictionary<string,object> newUser = new Dictionary<string,object>();
        //     Dictionary<string,string> userProp = new Dictionary<string,string>();
        //     userProp.Add("name", userName);
        //     userProp.Add("role", "agent");
        //     userProp.Add("custom_role_id", zendeskTeamLeaderRoleId);
        //     userProp.Add("email", userName.Replace(" ", "_") + "@example.com");
        //     newUser.Add("user", userProp);
        //     var createUserApi =  "/api/v2/users.json";
        //     string response = callingApi.callApiPost(JsonConvert.SerializeObject(newUser), createUserApi);
        //     return response;
        // }

        // public static List<Dictionary<string,string>> createGroupMembership (string groupMm, string userMm, string groupDh, string userDh, /*string groupBm, string userBm,*/ string groupAh, string userAh) {
        //     CallingApi callingApi = new CallingApi();
        //     var createGroupMembershipApi = "/api/v2/group_memberships/create_many.json";

        //     Dictionary<string,string> groupMembers = new Dictionary<string,string>();
        //     List<Dictionary<string,string>> groupMembersList = new List<Dictionary<string,string>>();
        //     Dictionary<string, List<Dictionary<string,string>>> groupMemberships = new Dictionary<string, List<Dictionary<string,string>>>();

        //     groupMembers.Add("user_id", userMm);
        //     groupMembers.Add("group_id", groupMm);
        //     doneList.Add(groupMembers);
        //     groupMembersList.Add(groupMembers);
        //     groupMembers = new Dictionary<string,string>();
        //     groupMembers.Add("user_id", userDh);
        //     groupMembers.Add("group_id", groupDh);
        //     doneList.Add(groupMembers);
        //     groupMembersList.Add(groupMembers);
        //     groupMembers = new Dictionary<string,string>();
        //     groupMembers.Add("user_id", userAh);
        //     groupMembers.Add("group_id", groupAh);
        //     doneList.Add(groupMembers);
        //     groupMembersList.Add(groupMembers);
        //     groupMembers = new Dictionary<string,string>();

        //     groupMemberships.Add("group_memberships", groupMembersList);
        //     string createMembershipResponse =  callingApi.callApiPost(JsonConvert.SerializeObject(groupMemberships), createGroupMembershipApi);
        //     return groupMembersList;
        // }

        // public string doCreateGroup(string groupName){
        //     CallingApi callingApi = new CallingApi();
        //     Dictionary<string,string> groupJson = new Dictionary<string,string>();
        //     Dictionary<string,Dictionary<string,string>> groupParameter = new Dictionary<string,Dictionary<string,string>>();

        //     groupJson.Add("name", groupName);
        //     groupParameter.Add("group", groupJson);
        //     var createGroupAPI =  "/api/v2/groups.json";
        //     string content = callingApi.callApiPost(JsonConvert.SerializeObject(groupParameter), createGroupAPI);
        //     return content;
        // }

        public static void initiate () {
            Console.WriteLine(zendeskUsername);
            Console.WriteLine(zendeskPassword);
            getAllGroups("null");
        }

        public static void getAllGroups (string nextPage) {
            CallingApi callingApi = new CallingApi();
            var getGroupApi =  "/api/v2/groups.json";
            string response = "";
            if (nextPage == "null") {
                response = callingApi.callApi(getGroupApi, zendeskUsername, zendeskPassword);
            } else {
                response = callingApi.callApi(nextPage, zendeskUsername, zendeskPassword);
            }

            // JObject joResponse = JObject.Parse(response);
            // JValue nextPageUrl = (JValue)joResponse["next_page"];
            // JArray groupsList = (JArray)joResponse["groups"];

            // for (int i=0; i<groupsList.Count; i++) {
            //     if (groupsList[i]["name"].ToString() == "Support") {
            //         supportGroupId = groupsList[i]["id"].ToString();
            //     }
            //     allGroups.Add(groupsList[i]);
            // }

            // if (joResponse["next_page"].ToString() != String.Empty) {
            //     getAllGroups(joResponse["next_page"].ToString());
            // }
        }

        public static void deleteGroups() {
            /*for (int i=0; i<allGroups.Count; i++) {
                if (allGroups[i]["name"].ToString() == "Support") {
                    Console.WriteLine(allGroups[i]);
                } else {
                    string deleteGroupApi = "/api/v2/groups/" + allGroups[i]["id"] + ".json";
                    Console.WriteLine("CALL DELETE: " + deleteGroupApi);

                    var client = new RestClient(deleteGroupApi);
                    client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

                    var request = new RestRequest("", Method.DELETE);

                    IRestResponse response = client.Execute(request);
                    var content = response.Content;
                    // return content;
                }
            }*/
        }

        // public void listGroupMembership (string groupId) {
        //     string showMembershipApi = "/api/v2/groups/" + groupId + "/memberships.json";
        //     CallingApi callingApi = new CallingApi();
        //     string membershipList = callingApi.callApi(showMembershipApi);
        //     JObject membershipResponse = JObject.Parse(membershipList);
        //     JArray memberships = (JArray)membershipResponse["group_memberships"];
        //     StringBuilder deleteParameter = new StringBuilder();
        //     if (memberships.Count > 0) {
        //         for (int i=0 ;i<memberships.Count; i++) {
        //             deleteParameter.Append(memberships[i]["id"].ToString()).Append(",");
        //         }
        //         deleteGroupMembership(deleteParameter.ToString());
        //     }
        // }

        // public void deleteGroupMembership (string deleteParameter) {
        //     string deleteManyMembershipApi = "/api/v2/group_memberships/destroy_many.json?ids=" + deleteParameter;
        //     CallingApi callingApi = new CallingApi();
        //     callingApi.callApiDelete(deleteManyMembershipApi);
        // }
    }
}
