using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Builder.Dialogs;
using Newtonsoft.Json.Linq;
using System.Net;

namespace Common
{
    public class ProjectServerTeamAPI
    {
        private string _userName;
        private string _userPassword;
        private string _userNameAdmin = ConfigurationManager.AppSettings["DomainAdmin"];
        private string _userPasswordAdmin = ConfigurationManager.AppSettings["DomainAdminPassword"];

        private string _userLoggedInName;
        private string _siteUri;
        public ProjectServerTeamAPI(string userName, string password, string UserLoggedInName)
        {
            _userName = userName;
            _userPassword = password;
            _userLoggedInName = UserLoggedInName;
            _siteUri = ConfigurationManager.AppSettings["PPMServerURL"];
        }


        public IMessageActivity GetMSProjects(IDialogContext dialogContext, int SIndex, bool showCompletion, bool ProjectDates, bool PDuration, bool projectManager, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            string AdminAPI = "/_api/ProjectData/Projects";
            string PMAPI = "/_api/ProjectData/Projects?$filter=ProjectOwnerName eq '" + _userLoggedInName + "'";
            Uri endpointUri = null;
            int ProjectCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    reply = GetAllProjects(dialogContext, jArrays, SIndex, showCompletion, ProjectDates, PDuration, projectManager, out ProjectCounter);
                }
                else if (GetUserGroup("Web Administrators (Project Web App Synchronized)") || GetUserGroup("Administrators for Project Web App") || GetUserGroup("Portfolio Managers for Project Web App") || GetUserGroup("Portfolio Viewers for Project Web App") || GetUserGroup("Portfolio Viewers for Project Web App") || GetUserGroup("Resource Managers for Project Web App"))
                {
                    endpointUri = new Uri(webUri + AdminAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    reply = GetAllProjects(dialogContext, jArrays, SIndex, showCompletion, ProjectDates, PDuration, projectManager, out ProjectCounter);
                }
            }
            Counter = ProjectCounter;
            return reply;
        }

        public bool GetUserGroup(string groupName)
        {
            bool exist = false;

            string UserLoggedInName = string.Empty;
                using (ProjectContext ctx = new ProjectContext(_siteUri))
                {

                    SecureString passWord = new SecureString();
                    foreach (char c in _userPasswordAdmin) passWord.AppendChar(c);
                    ctx.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);

                User user = ctx.Web.EnsureUser(_userName);
                ctx.Load(user);
                ctx.ExecuteQuery();

                if (user != null)
                {
                    ctx.Load(user.Groups);
                    ctx.ExecuteQuery();
                    GroupCollection group = user.Groups;

                    if (group.Count>0)
                    {
                        IEnumerable<Group> usergroup = ctx.LoadQuery(user.Groups.Where(p => p.Title == groupName));
                        ctx.ExecuteQuery();
                        if (!usergroup.Any())
                        {
                            exist = false;
                        }
                        else
                        {
                            exist = true;

                            UserLoggedInName = user.Title;
                        }
                    }
                }
            }
            return exist;
        }
       public IMessageActivity GetAllProjects(IDialogContext context, List<JToken> jArrays, int SIndex, bool showCompletion, bool ProjectDates, bool PDuration, bool projectManager, out int Counter)
        {
            IMessageActivity reply = null;
            reply = context.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = jArrays.Count;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;

            DateTime ProjectFinishDate = new DateTime();
            DateTime ProjectStartDate = new DateTime();

            string ProjectName = string.Empty;
            string ProjectWorkspaceInternalUrl = string.Empty;
            string ProjectPercentCompleted = string.Empty;
            string ProjectDuration = string.Empty;
            string ProjectOwnerName = string.Empty;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    //foreach (var item in jArrays)
                    // {
                    var item = jArrays[startIndex];
                    string SubtitleVal = "";
                    if (item["ProjectName"] != null)
                        ProjectName = (string)item["ProjectName"];

                    if (item["ProjectWorkspaceInternalUrl"] != null)
                        ProjectWorkspaceInternalUrl = (string)item["ProjectWorkspaceInternalUrl"];

                    if (item["ProjectPercentCompleted"] != null)
                        ProjectPercentCompleted = (string)item["ProjectPercentCompleted"];

                    if (item["ProjectFinishDate"] != null && (string)item["ProjectFinishDate"] != null && (string)item["ProjectFinishDate"] != "")
                         ProjectFinishDate = (DateTime)item["ProjectFinishDate"];

                    if (item["ProjectStartDate"] != null && (string)item["ProjectStartDate"] != null && (string)item["ProjectStartDate"] != "")
                        ProjectStartDate = (DateTime)item["ProjectStartDate"];

                    if (item["ProjectDuration"] != null)
                        ProjectDuration = (string)item["ProjectDuration"];

                    if (item["ProjectOwnerName"] != null)
                        ProjectOwnerName = (string)item["ProjectOwnerName"];






                    if (showCompletion == false && ProjectDates == false && PDuration == false && projectManager == false)
                    {
                        SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                        if (item["ProjectFinishDate"] != null && (string)item["ProjectFinishDate"] != null && (string)item["ProjectFinishDate"] != "")
                            SubtitleVal += "Finish Date :\n" + ProjectFinishDate.ToString() + "</br>";
                        else
                            SubtitleVal += "Finish Date :\n</br>";

                        if (item["ProjectStartDate"] != null && (string)item["ProjectStartDate"] != null && (string)item["ProjectStartDate"] != "")
                            SubtitleVal += "Start Date :\n" + ProjectStartDate.ToString() + "</br>";
                        else
                            SubtitleVal += "Start Date :\n</br>";


                        SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                        SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                    }

                    else if (ProjectDates == true)
                    {
                        if (item["ProjectFinishDate"] != null && (string)item["ProjectFinishDate"] != null && (string)item["ProjectFinishDate"] != "")
                            SubtitleVal += "Finish Date :\n" + ProjectFinishDate.ToString() + "</br>";
                        else
                            SubtitleVal += "Finish Date :\n</br>";

                        if (item["ProjectStartDate"] != null && (string)item["ProjectStartDate"] != null && (string)item["ProjectStartDate"] != "")
                            SubtitleVal += "Start Date :\n" + ProjectStartDate.ToString() + "</br>";
                        else
                            SubtitleVal += "Start Date :\n</br>";

                    }
                    else if (PDuration == true)
                    {
                        SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                    }
                    else if (projectManager == true)
                    {
                        SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                    }
                    string ImageURL = "http://02-code.com/images/logo.jpg";
                    List<CardImage> cardImages = new List<CardImage>();
                    List<CardAction> cardactions = new List<CardAction>();
                    cardImages.Add(new CardImage(url: ImageURL));
                    CardAction btnWebsite = new CardAction()
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "Open",
                        Value = ProjectWorkspaceInternalUrl + "?redirect_uri={" + ProjectWorkspaceInternalUrl + "}",
                    };
                    CardAction btnTasks = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = "Tasks",
                        Value = "show a list of " + ProjectName + " tasks",
                        //  DisplayText = "show a list of " + ProjectName + " tasks",
                        Text = "show a list of " + ProjectName + " tasks",
                    };
                    cardactions.Add(btnTasks);

                    CardAction btnIssues = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = "Issues",
                        Value = "show a list of " + ProjectName + " issues",
                        Text = "show a list of " + ProjectName + " issues"
                    };
                    cardactions.Add(btnIssues);

                    CardAction btnRisks = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = "Risks",
                        Value = "Show risks and the assigned resources of " + ProjectName,
                        Text = "Show risks and the assigned resources of " + ProjectName,

                    };
                    cardactions.Add(btnRisks);

                    CardAction btnDeliverables = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = "Deliverables",
                        Value = "Show " + ProjectName + " deliverables",
                        Text = "Show " + ProjectName + " deliverables",
                    };
                    cardactions.Add(btnDeliverables);

                    CardAction btnAssignments = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = "Assignments",
                        Value = "get " + ProjectName + " assignments",
                        Text = "get " + ProjectName + " assignments",

                    };
                    cardactions.Add(btnAssignments);

                    CardAction btnMilestones = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = "Milestones",
                        Value = "get " + ProjectName + " milestones",
                        Text = "get " + ProjectName + " milestones",

                    };
                    cardactions.Add(btnMilestones);

                    HeroCard plCard = new HeroCard()
                    {
                        Title = ProjectName,
                        Subtitle = SubtitleVal,
                        Images = cardImages,
                        Buttons = cardactions,
                        Tap = btnTasks,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());
                }

            }

            return reply;
        }

        public IMessageActivity GetProjectTasks(IDialogContext dialogContext, int itemStartIndex, string pName, bool Completed, bool NotCompleted, bool delayed, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            string PMAPI = "";
           

            Uri endpointUri = null;
            int TaskCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");


                //if (Completed == true)
                //{
                //    PMAPI = "/_api/ProjectData/Tasks?$filter=ProjectName%20eq%20%27" + pName + "%27&TaskPercentCompleted%20eq%201";
                //}
                //else if (NotCompleted == true)
                //{
                //    PMAPI = "/_api/ProjectData/Tasks?$filter=ProjectName%20eq%20%27" + pName + "%27&TaskPercentCompleted%20Nq%201";
                //}
                //else if (delayed == true)
                //{
                //    PMAPI = "/_api/ProjectData/Tasks?$filter=ProjectName%20eq%20%27" + pName + "%27&ActualDuration%20lt%20ScheduledDuration";
                //}
                //else
                pName = ProjectNameStr(pName);
                    PMAPI = "/_api/ProjectData/Tasks?$filter=ProjectName eq '" + pName + "'";

                if (GetUserGroup("Team Members (Project Web App Synchronized)") || GetUserGroup("Team Leads for Project Web App"))
                {
                   // reply = GetResourceLoggedInTasks(dialogContext, itemStartIndex, context, project, Completed, NotCompleted, delayed, out TaskCounter);
                }
                else if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    if(_userLoggedInName.ToLower() == GetProjectPMName(pName).ToLower())
                    {
                        endpointUri = new Uri(webUri + PMAPI);
                        var responce = client.DownloadString(endpointUri);
                        var t = JToken.Parse(responce);
                        JObject results = JObject.Parse(t["d"].ToString());
                        List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                        reply = GetAllTasks(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                    }
                    else
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to access this project",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());

                    }
                }
                else 
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    reply = GetAllTasks(dialogContext, itemStartIndex, jArrays, out TaskCounter);                  
                }
               
            }
        
            Counter = TaskCounter;
            return reply;
        }

        public IMessageActivity GetProjectIssues(IDialogContext dialogContext, int itemStartIndex, string pName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            pName = ProjectNameStr(pName);

            string PMAPI = "/_api/ProjectData/Issues?$filter=ProjectName  eq '" + pName + "'";
            


            Uri endpointUri = null;
            int TaskCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                if (GetUserGroup("Team Members (Project Web App Synchronized)") || GetUserGroup("Team Leads for Project Web App"))
                {
                    // reply = GetResourceLoggedInTasks(dialogContext, itemStartIndex, context, project, Completed, NotCompleted, delayed, out TaskCounter);
                }
                else if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    if (_userLoggedInName.ToLower() == GetProjectPMName(pName).ToLower())
                    {
                        endpointUri = new Uri(webUri + PMAPI);
                        var responce = client.DownloadString(endpointUri);
                        var t = JToken.Parse(responce);
                        JObject results = JObject.Parse(t["d"].ToString());
                        List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                        reply = GetAllIssues(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                    }
                    else
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to access this project",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());

                    }
                }
                else
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    reply = GetAllIssues(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                }
            }


            Counter = TaskCounter;
            return reply;
        }

        public IMessageActivity GetProjectRisks(IDialogContext dialogContext, int itemStartIndex, string pName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            pName = ProjectNameStr(pName);

            string PMAPI = "/_api/ProjectData/Risks?$filter=ProjectName eq '" + pName + "'";



            Uri endpointUri = null;
            int TaskCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                if (GetUserGroup("Team Members (Project Web App Synchronized)") || GetUserGroup("Team Leads for Project Web App"))
                {
                    // reply = GetResourceLoggedInRisks(dialogContext, itemsRisk, itemStartIndex, out TaskCounter);
                }
                else if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    if (_userLoggedInName.ToLower() == GetProjectPMName(pName).ToLower())
                    {
                        endpointUri = new Uri(webUri + PMAPI);
                        var responce = client.DownloadString(endpointUri);
                        var t = JToken.Parse(responce);
                        JObject results = JObject.Parse(t["d"].ToString());
                        List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                        reply = GetAllRisks(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                    }
                    else
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to access this project",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());

                    }
                }
                else
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();

                    reply = GetAllRisks(dialogContext,itemStartIndex, jArrays, out TaskCounter);
                }
            }

            Counter = TaskCounter;

            return reply;
        }

        public IMessageActivity GetProjectDeliverables(IDialogContext dialogContext, int itemStartIndex, string pName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            pName = ProjectNameStr(pName);

            string PMAPI = "/_api/ProjectData/Deliverables?$filter=ProjectName eq '" + pName + "'";



            Uri endpointUri = null;
            int TaskCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                if (GetUserGroup("Team Members (Project Web App Synchronized)") || GetUserGroup("Team Leads for Project Web App"))
                {
                }
                else if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    if (_userLoggedInName.ToLower() == GetProjectPMName(pName).ToLower())
                    {
                        endpointUri = new Uri(webUri + PMAPI);
                        var responce = client.DownloadString(endpointUri);
                        var t = JToken.Parse(responce);
                        JObject results = JObject.Parse(t["d"].ToString());
                        List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                        reply = GetAllDeliverabels(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                    }
                    else
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to access this project",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());

                    }
                }

                else
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();

                    reply = GetAllDeliverabels(dialogContext,itemStartIndex, jArrays, out TaskCounter);
                }
            }
            Counter = TaskCounter;
            return reply;
        }

        public IMessageActivity GetProjectAssignments(IDialogContext dialogContext, int itemStartIndex, string pName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            pName = ProjectNameStr(pName);

            string PMAPI = "/_api/ProjectData/Assignments?$filter=ProjectName eq '" + pName + "'";
            Uri endpointUri = null;
            int TaskCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                if (GetUserGroup("Team Members (Project Web App Synchronized)") || GetUserGroup("Team Leads for Project Web App"))
                {
                }
                else if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    if (_userLoggedInName.ToLower() == GetProjectPMName(pName).ToLower())
                    {
                        endpointUri = new Uri(webUri + PMAPI);
                        var responce = client.DownloadString(endpointUri);
                        var t = JToken.Parse(responce);
                        JObject results = JObject.Parse(t["d"].ToString());
                        List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                        reply = GetAllAssignments(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                    }
                    else
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to access this project",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());

                    }
                }
                else
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();


                    reply = GetAllAssignments(dialogContext,itemStartIndex, jArrays ,  out TaskCounter);
                }



            }
            Counter = TaskCounter;
            return reply;
        }

        public IMessageActivity GetProjectMilestones(IDialogContext dialogContext, int itemStartIndex, string pName, out int Counter)
        {
            IMessageActivity reply = dialogContext.MakeMessage();

            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            pName = ProjectNameStr(pName);

            string PMAPI = "/_api/ProjectData/Tasks?$filter=ProjectName eq '" + pName + "'";
            Uri endpointUri = null;
            int TaskCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                if (GetUserGroup("Team Members (Project Web App Synchronized)") || GetUserGroup("Team Leads for Project Web App"))
                {
                }
                else if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    if (_userLoggedInName.ToLower() == GetProjectPMName(pName).ToLower())
                    {
                        endpointUri = new Uri(webUri + PMAPI);
                        var responce = client.DownloadString(endpointUri);
                        var t = JToken.Parse(responce);
                        JObject results = JObject.Parse(t["d"].ToString());
                        List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                        reply = GetAllProjectMilestones(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                    }
                    else
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to access this project",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());

                    }
                }
                else
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();


                     reply = GetAllProjectMilestones(dialogContext, itemStartIndex, jArrays, out TaskCounter);
                }



            }

            Counter = TaskCounter;
            return reply;
        }

        private IMessageActivity GetAllTasks(IDialogContext dialogContext, int SIndex, List<JToken> jArrays, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = jArrays.Count;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var tsk = jArrays[startIndex];
                    var SubtitleVal = "";
                    string TaskName = (string)tsk["TaskName"];
                    string TaskDuration = (string)tsk["TaskDuration"];
                    string TaskPercentCompleted = (string)tsk["TaskPercentCompleted"];
                    string TaskStartDate = (string)tsk["TaskStartDate"];
                    string TaskFinishDate = (string)tsk["TaskFinishDate"];
                    SubtitleVal += "Task Duration\n" + TaskDuration + "</br>";
                    SubtitleVal += "Task Percent Completed\n" + TaskPercentCompleted + "</br>";
                    SubtitleVal += "Task Start Date\n" + TaskStartDate + "</br>";
                    SubtitleVal += "Task Finish Date\n" + TaskFinishDate + "</br>";
                    HeroCard plCard = new HeroCard()
                    {
                        Title = TaskName,
                        Subtitle = SubtitleVal
                    };
                    reply.Attachments.Add(plCard.ToAttachment());
                }
            }
            return reply;
        }

        private IMessageActivity GetAllIssues(IDialogContext dialogContext, int SIndex, List<JToken> jArrays , out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = jArrays.Count;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    var itemsIssue = jArrays[startIndex];
                    string IssueName = (string)itemsIssue["Title"];
                    string IssueStatus = (string)itemsIssue["Status"];
                    string IssuePriority = (string)itemsIssue["Priority"];
                    SubtitleVal += "Status\n" + IssueStatus + "</br>";
                    SubtitleVal += "Priority\n" + IssuePriority + "</br>";
                    HeroCard plCard = new HeroCard()
                    {
                        Title = IssueName,
                        Subtitle = SubtitleVal,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());
                }
            }
            return reply;
        }

        private IMessageActivity GetAllRisks(IDialogContext dialogContext, int SIndex, List<JToken> jArrays , out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string RiskName = string.Empty;
            string ResourceName = string.Empty;
            string riskStatus = string.Empty;
            string riskImpact = string.Empty;
            string riskProbability = string.Empty;
            string riskCostExposure = string.Empty;


            Counter = jArrays.Count;

            int inDexToVal = SIndex + 10;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    var item = jArrays[startIndex];
                    RiskName = (string)item["Title"];
                    if (item["AssignedToResource"] != null)
                    {
                        string AssignedToResource = (string)item["AssignedToResource"];
                        SubtitleVal += "Assigned To Resource\n" + AssignedToResource + "</br>";

                    }
                    else
                        SubtitleVal += "Assigned To Resource :\n" + "Not assigned" + "</br>";

                    if (item["Status"] != null)
                        riskStatus = (string)item["Status"];
                    SubtitleVal += "Risk Status\n" + riskStatus + "</br>";

                    if (item["Impact"] != null)
                        riskImpact = item["Impact"].ToString();
                    SubtitleVal += "Risk Impact\n" + riskImpact + "</br>";

                    if (item["Probability"] != null)
                        riskProbability = item["Probability"].ToString();
                    SubtitleVal += "Risk Probability\n" + riskProbability + "</br>";

                    if (item["Exposure"] != null)
                        riskCostExposure = item["Exposure"].ToString();
                    SubtitleVal += "Risk CostExposure\n" + riskCostExposure + "</br>";


                    HeroCard plCard = new HeroCard()
                    {
                        Title = RiskName,
                        Subtitle = SubtitleVal,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());

                }

            }

            return reply;
        }

        private IMessageActivity GetAllDeliverabels(IDialogContext dialogContext, int SIndex, List<JToken> jArrays , out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string DeliverableName = string.Empty;
            string DeliverableStart = string.Empty;
            string DeliverableFinish = string.Empty;
            string CreateByResource = string.Empty;


            Counter = jArrays.Count;

            int inDexToVal = SIndex + 10;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    var item = jArrays[startIndex];

                    if (item["Title"] != null)
                    {
                        DeliverableName = (string)item["Title"];
                        SubtitleVal += "Deliverable Name\n" + DeliverableName + "</br>";
                    }

                    if (item["CreateByResource"] != null)
                    {
                        DeliverableStart = item["CreateByResource"].ToString();
                        SubtitleVal += "Start Date :\n" + DeliverableStart + "</br>";

                    }

                    if (item["StartDate"] != null)
                    {
                        DeliverableStart = item["StartDate"].ToString();
                        SubtitleVal += "Start Date :\n" + DeliverableStart + "</br>";
                    }

                    if (item["FinishDate"] != null)
                    {
                        DeliverableFinish = item["FinishDate"].ToString();
                        SubtitleVal += "Finish Date :\n" + DeliverableFinish + "</br>";
                    }


                    HeroCard plCard = new HeroCard()
                    {
                        Title = DeliverableName,
                        Subtitle = SubtitleVal,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());

                }

            }

            return reply;
        }

        private IMessageActivity GetAllAssignments(IDialogContext dialogContext, int SIndex, List<JToken> jArrays, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = jArrays.Count;

            int inDexToVal = SIndex + 10;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    var item = jArrays[startIndex];

                    string TaskName = (string)item["TaskName"];
                    string ResourceName = (string)item["ResourceName"];
                    string Start = (string)item["AssignmentStartDate"];
                    string Finish = (string)item["AssignmentFinishDate"];

                    SubtitleVal += "Resource Name :\n" + ResourceName + "</br>";
                    SubtitleVal += "Start Date\n" + Start + "</br>";
                    SubtitleVal += "Finish Date\n" + Finish + "</br>";


                    HeroCard plCard = new HeroCard()
                    {
                        Title = TaskName,
                        Subtitle = SubtitleVal,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());
                }
            }
            return reply;
        }
        private IMessageActivity GetAllProjectMilestones(IDialogContext dialogContext, int SIndex, List<JToken> jArrays, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = 0;
            //if (inDexToVal >= jArrays.Count)
            //    inDexToVal = jArrays.Count;

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < jArrays.Count; startIndex++)
                {
                    var tsk = jArrays[startIndex];

                    if (tsk["TaskDuration"] != null)
                    {
                        if ((string)tsk["TaskDuration"] == "0.000000")
                        {
                            // string TaskDuration = (string)tsk["TaskDuration"];
                            Counter++;
                            var SubtitleVal = "";
                            string TaskName = (string)tsk["TaskName"];
                            string TaskPercentCompleted = (string)tsk["TaskPercentCompleted"];
                            string TaskStartDate = (string)tsk["TaskStartDate"];
                            string TaskFinishDate = (string)tsk["TaskFinishDate"];
                            SubtitleVal += "Task Percent Completed\n" + TaskPercentCompleted + "</br>";
                            SubtitleVal += "Task Start Date\n" + TaskStartDate + "</br>";
                            SubtitleVal += "Task Finish Date\n" + TaskFinishDate + "</br>";
                            HeroCard plCard = new HeroCard()
                            {
                                Title = TaskName,
                                Subtitle = SubtitleVal
                            };
                            reply.Attachments.Add(plCard.ToAttachment());
                        }
                    }

                    if (Counter == 10)
                        break;
                }
            }
            return reply;
        }

        public IMessageActivity FilterMSProjects(IDialogContext dialogContext, int SIndex, int completionpercentVal , string FilterType, string pStartDate, string PEndDate, string ProjectSEdateFlag, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            DateTime startdate = new DateTime();
            DateTime endate = new DateTime();
            if (!string.IsNullOrEmpty(pStartDate))
                startdate = DateTime.Parse(pStartDate);

            if (!string.IsNullOrEmpty(PEndDate))
                endate = DateTime.Parse(PEndDate);

            string formatedstartdate = startdate.ToString("yyyy-MM-ddT23:59:59Z");
            string formatedendate = endate.ToString("yyyy-MM-ddT23:59:59Z");

            //{13/08/2019 10:24:00}
            string formatedstartdatePM = startdate.ToString("dd/MM/yyyy 23:59:59");
            string formatedendatePM = endate.ToString("dd/MM/yyyy 23:59:59");

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            string AdminAPI = "/_api/ProjectData/Projects";
            string PMAPI = "/_api/ProjectData/Projects?$filter=ProjectOwnerName eq '" + _userLoggedInName + "'";
            Uri endpointUri = null;
            int ProjectCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");

                if (GetUserGroup("Project Managers (Project Web App Synchronized)"))
                {
                    endpointUri = new Uri(webUri + PMAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();

                    reply = GetFilteredProjects(dialogContext, jArrays, SIndex, completionpercentVal, FilterType, formatedstartdatePM, formatedendatePM, ProjectSEdateFlag, out ProjectCounter);

                }
                else if (GetUserGroup("Web Administrators (Project Web App Synchronized)") || GetUserGroup("Administrators for Project Web App") || GetUserGroup("Portfolio Managers for Project Web App") || GetUserGroup("Portfolio Viewers for Project Web App") || GetUserGroup("Portfolio Viewers for Project Web App") || GetUserGroup("Resource Managers for Project Web App"))
                {
                    if(completionpercentVal ==100)
                        AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectPercentCompleted eq "+ completionpercentVal;
                    if (completionpercentVal == 90)
                        AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectPercentCompleted lt 100";
                    if (ProjectSEdateFlag == "START")
                    {
                        if (FilterType.ToUpper() == "BEFORE" && pStartDate != "")
                            AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectStartDate le DateTime'"+ formatedstartdate + "'&$orderby=ProjectStartDate";

                        else if (FilterType.ToUpper() == "AFTER" && pStartDate != "")
                            AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectStartDate ge DateTime'" + formatedstartdate + "'&$orderby=ProjectStartDate";

                        else if (FilterType.ToUpper() == "BETWEEN" && pStartDate != "")
                            AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectStartDate ge DateTime'" + formatedstartdate + "' and ProjectStartDate le DateTime'" + formatedendate + "'&$orderby=ProjectStartDate";
                    }
                    else if (ProjectSEdateFlag == "Finish")
                    {
                        if (FilterType.ToUpper() == "BEFORE" && PEndDate != "")
                            AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectFinishDate le DateTime'" + formatedendate + "'&$orderby=ProjectFinishDate";
                        else if (FilterType.ToUpper() == "AFTER" && PEndDate != "")
                            AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectFinishDate ge DateTime'" + formatedendate + "'&$orderby=ProjectFinishDate";
                        else if (FilterType.ToUpper() == "BETWEEN" && PEndDate != "")
                            AdminAPI = "/_api/ProjectData/Projects?$filter=ProjectFinishDate ge DateTime'" + formatedstartdate + "' and ProjectFinishDate le DateTime'" + formatedendate + "'&$orderby=ProjectFinishDate";
                    }

                    endpointUri = new Uri(webUri + AdminAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    reply = GetAllProjects(dialogContext, jArrays, SIndex, false, false, false, false, out ProjectCounter);
                }
            }
            Counter = ProjectCounter;
            return reply;
        }

        

        public IMessageActivity GetProjectInfo(IDialogContext dialogContext, string pName, bool optionalDate = false, bool optionalDuration = false, bool optionalCompletion = false, bool optionalPM = false)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string SubtitleVal = "";

            //using (ProjectContext context = new ProjectContext(_siteUri))
            //{
            //    SecureString passWord = new SecureString();
            //    foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            //    context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);

            //    PublishedProject project = GetProjectByName(pName, context);

            //    if (project != null)
            //    {
            //        if (optionalDate == true)
            //        {
            //            SubtitleVal += "Start Date :\n" + project.StartDate + "</br>";
            //            SubtitleVal += "Finish Date :\n" + project.FinishDate + "</br>";
            //        }

            //        if (optionalDuration == true)
            //        {
            //            TimeSpan duration = project.FinishDate - project.StartDate;
            //            SubtitleVal += "Project Duration :\n" + duration.Days + "</br>";
            //        }

            //        if (optionalCompletion == true)
            //            SubtitleVal += "Project Completed Percentage :\n" + project.PercentComplete + "%</br>";

            //        if (optionalPM == true)
            //        {
            //            if (GetUserGroup("Team Members (Project Web App Synchronized)") == false)
            //            {
            //                context.Load(project.Owner);
            //                context.ExecuteQuery();
            //                SubtitleVal += "Project Manager Name :\n" + project.Owner.Title + "</br>";
            //            }
            //        }

            //        HeroCard plCard = new HeroCard()
            //        {
            //            Title = pName,
            //            Subtitle = SubtitleVal,
            //        };
            //        reply.Attachments.Add(plCard.ToAttachment());

            //    }
            //    else
            //    {
            //        HeroCard plCardNoData = new HeroCard()
            //        {
            //            Title = "Project Name Not Exist",
            //        };
            //        reply.Attachments.Add(plCardNoData.ToAttachment());

            //    }





            //}
            return reply;
        }


        public string GetProjectPMName(string ProjectName)
        {
            string ProjectPMName = string.Empty;
            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            string PMAPI = "/_api/ProjectData/Projects?$filter=ProjectName eq '"+ProjectName+"'";
            Uri endpointUri = null;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");
                endpointUri = new Uri(webUri + PMAPI);
                var responce = client.DownloadString(endpointUri);
                var t = JToken.Parse(responce);
                JObject results = JObject.Parse(t["d"].ToString());
                List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();

                if(jArrays !=null)
                {
                    if(jArrays.Count >0)
                    {
                        if (jArrays[0]["ProjectOwnerName"] != null)
                            ProjectPMName = (string)jArrays[0]["ProjectOwnerName"];
                    }
                }
            }
            return ProjectPMName;
        }

        public string ProjectNameStr(string ProjectName)
        {
            string pName = ProjectName;
            if (pName.Contains(" - "))
                pName = pName.Replace(" - ", "-");

            return pName;
        }



        public IMessageActivity GetFilteredProjects(IDialogContext dialogContext, List<JToken> jArrays, int SIndex, int completionpercentVal, string FilterType, string pStartDate, string PEndDate, string ProjectSEdateFlag, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;
            IEnumerable<JToken> jToken = null;
            DateTime ProjectFinishDate = new DateTime();
            DateTime ProjectStartDate = new DateTime();

            string ProjectName = string.Empty;
            string ProjectWorkspaceInternalUrl = string.Empty;
            string ProjectPercentCompleted = string.Empty;
            string ProjectDuration = string.Empty;
            string ProjectOwnerName = string.Empty;

            if (jArrays.Count > 0)
            {
                if (completionpercentVal == 100)
                     jToken = jArrays.Where(t => (int?)t["ProjectPercentCompleted"] == 100);
                else if (completionpercentVal == 90)
                    jToken = jArrays.Where(t => (int?)t["ProjectPercentCompleted"] < 100);

                if (ProjectSEdateFlag == "START")
                {
                    if (FilterType.ToUpper() == "BEFORE" && pStartDate != "")
                        jToken = jArrays.Where(t => (DateTime)t["ProjectStartDate"] <= DateTime.Parse(pStartDate));

                //    else if (FilterType.ToUpper() == "AFTER" && pStartDate != "")
                //        jToken = jArrays.Where(t => (DateTime?)t["ProjectStartDate"] >= DateTime.Parse(pStartDate));

                //    else if (FilterType.ToUpper() == "BETWEEN" && pStartDate != "")
                //        jToken = jArrays.Where(t => (DateTime?)t["ProjectStartDate"] >= DateTime.Parse(pStartDate) && (DateTime?)t["ProjectStartDate"] <= DateTime.Parse(PEndDate));
                }
                else if (ProjectSEdateFlag == "Finish")
                {
                //    if (FilterType.ToUpper() == "BEFORE" && PEndDate != "")
                //        jToken = jArrays.Where(t => (DateTime?)t["ProjectFinishDate"] <= DateTime.Parse(PEndDate));

                //    else if (FilterType.ToUpper() == "AFTER" && PEndDate != "")
                //        jToken = jArrays.Where(t => (DateTime?)t["ProjectFinishDate"] >= DateTime.Parse(PEndDate));
                //    else if (FilterType.ToUpper() == "BETWEEN" && PEndDate != "")
                //        jToken = jArrays.Where(t => (DateTime?)t["ProjectFinishDate"] >= DateTime.Parse(pStartDate) && (DateTime?)t["ProjectFinishDate"] <= DateTime.Parse(PEndDate));
                }

                if (jToken !=null)
                {
                    if (jToken.Count() > 0)
                    {
                        int inDexToVal = SIndex + 10;
                        Counter = jToken.Count();
                        if (inDexToVal >= jToken.Count())
                            inDexToVal = jToken.Count();

                        //        for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                        //        {
                        //            var item = jToken.ElementAt(startIndex);
                        //            string SubtitleVal = "";


                        //            if (item["ProjectName"] != null)
                        //                ProjectName = (string)item["ProjectName"];

                        //            if (item["ProjectWorkspaceInternalUrl"] != null)
                        //                ProjectWorkspaceInternalUrl = (string)item["ProjectWorkspaceInternalUrl"];

                        //            if (item["ProjectPercentCompleted"] != null)
                        //                ProjectPercentCompleted = (string)item["ProjectPercentCompleted"];

                        //            if (item["ProjectFinishDate"] != null)
                        //                ProjectFinishDate = (DateTime)item["ProjectFinishDate"];

                        //            if (item["ProjectStartDate"] != null)
                        //                ProjectStartDate = (DateTime)item["ProjectStartDate"];

                        //            if (item["ProjectDuration"] != null)
                        //                ProjectDuration = (string)item["ProjectDuration"];

                        //            if (item["ProjectOwnerName"] != null)
                        //                ProjectOwnerName = (string)item["ProjectOwnerName"];

                        //            SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                        //            SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                        //            SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                        //            SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                        //            SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";

                        //            string ImageURL = "http://02-code.com/images/logo.jpg";
                        //            List<CardImage> cardImages = new List<CardImage>();
                        //            List<CardAction> cardactions = new List<CardAction>();
                        //            cardImages.Add(new CardImage(url: ImageURL));
                        //            CardAction btnWebsite = new CardAction()
                        //            {
                        //                Type = ActionTypes.OpenUrl,
                        //                Title = "Open",
                        //                Value = ProjectWorkspaceInternalUrl + "?redirect_uri={" + ProjectWorkspaceInternalUrl + "}",
                        //            };
                        //            CardAction btnTasks = new CardAction()
                        //            {
                        //                Type = ActionTypes.PostBack,
                        //                Title = "Tasks",
                        //                Value = "show a list of " + ProjectName + " tasks",
                        //                //  DisplayText = "show a list of " + ProjectName + " tasks",
                        //                Text = "show a list of " + ProjectName + " tasks",
                        //            };
                        //            cardactions.Add(btnTasks);

                        //            CardAction btnIssues = new CardAction()
                        //            {
                        //                Type = ActionTypes.PostBack,
                        //                Title = "Issues",
                        //                Value = "show a list of " + ProjectName + " issues",
                        //                Text = "show a list of " + ProjectName + " issues"
                        //            };
                        //            cardactions.Add(btnIssues);

                        //            CardAction btnRisks = new CardAction()
                        //            {
                        //                Type = ActionTypes.PostBack,
                        //                Title = "Risks",
                        //                Value = "Show risks and the assigned resources of " + ProjectName,
                        //                Text = "Show risks and the assigned resources of " + ProjectName,

                        //            };
                        //            cardactions.Add(btnRisks);

                        //            CardAction btnDeliverables = new CardAction()
                        //            {
                        //                Type = ActionTypes.PostBack,
                        //                Title = "Deliverables",
                        //                Value = "Show " + ProjectName + " deliverables",
                        //                Text = "Show " + ProjectName + " deliverables",
                        //            };
                        //            cardactions.Add(btnDeliverables);

                        //            CardAction btnAssignments = new CardAction()
                        //            {
                        //                Type = ActionTypes.PostBack,
                        //                Title = "Assignments",
                        //                Value = "get " + ProjectName + " assignments",
                        //                Text = "get " + ProjectName + " assignments",

                        //            };
                        //            cardactions.Add(btnAssignments);

                        //            CardAction btnMilestones = new CardAction()
                        //            {
                        //                Type = ActionTypes.PostBack,
                        //                Title = "Milestones",
                        //                Value = "get " + ProjectName + " milestones",
                        //                Text = "get " + ProjectName + " milestones",

                        //            };
                        //            cardactions.Add(btnMilestones);

                        //            HeroCard plCard = new HeroCard()
                        //            {
                        //                Title = ProjectName,
                        //                Subtitle = SubtitleVal,
                        //                Images = cardImages,
                        //                Buttons = cardactions,
                        //                Tap = btnTasks,
                        //            };
                        //            reply.Attachments.Add(plCard.ToAttachment());
                        //        }

                    }
                }

               

            }

            HeroCard plCard2 = new HeroCard()
            {
                Title = pStartDate,

            };
            reply.Attachments.Add(plCard2.ToAttachment());

            return reply;
        }

        public IMessageActivity GetResourceAssignments(IDialogContext dialogContext, int SIndex, string ResourceName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;

            SecureString passWord = new SecureString();
            foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
            var webUri = new Uri(_siteUri);
            ResourceName = ResourceName.Replace(" ", String.Empty);
            string AdminAPI = "/_api/ProjectData/Resources?$filter=ResourceEmailAddress eq '" + ResourceName + "'";
            //string AdminAPI = "/_api/ProjectData/Resources";

            Uri endpointUri = null;
            int RCounter = 0;

          
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");               
                if (GetUserGroup("Web Administrators (Project Web App Synchronized)") || GetUserGroup("Administrators for Project Web App") || GetUserGroup("Portfolio Managers for Project Web App") || GetUserGroup("Portfolio Viewers for Project Web App") || GetUserGroup("Portfolio Viewers for Project Web App") || GetUserGroup("Resource Managers for Project Web App"))
                {
                    endpointUri = new Uri(webUri + AdminAPI);
                    var responce = client.DownloadString(endpointUri);
                    var t = JToken.Parse(responce);
                    JObject results = JObject.Parse(t["d"].ToString());
                    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    reply = GetAllResourceAssignments(dialogContext,SIndex, jArrays, out RCounter);
                }
            }
            Counter = RCounter;
            return reply;

       
          
        }
        public IMessageActivity GetAllResourceAssignments(IDialogContext dialogContext, int SIndex , List<JToken> jArrays, out int RCounter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            int inDexToVal = SIndex + 10;
            RCounter = jArrays.Count;
            if (inDexToVal >= jArrays.Count)
                inDexToVal = jArrays.Count;



            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var item = jArrays[startIndex];

                    string SubtitleVal = "";
                    string ResourceName = (string)item["ResourceName"];
                    string ResourceBaseCalendar = (string)item["ResourceBaseCalendar"];
                    string ResourceIsActive = (string)item["ResourceIsActive"];
                    string Role = (string)item["Role"];

                    SubtitleVal += "Resource Base Calendar :\n" + ResourceBaseCalendar + "</br>";
                    SubtitleVal += "Resource Is Active :\n" + ResourceIsActive + "</br>";
                    SubtitleVal += "Role :\n" + Role + "</br>";


                    HeroCard plCard = new HeroCard()
                    {
                        Title = ResourceName,
                        Subtitle = SubtitleVal,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());
                }

            }

           




            return reply;
        }
        
        public IMessageActivity TotalCountGeneralMessage(IDialogContext dialogContext, int SIndex, int Counter, string ListName)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            if (ListName == Enums.ListName.Projects.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 projects";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 projects";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Projects :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Projects :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  Projects\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == Enums.ListName.Tasks.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Tasks";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Tasks";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Tasks :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Tasks :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Tasks\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == Enums.ListName.Issues.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Issues";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Issues";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Issues :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Issues :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  Issues\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == Enums.ListName.Assignments.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Assignments";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Assignments";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Assignments :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Assignments :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Assignments\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == Enums.ListName.Risks.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Risks";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Risks";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Risks :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Risks :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  Risks\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == Enums.ListName.Deliverables.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Deliverables";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Deliverables";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Deliverables :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Deliverables :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  Deliverables\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == "FilterProjects")
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Projects";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Projects";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Projects :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Projects :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  Projects\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == "UserAssignments")
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 assignments";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 assignments";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of user assignments :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of user assignments :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  assignments\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == Enums.ListName.Milestones.ToString())
            {
                if (Counter > 0)
                {
                    if (Counter >= 10)
                    {
                        string subTitle = string.Empty;
                        if (SIndex == 0)
                            subTitle = "You are viwing the first page , each page view 10 Milestones";
                        else if (SIndex > 0)
                        {
                            int pagenumber = SIndex / 10 + 1;
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Milestones";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Milestones :\n" + Counter,
                            Subtitle = subTitle,
                            //  Buttons = cardButtons,
                        };
                        reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                    else
                    {
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available  Milestones :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available  Milestones\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else
            {
                HeroCard plCardNoData = new HeroCard()
                { Title = "No data returned \n\n" };
                reply.Attachments.Add(plCardNoData.ToAttachment());
            }
            return reply;
        }

        public IMessageActivity CreateButtonsPager(IDialogContext dialogContext, int totalCount, string ListName, string projectName, string query)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string valuebutton = string.Empty;
            List<CardAction> cardButtons = new List<CardAction>();
            double p = totalCount * 0.1;
            double result = Math.Ceiling(p);
            int pagenumber = int.Parse(result.ToString());

            if (query.Contains("at index"))
                query = query.Substring(0, query.IndexOf("at index"));
            if (totalCount > 10)
            {
                

              
                for (int i = 0; i < pagenumber; i++)
                {
                    string CurrentNumber = Convert.ToString(i);
                    if (ListName == Enums.ListName.Projects.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "get all projects at index 0";
                        }
                        else
                            valuebutton = "get all projects at index " + i * 10;
                    }
                    else if (ListName == Enums.ListName.Tasks.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "show a list of " + projectName + " tasks at index 0";
                        }
                        else
                            valuebutton = "show a list of " + projectName + " tasks at index " + i * 10;

                    }
                    else if (ListName == Enums.ListName.Issues.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "show a list of " + projectName + " issues at index 0";
                        }
                        else
                            valuebutton = "show a list of " + projectName + " issues at index " + i * 10;

                    }
                    else if (ListName == Enums.ListName.Deliverables.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "show a list of " + projectName + " deliverables at index 0";
                        }
                        else
                            valuebutton = "show a list of " + projectName + " deliverables at index " + i * 10;

                    }
                    else if (ListName == Enums.ListName.Risks.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "show a list of " + projectName + " risks at index 0";
                        }
                        else
                            valuebutton = "show a list of " + projectName + " risks at index " + i * 10;

                    }
                    else if (ListName == Enums.ListName.Assignments.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "get " + projectName + " assignments at index 0";
                        }
                        else
                            valuebutton = "get " + projectName + " assignments at index " + i * 10;

                    }
                    else if (ListName == "FilterProjects" && query != "")
                    {
                        if (i == 0)
                        {
                            valuebutton = query + " at index 0";
                        }
                        else
                            valuebutton = query + " at index " + i * 10;

                    }
                    else if (ListName == "UserAssignments" && query != "")
                    {
                        if (i == 0)
                        {
                            valuebutton = query + " at index 0";
                        }
                        else
                            valuebutton = query + " at index " + i * 10;

                    }
                    else if (ListName == Enums.ListName.Milestones.ToString())
                    {
                        if (i == 0)
                        {
                            valuebutton = "show a list of " + projectName + " milestones at index 0";
                        }
                        else
                            valuebutton = "show a list of " + projectName + " milestones at index " + i * 10;

                    }
                    CurrentNumber = Convert.ToString(i + 1);
                    CardAction CardButton = new CardAction()
                    {
                        Type = ActionTypes.PostBack,
                        Title = CurrentNumber,
                        Value = valuebutton,
                        Text = valuebutton,
                    };

                    ThumbnailCard plCardCounter = new ThumbnailCard()
                    {
                        Title = "Page" + CurrentNumber,
                        //   Images = cardImages,
                        Tap = CardButton,

                    };

                    reply.Attachments.Add(plCardCounter.ToAttachment());
                }
            }




            return reply;
        }

        public IMessageActivity DataSuggestions(IDialogContext dialogContext, string ListName, string ProjectName)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            if (ListName == Enums.ListName.Tasks.ToString())
            {
                List<CardAction> cardactions = new List<CardAction>();
                CardAction btnNotCompltedTasks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Not Completed Tasks",
                    Value = "show a list of " + ProjectName + " not completed tasks",
                    Text = "show a list of " + ProjectName + " not completed tasks",
                };
                cardactions.Add(btnNotCompltedTasks);

                CardAction btnCompltedTasks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Completed Tasks",
                    Value = "show a list of " + ProjectName + " completed tasks",
                    Text = "show a list of " + ProjectName + " completed tasks",
                };
                cardactions.Add(btnCompltedTasks);
                CardAction btnDelayedTasks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Delayed Tasks",
                    Value = "show a list of " + ProjectName + " delayed tasks",
                    Text = "show a list of " + ProjectName + " delayed tasks",
                };
                cardactions.Add(btnDelayedTasks);
                HeroCard plCard = new HeroCard()
                {
                    Title = "Suggestions",
                    Subtitle = "Because of large numnber of tasks, We have the below suggestions to get your tasks",
                    Buttons = cardactions,
                };
                reply.Attachments.Add(plCard.ToAttachment());
            }
            else if (ListName == Enums.ListName.Projects.ToString())
            {
                List<CardAction> cardactions = new List<CardAction>();
                CardAction btnClosedProjects = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Closed Projects",
                    Value = "get all projects where compeleted percentage is 100%",
                    Text = "get all projects where compeleted percentage is 100%",
                };
                cardactions.Add(btnClosedProjects);

                CardAction btnPendingProjects = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Pending Projects",
                    Value = "get all projects where compeleted percentage is 90%",
                    Text = "get all projects where compeleted percentage is 90%",
                };
                cardactions.Add(btnPendingProjects);

                //CardAction btnCompltedPcurrentYear = new CardAction()
                //{
                //    Type = ActionTypes.PostBack,
                //    Title = "Completed Projects This Year",
                //    Value = "get all projects closed this year",
                //    Text = "get all projects closed this year",
                //};
                //cardactions.Add(btnCompltedPcurrentYear);
                //CardAction btnStartedPcurrentYear = new CardAction()
                //{
                //    Type = ActionTypes.PostBack,
                //    Title = "Started Projects This Year",
                //    Value = "get all projects started this year",
                //    Text = "get all projects started this year",
                //};
                //cardactions.Add(btnStartedPcurrentYear);
                HeroCard plCard = new HeroCard()
                {
                    Title = "Suggestions",
                    Subtitle = "Because of large numnber of projects, We have the below suggestions to get your projects",
                    Buttons = cardactions,
                };
                reply.Attachments.Add(plCard.ToAttachment());
            }

            return reply;
        }
        //private IMessageActivity GetResourceLoggedInTasks(IDialogContext dialogContext, int SIndex, ProjectContext context, PublishedProject proj, bool Completed, bool NotCompleted, bool delayed, out int Counter)
        //{
        //    var SubtitleVal = "";
        //    IMessageActivity reply = null;
        //    reply = dialogContext.MakeMessage();
        //    context.Load(proj.Assignments, da => da.Where(a => a.Resource.Email != string.Empty && a.Resource.Email == _userName));
        //    context.ExecuteQuery();
        //    Counter = 0;



        //    if (proj.Assignments != null)
        //    {
        //        PublishedAssignmentCollection proAssignment = proj.Assignments;

        //        int inDexToVal = SIndex + 10;
        //        Counter = proAssignment.Count;
        //        if (inDexToVal >= proAssignment.Count)
        //            inDexToVal = proAssignment.Count;

        //        for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
        //        {
        //            PublishedAssignment ass = proAssignment[startIndex];
        //            context.Load(ass.Task);
        //            context.Load(ass.Resource);

        //            context.ExecuteQuery();
        //            var tsk = ass.Task;
        //            string TaskName = tsk.Name;
        //            string TaskDuration = tsk.Duration;
        //            string TaskPercentCompleted = tsk.PercentComplete.ToString();
        //            string TaskStartDate = tsk.Start.ToString();
        //            string TaskFinishDate = tsk.Finish.ToString();

        //            SubtitleVal += "Task Duration\n" + TaskDuration + "</br>";
        //            SubtitleVal += "Task Percent Completed\n" + TaskPercentCompleted + "</br>";
        //            SubtitleVal += "Task Start Date\n" + TaskStartDate + "</br>";
        //            SubtitleVal += "Task Finish Date\n" + TaskFinishDate + "</br>";

        //            HeroCard plCard = new HeroCard()
        //            {
        //                Title = TaskName,
        //                Subtitle = SubtitleVal,

        //            };
        //            reply.Attachments.Add(plCard.ToAttachment());
        //        }
        //    }
        //    return reply;
        //}


        //private IMessageActivity GetResourceLoggedInIssues(IDialogContext dialogContext, ListItemCollection itemsIssue, int SIndex, out int Counter)
        //{
        //    IMessageActivity reply = null;
        //    reply = dialogContext.MakeMessage();
        //    Counter = 0;



        //    int inDexToVal = SIndex + 10;
        //    if (inDexToVal >= itemsIssue.Count)
        //        inDexToVal = itemsIssue.Count;


        //    if (itemsIssue.Count > 0)
        //    {
        //        int count = 0;
        //        for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
        //        {
        //            ListItem item = itemsIssue[startIndex];

        //            if (item["AssignedTo"] != null)
        //            {
        //                count++;
        //                FieldUserValue fuv = (FieldUserValue)item["AssignedTo"];
        //                if (fuv.Email == _userName)
        //                {
        //                    string SubtitleVal = "";
        //                    string IssueName = string.Empty;
        //                    string IssueStatus = string.Empty;
        //                    string IssuePriority = string.Empty;

        //                    if (item["Title"] != null)
        //                        IssueName = (string)item["Title"];
        //                    if (item["Status"] != null)
        //                        IssueStatus = (string)item["Status"];
        //                    if (item["Priority"] != null)
        //                        IssuePriority = (string)item["Priority"];
        //                    SubtitleVal += "Status\n" + IssueStatus + "</br>";
        //                    SubtitleVal += "Priority\n" + IssuePriority + "</br>";
        //                    HeroCard plCard = new HeroCard()
        //                    {
        //                        Title = IssueName,
        //                        Subtitle = SubtitleVal,
        //                    };
        //                    reply.Attachments.Add(plCard.ToAttachment());
        //                }
        //            }
        //        }
        //        Counter = count;
        //    }
        //    return reply;
        //}



        //private IMessageActivity GetResourceLoggedInRisks(IDialogContext dialogContext, ListItemCollection itemsRisk, int SIndex, out int Counter)
        //{
        //    IMessageActivity reply = null;
        //    reply = dialogContext.MakeMessage();
        //    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
        //    Counter = 0;
        //    string RiskName = string.Empty;
        //    string ResourceName = string.Empty;
        //    string riskStatus = string.Empty;
        //    string riskImpact = string.Empty;
        //    string riskProbability = string.Empty;
        //    string riskCostExposure = string.Empty;
        //    if (itemsRisk.Count > 0)
        //    {
        //        int count = 0;

        //        int inDexToVal = SIndex + 10;
        //        if (inDexToVal >= itemsRisk.Count)
        //            inDexToVal = itemsRisk.Count;
        //        for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
        //        {
        //            var SubtitleVal = "";
        //            ListItem item = itemsRisk[startIndex];


        //            if (item["AssignedTo"] != null)
        //            {
        //                count++;
        //                FieldUserValue fuv = (FieldUserValue)item["AssignedTo"];
        //                if (fuv.Email == _userName)
        //                {
        //                    if (item["Title"] != null)
        //                        RiskName = (string)item["Title"];
        //                    SubtitleVal += "Risk Title\n" + RiskName + "</br>";

        //                    SubtitleVal += "Assigned To Resource\n" + fuv.LookupValue + "</br>";
        //                    if (item["Status"] != null)
        //                        riskStatus = (string)item["Status"];
        //                    SubtitleVal += "Risk Status\n" + riskStatus + "</br>";

        //                    if (item["Impact"] != null)
        //                        riskImpact = item["Impact"].ToString();
        //                    SubtitleVal += "Risk Impact\n" + riskImpact + "</br>";

        //                    if (item["Probability"] != null)
        //                        riskProbability = item["Probability"].ToString();
        //                    SubtitleVal += "Risk Probability\n" + riskProbability + "</br>";

        //                    if (item["Exposure"] != null)
        //                        riskCostExposure = item["Exposure"].ToString();
        //                    SubtitleVal += "Risk CostExposure\n" + riskCostExposure + "</br>";

        //                    HeroCard plCard = new HeroCard()
        //                    {
        //                        Title = RiskName,
        //                        Subtitle = SubtitleVal,
        //                    };
        //                    reply.Attachments.Add(plCard.ToAttachment());

        //                }

        //            }
        //            Counter = count;


        //        }

        //    }

        //    return reply;
        //}





        //public IMessageActivity GetResourceLoggedInAssignments(IDialogContext dialogContext, ProjectContext context, PublishedAssignmentCollection itemsAssignments, int SIndex, string ResourceName, out int Counter)
        //{
        //    IMessageActivity reply = null;
        //    reply = dialogContext.MakeMessage();
        //    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
        //    Counter = 0;
        //    int count = 0;
        //    if (itemsAssignments.Count > 0)
        //    {
        //        int inDexToVal = SIndex + 10;
        //        if (inDexToVal >= itemsAssignments.Count)
        //            inDexToVal = itemsAssignments.Count;

        //        for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
        //        {
        //            PublishedAssignment ass = itemsAssignments[startIndex];

        //            context.Load(ass.Task);
        //            context.Load(ass.Resource);
        //            context.ExecuteQuery();

        //            if (ass.Resource.Email == ResourceName)
        //            {
        //                count++;
        //                string SubtitleVal = "";
        //                string TaskName = ass.Task.Name;
        //                SubtitleVal += "Resource Name :\n" + ass.Resource.Name + "</br>";
        //                SubtitleVal += "Start Date\n" + ass.Start + "</br>";
        //                SubtitleVal += "Finish Date\n" + ass.Finish + "</br>";


        //                HeroCard plCard = new HeroCard()
        //                {
        //                    Title = TaskName,
        //                    Subtitle = SubtitleVal,
        //                };
        //                reply.Attachments.Add(plCard.ToAttachment());
        //            }
        //        }
        //        Counter = count;
        //    }
        //    return reply;

        //}
        //public bool UserHavePermissionOnaProjects(string siteUrl, string subSiteTitle, ProjectContext context)
        //{

        //    var web = context.Web;
        //    bool exist = false;
        //    context.Load(web, w => w.Webs);
        //    context.ExecuteQuery();
        //    foreach (Web subWeb in web.Webs)
        //    {
        //        if (subWeb.Title.ToLower() == subSiteTitle.ToLower())
        //        {
        //            var user = subWeb.EnsureUser(_userName);
        //            context.Load(user);
        //            context.ExecuteQuery();

        //            if (null != user)
        //            {
        //                ClientResult<BasePermissions> permissions = subWeb.GetUserEffectivePermissions(user.LoginName);
        //                context.ExecuteQuery();


        //                if (permissions.Value.Has(PermissionKind.ViewListItems))
        //                {
        //                    exist = true;
        //                    break;
        //                }
        //                else
        //                    exist = false;


        //            }
        //            else
        //                exist = false;




        //        }
        //    }

        //    return exist;
        //}



        //private static PublishedProject GetProjectByName(string name, ProjectContext context)
        //{
        //    if (name.Contains(" - "))
        //        name = name.Replace(" - ", "-");
        //    IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Name == name));
        //    context.ExecuteQuery();
        //    if (!projs.Any())       // no project found
        //    {
        //        return null;
        //    }
        //    return projs.FirstOrDefault();

        //}

        //private static Web GetProjectWEB(string siteurl, ProjectContext context)
        //{
        //    IEnumerable<Web> webs = context.LoadQuery(context.Web.Webs.Where(p => p.Url == siteurl));
        //    context.ExecuteQuery();
        //    if (!webs.Any())       // no project found
        //    {
        //        return null;
        //    }
        //    return webs.FirstOrDefault();

        //}
        //public IMessageActivity GetLoggedInPMProjects(IDialogContext dialogContext, ProjectContext context, ProjectCollection projectDetails, int SIndex, bool showCompletion, bool ProjectDates, bool PDuration, bool projectManager, out int Counter)
        //{
        //    IMessageActivity reply = null;
        //    reply = dialogContext.MakeMessage();
        //    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

        //    int inDexToVal = SIndex + 10;
        //    Counter = 0;
        //    //Counter = projectDetails.Count;
        //    //if (inDexToVal >= projectDetails.Count)
        //    //    inDexToVal = projectDetails.Count;

        //    foreach (PublishedProject pro in projectDetails)
        //    {
        //        //  PublishedProject pro = context.Projects[startIndex];
        //        context.Load(pro.Owner);
        //        context.Load(pro, p => p.ProjectSiteUrl);
        //        context.ExecuteQuery();


        //        if (pro.Owner.Email == _userName)
        //        {
        //            Counter++;
        //            string ProjectName = pro.Name;
        //            string ProjectWorkspaceInternalUrl = pro.ProjectSiteUrl;
        //            string ProjectPercentCompleted = pro.PercentComplete.ToString();
        //            string ProjectFinishDate = pro.FinishDate.ToString();
        //            string ProjectStartDate = pro.StartDate.ToString();
        //            TimeSpan duration = pro.FinishDate - pro.StartDate;
        //            string ProjectDuration = duration.Days.ToString();
        //            string ProjectOwnerName = pro.Owner.Title;
        //            string SubtitleVal = "";
        //            if (showCompletion == false && ProjectDates == false && PDuration == false && projectManager == false)
        //            {
        //                SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
        //                SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
        //                SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
        //                SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
        //                SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
        //            }

        //            else if (ProjectDates == true)
        //            {
        //                SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
        //                SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
        //            }
        //            else if (PDuration == true)
        //            {
        //                SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
        //            }
        //            else if (projectManager == true)
        //            {
        //                SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
        //            }
        //            string ImageURL = "http://02-code.com/images/logo.jpg";
        //            List<CardImage> cardImages = new List<CardImage>();
        //            List<CardAction> cardactions = new List<CardAction>();
        //            cardImages.Add(new CardImage(url: ImageURL));
        //            CardAction btnWebsite = new CardAction()
        //            {
        //                Type = ActionTypes.OpenUrl,
        //                Title = "Open",
        //                Value = ProjectWorkspaceInternalUrl + "?redirect_uri={" + ProjectWorkspaceInternalUrl + "}",
        //            };
        //            CardAction btnTasks = new CardAction()
        //            {
        //                Type = ActionTypes.PostBack,
        //                Title = "Tasks",
        //                Value = "show a list of " + ProjectName + " tasks",
        //                //  DisplayText = "show a list of " + ProjectName + " tasks",
        //                Text = "show a list of " + ProjectName + " tasks",
        //            };
        //            cardactions.Add(btnTasks);

        //            CardAction btnIssues = new CardAction()
        //            {
        //                Type = ActionTypes.PostBack,
        //                Title = "Issues",
        //                Value = "show a list of " + ProjectName + " issues",
        //                Text = "show a list of " + ProjectName + " issues"
        //            };
        //            cardactions.Add(btnIssues);

        //            CardAction btnRisks = new CardAction()
        //            {
        //                Type = ActionTypes.PostBack,
        //                Title = "Risks",
        //                Value = "Show risks and the assigned resources of " + ProjectName,
        //                Text = "Show risks and the assigned resources of " + ProjectName,

        //            };
        //            cardactions.Add(btnRisks);

        //            CardAction btnDeliverables = new CardAction()
        //            {
        //                Type = ActionTypes.PostBack,
        //                Title = "Deliverables",
        //                Value = "Show " + ProjectName + " deliverables",
        //                Text = "Show " + ProjectName + " deliverables",
        //            };
        //            cardactions.Add(btnDeliverables);

        //            CardAction btnAssignments = new CardAction()
        //            {
        //                Type = ActionTypes.PostBack,
        //                Title = "Assignments",
        //                Value = "get " + ProjectName + " assignments",
        //                Text = "get " + ProjectName + " assignments",

        //            };
        //            cardactions.Add(btnAssignments);

        //            CardAction btnMilestones = new CardAction()
        //            {
        //                Type = ActionTypes.PostBack,
        //                Title = "Milestones",
        //                Value = "get " + ProjectName + " milestones",
        //                Text = "get " + ProjectName + " milestones",

        //            };
        //            cardactions.Add(btnMilestones);

        //            HeroCard plCard = new HeroCard()
        //            {
        //                Title = ProjectName,
        //                Subtitle = SubtitleVal,
        //                Images = cardImages,
        //                Buttons = cardactions,
        //                Tap = btnTasks,
        //            };
        //            reply.Attachments.Add(plCard.ToAttachment());
        //            if (Counter == 10)
        //                break;
        //        }
        //    }

        //    return reply;
        //}
        //public IMessageActivity GetResourceLoggedInProjects(IDialogContext dialogContext, ProjectContext context, ProjectCollection projectDetails, int SIndex, bool showCompletion, bool ProjectDates, bool PDuration, bool projectManager, out int Counter)
        //{
        //    IMessageActivity reply = null;
        //    reply = dialogContext.MakeMessage();
        //    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

        //    int inDexToVal = SIndex + 10;
        //    Counter = projectDetails.Count;
        //    if (inDexToVal >= projectDetails.Count)
        //        inDexToVal = projectDetails.Count;

        //    for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
        //    {
        //        PublishedProject pro = context.Projects[startIndex];
        //        context.Load(pro.Owner);
        //        //  context.Load(pro, p => p.ProjectSiteUrl);
        //        context.ExecuteQuery();

        //        string ProjectName = pro.Name;
        //        //     string ProjectWorkspaceInternalUrl = pro.ProjectSiteUrl;
        //        string ProjectPercentCompleted = pro.PercentComplete.ToString();
        //        string ProjectFinishDate = pro.FinishDate.ToString();
        //        string ProjectStartDate = pro.StartDate.ToString();
        //        TimeSpan duration = pro.FinishDate - pro.StartDate;
        //        string ProjectDuration = duration.Days.ToString();
        //        // string ProjectOwnerName = pro.Owner.Title;
        //        string SubtitleVal = "";
        //        if (showCompletion == false && ProjectDates == false && PDuration == false && projectManager == false)
        //        {
        //            SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
        //            SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
        //            SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
        //            SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
        //            //   SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
        //        }
        //        else if (showCompletion == true)
        //            SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
        //        else if (ProjectDates == true)
        //        {
        //            SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
        //            SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
        //        }
        //        else if (PDuration == true)
        //        {
        //            SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
        //        }
        //        //else if (projectManager == true)
        //        //{
        //        //    SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
        //        //}
        //        string ImageURL = "http://02-code.com/images/logo.jpg";
        //        List<CardImage> cardImages = new List<CardImage>();
        //        List<CardAction> cardactions = new List<CardAction>();
        //        cardImages.Add(new CardImage(url: ImageURL));
        //        //CardAction btnWebsite = new CardAction()
        //        //{
        //        //    Type = ActionTypes.OpenUrl,
        //        //    Title = "Open",
        //        //    Value = ProjectWorkspaceInternalUrl + "?redirect_uri={" + ProjectWorkspaceInternalUrl + "}",
        //        //};
        //        CardAction btnTasks = new CardAction()
        //        {
        //            Type = ActionTypes.PostBack,
        //            Title = "Tasks",
        //            Value = "show a list of " + ProjectName + " tasks",
        //        };
        //        cardactions.Add(btnTasks);

        //        CardAction btnIssues = new CardAction()
        //        {
        //            Type = ActionTypes.PostBack,
        //            Title = "Issues",
        //            Value = "show a list of " + ProjectName + " issues",
        //        };
        //        cardactions.Add(btnIssues);

        //        CardAction btnRisks = new CardAction()
        //        {
        //            Type = ActionTypes.PostBack,
        //            Title = "Risks",
        //            Value = "Show risks and the assigned resources of " + ProjectName,
        //        };
        //        cardactions.Add(btnRisks);

        //        //CardAction btnDeliverables = new CardAction()
        //        //{
        //        //    Type = ActionTypes.PostBack,
        //        //    Title = "Deliverables",
        //        //    Value = "Show " + ProjectName + " deliverables",
        //        //};
        //        //cardactions.Add(btnDeliverables);

        //        CardAction btnDAssignments = new CardAction()
        //        {
        //            Type = ActionTypes.PostBack,
        //            Title = "Assignments",
        //            Value = "get " + ProjectName + " assignments",
        //        };
        //        cardactions.Add(btnDAssignments);

        //        HeroCard plCard = new HeroCard()
        //        {
        //            Title = ProjectName,
        //            Subtitle = SubtitleVal,
        //            Images = cardImages,
        //            Buttons = cardactions,
        //            Tap = btnTasks,
        //        };
        //        reply.Attachments.Add(plCard.ToAttachment());
        //    }

        //    return reply;
        //}
    }
}
