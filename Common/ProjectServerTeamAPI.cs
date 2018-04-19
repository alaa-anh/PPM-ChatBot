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
            string PMAPI = "/_api/ProjectData/Projects?$filter=ProjectOwnerName%20eq%20%27" + _userLoggedInName + "%27";
            Uri endpointUri = null;
            int ProjectCounter = 0;
            using (var client = new WebClient())
            {
                client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                client.Credentials = credentials;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json;odata=verbose");
                client.Headers.Add(HttpRequestHeader.Accept, "application/json;odata=verbose");

                if (GetUserGroupAPI("Project Managers (Project Web App Synchronized)"))
                {
                    //endpointUri = new Uri(webUri + PMAPI);
                    //var responce = client.DownloadString(endpointUri);
                    //var t = JToken.Parse(responce);
                    //JObject results = JObject.Parse(t["d"].ToString());


                    //List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                    //reply = GetAllProjects(dialogContext, jArrays, SIndex, showCompletion, ProjectDates, PDuration, projectManager, out ProjectCounter);

                    HeroCard plCard = new HeroCard()
                    {
                        Title = "Test PM",

                    };
                    reply.Attachments.Add(plCard.ToAttachment());


                }
                //else if (GetUserGroupAPI("Web Administrators (Project Web App Synchronized)") || GetUserGroupAPI("Administrators for Project Web App") || GetUserGroupAPI("Portfolio Managers for Project Web App") || GetUserGroupAPI("Portfolio Viewers for Project Web App") || GetUserGroupAPI("Portfolio Viewers for Project Web App") || GetUserGroupAPI("Resource Managers for Project Web App"))
                //{
                //    endpointUri = new Uri(webUri + AdminAPI);
                //    var responce = client.DownloadString(endpointUri);
                //    var t = JToken.Parse(responce);
                //    JObject results = JObject.Parse(t["d"].ToString());


                //    List<JToken> jArrays = ((Newtonsoft.Json.Linq.JContainer)((Newtonsoft.Json.Linq.JContainer)t["d"]).First).First.ToList();
                //    reply = GetAllProjects(dialogContext, jArrays, SIndex, showCompletion, ProjectDates, PDuration, projectManager, out ProjectCounter);


                //}




            }

           

            Counter = ProjectCounter;
            return reply;
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

            if (jArrays.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    //foreach (var item in jArrays)
                    // {
                    var item = jArrays[startIndex];
                    string SubtitleVal = "";
                    string ProjectName = (string)item["ProjectName"];
                    string ProjectWorkspaceInternalUrl = (string)item["ProjectWorkspaceInternalUrl"];
                    string ProjectPercentCompleted = (string)item["ProjectPercentCompleted"];
                    string ProjectFinishDate = (string)item["ProjectFinishDate"];
                    string ProjectStartDate = (string)item["ProjectStartDate"];
                    string ProjectDuration = (string)item["ProjectDuration"];
                    string ProjectOwnerName = (string)item["ProjectOwnerName"];



                    if (showCompletion == false && ProjectDates == false && PDuration == false && projectManager == false)
                    {
                        SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                        SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                        SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                        SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                        SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                    }

                    else if (ProjectDates == true)
                    {
                        SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                        SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
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
                        Text = "get " + ProjectName + " milesones",

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


        public IMessageActivity FilterMSProjects(IDialogContext dialogContext, int SIndex, int completionpercentVal, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);

                context.Load(context.Projects);
                context.ExecuteQuery();

                int ProjectCounter = 0;

                if (completionpercentVal > 0)
                {
                    IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.PercentComplete == completionpercentVal));
                    context.ExecuteQuery();

                    if (projs.Any())
                    {
                        if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                        {
                            //reply = GetResourceLoggedInFilProjects(dialogContext, context, projectDetails, SIndex, showCompletion, ProjectDates, PDuration, projectManager, out ProjectCounter);
                        }
                        else
                        {
                            reply = GetFilteredProjects(dialogContext, context, projs, SIndex, completionpercentVal, out ProjectCounter);
                        }

                        Counter = ProjectCounter;
                    }

                }
                //else
                //{
                //    ProjectCollection projectDetails = context.Projects;
                //    if (context.Projects.Count > 0)
                //    {
                //        if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                //        {
                //            //reply = GetResourceLoggedInFilProjects(dialogContext, context, projectDetails, SIndex, showCompletion, ProjectDates, PDuration, projectManager, out ProjectCounter);
                //        }
                //        else
                //        {
                //            reply = GetAllProjects(dialogContext, context, projectDetails, SIndex, completionpercentVal, out ProjectCounter);
                //        }

                //        Counter = ProjectCounter;
                //    }

                //}
            }

            return reply;
        }



        public IMessageActivity GetFilteredProjects(IDialogContext dialogContext, ProjectContext context, IEnumerable<PublishedProject> publishedProject, int SIndex, int completionpercentVal, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = publishedProject.Count();
            if (inDexToVal >= publishedProject.Count())
                inDexToVal = publishedProject.Count();


            foreach (PublishedProject pro in publishedProject)
            {
                context.Load(pro.Owner);
                context.Load(pro, p => p.ProjectSiteUrl);
                context.ExecuteQuery();

                string ProjectName = pro.Name;
                string ProjectWorkspaceInternalUrl = pro.ProjectSiteUrl;
                string ProjectPercentCompleted = pro.PercentComplete.ToString();
                string ProjectFinishDate = pro.FinishDate.ToString();
                string ProjectStartDate = pro.StartDate.ToString();
                TimeSpan duration = pro.FinishDate - pro.StartDate;
                string ProjectDuration = duration.Days.ToString();
                string ProjectOwnerName = pro.Owner.Title;
                string SubtitleVal = "";
                SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";

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

                CardAction btnDAssignments = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Assignments",
                    Value = "get " + ProjectName + " assignments",
                    Text = "get " + ProjectName + " assignments",

                };
                cardactions.Add(btnDAssignments);

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

            return reply;
        }

        public IMessageActivity GetLoggedInPMProjects(IDialogContext dialogContext, ProjectContext context, ProjectCollection projectDetails, int SIndex, bool showCompletion, bool ProjectDates, bool PDuration, bool projectManager, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = projectDetails.Count;
            if (inDexToVal >= projectDetails.Count)
                inDexToVal = projectDetails.Count;


            //context.Load(context.EnterpriseResources);
            //var resources = context.EnterpriseResources;

            //context.ExecuteQuery();


            var user = context.Web.EnsureUser(_userName);
            context.Load(user);
            context.ExecuteQuery();

            IEnumerable<PublishedProject> projs = projectDetails.Where(p => p.Owner == user);
            context.ExecuteQuery();

            for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
            {
                PublishedProject pro = context.Projects[startIndex];
                context.Load(pro.Owner);
                //  context.Load(pro, p => p.ProjectSiteUrl);
                context.ExecuteQuery();

                string ProjectName = pro.Name;
                //     string ProjectWorkspaceInternalUrl = pro.ProjectSiteUrl;
                string ProjectPercentCompleted = pro.PercentComplete.ToString();
                string ProjectFinishDate = pro.FinishDate.ToString();
                string ProjectStartDate = pro.StartDate.ToString();
                TimeSpan duration = pro.FinishDate - pro.StartDate;
                string ProjectDuration = duration.Days.ToString();
                // string ProjectOwnerName = pro.Owner.Title;
                string SubtitleVal = "";
                if (showCompletion == false && ProjectDates == false && PDuration == false && projectManager == false)
                {
                    SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                    SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                    SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                    SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                    //   SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                }
                else if (showCompletion == true)
                    SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                else if (ProjectDates == true)
                {
                    SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                    SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                }
                else if (PDuration == true)
                {
                    SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                }
                //else if (projectManager == true)
                //{
                //    SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                //}
                string ImageURL = "http://02-code.com/images/logo.jpg";
                List<CardImage> cardImages = new List<CardImage>();
                List<CardAction> cardactions = new List<CardAction>();
                cardImages.Add(new CardImage(url: ImageURL));
                //CardAction btnWebsite = new CardAction()
                //{
                //    Type = ActionTypes.OpenUrl,
                //    Title = "Open",
                //    Value = ProjectWorkspaceInternalUrl + "?redirect_uri={" + ProjectWorkspaceInternalUrl + "}",
                //};
                CardAction btnTasks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Tasks",
                    Value = "show a list of " + ProjectName + " tasks",
                };
                cardactions.Add(btnTasks);

                CardAction btnIssues = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Issues",
                    Value = "show a list of " + ProjectName + " issues",
                };
                cardactions.Add(btnIssues);

                CardAction btnRisks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Risks",
                    Value = "Show risks and the assigned resources of " + ProjectName,
                };
                cardactions.Add(btnRisks);

                //CardAction btnDeliverables = new CardAction()
                //{
                //    Type = ActionTypes.PostBack,
                //    Title = "Deliverables",
                //    Value = "Show " + ProjectName + " deliverables",
                //};
                //cardactions.Add(btnDeliverables);

                CardAction btnDAssignments = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Assignments",
                    Value = "get " + ProjectName + " assignments",
                };
                cardactions.Add(btnDAssignments);

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

            return reply;
        }
        public IMessageActivity GetResourceLoggedInProjects(IDialogContext dialogContext, ProjectContext context, ProjectCollection projectDetails, int SIndex, bool showCompletion, bool ProjectDates, bool PDuration, bool projectManager, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int inDexToVal = SIndex + 10;
            Counter = projectDetails.Count;
            if (inDexToVal >= projectDetails.Count)
                inDexToVal = projectDetails.Count;

            for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
            {
                PublishedProject pro = context.Projects[startIndex];
                context.Load(pro.Owner);
                //  context.Load(pro, p => p.ProjectSiteUrl);
                context.ExecuteQuery();

                string ProjectName = pro.Name;
                //     string ProjectWorkspaceInternalUrl = pro.ProjectSiteUrl;
                string ProjectPercentCompleted = pro.PercentComplete.ToString();
                string ProjectFinishDate = pro.FinishDate.ToString();
                string ProjectStartDate = pro.StartDate.ToString();
                TimeSpan duration = pro.FinishDate - pro.StartDate;
                string ProjectDuration = duration.Days.ToString();
                // string ProjectOwnerName = pro.Owner.Title;
                string SubtitleVal = "";
                if (showCompletion == false && ProjectDates == false && PDuration == false && projectManager == false)
                {
                    SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                    SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                    SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                    SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                    //   SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                }
                else if (showCompletion == true)
                    SubtitleVal += "Completed Percentage :\n" + ProjectPercentCompleted + "%</br>";
                else if (ProjectDates == true)
                {
                    SubtitleVal += "Start Date :\n" + ProjectStartDate + "</br>";
                    SubtitleVal += "Finish Date :\n" + ProjectFinishDate + "</br>";
                }
                else if (PDuration == true)
                {
                    SubtitleVal += "Project Duration :\n" + ProjectDuration + "</br>";
                }
                //else if (projectManager == true)
                //{
                //    SubtitleVal += "Project Manager :\n" + ProjectOwnerName + "</br>";
                //}
                string ImageURL = "http://02-code.com/images/logo.jpg";
                List<CardImage> cardImages = new List<CardImage>();
                List<CardAction> cardactions = new List<CardAction>();
                cardImages.Add(new CardImage(url: ImageURL));
                //CardAction btnWebsite = new CardAction()
                //{
                //    Type = ActionTypes.OpenUrl,
                //    Title = "Open",
                //    Value = ProjectWorkspaceInternalUrl + "?redirect_uri={" + ProjectWorkspaceInternalUrl + "}",
                //};
                CardAction btnTasks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Tasks",
                    Value = "show a list of " + ProjectName + " tasks",
                };
                cardactions.Add(btnTasks);

                CardAction btnIssues = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Issues",
                    Value = "show a list of " + ProjectName + " issues",
                };
                cardactions.Add(btnIssues);

                CardAction btnRisks = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Risks",
                    Value = "Show risks and the assigned resources of " + ProjectName,
                };
                cardactions.Add(btnRisks);

                //CardAction btnDeliverables = new CardAction()
                //{
                //    Type = ActionTypes.PostBack,
                //    Title = "Deliverables",
                //    Value = "Show " + ProjectName + " deliverables",
                //};
                //cardactions.Add(btnDeliverables);

                CardAction btnDAssignments = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Assignments",
                    Value = "get " + ProjectName + " assignments",
                };
                cardactions.Add(btnDAssignments);

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

            return reply;
        }

        public IMessageActivity GetProjectTasks(IDialogContext dialogContext, int itemStartIndex, string pName, bool Completed, bool NotCompleted, bool delayed, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int TaskCounter = 0;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                PublishedProject project = GetProjectByName(pName, context);


                if (project != null)
                {
                    context.Load(project.Tasks);
                    context.ExecuteQuery();
                    PublishedTaskCollection publishedTask = project.Tasks;
                    if (project.Tasks.Count > 0)
                    {
                        if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                        {
                            reply = GetResourceLoggedInTasks(dialogContext, itemStartIndex, context, project, Completed, NotCompleted, delayed, out TaskCounter);
                        }
                        else if (GetUserGroup(context, "Project Managers (Project Web App Synchronized)"))
                        {
                            context.Load(project.Owner);
                            context.ExecuteQuery();
                            if (project.Owner.Email == _userName) // if the logged in user is a project manager on this project
                            {
                                reply = GetAllTasks(dialogContext, itemStartIndex, publishedTask, project, Completed, NotCompleted, delayed, out TaskCounter);
                            }
                            else
                            {
                                reply = GetResourceLoggedInTasks(dialogContext, itemStartIndex, context, project, Completed, NotCompleted, delayed, out TaskCounter);
                            }
                        }
                        else
                        {
                            reply = GetAllTasks(dialogContext, itemStartIndex, publishedTask, project, Completed, NotCompleted, delayed, out TaskCounter);
                        }
                    }
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
            string projectsite = string.Empty;
            Web projectweb;
            int TaskCounter = 0;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                // SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                PublishedProject project = GetProjectByName(pName, context);
                if (project != null)
                {
                    context.Load(project, p => p.ProjectSiteUrl);
                    context.Load(project.Owner);

                    context.ExecuteQuery();
                    projectsite = project.ProjectSiteUrl;
                    projectweb = GetProjectWEB(projectsite, context);
                    var issues = projectweb.Lists.GetByTitle(Enums.ListName.Issues.ToString());
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection itemsIssue = issues.GetItems(query);
                    context.Load(issues);
                    context.Load(itemsIssue);
                    context.ExecuteQuery();
                    if (itemsIssue.Count() > 0)
                    {
                        if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                        {
                            reply = GetResourceLoggedInIssues(dialogContext, itemsIssue, itemStartIndex, out TaskCounter);
                        }
                        if (GetUserGroup(context, "Project Managers (Project Web App Synchronized)"))
                        {
                            if (project.Owner.Email == _userName)
                            {
                                reply = GetAllIssues(dialogContext, itemsIssue, itemStartIndex, out TaskCounter);
                            }
                            else
                            {
                                reply = GetResourceLoggedInIssues(dialogContext, itemsIssue, itemStartIndex, out TaskCounter);

                            }

                        }
                        else
                        {
                            reply = GetAllIssues(dialogContext, itemsIssue, itemStartIndex, out TaskCounter);
                        }
                    }
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
            string projectsite = string.Empty;
            Web projectweb;
            int TaskCounter = 0;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                //  SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                PublishedProject project = GetProjectByName(pName, context);

                if (project != null)
                {
                    context.Load(project, p => p.ProjectSiteUrl);
                    context.ExecuteQuery();

                    projectsite = project.ProjectSiteUrl;
                    projectweb = GetProjectWEB(projectsite, context);

                    var risks = projectweb.Lists.GetByTitle(Enums.ListName.Risks.ToString());
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection itemsRisk = risks.GetItems(query);

                    projectweb.Context.Load(risks);
                    projectweb.Context.Load(itemsRisk);
                    projectweb.Context.ExecuteQuery();

                    if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                    {
                        reply = GetResourceLoggedInRisks(dialogContext, itemsRisk, itemStartIndex, out TaskCounter);
                    }
                    else if (GetUserGroup(context, "Project Managers (Project Web App Synchronized)"))
                    {
                        context.Load(project.Owner);
                        context.ExecuteQuery();
                        if (project.Owner.Email == _userName) // if the logged in user is a project manager on this project
                        {
                            reply = GetAllRisks(dialogContext, itemsRisk, itemStartIndex, out TaskCounter);
                        }
                        else
                        {
                            reply = GetResourceLoggedInRisks(dialogContext, itemsRisk, itemStartIndex, out TaskCounter);

                        }
                    }
                    else
                    {
                        reply = GetAllRisks(dialogContext, itemsRisk, itemStartIndex, out TaskCounter);
                    }




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
            string projectsite = string.Empty;
            Web projectweb;
            int TaskCounter = 0;

            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                //      SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                PublishedProject project = GetProjectByName(pName, context);

                if (project != null)
                {


                    if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                    {
                        HeroCard plCard = new HeroCard()
                        {
                            Title = "You Don't have permission to view the deliverabels of this projects",
                        };
                        reply.Attachments.Add(plCard.ToAttachment());
                    }
                    else if (GetUserGroup(context, "Project Managers (Project Web App Synchronized)"))
                    {
                        context.Load(project, p => p.ProjectSiteUrl);
                        context.ExecuteQuery();

                        projectsite = project.ProjectSiteUrl;
                        projectweb = GetProjectWEB(projectsite, context);

                        var delive = projectweb.Lists.GetByTitle(Enums.ListName.Deliverables.ToString());
                        CamlQuery query = CamlQuery.CreateAllItemsQuery();
                        ListItemCollection itemsdelive = delive.GetItems(query);

                        projectweb.Context.Load(delive);
                        projectweb.Context.Load(itemsdelive);
                        projectweb.Context.ExecuteQuery();
                        context.Load(project.Owner);
                        context.ExecuteQuery();
                        if (project.Owner.Email == _userName) // if the logged in user is a project manager on this project
                        {
                            reply = GetAllDeliverabels(dialogContext, itemsdelive, itemStartIndex, out TaskCounter);
                        }
                        else
                        {
                            HeroCard plCard = new HeroCard()
                            {
                                Title = "You Don't have permission to view the deliverabels of this projects",
                            };
                            reply.Attachments.Add(plCard.ToAttachment());

                        }
                    }
                    else
                    {
                        context.Load(project, p => p.ProjectSiteUrl);
                        context.ExecuteQuery();

                        projectsite = project.ProjectSiteUrl;
                        projectweb = GetProjectWEB(projectsite, context);

                        var delive = projectweb.Lists.GetByTitle(Enums.ListName.Deliverables.ToString());
                        CamlQuery query = CamlQuery.CreateAllItemsQuery();
                        ListItemCollection itemsdelive = delive.GetItems(query);

                        projectweb.Context.Load(delive);
                        projectweb.Context.Load(itemsdelive);
                        projectweb.Context.ExecuteQuery();
                        reply = GetAllDeliverabels(dialogContext, itemsdelive, itemStartIndex, out TaskCounter);
                    }

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
            string projectsite = string.Empty;
            int TaskCounter = 0;

            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                //       SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                PublishedProject project = GetProjectByName(pName, context);
                Counter = 0;
                if (project != null)
                {
                    context.Load(project.Assignments);
                    context.ExecuteQuery();
                    PublishedAssignmentCollection itemsAssignments = project.Assignments;
                    if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                    {
                        reply = GetResourceLoggedInAssignments(dialogContext, context, itemsAssignments, itemStartIndex, _userName, out TaskCounter);
                    }
                    else if (GetUserGroup(context, "Project Managers (Project Web App Synchronized)"))
                    {
                        context.Load(project.Owner);
                        context.ExecuteQuery();
                        if (project.Owner.Email == _userName) // if the logged in user is a project manager on this project
                        {
                            reply = GetAllAssignments(dialogContext, context, itemsAssignments, itemStartIndex, out TaskCounter);
                        }
                        else
                        {
                            reply = GetResourceLoggedInAssignments(dialogContext, context, itemsAssignments, itemStartIndex, _userName, out TaskCounter);

                        }
                    }
                    else
                    {
                        reply = GetAllAssignments(dialogContext, context, itemsAssignments, itemStartIndex, out TaskCounter);
                    }
                }
            }
            Counter = TaskCounter;
            return reply;
        }

        public IMessageActivity GetProjectMilestones(IDialogContext dialogContext, int itemStartIndex, string pName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

            int TaskCounter = 0;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                PublishedProject project = GetProjectByName(pName, context);


                if (project != null)
                {
                    context.Load(project.Tasks);
                    context.ExecuteQuery();
                    PublishedTaskCollection publishedTask = project.Tasks;
                    if (project.Tasks.Count > 0)
                    {
                        if (GetUserGroup(context, "Team Members (Project Web App Synchronized)"))
                        {
                            // reply = GetResourceLoggedInMilestones(dialogContext, itemStartIndex, context, project, Completed, NotCompleted, delayed, out TaskCounter);
                        }
                        else if (GetUserGroup(context, "Project Managers (Project Web App Synchronized)"))
                        {
                            context.Load(project.Owner);
                            context.ExecuteQuery();
                            if (project.Owner.Email == _userName) // if the logged in user is a project manager on this project
                            {
                                reply = GetProjectMilestones(dialogContext, itemStartIndex, publishedTask, project, out TaskCounter);
                            }
                            else
                            {
                                //reply = GetResourceLoggedInTasks(dialogContext, itemStartIndex, context, project, Completed, NotCompleted, delayed, out TaskCounter);
                            }
                        }
                        else
                        {
                            reply = GetProjectMilestones(dialogContext, itemStartIndex, publishedTask, project, out TaskCounter);
                        }
                    }
                }
            }
            Counter = TaskCounter;
            return reply;
        }


        public IMessageActivity FilterProjectsByDate(IDialogContext dialogContext, string FilterType, string pStartDate, string PEndDate, string ProjectSEdateFlag, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;
            IEnumerable<PublishedProject> retrivedProjects = null; ;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                //      SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);
                DateTime startdate = new DateTime();
                DateTime endate = new DateTime();

                if (!string.IsNullOrEmpty(pStartDate))
                    startdate = DateTime.Parse(pStartDate);

                if (!string.IsNullOrEmpty(PEndDate))
                    endate = DateTime.Parse(PEndDate);


                if (ProjectSEdateFlag == "START")
                {

                    if (FilterType.ToUpper() == "BEFORE" && pStartDate != "")
                    {
                        var pubProjects = context.Projects
                                  .Where(p => (p.StartDate <= startdate))
                                  .Select(p => p);
                        retrivedProjects = context.LoadQuery(pubProjects);


                    }
                    else if (FilterType.ToUpper() == "AFTER" && pStartDate != "")
                    {
                        var pubProjects = context.Projects
                            .Where(p => p.IsEnterpriseProject == true
                            && p.StartDate >= startdate);
                        retrivedProjects = context.LoadQuery(pubProjects);

                    }

                    else if (FilterType.ToUpper() == "BETWEEN" && pStartDate != "")
                    {
                        var pubProjects = context.Projects
                            .Where(p => p.IsEnterpriseProject == true
                            && p.StartDate >= startdate && p.StartDate <= endate);
                        retrivedProjects = context.LoadQuery(pubProjects);

                    }
                }
                else
                {

                    if (FilterType.ToUpper() == "BEFORE" && PEndDate != "")
                    {
                        var pubProjects = context.Projects
                            .Where(p => p.IsEnterpriseProject == true
                            && p.FinishDate <= endate);
                        retrivedProjects = context.LoadQuery(pubProjects);

                    }

                    else if (FilterType.ToUpper() == "AFTER" && PEndDate != "")
                    {
                        var pubProjects = context.Projects
                            .Where(p => p.IsEnterpriseProject == true
                            && p.StartDate >= endate);
                        retrivedProjects = context.LoadQuery(pubProjects);

                    }
                    else if (FilterType.ToUpper() == "BETWEEN" && PEndDate != "")
                    {
                        var pubProjects = context.Projects
                            .Where(p => p.IsEnterpriseProject == true
                            && p.FinishDate >= startdate && p.FinishDate <= endate);
                        retrivedProjects = context.LoadQuery(pubProjects);

                    }
                }

                context.ExecuteQuery();
                if (retrivedProjects.Count() > 0)
                {
                    Counter = retrivedProjects.Count();
                    foreach (var item in retrivedProjects)
                    {
                        string SubtitleVal = "";


                        string ProjectName = item.Name;

                        SubtitleVal += "Start Date\n" + item.StartDate + "</br>";
                        SubtitleVal += "Finish Date\n" + item.FinishDate + "</br>";
                        SubtitleVal += "Actual Cost\n" + item.DefaultFixedCostAccrual.ToString() + "</br>";

                        string ImageURL = "http://02-code.com/images/logo.jpg";
                        List<CardImage> cardImages = new List<CardImage>();
                        List<CardAction> cardactions = new List<CardAction>();
                        cardImages.Add(new CardImage(url: ImageURL));

                        CardAction btnTasks = new CardAction()
                        {
                            Type = ActionTypes.PostBack,
                            Title = "Tasks",
                            Value = "show a list of " + ProjectName + " tasks",
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

                        CardAction btnDAssignments = new CardAction()
                        {
                            Type = ActionTypes.PostBack,
                            Title = "Assignments",
                            Value = "get " + ProjectName + " assignments",
                            Text = "get " + ProjectName + " assignments",

                        };
                        cardactions.Add(btnDAssignments);

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
            }





            return reply;
        }

        public IMessageActivity GetProjectInfo(IDialogContext dialogContext, string pName, bool optionalDate = false, bool optionalDuration = false, bool optionalCompletion = false, bool optionalPM = false)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string SubtitleVal = "";

            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);

                PublishedProject project = GetProjectByName(pName, context);

                if (project != null)
                {
                    if (optionalDate == true)
                    {
                        SubtitleVal += "Start Date :\n" + project.StartDate + "</br>";
                        SubtitleVal += "Finish Date :\n" + project.FinishDate + "</br>";
                    }

                    if (optionalDuration == true)
                    {
                        TimeSpan duration = project.FinishDate - project.StartDate;
                        SubtitleVal += "Project Duration :\n" + duration.Days + "</br>";
                    }

                    if (optionalCompletion == true)
                        SubtitleVal += "Project Completed Percentage :\n" + project.PercentComplete + "%</br>";

                    if (optionalPM == true)
                    {
                        if (GetUserGroup(context, "Team Members (Project Web App Synchronized)") == false)
                        {
                            context.Load(project.Owner);
                            context.ExecuteQuery();
                            SubtitleVal += "Project Manager Name :\n" + project.Owner.Title + "</br>";
                        }
                    }

                    HeroCard plCard = new HeroCard()
                    {
                        Title = pName,
                        Subtitle = SubtitleVal,
                    };
                    reply.Attachments.Add(plCard.ToAttachment());

                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    {
                        Title = "Project Name Not Exist",
                    };
                    reply.Attachments.Add(plCardNoData.ToAttachment());

                }





            }
            return reply;
        }




        private IMessageActivity GetAllTasks(IDialogContext dialogContext, int SIndex, PublishedTaskCollection tskcoll, PublishedProject project, bool Completed, bool NotCompleted, bool delayed, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;


            int inDexToVal = SIndex + 10;
            Counter = project.Tasks.Count;
            if (inDexToVal >= project.Tasks.Count)
                inDexToVal = project.Tasks.Count;

            tskcoll = project.Tasks;



            if (tskcoll.Count > 0)
            {
                if (Completed == true)
                {
                    IEnumerable<PublishedTask> completedlist = project.Tasks.Where(p => p.PercentComplete == 100).ToList();
                    if (completedlist.Count() > 0)
                    {
                        inDexToVal = SIndex + 10;
                        Counter = completedlist.Count();
                        if (inDexToVal >= completedlist.Count())
                            inDexToVal = completedlist.Count();

                        int startIndex = SIndex;
                        foreach (PublishedTask tsk in completedlist)
                        {
                            if (startIndex >= inDexToVal)
                                break;
                            var SubtitleVal = "";
                            string TaskName = tsk.Name;
                            string TaskDuration = tsk.Duration;
                            string TaskPercentCompleted = tsk.PercentComplete.ToString();
                            string TaskStartDate = tsk.Start.ToString();
                            string TaskFinishDate = tsk.Finish.ToString();

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
                            startIndex++;

                        }
                    }
                }
                else if (NotCompleted == true)
                {
                    IEnumerable<PublishedTask> completedlist = project.Tasks.Where(p => p.PercentComplete < 100).ToList();
                    if (completedlist.Count() > 0)
                    {
                        inDexToVal = SIndex + 10;
                        Counter = completedlist.Count();
                        if (inDexToVal >= completedlist.Count())
                            inDexToVal = completedlist.Count();

                        int startIndex = SIndex;
                        foreach (PublishedTask tsk in completedlist)
                        {
                            if (startIndex >= inDexToVal)
                                break;
                            var SubtitleVal = "";
                            string TaskName = tsk.Name;
                            string TaskDuration = tsk.Duration;
                            string TaskPercentCompleted = tsk.PercentComplete.ToString();
                            string TaskStartDate = tsk.Start.ToString();
                            string TaskFinishDate = tsk.Finish.ToString();

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
                            startIndex++;

                        }
                    }
                }
                else if (delayed == true)
                {
                    IEnumerable<PublishedTask> completedlist = project.Tasks.Where(p => p.ActualDuration == p.ScheduledDuration).ToList();
                    if (completedlist.Count() > 0)
                    {
                        inDexToVal = SIndex + 10;
                        Counter = completedlist.Count();
                        if (inDexToVal >= completedlist.Count())
                            inDexToVal = completedlist.Count();

                        int startIndex = SIndex;
                        foreach (PublishedTask tsk in completedlist)
                        {
                            if (startIndex >= inDexToVal)
                                break;
                            var SubtitleVal = "";
                            string TaskName = tsk.Name;
                            string TaskDuration = tsk.Duration;
                            string TaskPercentCompleted = tsk.PercentComplete.ToString();
                            string TaskStartDate = tsk.Start.ToString();
                            string TaskFinishDate = tsk.Finish.ToString();

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
                            startIndex++;

                        }
                    }
                }
                else
                {

                    for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                    {
                        var SubtitleVal = "";
                        PublishedTask tsk = tskcoll[startIndex];
                        string TaskName = tsk.Name;
                        string TaskDuration = tsk.Duration;
                        string TaskPercentCompleted = tsk.PercentComplete.ToString();
                        string TaskStartDate = tsk.Start.ToString();
                        string TaskFinishDate = tsk.Finish.ToString();



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

            }

            return reply;
        }

        private IMessageActivity GetProjectMilestones(IDialogContext dialogContext, int SIndex, PublishedTaskCollection tskcoll, PublishedProject project, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;


            int inDexToVal = SIndex + 10;
            Counter = project.Tasks.Count;
            if (inDexToVal >= project.Tasks.Count)
                inDexToVal = project.Tasks.Count;

            tskcoll = project.Tasks;
            if (tskcoll.Count > 0)
            {
                IEnumerable<PublishedTask> cmilestonelist = project.Tasks.Where(p => p.Duration == "0d").ToList();
                if (cmilestonelist.Count() > 0)
                {
                    inDexToVal = SIndex + 10;
                    Counter = cmilestonelist.Count();
                    if (inDexToVal >= cmilestonelist.Count())
                        inDexToVal = cmilestonelist.Count();

                    int startIndex = SIndex;
                    foreach (PublishedTask tsk in cmilestonelist)
                    {
                        if (startIndex >= inDexToVal)
                            break;
                        var SubtitleVal = "";
                        string TaskName = tsk.Name;
                        string TaskPercentCompleted = tsk.PercentComplete.ToString();
                        string TaskStartDate = tsk.Start.ToString();
                        string TaskFinishDate = tsk.Finish.ToString();

                        SubtitleVal += "Task Percent Completed\n" + TaskPercentCompleted + "</br>";
                        SubtitleVal += "Task Start Date\n" + TaskStartDate + "</br>";
                        SubtitleVal += "Task Finish Date\n" + TaskFinishDate + "</br>";

                        HeroCard plCard = new HeroCard()
                        {
                            Title = TaskName,
                            Subtitle = SubtitleVal
                        };

                        reply.Attachments.Add(plCard.ToAttachment());
                        startIndex++;

                    }
                }
            }




            return reply;
        }

        private IMessageActivity GetResourceLoggedInTasks(IDialogContext dialogContext, int SIndex, ProjectContext context, PublishedProject proj, bool Completed, bool NotCompleted, bool delayed, out int Counter)
        {
            var SubtitleVal = "";
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            context.Load(proj.Assignments, da => da.Where(a => a.Resource.Email != string.Empty && a.Resource.Email == _userName));
            context.ExecuteQuery();
            Counter = 0;



            if (proj.Assignments != null)
            {
                PublishedAssignmentCollection proAssignment = proj.Assignments;

                int inDexToVal = SIndex + 10;
                Counter = proAssignment.Count;
                if (inDexToVal >= proAssignment.Count)
                    inDexToVal = proAssignment.Count;

                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    PublishedAssignment ass = proAssignment[startIndex];
                    context.Load(ass.Task);
                    context.Load(ass.Resource);

                    context.ExecuteQuery();
                    var tsk = ass.Task;
                    string TaskName = tsk.Name;
                    string TaskDuration = tsk.Duration;
                    string TaskPercentCompleted = tsk.PercentComplete.ToString();
                    string TaskStartDate = tsk.Start.ToString();
                    string TaskFinishDate = tsk.Finish.ToString();

                    SubtitleVal += "Task Duration\n" + TaskDuration + "</br>";
                    SubtitleVal += "Task Percent Completed\n" + TaskPercentCompleted + "</br>";
                    SubtitleVal += "Task Start Date\n" + TaskStartDate + "</br>";
                    SubtitleVal += "Task Finish Date\n" + TaskFinishDate + "</br>";

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

        private IMessageActivity GetAllIssues(IDialogContext dialogContext, ListItemCollection itemsIssue, int SIndex, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string IssueName = string.Empty;
            string IssueStatus = string.Empty;
            string IssuePriority = string.Empty;

            Counter = itemsIssue.Count;

            int inDexToVal = SIndex + 10;
            if (inDexToVal >= itemsIssue.Count)
                inDexToVal = itemsIssue.Count;

            if (itemsIssue.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    ListItem item = itemsIssue[startIndex];
                    if (item["Title"] != null)
                        IssueName = (string)item["Title"];
                    if (item["Status"] != null)
                        IssueStatus = (string)item["Status"];
                    if (item["Priority"] != null)
                        IssuePriority = (string)item["Priority"];
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

        private IMessageActivity GetResourceLoggedInIssues(IDialogContext dialogContext, ListItemCollection itemsIssue, int SIndex, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            Counter = 0;



            int inDexToVal = SIndex + 10;
            if (inDexToVal >= itemsIssue.Count)
                inDexToVal = itemsIssue.Count;


            if (itemsIssue.Count > 0)
            {
                int count = 0;
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    ListItem item = itemsIssue[startIndex];

                    if (item["AssignedTo"] != null)
                    {
                        count++;
                        FieldUserValue fuv = (FieldUserValue)item["AssignedTo"];
                        if (fuv.Email == _userName)
                        {
                            string SubtitleVal = "";
                            string IssueName = string.Empty;
                            string IssueStatus = string.Empty;
                            string IssuePriority = string.Empty;

                            if (item["Title"] != null)
                                IssueName = (string)item["Title"];
                            if (item["Status"] != null)
                                IssueStatus = (string)item["Status"];
                            if (item["Priority"] != null)
                                IssuePriority = (string)item["Priority"];
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
                }
                Counter = count;
            }
            return reply;
        }

        private IMessageActivity GetAllRisks(IDialogContext dialogContext, ListItemCollection itemsRisk, int SIndex, out int Counter)
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


            Counter = itemsRisk.Count;

            int inDexToVal = SIndex + 10;
            if (inDexToVal >= itemsRisk.Count)
                inDexToVal = itemsRisk.Count;

            if (itemsRisk.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    ListItem item = itemsRisk[startIndex];


                    if (item["Title"] != null)
                        RiskName = (string)item["Title"];
                    if (item["AssignedTo"] != null)
                    {
                        FieldUserValue fuv = (FieldUserValue)item["AssignedTo"];
                        SubtitleVal += "Assigned To Resource\n" + fuv.LookupValue + "</br>";

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
        private IMessageActivity GetResourceLoggedInRisks(IDialogContext dialogContext, ListItemCollection itemsRisk, int SIndex, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;
            string RiskName = string.Empty;
            string ResourceName = string.Empty;
            string riskStatus = string.Empty;
            string riskImpact = string.Empty;
            string riskProbability = string.Empty;
            string riskCostExposure = string.Empty;
            if (itemsRisk.Count > 0)
            {
                int count = 0;

                int inDexToVal = SIndex + 10;
                if (inDexToVal >= itemsRisk.Count)
                    inDexToVal = itemsRisk.Count;
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    ListItem item = itemsRisk[startIndex];


                    if (item["AssignedTo"] != null)
                    {
                        count++;
                        FieldUserValue fuv = (FieldUserValue)item["AssignedTo"];
                        if (fuv.Email == _userName)
                        {
                            if (item["Title"] != null)
                                RiskName = (string)item["Title"];
                            SubtitleVal += "Risk Title\n" + RiskName + "</br>";

                            SubtitleVal += "Assigned To Resource\n" + fuv.LookupValue + "</br>";
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
                    Counter = count;


                }

            }

            return reply;
        }

        private IMessageActivity GetAllDeliverabels(IDialogContext dialogContext, ListItemCollection itemsDeliverabels, int SIndex, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            string DeliverableName = string.Empty;
            string DeliverableStart = string.Empty;
            string DeliverableFinish = string.Empty;


            Counter = itemsDeliverabels.Count;

            int inDexToVal = SIndex + 10;
            if (inDexToVal >= itemsDeliverabels.Count)
                inDexToVal = itemsDeliverabels.Count;

            if (itemsDeliverabels.Count > 0)
            {
                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    var SubtitleVal = "";
                    ListItem item = itemsDeliverabels[startIndex];


                    if (item["Title"] != null)
                        DeliverableName = (string)item["Title"];
                    SubtitleVal += "Deliverable Name\n" + DeliverableName + "</br>";

                    if (item["Author"] != null)
                    {
                        FieldUserValue fuv = (FieldUserValue)item["Author"];
                        SubtitleVal += "Create By Resource :\n" + fuv.LookupValue + "</br>";

                    }

                    if (item["CommitmentStart"] != null)
                        DeliverableStart = item["CommitmentStart"].ToString();
                    SubtitleVal += "Start Date :\n" + DeliverableStart + "</br>";

                    if (item["CommitmentFinish"] != null)
                        DeliverableFinish = item["CommitmentFinish"].ToString();
                    SubtitleVal += "Finish Date :\n" + DeliverableFinish + "</br>";


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



        private IMessageActivity GetAllAssignments(IDialogContext dialogContext, ProjectContext context, PublishedAssignmentCollection itemsAssignments, int SIndex, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = itemsAssignments.Count;
            if (itemsAssignments.Count > 0)
            {



                int inDexToVal = SIndex + 10;
                if (inDexToVal >= itemsAssignments.Count)
                    inDexToVal = itemsAssignments.Count;

                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    PublishedAssignment ass = itemsAssignments[startIndex];
                    string SubtitleVal = "";
                    context.Load(ass.Task);
                    context.Load(ass.Resource);
                    context.ExecuteQuery();

                    string TaskName = ass.Task.Name;
                    SubtitleVal += "Resource Name :\n" + ass.Resource.Name + "</br>";
                    SubtitleVal += "Start Date\n" + ass.Start + "</br>";
                    SubtitleVal += "Finish Date\n" + ass.Finish + "</br>";


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


        public IMessageActivity GetResourceAssignments(IDialogContext dialogContext, int SIndex, string ResourceName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            int counttotalAss = 0;
            Counter = 0;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {


                SecureString passWord = new SecureString();
                foreach (char c in _userPassword.ToCharArray()) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(_userName, passWord);


                context.Load(context.Projects);
                context.ExecuteQuery();
                ProjectCollection projcoll = context.Projects;


                context.Load(context.EnterpriseResources);
                var resources = context.EnterpriseResources;

                context.ExecuteQuery();


                ResourceName = ResourceName.Replace(" ", String.Empty);

                string fullEmail = string.Concat(ResourceName, ConfigurationManager.AppSettings["DomainEmail"]);

                var user = context.Web.EnsureUser(ResourceName);
                context.Load(user);
                context.ExecuteQuery();

                if (user != null)
                {
                    var resource = resources.FirstOrDefault(i => i.Email == user.Email);
                    if (resource != null)
                    {

                        foreach (PublishedProject proj in projcoll)
                        {
                            context.Load(proj.Assignments, da => da.Where(a => a.Resource.Email == user.Email));

                            context.ExecuteQuery();
                            PublishedAssignmentCollection itemsAssignments = proj.Assignments;

                            if (proj.Assignments != null)
                            {
                                int inDexToVal = SIndex + 10;
                                if (inDexToVal >= itemsAssignments.Count)
                                    inDexToVal = itemsAssignments.Count;

                                PublishedAssignmentCollection proAssignment = proj.Assignments;
                                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                                {
                                    PublishedAssignment ass = itemsAssignments[startIndex];
                                    string SubtitleVal = "";
                                    context.Load(ass.Task);
                                    context.ExecuteQuery();
                                    var tsk = ass.Task;
                                    SubtitleVal += "Assignment Start Date :\n" + ass.Start + "</br>";
                                    SubtitleVal += "Task Name :\n" + tsk.Name + "</br>";
                                    counttotalAss++;


                                    HeroCard plCard = new HeroCard()
                                    {
                                        Title = proj.Name,
                                        Subtitle = SubtitleVal,
                                    };
                                    reply.Attachments.Add(plCard.ToAttachment());
                                }
                                Counter = counttotalAss;

                            }
                        }


                    }

                }




                return reply;
            }

        }


        public IMessageActivity GetResourceLoggedInAssignments(IDialogContext dialogContext, ProjectContext context, PublishedAssignmentCollection itemsAssignments, int SIndex, string ResourceName, out int Counter)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            Counter = 0;
            int count = 0;
            if (itemsAssignments.Count > 0)
            {
                int inDexToVal = SIndex + 10;
                if (inDexToVal >= itemsAssignments.Count)
                    inDexToVal = itemsAssignments.Count;

                for (int startIndex = SIndex; startIndex < inDexToVal; startIndex++)
                {
                    PublishedAssignment ass = itemsAssignments[startIndex];

                    context.Load(ass.Task);
                    context.Load(ass.Resource);
                    context.ExecuteQuery();

                    if (ass.Resource.Email == ResourceName)
                    {
                        count++;
                        string SubtitleVal = "";
                        string TaskName = ass.Task.Name;
                        SubtitleVal += "Resource Name :\n" + ass.Resource.Name + "</br>";
                        SubtitleVal += "Start Date\n" + ass.Start + "</br>";
                        SubtitleVal += "Finish Date\n" + ass.Finish + "</br>";


                        HeroCard plCard = new HeroCard()
                        {
                            Title = TaskName,
                            Subtitle = SubtitleVal,
                        };
                        reply.Attachments.Add(plCard.ToAttachment());
                    }
                }
                Counter = count;
            }
            return reply;

        }
        public bool UserHavePermissionOnaProjects(string siteUrl, string subSiteTitle, ProjectContext context)
        {

            var web = context.Web;
            bool exist = false;
            context.Load(web, w => w.Webs);
            context.ExecuteQuery();
            foreach (Web subWeb in web.Webs)
            {
                if (subWeb.Title.ToLower() == subSiteTitle.ToLower())
                {
                    var user = subWeb.EnsureUser(_userName);
                    context.Load(user);
                    context.ExecuteQuery();

                    if (null != user)
                    {
                        ClientResult<BasePermissions> permissions = subWeb.GetUserEffectivePermissions(user.LoginName);
                        context.ExecuteQuery();


                        if (permissions.Value.Has(PermissionKind.ViewListItems))
                        {
                            exist = true;
                            break;
                        }
                        else
                            exist = false;


                    }
                    else
                        exist = false;




                }
            }

            return exist;
        }

        public bool GetUserGroup(ProjectContext context, string groupName)
        {
            bool exist = false;

            context.Load(context.Web);

            //  context.Load(web.SiteUsers);
            context.ExecuteQuery();

            Web web = context.Web;

            IEnumerable<User> user = context.LoadQuery(web.SiteUsers.Where(p => p.Email == _userName));
            context.ExecuteQuery();

            if (user.Any())
            {
                User userLogged = user.FirstOrDefault();

                context.Load(userLogged.Groups);
                context.ExecuteQuery();

                GroupCollection group = userLogged.Groups;

                IEnumerable<Group> usergroup = context.LoadQuery(userLogged.Groups.Where(p => p.Title == groupName));
                context.ExecuteQuery();
                if (!usergroup.Any())
                {
                    exist = false;
                }
                else
                    exist = true;
            }

            return exist;
        }

        private static PublishedProject GetProjectByName(string name, ProjectContext context)
        {
            if (name.Contains(" - "))
                name = name.Replace(" - ", "-");
            IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Name == name));
            context.ExecuteQuery();
            if (!projs.Any())       // no project found
            {
                return null;
            }
            return projs.FirstOrDefault();

        }

        private static Web GetProjectWEB(string siteurl, ProjectContext context)
        {
            IEnumerable<Web> webs = context.LoadQuery(context.Web.Webs.Where(p => p.Url == siteurl));
            context.ExecuteQuery();
            if (!webs.Any())       // no project found
            {
                return null;
            }
            return webs.FirstOrDefault();

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
                    { Title = "No Available  Tasks\n\n" };
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
                            Title = "Total Number Of Available Issues :\n" + Counter,
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
                            Title = "Total Number Of Available Issues :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Issues\n\n" };
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
                            Title = "Total Number Of Available Assignments :\n" + Counter,
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
                            Title = "Total Number Of Available Assignments :\n" + Counter,
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
                            Title = "Total Number Of Available Risks :\n" + Counter,
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
                            Title = "Total Number Of Available Risks :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Risks\n\n" };
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
                            Title = "Total Number Of Available Deliverables :\n" + Counter,
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
                            Title = "Total Number Of Available Deliverables :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Deliverables\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }
            else if (ListName == "FilterByDate")
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
                            subTitle = "You are viwing the page number " + pagenumber + " , each page view 10 Deliverables";
                        }
                        HeroCard plCardCounter = new HeroCard()
                        {
                            Title = "Total Number Of Available Projects :\n" + Counter,
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
                            Title = "Total Number Of Available Projects :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Projects\n\n" };
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
                    { Title = "No Available assignments\n\n" };
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
                            Title = "Total Number Of Available Milestones :\n" + Counter,
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
                            Title = "Total Number Of Available Milestones :\n" + Counter,
                        };
                        reply.Attachments.Add(plCardCounter.ToAttachment());
                    }
                }
                else
                {
                    HeroCard plCardNoData = new HeroCard()
                    { Title = "No Available Milestones\n\n" };
                    reply.Attachments.Add(plCardNoData.ToAttachment());
                }
            }

            return reply;
        }

        public IMessageActivity CreateButtonsPager(IDialogContext dialogContext, int totalCount, string ListName, string projectName, string query)
        {
            IMessageActivity reply = null;
            reply = dialogContext.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            if (totalCount > 10)
            {
                List<CardAction> cardButtons = new List<CardAction>();
                double p = totalCount * 0.1;
                double result = Math.Ceiling(p);
                int pagenumber = int.Parse(result.ToString());

                string valuebutton = string.Empty;
                for (int i = 0; i < pagenumber; i++)
                {
                    string CurrentNumber = Convert.ToString(i);
                    if (projectName == "")
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
                    else if (ListName == "FilterByDate" && query != "")
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

                CardAction btnCompltedPcurrentYear = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Completed Projects This Year",
                    Value = "get all projects closed this year",
                    Text = "get all projects closed this year",
                };
                cardactions.Add(btnCompltedPcurrentYear);
                CardAction btnStartedPcurrentYear = new CardAction()
                {
                    Type = ActionTypes.PostBack,
                    Title = "Started Projects This Year",
                    Value = "get all projects started this year",
                    Text = "get all projects started this year",
                };
                cardactions.Add(btnStartedPcurrentYear);
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

        public bool GetUserGroupAPI(string groupName)
        {
            bool exist = false;
            using (ProjectContext context = new ProjectContext(_siteUri))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _userPasswordAdmin.ToCharArray()) passWord.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(_userNameAdmin, passWord);

                context.Load(context.Web);
                context.ExecuteQuery();

                Web web = context.Web;

                IEnumerable<User> user = context.LoadQuery(web.SiteUsers.Where(p => p.Email == _userName));
                context.ExecuteQuery();

                if (user.Any())
                {
                    User userLogged = user.FirstOrDefault();

                    context.Load(userLogged.Groups);
                    context.ExecuteQuery();

                    GroupCollection group = userLogged.Groups;

                    IEnumerable<Group> usergroup = context.LoadQuery(userLogged.Groups.Where(p => p.Title == groupName));
                    context.ExecuteQuery();
                    if (!usergroup.Any())
                    {
                        exist = false;
                    }
                    else
                        exist = true;
                }
            }
            return exist;
        }
    }
}
