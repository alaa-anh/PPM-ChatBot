using System;
using System.Threading.Tasks;
using System.Web.Http;

using Microsoft.Bot.Connector;
using Microsoft.Bot.Builder.Dialogs;
using System.Web.Http.Description;
using System.Net.Http;
using System.Diagnostics;
using System.Collections.Generic;
using System.Web;
using Common;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Autofac;
//using LuisBot.Forms;
using Microsoft.Bot.Builder.FormFlow;
using PM.BotApplication.Dialogs;

namespace PM.BotApplication
{
    [BotAuthentication]
    public class PPMController : ApiController
    {

        protected string prompt { get; }
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        [ResponseType(typeof(void))]
        public virtual async Task<HttpResponseMessage> Post([FromBody] Activity activity)
        {
            if (activity.GetActivityType() == ActivityTypes.Message)
            {
                //if (activity.ChannelId == "webchat")
                //    await Conversation.SendAsync(activity, () => new PPMDialog(activity));
                //else if (activity.ChannelId == "telegram")
                //    await Conversation.SendAsync(activity, () => new PPMDialogTelegram(activity));
                //else if (activity.ChannelId == "emulator")
                //    await Conversation.SendAsync(activity, () => new PPMDialog(activity));
                //else if (activity.ChannelId == "team")
                    await Conversation.SendAsync(activity, () => new PPMDialogMSTeam(activity));
                //else
                //    await Conversation.SendAsync(activity, () => new PPMDialog(activity));

            }
            else
            {
                HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(System.Net.HttpStatusCode.OK);
            return response;
        }

        

        private Activity HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {
            }

            return null;
        }
    }
}