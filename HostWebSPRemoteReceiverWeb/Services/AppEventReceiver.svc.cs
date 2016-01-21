using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace HostWebSPRemoteReceiverWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            switch(properties.EventType)
            {
                case SPRemoteEventType.AppInstalled:
                    HandleAppInstalled(properties);
                    break;
                
                
            }

            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            //Handles asynchronus events
            switch(properties.EventType)
            {
                case SPRemoteEventType.ListAdded:
                    HandleListAdded(properties);
                    break;
            }
        }

        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {
            //Below code is used to register ListAdded event on Host Web 
            //Note the token helper method used - second paratemer is to whether to create context with App or not
            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false)) 
            {
                if(clientContext != null)
                {
                    new RemoteEventReceiverManager().AssociateRemoteEventsToHostWeb(clientContext);
                }
            }
        }

        private void HandleItemAdding(SPRemoteEventProperties properties)
        {
            using(ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if(null != clientContext)
                {
                    new RemoteEventReceiverManager().ItemAddingListEventHandler(clientContext, properties.ItemEventProperties.ListId, properties.ItemEventProperties.ListItemId);
                }
            }
        }

        private void HandleListAdded(SPRemoteEventProperties properties)
        {
            //Below code is called when list created.
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (null != clientContext)
                {
                    new RemoteEventReceiverManager().ListAddedEventHandler(clientContext, properties.ListEventProperties.ListId);
                }
            }
        }

    }
}
