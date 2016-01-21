using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace HostWebSPRemoteReceiverWeb
{
    public class RemoteEventReceiverManager
    {
        private const string RECEIVER_NAME = "ItemAddedEvent";
        private const string LIST_TITLE = "Announcements";

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {

            //get the title and event receiver list
            clientContext.Load(clientContext.Web.Lists,
                lists => lists.Include(list => list.Title, list => list.EventReceivers)
                    .Where(list => list.Title == LIST_TITLE));

            clientContext.ExecuteQuery();

            List targetList = clientContext.Web.Lists.FirstOrDefault();

            bool rerExists = false;
            if( null != targetList)
            {
                foreach(var rer in targetList.EventReceivers)
                {
                    if(rer.ReceiverName == RECEIVER_NAME)
                    {
                        rerExists = true;
                        System.Diagnostics.Trace.WriteLine("Found existing event receiver at " + rer.ReceiverUrl);
                    }
                }
            }
            if(!rerExists)
            {
                EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();

                receiver.EventType = EventReceiverType.ItemAdding;
                //Get WCF URL where this message was handled
                OperationContext op = OperationContext.Current;
                Message msg = op.RequestContext.RequestMessage;

                receiver.ReceiverUrl = msg.Headers.To.ToString();
                receiver.ReceiverName = RECEIVER_NAME;
                receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                targetList.EventReceivers.Add(receiver);
                clientContext.ExecuteQuery();

                System.Diagnostics.Trace.WriteLine("Added ItemAdding event at " + receiver.ReceiverUrl);
            }
        }

        public void ItemAddingListEventHandler(ClientContext clientContext, Guid listId, int itemId)
        {
            try
            {
                List targetList = clientContext.Web.Lists.GetById(listId);
                //ListItem item = targetList.GetItemById(itemId);

                clientContext.Load(targetList);
                clientContext.ExecuteQuery();

                if(targetList != null)
                {
                    targetList.EnableVersioning = true;
                    targetList.MajorVersionLimit = 3;
                    targetList.Update();
                    clientContext.ExecuteQuery();
                    System.Diagnostics.Trace.WriteLine("List version is enable and set to Major version with limit 3.");
                }
            }
            catch(Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }
    }
}