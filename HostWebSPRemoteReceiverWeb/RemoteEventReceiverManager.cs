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
        private const string RECEIVER_NAME = "LibraryAddedEvent";
        private const string LIST_TITLE = "Announcements";

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            try
            {
                //Getting Host Web
                var hostWeb = clientContext.Web;
                clientContext.Load(hostWeb);

                //Getting event receivers on the web
                clientContext.Load(hostWeb.EventReceivers);
                clientContext.ExecuteQuery();

                bool rerExists = false;
                if(null != hostWeb && hostWeb.EventReceivers.Count > 0)
                {
                    foreach(var rer in hostWeb.EventReceivers)
                    {
                        if(rer.ReceiverName == RECEIVER_NAME)
                        {
                            rerExists = true;
                            System.Diagnostics.Trace.WriteLine("Found existing event receiver at " + rer.ReceiverUrl);
                            break;
                        }
                    }
                }
                if(!rerExists)
                {
                    //registering ListAdding event on the web
                    EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();

                    receiver.EventType = EventReceiverType.ListAdded;
                    //Get WCF URL where this message was handled
                    OperationContext op = OperationContext.Current;
                    Message msg = op.RequestContext.RequestMessage;
                    receiver.ReceiverUrl = msg.Headers.To.ToString();
                    receiver.ReceiverName = RECEIVER_NAME;
                    receiver.Synchronization = EventReceiverSynchronization.Asynchronous;
                    hostWeb.EventReceivers.Add(receiver);
                    clientContext.ExecuteQuery();

                    System.Diagnostics.Trace.WriteLine("Added ListAdded event at " + receiver.ReceiverUrl);
                }
            }
            catch (Exception ex)
            {

            }
            //get the title and event receiver list
            /*clientContext.Load(clientContext.Web.Lists,
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
            }*/
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientContext">SharePoint Client Context Object</param>
        /// <param name="listId">List GUID</param>
        public void ListAddedEventHandler(ClientContext clientContext, Guid listId)
        {
            try
            {
                //Getting newly added list
                List targetList = clientContext.Web.Lists.GetById(listId);

                clientContext.Load(targetList);
                clientContext.ExecuteQuery();
                //Checking if newly added is of type "DocumentLibrary"
                if (targetList != null && targetList.BaseTemplate == (int)ListTemplateType.DocumentLibrary)
                {
                    targetList.EnableVersioning = true;
                    targetList.MajorVersionLimit = 3;
                    targetList.Update();
                    clientContext.ExecuteQuery();
                    System.Diagnostics.Trace.WriteLine("List version is enable and set to Major version with limit 3.");
                }
            }
            catch(Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
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

                if (targetList != null)
                {
                    targetList.EnableVersioning = true;
                    targetList.MajorVersionLimit = 3;
                    targetList.Update();
                    clientContext.ExecuteQuery();
                    System.Diagnostics.Trace.WriteLine("List version is enable and set to Major version with limit 3.");
                }
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }
    }
}