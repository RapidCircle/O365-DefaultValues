using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace DefaultValuesWeb.Services
{
    public class RemoteEventReceiverManager
    {
        private const string RECEIVER_NAME_ADDED = "DefaultValues";
        private const string LIST_TITLE = "documents";


        public void ItemUpdatedToListEventHandler(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                UpdateDefaultValues(clientContext, listId, listItemId);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemAddingToListEventHandler(SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            try
            {
                UpdateDefaultValuesAdding(properties, result);
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            var rerList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(rerList);
            clientContext.ExecuteQuery();

            bool rerExists = false;
            if (!rerExists)
            {
                //Get WCF URL where this message was handled
                OperationContext op = OperationContext.Current;
                Message msg = op.RequestContext.RequestMessage;

                EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();

                receiver.EventType = EventReceiverType.ItemAdding;
                receiver.ReceiverUrl = msg.Headers.To.ToString();
                receiver.ReceiverName = EventReceiverType.ItemAdding.ToString();
                receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                receiver.SequenceNumber = 350;
                rerList.EventReceivers.Add(receiver);
                clientContext.ExecuteQuery();
                System.Diagnostics.Trace.WriteLine("Added ItemAdded receiver at " + receiver.ReceiverUrl);

                receiver = new EventReceiverDefinitionCreationInformation();
                receiver.EventType = EventReceiverType.ItemUpdated;
                receiver.ReceiverUrl = msg.Headers.To.ToString();
                receiver.ReceiverName = EventReceiverType.ItemUpdated.ToString();
                receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                receiver.SequenceNumber = 350;
                rerList.EventReceivers.Add(receiver);
                clientContext.ExecuteQuery();

                System.Diagnostics.Trace.WriteLine("Added ItemUpdated receiver at " + receiver.ReceiverUrl);
            }
        }

        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext)
        {
            List myList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
            clientContext.Load(myList, p => p.EventReceivers);
            clientContext.ExecuteQuery();

            var rer = myList.EventReceivers.Where(e => e.ReceiverName == RECEIVER_NAME_ADDED).FirstOrDefault();

            try
            {
                System.Diagnostics.Trace.WriteLine("Removing receiver at "
                        + rer.ReceiverUrl);

                var rerList = myList.EventReceivers.Where(e => e.ReceiverUrl == rer.ReceiverUrl).ToList<EventReceiverDefinition>();

                foreach (var rerFromUrl in rerList)
                {
                    //This will fail when deploying via F5, but works
                    //when deployed to production
                    rerFromUrl.DeleteObject();
                }
                clientContext.ExecuteQuery();
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

            clientContext.ExecuteQuery();
        }

        ///// <summary>
        ///// Handles the ItemAdded event by modifying the Description
        ///// field of the item.
        ///// </summary>
        ///// <param name="properties"></param>
        private void UpdateDefaultValuesAdding(SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    var itemProperties = properties.ItemEventProperties;
                    var _userLoginName = properties.ItemEventProperties.UserLoginName;
                    var _afterProperites = itemProperties.AfterProperties;
                    string phase = "";

                    List list = clientContext.Web.Lists.GetById(properties.ItemEventProperties.ListId);
                    //ListItem item = list.GetItemById(properties.ItemEventProperties.ListItemId);
                    FieldCollection fields = list.Fields;
                    //clientContext.Load(item);
                    clientContext.Load(fields);
                    clientContext.ExecuteQuery();

                    string[] folderNameSplit = itemProperties.BeforeUrl.ToString().Split('/');
                    string folderName = folderNameSplit[folderNameSplit.Length - 2];
                    string code = folderName.Split('_').First();

                    string lookupFieldName = "Phase";
                    FieldLookup lookupField = clientContext.CastTo<FieldLookup>(fields.GetByInternalNameOrTitle(lookupFieldName));
                    clientContext.Load(lookupField);
                    clientContext.ExecuteQuery();

                    Guid parentWeb = lookupField.LookupWebId;
                    string parentListGUID = lookupField.LookupList;

                    Web lookupListWeb = clientContext.Site.OpenWebById(parentWeb);
                    List parentList = lookupListWeb.Lists.GetById(new Guid(parentListGUID));

                    CamlQuery cq = new CamlQuery();
                    cq.ViewXml = "<ViewFields><FieldRef Name='Title' /><FieldRef Name='ID' /></ViewFields>";

                    ListItemCollection litems = parentList.GetItems(cq);
                    clientContext.Load(parentList);
                    clientContext.Load(litems);
                    clientContext.ExecuteQuery();

                    FieldLookupValue flv = new FieldLookupValue();
                    foreach (ListItem li in litems)
                    {
                        string phaseItemCode = li["Title"].ToString().TrimEnd(')').Split('(').Last();
                        if (phaseItemCode == code)
                        {
                            if (itemProperties.AfterProperties.ContainsKey("Phase") && !String.IsNullOrEmpty(itemProperties.AfterProperties["Phase"].ToString()))
                            {
                                string[] stringSeparators = new string[] { ";#" };
                                if (itemProperties.AfterProperties["Phase"].ToString().Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)[0] == li["ID"].ToString()) break;
                            }

                            phase = li.Id + ";#" + li["Title"];
                            break;
                        }
                    }

                    result.ChangedItemProperties.Add("Phase", phase);
                    result.Status = SPRemoteEventServiceStatus.Continue;
                }
            }
        }

        public void UpdateDefaultValues(ClientContext clientContext, Guid listId, int listItemId)
        {
            List list = clientContext.Web.Lists.GetById(listId);
            ListItem item = list.GetItemById(listItemId);
            FieldCollection fields = list.Fields;
            clientContext.Load(item);
            clientContext.Load(fields);
            clientContext.ExecuteQuery();

            if (item["Phase"] != null && !string.IsNullOrEmpty(item["Phase"].ToString()))
                return;

            string folderName = item["FileDirRef"].ToString().Split('/').Last();
            string code = folderName.Split('_').First();

            string lookupFieldName = "Phase";
            FieldLookup lookupField = clientContext.CastTo<FieldLookup>(fields.GetByInternalNameOrTitle(lookupFieldName));
            clientContext.Load(lookupField);
            clientContext.ExecuteQuery();

            Guid parentWeb = lookupField.LookupWebId;
            string parentListGUID = lookupField.LookupList;

            Web lookupListWeb = clientContext.Site.OpenWebById(parentWeb);
            List parentList = lookupListWeb.Lists.GetById(new Guid(parentListGUID));

            CamlQuery cq = new CamlQuery();
            cq.ViewXml = "<ViewFields><FieldRef Name='Title' /><FieldRef Name='ID' /></ViewFields>";

            ListItemCollection litems = parentList.GetItems(cq);
            clientContext.Load(parentList);
            clientContext.Load(litems);
            clientContext.ExecuteQuery();

            foreach (ListItem li in litems)
            {
                string phaseItemCode = li["Title"].ToString().TrimEnd(')').Split('(').Last();
                if (phaseItemCode == code)
                {
                    string value = item.Id + ";#" + item["Title"];
                    if (item["Phase"] != null)
                    {
                        FieldLookupValue flv = (FieldLookupValue)item["Phase"];
                        string existingCode = flv.LookupValue.ToString().TrimEnd(')').Split('(').Last();
                        if (existingCode == code) break;
                    }

                    item["Phase"] = li.Id + ";#" + li["Title"];
                    item.Update();
                    clientContext.ExecuteQuery();
                    break;
                }
            }
        }
    }
}