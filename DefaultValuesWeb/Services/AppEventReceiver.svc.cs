using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Runtime.Caching;

namespace DefaultValuesWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {

        //the cache needs an unique name
        private static MemoryCache _cacheObject = new MemoryCache("26397979caf840b39c36368fa9b1bb9f");

        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult _result = new SPRemoteEventResult();

            try
            {
                switch (properties.EventType)
                {
                    case SPRemoteEventType.AppInstalled:
                        HandleAppInstalled(properties);
                        break;
                    case SPRemoteEventType.AppUninstalling:
                        HandleAppUninstalling(properties);
                        break;
                    case SPRemoteEventType.ItemAdding:
                        new RemoteEventReceiverManager().ItemAddingToListEventHandler(properties, _result);
                        break;
                    case SPRemoteEventType.ItemUpdated:
                        HandleItemUpdated(properties);
                        break;
                }
                _result.Status = SPRemoteEventServiceStatus.Continue;
            }
            catch (Exception ex)
            {
                //You should log here.               
            }
            return _result;
        }

        /// <summary>
        /// Handles the ItemUpdating event by check the Description
        /// field of the item.
        /// </summary>
        /// <param name="properties"></param>
        private void HandleItemUpdated(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)// && EventFiringEnabled(clientContext, properties, "2"))
                {
                    new RemoteEventReceiverManager().ItemUpdatedToListEventHandler(clientContext, properties.ItemEventProperties.ListId, properties.ItemEventProperties.ListItemId);
                }
            }
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
        }

        /// <summary>
        /// Handles when an app is installed.  Activates a feature in the
        /// host web.  The feature is not required.  
        /// Next, if the Jobs list is
        /// not present, creates it.  Finally it attaches a remote event
        /// receiver to the list.  
        /// </summary>
        /// <param name="properties"></param>
        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().AssociateRemoteEventsToHostWeb(clientContext);
                }
            }
        }

        /// <summary>
        /// Removes the remote event receiver from the list and 
        /// adds a new item to the list.
        /// </summary>
        /// <param name="properties"></param>
        private void HandleAppUninstalling(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext =
                TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext);
                }
            }
        }

        /// <summary>
        /// Used to chek if the event should processed or not. we use here a local MemoryCache object 
        /// or if the RER is loadbalanced a Redis Server on AZURE or onPrem
        /// </summary>
        /// <param name="ctx">the current context as ClientContext</param>
        /// <param name="properties">the event properies as SPRemoteEventProperties</param>
        /// <param name="eventType">identifier for event receiver type (~ing=1, ~ed=2) as String</param>
        /// <returns></returns>
        private static bool EventFiringEnabled(ClientContext ctx, SPRemoteEventProperties properties, string eventType)
        {
            try
            {
                // set the correlation id for the next roundtrip to context
                ctx.TraceCorrelationId = properties.CorrelationId.ToString();

                var key = string.Concat(properties.CorrelationId.ToString("N"), eventType);
                var ts = TimeSpan.FromSeconds(20);

                return _cacheObject.Add(key, properties.CorrelationId, new CacheItemPolicy
                {
                    Priority = CacheItemPriority.Default,
                    SlidingExpiration = ts
                });
            }
            catch (Exception oops)
            {
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

            return false;
        }

    }
}
