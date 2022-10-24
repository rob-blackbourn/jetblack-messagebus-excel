﻿using System.Collections.Generic;
using System.Runtime.InteropServices;

using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace Jetblack.MessageBus.ExcelAddin
{
    [ComVisible(true)]
    [ProgId(Subscriber.ServerProgId)]
    public class Subscriber : ExcelRtdServer
    {
        public const string ServerProgId = "MessageBus.Subscriber";

        private readonly IDictionary<ClientKey, CacheableClient> _clients = new Dictionary<ClientKey, CacheableClient>();
        private readonly object _gate = new object();

        protected override bool ServerStart()
        {
            return true;
        }

        protected override void ServerTerminate()
        {
            lock (_gate)
            {
                foreach (var client in _clients.Values)
                    client.Dispose();

                _clients.Clear();
            }
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            if (topicInfo.Count != 3)
                return ExcelError.ExcelErrorValue;

            var feed = topicInfo[0];
            var subject = topicInfo[1];

            if (string.IsNullOrWhiteSpace(feed) || string.IsNullOrWhiteSpace(subject))
                return ExcelError.ExcelErrorValue;

            var endpoint = AddinConfig.MakeEndPoint(topicInfo[2]);

            var client = GetClient(endpoint);
            var token = client.RegisterTopic(topic, feed, subject);

            return token;
        }

        protected override void DisconnectData(Topic topic)
        {
            lock ( _gate )
            {
                foreach (var client in _clients.Values)
                {
                    if (client.HasTopic(topic))
                    {
                        client.UnregisterTopic(topic);
                        break;
                    }
                }
            }
        }

        private CacheableClient GetClient(EndPoint endPoint)
        {
            lock (_gate)
            {
                var key = new ClientKey(endPoint.Host, endPoint.Port);
                if (!_clients.TryGetValue(key, out var client))
                    _clients[key] = client = new CacheableClient(endPoint.Host, endPoint.Port);

                return client;
            }
        }
    }
}