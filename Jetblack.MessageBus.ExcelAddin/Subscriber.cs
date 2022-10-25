using System.Collections.Generic;
using System.Linq;
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

        private readonly IDictionary<string, CacheableClient> _clients = new Dictionary<string, CacheableClient>();
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

            var topicArgs = TopicArgs.Parse(topicInfo);

            if (topicArgs == null)
                return ExcelError.ExcelErrorValue;

            var client = GetClient(topicArgs.Endpoint);
            var token = client.RegisterTopic(topic, topicArgs.Feed, topicArgs.Topic);

            return token;
        }

        protected override void DisconnectData(Topic topic)
        {
            lock ( _gate )
            {
                var client = _clients.Values.FirstOrDefault(x => x.HasTopic(topic));
                client?.UnregisterTopic(topic);
            }
        }

        private CacheableClient GetClient(string value)
        {
            if (value == null || !EndPoint.TryParse(value, out var endpoint))
                endpoint = AddinConfig.DefaultEndPoint;

            lock (_gate)
            {
                var key = endpoint.ToString();
                if (!_clients.TryGetValue(key, out var client))
                    _clients[key] = client = new CacheableClient(endpoint);

                return client;
            }
        }

        private class TopicArgs
        {
            public string Feed;
            public string Topic;
            public string Endpoint;

            private TopicArgs(string feed, string topic, string endpoint)
            {
                Feed = feed;
                Topic = topic;
                Endpoint = endpoint;
            }

            public static TopicArgs Parse(IList<string> topicInfo)
            {
                if (string.IsNullOrWhiteSpace(topicInfo[0]) || string.IsNullOrWhiteSpace(topicInfo[1]))
                    return null;

                return new TopicArgs(topicInfo[0], topicInfo[1], topicInfo[2]);
            }
        }
    }
}
