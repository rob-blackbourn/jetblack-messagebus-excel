using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ExcelDna.Integration.Rtd;

using Newtonsoft.Json;

using JetBlack.MessageBus.Adapters;

namespace Jetblack.MessageBus.ExcelAddin
{
    internal class CacheableClient : IDisposable
    {
        private readonly Client _client;
        private readonly IDictionary<SubscriptionKey, IList<ExcelRtdServer.Topic>> _subscriptions = new Dictionary<SubscriptionKey, IList<ExcelRtdServer.Topic>>();
        private readonly IDictionary<ExcelRtdServer.Topic, SubscriptionKey> _topics = new Dictionary<ExcelRtdServer.Topic, SubscriptionKey>();
        private readonly IDictionary<SubscriptionKey, IDictionary<string, IDictionary<string, object>>> _cache = new Dictionary<SubscriptionKey, IDictionary<string, IDictionary<string, object>>>();
        private readonly object _gate = new object();

        public CacheableClient(string host, int port)
        {
            _client = Client.SspiCreate(host, port);
            _client.OnDataReceived += OnDataReceived;
            _client.OnConnectionChanged += OnConnectionChanged;
            _client.OnHeartbeat += OnHeartbeat;
        }

        private void OnHeartbeat(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.Print("Heartbeat");
        }

        private void OnConnectionChanged(object sender, ConnectionChangedEventArgs e)
        {
            System.Diagnostics.Debug.Print($"OnConnectionChanged: {e.State}");
        }

        private void OnDataReceived(object sender, DataReceivedEventArgs e)
        {
            System.Diagnostics.Debug.Print($"OnDataReceived: {e.Feed} {e.Topic}");

            lock(_gate)
            {
                var key = new SubscriptionKey(e.Feed, e.Topic);
                if (!_subscriptions.TryGetValue(key, out var topics))
                    return;

                if (!_cache.TryGetValue(key, out var data))
                    _cache[key] = data = new Dictionary<string, IDictionary<string, object>>();

                foreach (var dataPacket in e.DataPackets.Where(x => x.Data != null))
                {
                    var json = Encoding.UTF8.GetString(dataPacket.Data);
                    var updates = JsonConvert.DeserializeObject<Dictionary<string, IDictionary<string, object>>>(json);
                    data.Update(updates);
                }

                foreach (var topic in topics)
                {
                    var updateCount = AddinFunctions.Cache.Set(topic.TopicId, data);
                    topic.UpdateValue($"{topic.TopicId}:{updateCount}");
                }
            }
        }

        public string RegisterTopic(ExcelRtdServer.Topic topic, string feed, string subject)
        {
            System.Diagnostics.Debug.Print($"Register Topic: Feed=\"{feed}\", topic=\"{subject}\"");

            lock(_gate)
            {
                var key = new SubscriptionKey(feed, subject);
                if (!_subscriptions.TryGetValue(key, out var topics))
                {
                    topics = new List<ExcelRtdServer.Topic>();
                    _subscriptions[key] = topics;
                    _topics[topic] = key;
                    _client.AddSubscription(feed, subject);
                }

                topics.Add(topic);

                return $"{topic.TopicId}:0";
            }
        }

        public void UnregisterTopic(ExcelRtdServer.Topic topic)
        {
            lock(_gate)
            {
                if (!_topics.TryGetValue(topic, out var key))
                    return;
                _topics.Remove(topic);

                if (!_subscriptions.TryGetValue(key, out var topics))
                    return;

                if (topics.Remove(topic) && topics.Count == 0)
                    _client.RemoveSubscription(key.Feed, key.Topic);
            }
        }

        public bool HasTopic(ExcelRtdServer.Topic topic)
        {
            return _topics.ContainsKey(topic);
        }

        public void Dispose()
        {
            _client.Dispose();
        }
    }
}
