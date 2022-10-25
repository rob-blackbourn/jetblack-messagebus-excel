using System.Collections.Generic;
using System.Linq;

namespace Jetblack.MessageBus.ExcelAddin
{
    public class TopicCache
    {
        private readonly IDictionary<int, CachedItem> _cache = new Dictionary<int, CachedItem>();
        private readonly object _gate = new object();

        public IDictionary<string, IDictionary<string, object>> Get(int topicId)
        {
            lock (_gate)
            {
                if (!_cache.TryGetValue(topicId, out var cachedItem))
                    _cache.Add(
                        topicId,
                        cachedItem = new CachedItem(new Dictionary<string, IDictionary<string, object>>()));

                return cachedItem.Data;
            }
        }

        public int Set(int topicId, IDictionary<string, IDictionary<string, object>> data)
        {
            lock (_gate)
            {
                if (!_cache.TryGetValue(topicId, out var cachedItem))
                    _cache.Add(topicId, cachedItem = new CachedItem(data));
                else
                    cachedItem.Update(data);

                return cachedItem.UpdateCount;
            }
        }

        public void Clear(int topicId)
        {
            lock (_gate)
            {
                _cache.Remove(topicId);
            }
        }

        public IDictionary<string, IDictionary<string, object>> Get(string token)
        {
            if (TryParseTopicId(token, out var topicId))
                return Get(topicId);

            return null;
        }

        private static bool TryParseTopicId(string token, out int topicId)
        {
            if (token != null)
            {
                var index = token.IndexOf(':');
                if (index != -1 && int.TryParse(token.Substring(0, index), out topicId))
                    return true;
            }

            topicId = 0;
            return false;
        }

        private class CachedItem
        {
            public IDictionary<string, IDictionary<string, object>> Data { get; private set; }
            public int UpdateCount { get; private set; } = 0;

            public CachedItem(IDictionary<string, IDictionary<string, object>> data)
            {
                Data = data;
            }

            public int Update(IDictionary<string, IDictionary<string, object>> data)
            {
                Data = data;
                return ++UpdateCount;
            }
        }
    }
}
