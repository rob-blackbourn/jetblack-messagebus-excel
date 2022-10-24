using System.Collections.Generic;

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
                {
                    cachedItem = new CachedItem(new Dictionary<string, IDictionary<string, object>>());
                    _cache.Add(topicId, cachedItem);
                }

                return cachedItem.Data;
            }
        }

        public int Set(int topicId, IDictionary<string, IDictionary<string, object>> data)
        {
            lock (_gate)
            {
                if (!_cache.TryGetValue(topicId, out var cachedItem))
                {
                    cachedItem = new CachedItem(data);
                    _cache.Add(topicId, cachedItem);
                }
                else
                {
                    cachedItem.Update(data);
                }

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
