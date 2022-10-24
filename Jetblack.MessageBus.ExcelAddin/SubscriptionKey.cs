using System;

namespace Jetblack.MessageBus.ExcelAddin
{
    internal class SubscriptionKey : IEquatable<SubscriptionKey>
    {
        public readonly string Feed;
        public readonly string Topic;

        public SubscriptionKey(string feed, string topic)
        {
            Feed = feed;
            Topic = topic;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as SubscriptionKey);
        }

        public bool Equals(SubscriptionKey other)
        {
            return
                other != null
                && string.Equals(Feed, other.Feed, StringComparison.InvariantCultureIgnoreCase)
                && string.Equals(Topic, other.Topic, StringComparison.InvariantCultureIgnoreCase);
        }

        public override int GetHashCode()
        {
            return Feed .GetHashCode() ^ Topic.GetHashCode();
        }

        public override string ToString() => $"{Feed}:{Topic}";
    }
}
