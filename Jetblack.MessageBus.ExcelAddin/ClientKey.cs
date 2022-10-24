using System;

namespace Jetblack.MessageBus.ExcelAddin
{
    internal class ClientKey : IEquatable<ClientKey>
    {
        public readonly string Feed;
        public readonly int Port;

        public ClientKey(string feed, int port)
        {
            Feed = feed;
            Port = port;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as ClientKey);
        }

        public bool Equals(ClientKey other)
        {
            return
                other != null
                && string.Equals(Feed, other.Feed, StringComparison.InvariantCultureIgnoreCase)
                && Port == other.Port;
        }

        public override int GetHashCode()
        {
            return Feed .GetHashCode() ^ Port.GetHashCode();
        }

        public override string ToString() => $"{Feed}:{Port}";
    }
}
