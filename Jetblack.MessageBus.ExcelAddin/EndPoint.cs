using System;

namespace Jetblack.MessageBus.ExcelAddin
{
    internal class EndPoint
    {
        public readonly ClientScheme Scheme;
        public readonly string Host;
        public readonly int Port;

        public EndPoint(ClientScheme scheme, string host, int port)
        {
            Scheme = scheme;
            Host = host;
            Port = port;
        }

        public static bool TryParse(string value, out EndPoint endpoint)
        {
            if (Uri.TryCreate(value, UriKind.Absolute, out var uri)
                && Enum.TryParse<ClientScheme>(uri.Scheme, true, out var scheme))
            {
                endpoint = new EndPoint(scheme, uri.Host, uri.Port);
                return true;
            }

            endpoint = null;
            return false;
        }

        public override string ToString() => $"{Scheme}://{Host}:{Port}";
    }
}
