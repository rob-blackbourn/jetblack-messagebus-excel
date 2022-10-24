using System.Configuration;
using System.Net;

namespace Jetblack.MessageBus.ExcelAddin
{
    public struct EndPoint
    {
        public readonly string Host;
        public readonly int Port;

        public EndPoint(string host, int port)
        {
            Host = host;
            Port = port;
        }
    }

    public class AddinConfig
    {
        private static EndPoint? _defaultEndPoint;

        public static EndPoint DefaultEndPoint
        {
            get
            {
                if (_defaultEndPoint.HasValue)
                {
                    var host = ConfigurationManager.AppSettings["host"];
                    if (string.IsNullOrWhiteSpace(host))
                    {
                        host = Dns.GetHostEntry("LocalHost").HostName;
                    }

                    var portAsString = ConfigurationManager.AppSettings["port"];
                    var port = !string.IsNullOrWhiteSpace(portAsString) && int.TryParse(portAsString, out var portAsInt)
                        ? portAsInt : 0;

                    _defaultEndPoint = new EndPoint(host, port);
                }

                return _defaultEndPoint.Value;
            }
        }

        public static EndPoint MakeEndPoint(string endpoint)
        {
            if (!string.IsNullOrWhiteSpace(endpoint))
            {
                var parts = endpoint.Split(':');
                if (parts.Length == 2 && int.TryParse(parts[1], out var port))
                {
                    return new EndPoint(parts[0], port);
                }
            }

            return DefaultEndPoint;
        }
    }
}
