using System;
using System.Configuration;
using System.Net;

namespace Jetblack.MessageBus.ExcelAddin
{
    public enum ClientScheme
    {
        Tcp,
        Ssl,
        Sspi
    }

    public struct EndPoint
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
    }

    public class AddinConfig
    {
        private static EndPoint? _defaultEndPoint;

        public static EndPoint DefaultEndPoint
        {
            get
            {
                if (!_defaultEndPoint.HasValue)
                {
                    var schemeAsString = ConfigurationManager.AppSettings["scheme"];
                    var scheme = !string.IsNullOrWhiteSpace(schemeAsString) && Enum.TryParse<ClientScheme>(schemeAsString, out var schemeAsEnum)
                        ? schemeAsEnum : ClientScheme.Tcp;

                    var host = ConfigurationManager.AppSettings["host"];
                    if (string.IsNullOrWhiteSpace(host))
                        host = Dns.GetHostEntry("LocalHost").HostName;


                    var portAsString = ConfigurationManager.AppSettings["port"];
                    var port = !string.IsNullOrWhiteSpace(portAsString) && int.TryParse(portAsString, out var portAsInt)
                        ? portAsInt : 9002;

                    _defaultEndPoint = new EndPoint(scheme, host, port);
                }

                return _defaultEndPoint.Value;
            }
        }

        public static EndPoint MakeEndPoint(string endpoint)
        {
            if (!string.IsNullOrWhiteSpace(endpoint)
                && Uri.TryCreate(endpoint, UriKind.Absolute, out var uri)
                && Enum.TryParse<ClientScheme>(uri.Scheme, true, out var scheme))
            {
                return new EndPoint(scheme, uri.Host, uri.Port);
            }

            return DefaultEndPoint;
        }
    }
}
