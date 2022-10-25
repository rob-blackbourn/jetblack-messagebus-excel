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

    public class AddinConfig
    {
        private static EndPoint _defaultEndPoint = null;

        public static EndPoint DefaultEndPoint
        {
            get
            {
                if (_defaultEndPoint == null)
                {
                    var endpointAsString = ConfigurationManager.AppSettings["endpoint"];
                    _defaultEndPoint = endpointAsString != null && EndPoint.TryParse(endpointAsString, out var endpoint)
                        ? endpoint
                        : new EndPoint(ClientScheme.Tcp, "LocalHost", 9001);
                }

                return _defaultEndPoint;
            }
        }

        public static EndPoint MakeEndPoint(string value)
        {
            return value != null && EndPoint.TryParse(value, out var endpoint)
                ? endpoint
                : DefaultEndPoint;
        }
    }
}
