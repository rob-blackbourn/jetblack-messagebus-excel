using System.Configuration;

namespace Jetblack.MessageBus.ExcelAddin
{
    public class AddinConfig
    {
        private static EndPoint _defaultEndPoint = null;

        public static EndPoint DefaultEndPoint
        {
            get
            {
                if (_defaultEndPoint == null)
                    _defaultEndPoint = MakeDefaultEndpoint();

                return _defaultEndPoint;
            }
        }

        private static EndPoint MakeDefaultEndpoint()
        {
            var endpointAsString = ConfigurationManager.AppSettings["endpoint"];
            return endpointAsString != null && EndPoint.TryParse(endpointAsString, out var endpoint)
                ? endpoint
                : new EndPoint(ClientScheme.Tcp, "LocalHost", 9001);
        }
    }
}
