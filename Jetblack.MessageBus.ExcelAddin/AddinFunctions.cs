using System.Linq;

using ExcelDna.Integration;

namespace Jetblack.MessageBus.ExcelAddin
{
    public class AddinFunctions
    {
        public static readonly TopicCache Cache = new TopicCache();

        [ExcelFunction(Name = "rtmbSUBSCRIBE", Description = "Subscribe to message bus data", Category = "Message Bus")]
        public static object Subscribe(
            [ExcelArgument(Description = "The feed name")] string feed,
            [ExcelArgument(Description = "The topic name")] string topic,
            [ExcelArgument(Description = "The columns to return or all if omitted")] object[] columns,
            [ExcelArgument(Description = "The rows to return or all if omitted")] object[] rows,
            [ExcelArgument(Description = "Show column headers in the returned table")] object showColHeaders,
            [ExcelArgument(Description = "Show row headers in the returned table")] object showRowHeaders,
            [ExcelArgument(Description = "The server endpoint (e.g. \"example.com:9002\"")] object endpoint)
        {
            var token = XlCall.RTD(
                Subscriber.ServerProgId,
                null,
                feed,
                topic,
                endpoint.Optional(string.Empty));

            var dataFrame = Cache.Get(token as string);
            if (dataFrame == null)
                return ExcelError.ExcelErrorValue;

            return dataFrame.ToTable(
                columns,
                rows,
                showColHeaders.Check(false),
                showRowHeaders.Check(false));
        }
    }
}
