using Google.Apis.Calendar.v3;
using Google.Apis.Http;
using Google.Apis.PeopleService.v1;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Util;
using Serilog;
using System;
using System.Net;
using System.Threading.Tasks;

namespace GoContactSyncMod
{
    internal class GoogleServices
    {
        // TODO(obelix30): it is workaround for open issue here: https://github.com/google/google-api-dotnet-client/issues/608
        /// <summary>
        /// Creates CalendarService with custom back-off handler.
        /// </summary>
        /// <param name="initializer">The service initializer.</param>
        /// <returns>Created CalendarService with custom back-off configured.</returns>
        public static CalendarService CreateCalendarService(BaseClientService.Initializer initializer)
        {
            CalendarService service = null;
            try
            {
                initializer.DefaultExponentialBackOffPolicy = ExponentialBackOffPolicy.None;
                service = new CalendarService(initializer);
                var backOffHandler = new BackoffHandler(service, new ExponentialBackOff());
                service.HttpClient.MessageHandler.AddUnsuccessfulResponseHandler(backOffHandler);
                return service;
            }
            catch
            {
                if (service != null)
                {
                    service.Dispose();
                }

                throw;
            }
        }

        public static PeopleServiceService CreatePeopleService(BaseClientService.Initializer initializer)
        {
            PeopleServiceService service = null;
            try
            {
                initializer.DefaultExponentialBackOffPolicy = ExponentialBackOffPolicy.None;
                service = new PeopleServiceService(initializer);
                var backOffHandler = new BackoffHandler(service, new ExponentialBackOff());
                service.HttpClient.MessageHandler.AddUnsuccessfulResponseHandler(backOffHandler);
                return service;
            }
            catch
            {
                if (service != null)
                {
                    service.Dispose();
                }

                throw;
            }
        }

        /// <summary>
        /// Check if error returned from Google API is transient, i.e. could be retried.
        /// </summary>
        /// <param name="statusCode">Status code returned from API (e.g. 403, 500).</param>        
        /// <param name="reqError">Server error.</param>
        /// <returns>If error is transient.</returns>
        public static bool IsTransientError(HttpStatusCode statusCode, RequestError reqError)
        {
            if ((int)statusCode >= (int)HttpStatusCode.InternalServerError)
            {
                return true;
            }

            if (statusCode == HttpStatusCode.Forbidden)
            {
                if (reqError.Errors[0].Reason == "rateLimitExceeded")
                {
                    return true;
                }

                if (reqError.Errors[0].Reason == "userRateLimitExceeded")
                {
                    return true;
                }

                if (reqError.Errors[0].Reason == "dailyLimitExceeded")
                {
                    return true;
                }

                if (reqError.Errors[0].Reason == "quotaExceeded")
                {
                    return true;
                }
            }

            return false;
        }
        public const int BatchRequestSize = 50;
        public const int BatchRequestBackoffDelay = 1000;

        private static readonly Random random = new Random();

        public static TimeSpan GetExpotentialBackoffDelay(int currentRetry)
        {
            var v = (Math.Pow(2.0, (double)currentRetry - 1) * 1000) + random.Next(0, 250);
            return TimeSpan.FromMilliseconds(v);
        }
        /// <summary>
        /// Specific back-off implementation.
        /// </summary>
        public class BackoffHandler : IHttpUnsuccessfulResponseHandler
        {
            private readonly IBackOff _backoff;
            private readonly IClientService _service;
            private readonly TimeSpan _maxTimeSpan;

            /// <summary>
            /// Constructs a new custom back-off handler
            /// </summary>
            /// <param name="service">Client service</param>
            /// <param name="backoff">Back-off strategy (e.g. exponential backoff)</param>
            public BackoffHandler(IClientService service, IBackOff backoff)
            {
                _service = service;
                _backoff = backoff;
                _maxTimeSpan = TimeSpan.FromHours(1);
            }

            /// <summary>
            /// Handler for back-off, if retry is not possible it returns <c>false</c>. Otherwise it blocks for some time (miliseconds)
            /// and returns <c>true</c>, so call is retried.
            /// </summary>
            /// <param name="args">The arguments object to handler call.</param>
            /// <returns>If request could be retried.</returns>
            public async Task<bool> HandleResponseAsync(HandleUnsuccessfulResponseArgs args)
            {
                if (!args.SupportsRetry || _backoff.MaxNumOfRetries < args.CurrentFailedTry)
                {
                    return false;
                }

                if (IsTransientError(args.Response.StatusCode, _service.DeserializeError(args.Response).Result))
                {
                    var delay = _backoff.GetNextBackOff(args.CurrentFailedTry);
                    if (delay > _maxTimeSpan || delay < TimeSpan.Zero)
                    {
                        return false;
                    }

                    await Task.Delay(delay, args.CancellationToken);
                    Log.Debug("Back-Off waited " + delay.TotalMilliseconds + "ms before next retry...");

                    return true;
                }

                return false;
            }
        }
    }
}
