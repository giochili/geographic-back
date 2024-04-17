using System.Net;

namespace GeographicDynamicWebAPI.Wrappers
{
    public class Result<T>
    {
        public bool Success { get; set; }
        public T? Value { get; set; }
        public HttpStatusCode StatusCode { get; set; }
        public string? Message { get; set; }
        public List<T> Data { get; set; }

    }
}

