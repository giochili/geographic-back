using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace BotReestriClassLibrary.Wrapper
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
