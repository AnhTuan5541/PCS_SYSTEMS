namespace PCS_SYSTEMS.Response
{
    public class CommonResponse<T>
    {
        public string StatusCode { get; set; }
        public string Message { get; set; }
        public IEnumerable<T> Data { get; set; }
        public int size { get; set; }
    }
}
