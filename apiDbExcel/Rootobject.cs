
using System.Collections.Generic;


namespace ApiToDatabase
{
    class Rootobject
    {
        public List<Datum> data { get; set; }
        public int result_count { get; set; }
        public int page_size { get; set; }
        public int current_page { get; set; }
        public int total_pages { get; set; }
        public int api_call_credits { get; set; }
    }

    public class Datum
    {
        public string date { get; set; }
        public float open { get; set; }
        public float high { get; set; }
        public float low { get; set; }
        public float close { get; set; }
        public float volume { get; set; }
        public float ex_dividend { get; set; }
        public float split_ratio { get; set; }
        public float adj_open { get; set; }
        public float adj_high { get; set; }
        public float adj_low { get; set; }
        public float adj_close { get; set; }
        public float adj_volume { get; set; }
    }
}
