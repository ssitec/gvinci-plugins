using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.OpenAI
{
    public class input
    {
        public string prompt { get; set; }
        public short n { get; set; }
        public string size { get; set; }
    }

    // model for the image url
    public class Link
    {
        public string url { get; set; }
    }

    // model for the DALL E api response
    public class ResponseModel
    {
        public long created { get; set; }
        public List<Link> data { get; set; }
    }
}
