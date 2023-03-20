using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.OpenAI
{
    public class OpenAIResponse
    {
        public string id { get; set; }
        // Target
        public string @object { get; set; }
        public int created { get; set; }
        //Engine
        public string model { get; set; }
        public List<Choice> choices { get; set; }
        public Usage usage { get; set; }
    }

    public class Choice
    {
        public string text { get; set; }
        public int index { get; set; }
        public object logprobs { get; set; }
        public string finish_reason { get; set; }
    }

    public class Usage
    {
        public int prompt_tokens { get; set; }
        public int completion_tokens { get; set; }
        public int total_tokens { get; set; }
    }
}
