using Gvinci.Plugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginOpenAI : IGPlugin
    {
        public string ID => "F2A6241E-24E8-41EE-0000-740EBA94C988";

        public string Name => "OpenAI";

        public string Description => "";

        public string CompatibilityVersion => "2023";

        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionChatGPT(this),
                    new GPluginActionDallE(this)
                };
            }
            
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                new GPluginDependency() { FileName = "ChatGPTModel.cs"  , DestinationRelativePath = "App_Code", AllowReplace = true },
                new GPluginDependency() { FileName = "DallEModel.cs"    , DestinationRelativePath = "App_Code", AllowReplace = true },
                new GPluginDependency() { FileName = "OpenAIHelper.cs"  , DestinationRelativePath = "App_Code", AllowReplace = true },
                new GPluginDependency() { FileName = "RestSharp.dll"    , DestinationRelativePath = "bin", AllowReplace = true }
            };
        }
    }
}
