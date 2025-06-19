using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginAllZap : IGPlugin
    {
        public string ID => "00A5C13E-54B3-40D6-0000-E3E15C35451B";
        public string Name => "AllZapp Integration";
        public string Description => "Send message by whatsapp";
        public string CompatibilityVersion => "2022";
        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionAllZapSend(this),
                };
            }
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
               // new GPluginDependency() { FileName = "AWSSDK.Core.dll", DestinationRelativePath = "bin", AllowReplace = true },
               // new GPluginDependency() { FileName = "AWSSDK.S3.dll", DestinationRelativePath = "bin", AllowReplace = true },
               new GPluginDependency() { FileName = "AllZapHelper.cs", DestinationRelativePath = "App_Code", AllowReplace = true }
            };
        }

 



    }
}