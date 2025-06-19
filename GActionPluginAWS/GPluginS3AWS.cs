using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginS3AWS : IGPlugin
    {
        public string ID => "7A69D6A0-CF98-4063-0000-6F7E02EC5F91";
        public string Name => "AWS S3 Integration";
        public string Description => "Integration with AWS S3";
        public string CompatibilityVersion => "2022";
        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionS3AWSCreate(this),
                    new GPluginActionS3AWSUpload(this),
                };
            }
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                new GPluginDependency() { FileName = "AWSSDK.Core.dll", DestinationRelativePath = "bin", AllowReplace = true },
                new GPluginDependency() { FileName = "AWSSDK.S3.dll", DestinationRelativePath = "bin", AllowReplace = true },
                new GPluginDependency() { FileName = "S3Helper.cs", DestinationRelativePath = "App_Code", AllowReplace = true }
            };
        }

    }
}