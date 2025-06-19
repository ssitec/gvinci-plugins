using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginViaCep : IGPlugin
    {
        public string ID => "EE1BFD47-B14E-41F1-0000-08FAE6102548";

        public string Name => "ViaCep";

        public string Description => "";

        public string CompatibilityVersion => "2023";

        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionGetCep(this)
                };
            }
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                new GPluginDependency() { FileName = "ViacepModel.cs", DestinationRelativePath = "App_Code", AllowReplace = true },
                new GPluginDependency() { FileName = "ViacepHelper.cs", DestinationRelativePath = "App_Code", AllowReplace = true },
                new GPluginDependency() { FileName = "RestSharp.dll" , DestinationRelativePath = "bin", AllowReplace = true }
            };
        }
    }
}
