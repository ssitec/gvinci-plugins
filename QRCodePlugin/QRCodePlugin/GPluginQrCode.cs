using QRCodePlugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginQrCode : IGPlugin
    {
        public string ID => "FCCB254B-4920-472A-0000-BEB5226B3DF2";
    
        public string Name => "QrCode";

        public string Description => "";

        public string CompatibilityVersion => "2023";

        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionFixedGeneration(this),
                    new GPluginActionDynamicGeneration(this)
                };
            }

        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
                new GPluginDependency() { FileName = "QrCodeHelper.cs"  , DestinationRelativePath = "App_Code", AllowReplace = true },
                new GPluginDependency() { FileName = "QRCoder.dll"    , DestinationRelativePath = "bin", AllowReplace = true }
            };
        }
    }
}
