using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginAKB : IGPlugin
    {
        public string ID => "6CACD40E-4F99-4D96-0000-C827AF67A9C2";
        public string Name => "AKB Integration";
        public string Description => "Integration with AKB Function";
        public string CompatibilityVersion => "2022";
        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionAKBR11301(this),
                    new GPluginActionAKBR3098(this),
                    new GPluginActionAKBR4557(this),
                    new GPluginActionAKBR4554(this),
                    new GPluginActionAKBR4559(this),
                    new GPluginActionAKBR21393(this),
                    new GPluginActionAKBR6662(this),
                    new GPluginActionAKBR4717(this),
                    new GPluginActionAKBR4555(this),
                    new GPluginActionAKBR241(this),
                    new GPluginActionAKBR21392(this),
                    new GPluginActionAKBR16268(this),
                };
            }
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>();
        }

    }
}