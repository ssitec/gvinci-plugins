using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginWhatsapp : IGPlugin
    {
        public string ID => "690ADC37-E5FE-46EB-0000-E7D4E004823E";

        public string Name => "Whatsapp Contact";

        public string Description => "";

        public string CompatibilityVersion => "2022";

        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionShowWhatsapp(this)
                };
            }
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>();
        }
    }
}
