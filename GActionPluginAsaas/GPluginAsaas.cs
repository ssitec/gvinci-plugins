using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginAsaas : IGPlugin
    {
        public string ID => "1EB7342C-C15D-4FC6-0000-2A8702EF06AA";
        public string Name => "Asaas Integration";
        public string Description => "Criar Cobranças (boleto e cartão) (Beta)";
        public string CompatibilityVersion => "2022";
        public string ProjectType => "CSHARP";

        public List<GPluginAction> Actions
        {
            get
            {
                return new List<GPluginAction>()
                {
                    new GPluginActionAsaasCustomerCreate(this),
                    new GPluginActionAsaasBoletoCreate(this),
                    new GPluginActionAsaasBoletoCreateCustomer(this)
                };
            }
        }

        public List<GPluginDependency> GetDependenciesFiles()
        {
            return new List<GPluginDependency>()
            {
               new GPluginDependency() { FileName = "AsaasHelper.cs", DestinationRelativePath = "App_Code", AllowReplace = true },
               new GPluginDependency() { FileName = "AsaasModel.cs", DestinationRelativePath = "App_Code", AllowReplace = true },
               new GPluginDependency() { FileName = "RestSharp.dll", DestinationRelativePath = "bin", AllowReplace = true }
            };
        }

 



    }
}