using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionShowWhatsapp : GPluginAction
    {
        public override string ID => "690ADC37-E5FE-46EB-0001-E7D4E004823E";

        public override string Name => "Show Whatsapp Button";

        public override string Description => "";

        private List<GPluginActionParameter> _parameters;

        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _parameters;
            }
        }

        public GPluginActionShowWhatsapp(IGPlugin Plugin) : base(Plugin)
        {
            _parameters = new List<GPluginActionParameter>()
            {
                new GPluginActionParameter() {ID = 1, Name = "Whatsapp (Ex: 556140423650)", Description="",Required=true ,Type= PluginActionParameterTypeEnum.STRING },
                new GPluginActionParameter() {ID = 2, Name = "Mensagem", Description="",Required=true , Type= PluginActionParameterTypeEnum.STRING },
                new GPluginActionParameter() {ID = 3, Name = "Informação", Description="",Required=true , Type= PluginActionParameterTypeEnum.STRING }

            };
        }


        public override string GetActionCall()
        {
            return "WhatsappContact.RegisterButton(this, " + this.Form.Name + ", " + this.Parameters[0].Value + ", " + this.Parameters[1].Value + ", " + this.Parameters[2].Value + ");";
        }

        public override List<GPluginDependency> GetDependenciesFiles()
        {

            return new List<GPluginDependency>()
                {
                    new GPluginDependency() { FileName= "whatsappbutton.css", DestinationRelativePath="styles", AllowReplace= true },
                    new GPluginDependency() { FileName= "whatsappbutton.js", DestinationRelativePath="js", AllowReplace= true },
                    new GPluginDependency() { FileName= "WhatsappContact.cs", DestinationRelativePath="App_Code", AllowReplace= true }

                };
        }
    }
}
