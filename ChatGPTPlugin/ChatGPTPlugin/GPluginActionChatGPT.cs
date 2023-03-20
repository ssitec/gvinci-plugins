using Gvinci.Plugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public enum HowDisplay
    {
        show,
        hide
    }
    public class GPluginActionChatGPT : GPluginAction
    {
        public override string ID => "F2A6241E-24E8-41EE-0001-740EBA94C988";

        public override string Name => "ChatGPT";

        public override string Description => "";

        private List<GPluginActionParameter> _parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _parameters;
            }
        }

        public GPluginActionChatGPT(IGPlugin Plugin) : base(Plugin)
        {
            _parameters = new List<GPluginActionParameter>()
            {
                new GPluginActionParameter() {ID = 1, Name= "Api Key",              Description = "", Required = true, Type= PluginActionParameterTypeEnum.STRING },
                new GPluginActionParameter() {ID = 2, Name= "Tokens",               Description = "", Required = true, Type= PluginActionParameterTypeEnum.INTEGER },
                new GPluginActionParameter() {ID = 3, Name= "Input",                Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 4, Name= "Engine",               Description = "", Required = true, Type= PluginActionParameterTypeEnum.STRING },
                new GPluginActionParameter() {ID = 5, Name= "Temperature",          Description = "", Required = true, Type = PluginActionParameterTypeEnum.DECIMAL },
                new GPluginActionParameter() {ID = 6, Name= "TopP",                 Description = "", Required = true, Type= PluginActionParameterTypeEnum.INTEGER },
                new GPluginActionParameter() {ID = 7, Name= "Frequency Penalty",    Description = "", Required = true, Type = PluginActionParameterTypeEnum.DECIMAL },
                new GPluginActionParameter() {ID = 8, Name= "Presence Penalty",     Description = "", Required = true, Type = PluginActionParameterTypeEnum.DECIMAL },
                new GPluginActionParameter() {ID = 9, Name= "Response",             Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GLABEL", "GTEXTBOX" } }
            };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {
            var apiToken             = this.Parameters[0].Value.ToString(); 
            var token                = this.Parameters[1].Value.ToString();
            IGPluginControl input    = this.Parameters[2].Value as IGPluginControl;
            var engine               = this.Parameters[3].Value.ToString(); 
            var temperature          = this.Parameters[4].Value.ToString(); 
            var topP                 = this.Parameters[5].Value.ToString(); 
            var frequency            = this.Parameters[6].Value.ToString();
            var presence             = this.Parameters[7].Value.ToString();
            IGPluginControl response = this.Parameters[8].Value as IGPluginControl;


            string identStr = new string('\t', Identation);
            Builder.AppendLine(identStr + $"{response.Name}.Text = Gvinci.Plugin.OpenAI.OpenAIHelper.CallOpenAIChatGPT({apiToken},{token},{input.Name}.Text,{engine},{temperature},{topP},{frequency},{presence});");
        }
    }
}
