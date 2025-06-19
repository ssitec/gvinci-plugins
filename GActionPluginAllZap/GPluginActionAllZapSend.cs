using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAllZapSend : GPluginAction
    {
        public override string ID => "00A5C13E-54B3-40D6-0001-E3E15C35451B";

        public override string Name => "Enviar Mensagem Via AllZap - Beta (Não habilitado)";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAllZapSend(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 1, Name = "Token", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 2, Name = "Nome da Sessão", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 3, Name = "Numero do Telefone", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 4, Name = "Mensagem", Type = PluginActionParameterTypeEnum.STRING },
                    };
        }

        public override string GetActionCall()
        {
            string token = this.Parameters[0].Value.ToString();
            string sessionName = this.Parameters[1].Value.ToString();
            string number = $"Amazon.RegionEndpoint.GetBySystemName({this.Parameters[2].Value.ToString()})";
            string text = this.Parameters[3].Value.ToString();
  
            string code = $"var client = sendMessage("+ this.Parameters[0].Value.ToString() + ", " + this.Parameters[1].Value.ToString() + ", " + this.Parameters[2].Value.ToString() + ", " + this.Parameters[3].Value.ToString() + ");";
            return code;
        }
    }
}