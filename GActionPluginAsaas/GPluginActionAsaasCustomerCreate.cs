using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAsaasCustomerCreate : GPluginAction
    {
        public override string ID => "1EB7342C-C15D-4FC6-0001-2A8702EF06AA";

        public override string Name => "Cadastrar Cliente";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAsaasCustomerCreate(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 1, Name = "Token", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 2, Name = "Nome", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 3, Name = "CpfCnpj", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 4, Name = "Email", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 5, Name = "Fone", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 6, Name = "MobileFone", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 7, Name = "CEP", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 8, Name = "NumeroDoEndereco", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 9, Name = "ExternalReference", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 10, Name = "P|S - (Produçao, Sandbox)", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 11, Name = "ReturnControl", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX" } },
                    };
        }


        public override string GetActionCall()
        {
            string code = "";
            IGPluginControl controlReturn = this.Parameters[10].Value as IGPluginControl;

            if (controlReturn?.ControlType == "GCOMBOBOX")
            {
                code += controlReturn.Name + ".Value";
            }
            else if (controlReturn?.ControlType == "GTEXTBOX")
            {
                code += controlReturn.Name + ".Text";
            }
            else
            {
                code = "var result";
            }
            code += $" = GvinciAsaas.Assas_CustomerCreateOrUpdate({this.Parameters[0].Value}, {this.Parameters[1].Value}, {this.Parameters[2].Value}, {this.Parameters[3].Value}, {this.Parameters[4].Value}, {this.Parameters[5].Value}, {this.Parameters[6].Value}, {this.Parameters[7].Value}, {this.Parameters[8].Value}, {this.Parameters[9].Value});";
            return code;
        }
    }
}