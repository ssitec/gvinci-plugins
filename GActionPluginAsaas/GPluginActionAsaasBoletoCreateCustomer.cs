using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAsaasBoletoCreateCustomer : GPluginAction
    {
        public override string ID => "1EB7342C-C15D-4FC6-0002-2A8702EF06AA";

        public override string Name => "Criar Boleto Criando Usuario";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAsaasBoletoCreateCustomer(IGPlugin Plugin) : base(Plugin)
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
                        new GPluginActionParameter() { ID = 8, Name = "NumeroEndereco", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 9, Name = "ClienteRef", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 10, Name = "P|S - (Produçao, Sandbox)", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 11, Name = "Vencimento", Type = PluginActionParameterTypeEnum.DATETIME },
                        new GPluginActionParameter() { ID = 12, Name = "Valor", Type = PluginActionParameterTypeEnum.DECIMAL },
                        new GPluginActionParameter() { ID = 13, Name = "Descricao", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 14, Name = "DocRef", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 15, Name = "Controle de Retorno", Type = PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX", "GCOMBOBOX" } }
                    };
        }


        public override string GetActionCall()
        {
            string code = "";
            IGPluginControl controlReturn = this.Parameters[14].Value as IGPluginControl;
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
            code += $" = GvinciAsaas.Assas_CreateBoleto({this.Parameters[0].Value}, {this.Parameters[1].Value}, {this.Parameters[2].Value}, {this.Parameters[3].Value}, {this.Parameters[4].Value}, {this.Parameters[5].Value}, {this.Parameters[6].Value}, {this.Parameters[7].Value}, {this.Parameters[8].Value}, {this.Parameters[9].Value}, {this.Parameters[10].Value}, {this.Parameters[11].Value}, {this.Parameters[12].Value}, {this.Parameters[13].Value});";
            return code;
        }
    }
}