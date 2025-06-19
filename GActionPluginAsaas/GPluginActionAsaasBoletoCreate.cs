using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionAsaasBoletoCreate : GPluginAction
    {
        public override string ID => "801CB33E-C5CF-420E-8271-FD46EFF85D06";
        
        public override string Name => "Criar Boleto com Asaas Customer ID";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionAsaasBoletoCreate(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 1, Name = "Token", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 2, Name = "Asaas Customer ID", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 3, Name = "P|S - (Produçao, Sandbox)", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 4, Name = "Vencimento", Type = PluginActionParameterTypeEnum.DATETIME },
                        new GPluginActionParameter() { ID = 5, Name = "Valor", Type = PluginActionParameterTypeEnum.DECIMAL },
                        new GPluginActionParameter() { ID = 6, Name = "Descricao", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 7, Name = "DocRef", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 8, Name= "Control de Retorno", Description = "", Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GCOMBOBOX", "GTEXTBOX" } }
                    };
        }


        public override string GetActionCall()
        {
            IGPluginControl controlReturn = this.Parameters[7].Value as IGPluginControl;

            string code = "";
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
            code += $" = GvinciAsaas.Assas_CreateBoleto({this.Parameters[1].Value}, {this.Parameters[2].Value}, {this.Parameters[0].Value}, {this.Parameters[3].Value}, {this.Parameters[4].Value}, {this.Parameters[5].Value}, {this.Parameters[6].Value} );";
            return code;
        }
    }
}