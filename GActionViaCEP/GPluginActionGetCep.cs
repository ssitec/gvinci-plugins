using System.Collections.Generic;
using System.Text;

namespace Gvinci.Plugin.Action
{
    public enum HowDisplay
    {
        show,
        hide
    }

    internal class GPluginActionGetCep : GPluginAction
    {

        public override string ID => "EE1BFD47-B14E-41F1-0001-08FAE6102548";

        public override string Name => "Capturar Endereço do CEP (UF COMBOBOX DB/Virtual)";

        public override string Description => "";

        private List<GPluginActionParameter> _parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _parameters;
            }
        }
        public GPluginActionGetCep(IGPlugin Plugin) : base(Plugin)
        {
            _parameters = new List<GPluginActionParameter>()
            {
                new GPluginActionParameter() {ID = 1, Name= "CEP", Description = "", Required = true, Type= PluginActionParameterTypeEnum.STRING},
                new GPluginActionParameter() {ID = 2, Name= "City Textbox", Description = "", Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 3, Name= "District Textbox", Description = "", Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 4, Name= "Address Textbox", Description = "", Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 5, Name= "Complement Textbox", Description = "", Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 6, Name= "State Combobox", Description = "", Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GCOMBOBOX", "GTEXTBOX" } }
            };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {
            IGPluginControl cityTextboxControl = this.Parameters[1].Value as IGPluginControl;
            IGPluginControl districtTextboxControl = this.Parameters[2].Value as IGPluginControl;
            IGPluginControl addressTextboxControl = this.Parameters[3].Value as IGPluginControl;
            IGPluginControl ComplementTextboxControl = this.Parameters[4].Value as IGPluginControl;
            IGPluginControl StateComponent = this.Parameters[5].Value as IGPluginControl;


            string identStr = new string('\t', Identation);
            Builder.AppendLine(identStr + string.Format("var response = Gvinci.Plugin.Viacep.ViacepHelper.BuscarCep({0});", this.Parameters[0].Value.ToString()));
            Builder.AppendLine(identStr + "if(response.cep != null)");
            Builder.AppendLine(identStr + "{");

            Identation += 1;
            identStr = new string('\t', Identation);

            Builder.AppendLine(identStr +  districtTextboxControl.Name + ".Text = response.bairro;");
            Builder.AppendLine(identStr +  cityTextboxControl.Name + ".Text = response.localidade;");
            Builder.AppendLine(identStr +  addressTextboxControl.Name + ".Text = response.logradouro;");
            Builder.AppendLine(identStr +  ComplementTextboxControl.Name + ".Text = response.complemento;");

            if (cityTextboxControl.ControlType == "GCOMBOBOX")
            {
                Builder.AppendLine(identStr + string.Format("SelectComboItem({0}, PageProvider.{0}Provider, response.uf);", StateComponent.Name));
            }
            else
            {
                Builder.AppendLine(identStr + StateComponent.Name + ".Text = response.uf;");
            }

            Identation -= 1;
            identStr = new string('\t', Identation);

            Builder.AppendLine(identStr + "}");
        }
    }
}
