using Gvinci.Plugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionDallE : GPluginAction
    {
        public override string ID => "F2A6241E-24E8-41EE-0002-740EBA94C988";

        public override string Name => "DallE (Gerador de Imagens)";

        public override string Description => "";

        private List<GPluginActionParameter> _parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _parameters;
            }
        }

        public GPluginActionDallE(IGPlugin Plugin) : base(Plugin)
        {
            _parameters = new List<GPluginActionParameter>()
            {             
                new GPluginActionParameter() {ID = 1, Name= "Api Key",              Description = "", Required = true, Type= PluginActionParameterTypeEnum.STRING },
                new GPluginActionParameter() {ID = 2, Name= "Input",                Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 3, Name= "Number",               Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 4, Name= "Size (n x n)",         Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GTEXTBOX" } },
                new GPluginActionParameter() {ID = 5, Name= "Response",             Description = "", Required = true, Type= PluginActionParameterTypeEnum.CONTROL, AllowedControlTypes = new string[] { "GLAYOUTCOL", "GCOMBOBOX" } },
            };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {
            var apiToken                = this.Parameters[0].Value.ToString(); 
            IGPluginControl input       = this.Parameters[1].Value as IGPluginControl;
            IGPluginControl number      = this.Parameters[2].Value as IGPluginControl;
            IGPluginControl size        = this.Parameters[3].Value as IGPluginControl;
            IGPluginControl response    = this.Parameters[4].Value as IGPluginControl;

            string identStr = new string('\t', Identation);

            if (size.ControlType == "GCOMBOBOX")
            {
                Builder.AppendLine(identStr + $"var sizes = {size.Name}.SelectedValue.Split('x', 'X');");
            }
            else
            {
                Builder.AppendLine(identStr + $"var sizes = {size.Name}.Text.Split('x', 'X');");
            }


            Builder.AppendLine(identStr + $"var dalleResponse = Gvinci.Plugin.OpenAI.OpenAIHelper.CallOpenAiDallE({apiToken},{input.Name}.Text,Convert.ToInt32({number.Name}.Text),{size.Name}.Text);");
            Builder.AppendLine(identStr + "for (int i = 0; i < dalleResponse.data.Count; i++)");
            Builder.AppendLine(identStr + "{");
            Builder.AppendLine(identStr + "\t HtmlImage img = new HtmlImage();");
            Builder.AppendLine(identStr + "\t img.Src = dalleResponse.data[i].url.ToString();");
            Builder.AppendLine(identStr + "\t img.Height = Convert.ToInt32(sizes[0]);");
            Builder.AppendLine(identStr + "\t img.Width  = Convert.ToInt32(sizes[1]);");
            Builder.AppendLine(identStr + $"\t {response.Name}.Controls.Add(img);");
            Builder.AppendLine(identStr + "}");
        }
    }
}
