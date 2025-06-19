using System.Collections.Generic;
using System.Text;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionS3AWSUpload : GPluginAction
    {
        public override string ID => "7A69D6A0-CF98-4063-0002-6F7E02EC5F91";

        public override string Name => "Update File to S3";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionS3AWSUpload(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 1, Name = "S3AccessKey", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 2, Name = "S3SecretKey", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 3, Name = "Region", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 4, Name = "BucketName", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 5, Name = "Path", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 6, Name = "FileContent", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 7, Name = "IsBase64Content", Type = PluginActionParameterTypeEnum.BOOL },
                        new GPluginActionParameter() { ID = 8, Name = "IsPublic", Type = PluginActionParameterTypeEnum.BOOL },
                    };
        }

        public override void WriteActionCall(StringBuilder Builder, int Identation, int ActionSequence)
        {
            string idenStr = new string('\t', Identation);
            Builder.AppendLine($"{idenStr}string accessKey = {this.Parameters[0].Value};");
            Builder.AppendLine($"{idenStr}string secretKey = {this.Parameters[1].Value};");
            Builder.AppendLine($"{idenStr}Amazon.RegionEndpoint region = Amazon.RegionEndpoint.GetBySystemName({this.Parameters[2].Value});");
            Builder.AppendLine($"{idenStr}string bucketName = {this.Parameters[3].Value};");
            Builder.AppendLine($"{idenStr}string key = {this.Parameters[4].Value};");
            Builder.AppendLine($"{idenStr}string contentType = MimeMapping.GetMimeMapping(key);");
            Builder.AppendLine($"{idenStr}byte[] content;");
            Builder.AppendLine($"{idenStr}if ({this.Parameters[6].Value})");
            Builder.AppendLine($"{idenStr}{{");
            Builder.AppendLine($"{idenStr}\tcontent = Convert.FromBase64String({this.Parameters[5].Value});");
            Builder.AppendLine($"{idenStr}}}");
            Builder.AppendLine($"{idenStr}else");
            Builder.AppendLine($"{idenStr}{{");
            Builder.AppendLine($"{idenStr}\tcontent = System.Text.Encoding.GetEncoding(1252).GetBytes({this.Parameters[5].Value});");
            Builder.AppendLine($"{idenStr}}}");
            Builder.AppendLine($"{idenStr}MemoryStream contentStream = new MemoryStream(content);");
            Builder.AppendLine($"{idenStr}Gvinci.Plugin.AWS.S3.S3Helper.UploadFile(accessKey, secretKey, region, bucketName, contentStream, key, contentType, {this.Parameters[7].Value});");
        }

    }
}