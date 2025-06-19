using System.Collections.Generic;

namespace Gvinci.Plugin.Action
{
    public class GPluginActionS3AWSCreate : GPluginAction
    {
        public override string ID => "7A69D6A0-CF98-4063-0001-6F7E02EC5F91";

        public override string Name => "Create File in S3";

        public override string Description => "";

        private List<GPluginActionParameter> _Parameters;
        public override List<GPluginActionParameter> Parameters
        {
            get
            {
                return _Parameters;
            }
        }

        public GPluginActionS3AWSCreate(IGPlugin Plugin) : base(Plugin)
        {
            _Parameters = new List<GPluginActionParameter>()
                    {
                        new GPluginActionParameter() { ID = 1, Name = "S3AccessKey", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 2, Name = "S3SecretKey", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 3, Name = "Region", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 4, Name = "BucketName", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 5, Name = "Path", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 6, Name = "ContentType", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 7, Name = "FileContent", Type = PluginActionParameterTypeEnum.STRING },
                        new GPluginActionParameter() { ID = 8, Name = "IsPublic", Type = PluginActionParameterTypeEnum.BOOL },
                    };
        }

        public override string GetActionCall()
        {
            string accessKey = this.Parameters[0].Value.ToString();
            string secretKey = this.Parameters[1].Value.ToString();
            string region = $"Amazon.RegionEndpoint.GetBySystemName({this.Parameters[2].Value.ToString()})";
            string bucketName = this.Parameters[3].Value.ToString();
            string key = this.Parameters[4].Value.ToString();
            string ContentType = this.Parameters[5].Value.ToString();
            string contentStream = $"new MemoryStream(System.Text.Encoding.GetEncoding(1252).GetBytes({this.Parameters[6].Value.ToString()}))";
            return $"Gvinci.Plugin.AWS.S3.S3Helper.UploadFile({accessKey}, {secretKey}, {region}, {bucketName}, {contentStream}, {key}, {ContentType}, {this.Parameters[7].Value.ToString()});";
        }

    }
}