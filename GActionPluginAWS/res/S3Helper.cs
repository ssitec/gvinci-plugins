using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Amazon.S3.Transfer;
using Amazon.S3;
using Amazon.S3.Model;
using System.Web;
using Amazon.S3.IO;
using Amazon;

namespace Gvinci.Plugin.AWS.S3
{
    
    public struct S3FileInfoStruct
    {
        public string Name;
        public string FullName;
        public long Length;
        public bool Exists;
    }

    public static class S3Helper
    {

        public static string UploadFile(string S3AccessKey, string S3SecretKey, Amazon.RegionEndpoint Region, string BucketName, Stream ContentStream, string Key, string ContentType, bool IsPublic)
        {
            var uploadRequest = new TransferUtilityUploadRequest
            {
                BucketName = BucketName,
                InputStream = ContentStream,
                Key = Key,
                ContentType = ContentType
            };

            if (IsPublic) uploadRequest.CannedACL = S3CannedACL.PublicRead;

            TransferUtility utility = new TransferUtility(S3AccessKey, S3SecretKey, Region);
            utility.Upload(uploadRequest);

            return string.Format("https://{0}.s3.{1}.amazonaws.com/{1}", BucketName, Region.SystemName, Key);
        }

        public static bool MoveFile(string S3AccessKey, string S3SecretKey, Amazon.RegionEndpoint Region, string BucketName, string SourceFileUrl, string TargetFileUrl)
        {
            using (var client = new AmazonS3Client(S3AccessKey, S3SecretKey, RegionEndpoint.SAEast1))
            {
                string fileKey = SourceFileUrl;
                if (SourceFileUrl.StartsWith("http"))
                {
                    Uri fInfo = new Uri(SourceFileUrl);
                    fileKey = fInfo.LocalPath.TrimStart('/');
                }
                var info = new Amazon.S3.IO.S3FileInfo(client, BucketName, fileKey);
                return info.MoveTo(BucketName, TargetFileUrl).Exists;
            }
        }

    }

}
