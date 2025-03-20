using DevExpress.Office.Services;

namespace DocumentProcessingWebAPI.BusinessObjects
{
    public class MailMergeUriStreamProvider : IUriStreamProvider
    {
        const string Prefix = "dbimg://";

        Stream? IUriStreamProvider.GetStream(string uri)
        {
            uri = uri.Trim();
            if (!uri.StartsWith(Prefix))
                return null;
            string strId = uri.Substring(Prefix.Length).Trim();
            int id;
            if (!int.TryParse(strId, out id))
                return null;
            string pictureRoot = "Documents/";
            string fileName = string.Format("{0}Photo{1}.jpeg", pictureRoot, id);
            byte[] bytes = new byte[0];
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                int length = (int)fs.Length;
                bytes = new byte[length];
                fs.Read(bytes, 0, length);
            }
            return new MemoryStream(bytes);
        }
    }


}
