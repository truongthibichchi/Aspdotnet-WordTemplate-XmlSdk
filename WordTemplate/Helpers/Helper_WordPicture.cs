using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WordTemplate.Helpers
{
    class Helper_WordPicture : Helper_WordBase
    {
        string image_path = @"F:\image.jpg";

        public Helper_WordPicture(SdtElement content_control, string url)
        {
            this.content_control = content_control;
            if (this.content_control == null)
                StaticValues.logs += "[Error]Can't get the content control" + Environment.NewLine;
            image_path = url;
        }

        /// <summary>
        /// Get the picture and add it into the content control
        /// </summary>
        public void AddPictureFromUri()
        {
            string pic_id = null;
            Drawing dr = content_control.Descendants<Drawing>().FirstOrDefault();
            if (dr != null)
            {
                Blip blip = dr.Descendants<Blip>().FirstOrDefault();
                if (blip != null)
                    pic_id = blip.Embed;
            }

            if (pic_id != null)
            {
                IdPartPair idpp = WordTemplateManager.document.MainDocumentPart.Parts
                    .Where(pa => pa.RelationshipId == pic_id).FirstOrDefault();
                if (idpp != null)
                {
                    ImagePart ip = (ImagePart)idpp.OpenXmlPart;
                    if (IsImageUrl(image_path))
                    {
                        Stream stream = DownloadData(image_path);
                        ip.FeedData(stream);
                        stream.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Download data from url
        /// </summary>
        /// <param name="url">A link</param>
        /// <returns></returns>
        protected Stream DownloadData(string url)
        {
            WebRequest req = WebRequest.Create(url);
            WebResponse response = req.GetResponse();
            Stream stream = response.GetResponseStream();
            return stream;
        }

        /// <summary>
        /// Detect image url
        /// </summary>
        /// <param name="url">A link</param>
        /// <returns></returns>
        public static bool IsImageUrl(string url)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
                request.Method = "HEAD";
                using (WebResponse respond = request.GetResponse())
                {
                    return respond.ContentType.ToLower(CultureInfo.InvariantCulture)
                               .StartsWith("image/");
                }
            }
            catch (Exception)
            {
                return false;
            }

        }

        /// <summary>
        /// Detect url
        /// </summary>
        /// <param name="url">A link</param>
        /// <returns></returns>
        public static bool IsUrl(string url)
        {
            return Uri.IsWellFormedUriString(url, UriKind.RelativeOrAbsolute);
            /*
            try
            {
                Uri uriResult;
                bool result = Uri.TryCreate(url, UriKind.Absolute, out uriResult)
                    && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            */
        }
    }
}
