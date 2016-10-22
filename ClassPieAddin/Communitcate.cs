using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ClassPieAddin {
    public static class Communitcate {
        public static string SendAndReceiveStr(string message, string url) {
            try {
                using (WebClient client = new WebClient()) {
                    byte[] buffer = client.DownloadData(url + message);
                    string str = Encoding.GetEncoding("UTF-8").GetString(buffer, 0, buffer.Length);
                    return str;
                }
            }
            catch (System.Net.WebException) {
                return "网络未连接。";
            }
        }

        public static bool SendAndSavePic(string url) {
            try {
                using (WebClient client = new WebClient()) {
                    byte[] buffer = client.DownloadData(url);
                    Image image = (Bitmap)((new ImageConverter()).ConvertFrom(buffer));
                    image.Save(Path.GetTempPath() + "\\classpieaddin.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
                }
                return true;
            }
            catch {
                return false;
            }
        }
    }
}
