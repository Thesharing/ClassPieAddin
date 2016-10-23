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

        public static string SendAndSavePic(string url) {
            System.Diagnostics.Debug.WriteLine(url);
            string fileName = null;
            //try {
                using (WebClient client = new WebClient()) {
                    byte[] buffer = client.DownloadData(url);
                    Image image = (Bitmap)((new ImageConverter()).ConvertFrom(buffer));
                    Random random = new Random();
                    fileName = Path.GetTempPath() + "\\classpieaddin" + (random.Next() % 1000).ToString() + ".png";
                    image.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                }
                return fileName;
            //}
            //catch {
                //return false;
            //}
        }
    }
}
