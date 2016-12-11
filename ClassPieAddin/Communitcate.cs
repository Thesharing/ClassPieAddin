using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Timers;

namespace ClassPieAddin {
    public static class Communitcate {

        class UploadFileRequest {
            public string url;
            public string file;
            public NameValueCollection data;

            public UploadFileRequest(string url, string file, NameValueCollection data) {
                this.url = url;
                this.file = file;
                this.data = data;
            }
        }

        public static BackgroundWorker uploadFileBW = new BackgroundWorker();
        public static BackgroundWorker sendMessageBW = new BackgroundWorker();
        public static Timer uploadFileTimer = new Timer();
        public static Timer sendMessageTimer = new Timer();

        private static UploadFileRequest requestCache;
        private static bool hasRequestInCache;

        public static void Init() {
            uploadFileBW.WorkerReportsProgress = true;
            uploadFileBW.WorkerSupportsCancellation = true;
            uploadFileBW.DoWork += new DoWorkEventHandler(UploadFileBW_DoWork);
            uploadFileBW.ProgressChanged += new ProgressChangedEventHandler(UploadFileBW_ProgressChanged);
            uploadFileBW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(UploadFileBW_RunWorkerCompleted);

            uploadFileTimer.Elapsed += new ElapsedEventHandler(UploadTimeOut);
            uploadFileTimer.Interval = 10000;
            uploadFileTimer.AutoReset = false; // 不会自动重置计时器，即只计时一次

            sendMessageBW.WorkerReportsProgress = true;
            sendMessageBW.WorkerSupportsCancellation = true;
            sendMessageBW.DoWork += new DoWorkEventHandler(SendMessageBW_DoWork);
            sendMessageBW.ProgressChanged += new ProgressChangedEventHandler(SendMessageBW_ProgressChanged);
            sendMessageBW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(SendMessageBW_RunWorkerCompleted);

            sendMessageTimer.Elapsed += new ElapsedEventHandler(SendMessageTimeOut);
            sendMessageTimer.Interval = 5000;
            sendMessageTimer.AutoReset = false;
        }

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


        private static readonly Encoding DEFAULTENCODE = Encoding.UTF8;

        /// <summary>
        /// HttpUploadFile
        /// </summary>
        /// <param name="url"></param>
        /// <param name="file"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string HttpUploadFile(string url, string file, NameValueCollection data) {
            return HttpUploadFile(url, file, data, DEFAULTENCODE);
        }

        /// <summary>
        /// HttpUploadFile
        /// </summary>
        /// <param name="url"></param>
        /// <param name="file"></param>
        /// <param name="data"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static string HttpUploadFile(string url, string file, NameValueCollection data, Encoding encoding) {
            return HttpUploadFile(url, new string[] { file }, data, encoding);
        }

        /// <summary>
        /// HttpUploadFile
        /// </summary>
        /// <param name="url"></param>
        /// <param name="files"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string HttpUploadFile(string url, string[] files, NameValueCollection data) {
            return HttpUploadFile(url, files, data, DEFAULTENCODE);
        }

        /// <summary>
        /// HttpUploadFile
        /// </summary>
        /// <param name="url"></param>
        /// <param name="files"></param>
        /// <param name="data"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static string HttpUploadFile(string url, string[] files, NameValueCollection data, Encoding encoding) {
            try {
                string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
                byte[] boundarybytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");
                byte[] endbytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");

                //1.HttpWebRequest
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.ContentType = "multipart/form-data; boundary=" + boundary;
                request.Method = "POST";
                request.KeepAlive = true;
                request.Credentials = CredentialCache.DefaultCredentials;

                using (Stream stream = request.GetRequestStream()) {
                    //1.1 key/value
                    string formdataTemplate = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";
                    if (data != null) {
                        foreach (string key in data.Keys) {
                            stream.Write(boundarybytes, 0, boundarybytes.Length);
                            string formitem = string.Format(formdataTemplate, key, data[key]);
                            byte[] formitembytes = encoding.GetBytes(formitem);
                            stream.Write(formitembytes, 0, formitembytes.Length);
                        }
                    }

                    //1.2 file
                    string headerTemplate = "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: image/pjpeg\r\n\r\n";
                    byte[] buffer = new byte[4096];
                    int bytesRead = 0;
                    for (int i = 0; i < files.Length; i++) {
                        stream.Write(boundarybytes, 0, boundarybytes.Length);
                        string header = string.Format(headerTemplate, "file", Path.GetFileName(files[i]));
                        byte[] headerbytes = encoding.GetBytes(header);
                        stream.Write(headerbytes, 0, headerbytes.Length);
                        using (FileStream fileStream = new FileStream(files[i], FileMode.Open, FileAccess.Read)) {
                            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0) {
                                stream.Write(buffer, 0, bytesRead);
                            }
                        }
                    }
                    //1.3 form end
                    stream.Write(endbytes, 0, endbytes.Length);
                }
                //2.WebResponse
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (StreamReader stream = new StreamReader(response.GetResponseStream())) {
                    return stream.ReadToEnd();
                }
            }
            catch(Exception e) {
                System.Windows.MessageBox.Show(e.Message);
                return "";
            }
        }

        public static void HttpUploadFileBackground(string url, string file, NameValueCollection data) {
            UploadFileRequest request = new UploadFileRequest(url, file, data);
            if (uploadFileBW.IsBusy) {
                requestCache = request;
                hasRequestInCache = true;
                uploadFileBW.CancelAsync();
            }
            else {
                uploadFileBW.RunWorkerAsync(request);
                uploadFileTimer.Start();
            }
        }

        private static void UploadFileBW_DoWork(Object sender, DoWorkEventArgs e) {
            BackgroundWorker backgroundWorker = sender as BackgroundWorker;
            UploadFileRequest request = (UploadFileRequest)e.Argument;
            e.Result = HttpUploadFile(request.url, request.file, request.data);
            backgroundWorker.ReportProgress(100); // 当Dowork完成时直接将进度设为100%，执行RunWorkerCompleted
        }

        private static void UploadFileBW_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            return;
        }

        private static void UploadFileBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            if (e.Cancelled == true || e.Error != null) {
                System.Diagnostics.Debug.WriteLine("Wrong at Uploading Picture.");
            }
            uploadFileTimer.Stop();
            if(hasRequestInCache == true) {
                hasRequestInCache = false;
                uploadFileBW.RunWorkerAsync(requestCache);
            }
        }

        private static void SendMessageBW_DoWork(object sender, DoWorkEventArgs e) {
            BackgroundWorker backgroundWorker = sender as BackgroundWorker;
            string url = e.Argument as string;
            try {
                using (WebClient client = new WebClient()) {
                    byte[] buffer = client.DownloadData(url);
                    e.Result = buffer.ToString();
                }
            }
            catch (WebException error) {
                System.Diagnostics.Debug.WriteLine("Error when End.");
                backgroundWorker.CancelAsync();
            }
            backgroundWorker.ReportProgress(100);
        }

        private static void SendMessageBW_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            return;
        }

        private static void SendMessageBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            if (e.Cancelled == true || e.Error != null) {
                System.Diagnostics.Debug.WriteLine("Wrong at sending message.");
            }
            sendMessageTimer.Stop();
        }

        private static void SendMessageTimeOut(object sender, EventArgs e) {
            sendMessageBW.CancelAsync();
        }

        private static void UploadTimeOut(object sender, EventArgs e) {
            uploadFileBW.CancelAsync();
        }

        public static void EndSlideShow(string url) {
            sendMessageBW.RunWorkerAsync(url);
            sendMessageTimer.Start();
        }
    }
}
