using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using vsto = Microsoft.Office.Tools;
using System.Timers;
using System.Drawing;
using System.Windows;
using System.Net;
using Newtonsoft.Json;

// TODO:   按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MainRibbon();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace ClassPieAddin {
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility {
        public Office.IRibbonUI ribbon;

        public bool isDanmakuOn = false;
        public bool hasStartAnswer = false;

        public MainRibbon() {
            Setting.mainRibbon = this;
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("ClassPieAddin.MainRibbon.xml");
        }

        #endregion

        #region 功能区回调
        //在此创建回调方法。有关添加回调方法的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        public void OnAddQuestionButton_Click(Office.IRibbonControl control) {
            ThisAddIn main = Globals.ThisAddIn;
            foreach (vsto.CustomTaskPane ctp in main.CustomTaskPanes) {
                if(ctp.Control is AddQuestionForm || ctp.Control is ModifyQuestionForm) {
                    main.CustomTaskPanes.Remove(ctp);
                    break;
                }
            }
            vsto.CustomTaskPane ct = main.CustomTaskPanes.Add(new AddQuestionForm(), "添加问题");
            ct.Visible = true;
            ct.Width = 270;
        }

        public void OnDanmakuButton_Click(Office.IRibbonControl control, bool pressed) {
            isDanmakuOn = pressed;
            ribbon.InvalidateControl("danmakuButton");
        }

        public void OnUploadQuestionButton_Click(Office.IRibbonControl control) {
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            Problem problem = new Problem();
            foreach (PowerPoint.Slide slide in slides) {
                try {
                    if(slide.Tags["Question"] == "Yes") {
                        Question question = new Question(int.Parse(slide.Tags["Number"]),slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text);
                        try {
                            for (int i = 1; i < 26; i++) {
                                question.Add(slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text);
                            }
                        }
                        catch (COMException innerError) {
                        }
                        problem.Add(question);
                    }
                }
                catch (COMException outerError) {

                }
            }
            if (problem.count > 0) {
                string message = JsonConvert.SerializeObject(problem);
                //MessageBox.Show(message);
                string str = Communitcate.SendAndReceiveStr(message + "}", "http://www.zhengzi.me/classpie/controller/appctrl.php?func=sendProblem&para={\"problemInfo\":");
                System.Diagnostics.Debug.WriteLine(str);
                try {
                    Setting.problemNumber = int.Parse(str);
                    PowerPoint.Application Application = Globals.ThisAddIn.Application;
                    PowerPoint.Slide slide;
                    if (Application.SlideShowWindows.Count > 0) {
                        slide = slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                        slide.Select();
                    }
                    else {
                        try {
                            slide = slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                            slide.Select();
                        }
                        catch (COMException error) {
                            slide = slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                            slide.Select();
                        }
                    }
                    string fileName = Communitcate.SendAndSavePic("http://www.zhengzi.me/classpie/controller/appctrl.php?func=getQrCode&para={\"problemId\":" + Setting.problemNumber.ToString() + "}");
                    if (fileName != null) {
                        Image image = Image.FromFile(fileName);
                        float screenHeight = (float)SystemParameters.PrimaryScreenHeight;
                        float screenWidth = (float)SystemParameters.PrimaryScreenWidth;
                        slide.Shapes.AddPicture(fileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 280, 100, 400, 400);
                        var textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 180, 50, 800, 50);
                        textBoxAnswer.Name = "tipsTextBox";
                        textBoxAnswer.TextFrame.TextRange.Text = "Scan the QR code to answer questions.";
                        textBoxAnswer.TextFrame.TextRange.Font.Size = 36;
                    }
                    else {
                        MessageBox.Show("Communication with server has encountered \nsome problem (Error Code 1).");
                    }
                }
                catch {
                    MessageBox.Show("Communication with server has encountered \nsome problem (Error Code 1)."+"\nReceive: "+str);
                }
            }
        }

        public void OnModifyQuestionButton_Click(Office.IRibbonControl control) {
            ThisAddIn main = Globals.ThisAddIn;
            foreach (vsto.CustomTaskPane ctp in main.CustomTaskPanes) {
                if (ctp.Control is ModifyQuestionForm || ctp.Control is AddQuestionForm) {
                    main.CustomTaskPanes.Remove(ctp);
                    break;
                }
            }
            vsto.CustomTaskPane ct = main.CustomTaskPanes.Add(new ModifyQuestionForm(), "修改问题");
            ct.Visible = true;
            ct.Width = 270;
        }

        #endregion

        #region Functions

        public Boolean GetBeginButtonEnabled(Office.IRibbonControl control) {
            PowerPoint.Application Application = Globals.ThisAddIn.Application;
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            PowerPoint.Slide slide;
            if (Application.SlideShowWindows.Count > 0) {
                slide = slides._Index(Application.SlideShowWindows[1].View.Slide.SlideIndex);
            }
            else {
                slide = slides._Index(Application.ActiveWindow.Selection.SlideRange.SlideIndex);
            }
            try {
                if(slide.Tags["Question"] == "Yes") {
                    return true & !hasStartAnswer;
                }
                else {
                    return false;
                }
            }
            catch (COMException error){
                return false;
            }
        }

        public Boolean GetModifyQuestionButtonEnabled(Office.IRibbonControl control) {
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            PowerPoint.Slide slide;
            if (Globals.ThisAddIn.Application.SlideShowWindows.Count > 0) {
                slide = slides._Index(Globals.ThisAddIn.Application.SlideShowWindows[1].View.Slide.SlideIndex);
            }
            else {
                try {
                    slide = slides._Index(Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex);
                }
                catch {
                    ThisAddIn main = Globals.ThisAddIn;
                    foreach (vsto.CustomTaskPane ctp in main.CustomTaskPanes) {
                        if (ctp.Control is ModifyQuestionForm) {
                            main.CustomTaskPanes.Remove(ctp);
                            break;
                        }
                    }
                    return false;
                }
            }
            try {
                if (slide.Tags["Question"] == "Yes") {
                    return true;
                }
                else {
                    return false;
                }
            }
            catch {
                return false;
            }
        }

        public Boolean GetEndButtonEnabled(Office.IRibbonControl control) {
            return hasStartAnswer;
        }

        public Bitmap GetUploadButtonImage(Office.IRibbonControl control) {
            return IconResource.upload;
        }

        public Bitmap GetAddQuestionImage(Office.IRibbonControl control) {
            return IconResource.addQuestion;
        }

        public Bitmap GetModifyQuestionImage(Office.IRibbonControl control) {
            return IconResource.modifyQuestion;
        }

        public string GetDanmakuLabel(Office.IRibbonControl control) {
            if(isDanmakuOn == true) {
                return "弹幕 开";
            }
            else {
                return "弹幕 关";
            }
        }

        public Bitmap GetDanmakuImage(Office.IRibbonControl control) {
            if(isDanmakuOn == true) {
                return IconResource.danmakuOn;
            }
            else {
                return IconResource.danmakuOff;
            }
        }

        public void BeginButton_Click(Office.IRibbonControl control) {
            if (Setting.problemNumber == -1) {
                MessageBox.Show("请先上传并生成问卷。");
            }
            else {
                hasStartAnswer = true;
            }
        }

        public void EndButton_Click(Office.IRibbonControl control) {
            if(hasStartAnswer == true) {
                PowerPoint.Application Application = Globals.ThisAddIn.Application;
                PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
                PowerPoint.Slide slide;
                if (Application.SlideShowWindows.Count > 0) {
                    slide = slides._Index(Application.SlideShowWindows[1].View.Slide.SlideIndex);
                }
                else {
                    slide = slides._Index(Application.ActiveWindow.Selection.SlideRange.SlideIndex);
                }
                if (Setting.problemNumber != -1) {
                    string fileName = Communitcate.SendAndSavePic("http://www.zhengzi.me/classpie/controller/appctrl.php?func=getChart&para={\"problemId\":" + Setting.problemNumber.ToString() + ", \"questionId\":" + slide.Tags["Number"] + "}");
                    if (fileName != null){
                        Image image = Image.FromFile(fileName);
                        float screenHeight = (float)SystemParameters.PrimaryScreenHeight;
                        float screenWidth = (float)SystemParameters.PrimaryScreenWidth;
                        if (Application.SlideShowWindows.Count > 0) {
                            slide = slides.Add(Application.SlideShowWindows[1].View.Slide.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                            Globals.ThisAddIn.Application.SlideShowWindows[1].View.Next();
                        }
                        else {
                            try {
                                slide = slides.Add(Application.ActiveWindow.Selection.SlideRange.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                                slide.Select();
                            }
                            catch (COMException error) {
                                slide = slides.Add(slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                                slide.Select();
                            }
                        }
                        slide.Shapes.AddPicture(fileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Math.Max(screenWidth - image.Width, 0) / 2, Math.Max(screenHeight - image.Height, 0) / 2);
                        var textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 25, 800, 50);
                        textBoxAnswer.Name = "titleTextBox";
                        textBoxAnswer.TextFrame.TextRange.Text = "Student Answers:";
                        textBoxAnswer.TextFrame.TextRange.Font.Size = 36;
                        hasStartAnswer = false;
                    }
                    else {
                        MessageBox.Show("Communication with server has encountered \nsome problem (Error Code 1).");
                    }
                }
                else {
                    MessageBox.Show("Please upload your questions first.");
                }
            }
        }

        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
