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
        private Office.IRibbonUI ribbon;
        private Timer timer;

        public bool hasStartAnswer = false;

        public MainRibbon() {
            timer = new Timer(3000);
            timer.Elapsed += Timer_Elapsed;
            timer.AutoReset = false;
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e) {
            
            
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

        public void OnAddSingleQuestionButton_Click(Office.IRibbonControl control) {
            ThisAddIn main = Globals.ThisAddIn;
            foreach (vsto.CustomTaskPane ctp in main.CustomTaskPanes) {
                if(ctp.Control is AddQuestionForm) {
                    main.CustomTaskPanes.Remove(ctp);
                    break;
                }
            }
            vsto.CustomTaskPane ct = main.CustomTaskPanes.Add(new AddQuestionForm(), "Add a question");
            ct.Visible = true;
            ct.Width = 250;
        }

        public void OnDanmakuButton_Click(Office.IRibbonControl control) {
            if (Setting.problemNumber != -1) {
                if (Communitcate.SendAndSavePic("http://www.zhengzi.me/classpie/controller/appctrl.php?func=getChart&para={\"problemId\":" + Setting.problemNumber.ToString() + "}") == true) {
                    PowerPoint.Application Application = Globals.ThisAddIn.Application;
                    PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
                    PowerPoint.Slide slide;
                    if (Application.SlideShowWindows.Count > 0) {
                        slide = slides._Index(Application.SlideShowWindows[1].View.Slide.SlideIndex);
                    }
                    else {
                        slide = slides._Index(Application.ActiveWindow.Selection.SlideRange.SlideIndex);
                    }
                    Image image = Image.FromFile(Path.GetTempPath() + "\\classpieaddin.bmp");
                    float screenHeight = (float)SystemParameters.PrimaryScreenHeight;
                    float screenWidth = (float)SystemParameters.PrimaryScreenWidth;
                    MessageBox.Show((Math.Max(screenWidth - image.Width, 0) / 2).ToString());
                    MessageBox.Show((Math.Max(screenHeight - image.Height, 0) / 2).ToString());
                    slide.Shapes.AddPicture(Path.GetTempPath() + "\\classpieaddin.bmp", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Math.Max(screenWidth - image.Width, 0) / 2, Math.Max(screenHeight - image.Height, 0) / 2);
                }
                else {
                    MessageBox.Show("Communication with server has encountered \nsome problem (Error Code 1).");
                }
            }
            else {
                MessageBox.Show("Please upload your questions first.");
            }
        }

        public void OnUploadQuestionButton_Click(Office.IRibbonControl control) {
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            List<Question> problem = new List<Question>();
            int number = 0;
            foreach (PowerPoint.Slide slide in slides) {
                try {
                    if(slide.Tags["Question"] == "Yes") {
                        Question question = new Question(++number,slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text);
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
            if (problem.Count > 0) {
                string message = JsonConvert.SerializeObject(problem);
                MessageBox.Show(message);
                string str = Communitcate.SendAndReceiveStr(message + "}", "http://www.zhengzi.me/classpie/controller/appctrl.php?func=sendProblem&para={\"problemInfo\":");
                try {
                    Setting.problemNumber = int.Parse(str);
                }
                catch {
                    MessageBox.Show("Communication with server has encountered \nsome problem (Error Code 1).");
                }
            }
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

        public Boolean GetEndButtonEnabled(Office.IRibbonControl control) {
            return hasStartAnswer;
        }

        public void BeginButton_Click(Office.IRibbonControl control) {
            if (Setting.problemNumber == -1) {
                MessageBox.Show("Please upload your questions first.");
            }
            else {
                hasStartAnswer = true;
            }
        }

        public void EndButton_Click(Office.IRibbonControl control) {
            if(hasStartAnswer == true) {
                if (Setting.problemNumber != -1) {
                    if (Communitcate.SendAndSavePic("http://www.zhengzi.me/classpie/controller/appctrl.php?func=getChart&para={\"problemId\":" + Setting.problemNumber.ToString() + "}") == true) {
                        PowerPoint.Application Application = Globals.ThisAddIn.Application;
                        PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
                        PowerPoint.Slide slide;
                        if (Application.SlideShowWindows.Count > 0) {
                            slide = slides._Index(Application.SlideShowWindows[1].View.Slide.SlideIndex);
                        }
                        else {
                            slide = slides._Index(Application.ActiveWindow.Selection.SlideRange.SlideIndex);
                        }
                        slide.Shapes.AddPicture(Path.GetTempPath() + "\\classpieaddin.bmp", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0);
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
