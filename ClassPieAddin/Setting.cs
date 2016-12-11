using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ClassPieAddin {
    public static class Setting {
        public static int problemNumber = -1;
        public static int speed = 5;
        public static int questionCount = 0;
        public static int nowScreen = 0;
        public static int num = 5;
        public static int fontSize = 24;
        public static MainRibbon mainRibbon;
        public static bool isLesson = false;

        public static Problem GetCurrentProblemSet() {
            Microsoft.Office.Interop.PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            Problem problem = new Problem();
            foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in slides) {
                try {
                    if (slide.Tags["Question"] == "Yes") {
                        Question question = new Question(int.Parse(slide.Tags["Number"]), slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text);
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
            if(problem.count > 0) {
                return problem;
            }
            else {
                return null;
            }
        }

        public static int GetCurrentQuestionNo(Microsoft.Office.Interop.PowerPoint.Slide slide) {
            try {
                System.Diagnostics.Debug.WriteLine(slide.Tags["Number"]);
                int number =  int.Parse(slide.Tags["Number"]);
                return number;
            }
            catch (COMException error) {
                return -1;
            }
            catch (FormatException error) {
                return -1;
            }
        }
    }
}
