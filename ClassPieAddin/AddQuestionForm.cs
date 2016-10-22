using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace ClassPieAddin {
    public partial class AddQuestionForm : UserControl {
        public List<TextBox> textboxArray = new List<TextBox>();
        public List<Label> labelArray = new List<Label>();
        public int count = 0;
        public PowerPoint.Slide slide;
        public AddQuestionForm() {
            InitializeComponent();
            CreateSingleQuestionPage();
            Init();
        }

        public void CreateSingleQuestionPage() {
            PowerPoint.Application Application = Globals.ThisAddIn.Application;
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
            PowerPoint.Slide slide;
            if (Application.SlideShowWindows.Count > 0) {
                slide = slides.Add(Application.SlideShowWindows[1].View.Slide.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            }
            else {
                try {
                    slide = slides.Add(Application.ActiveWindow.Selection.SlideRange.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    slide.Select();
                }
                catch(COMException error) {
                    slide = slides.Add(slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    slide.Select();
                }
            }
            if (slides.Count > 0) {
                try {
                    Setting.questionCount = int.Parse(slides[1].Tags["TotalCount"]);
                }
                catch {
                }
            }
            slide.Tags.Add("Question", "Yes");
            slide.Tags.Add("Number", (++Setting.questionCount).ToString());
            slide.Tags.Add("Selection", "Single");
            slide.Tags.Add("Answer", "Not specified.");
            slide.Tags.Add("Count", "4");
            if(slides.Count > 0) {
                slides[1].Tags.Add("TotalCount", Setting.questionCount.ToString());
            }

            var textBoxTitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 30, 600, 50);
            textBoxTitle.Name = "TextBoxTitle";
            textBoxTitle.TextFrame.TextRange.Text = "Question";//设置文本框的内容
            textBoxTitle.TextFrame.TextRange.Font.Size = 72;//设置文本字体大小
            textBoxTitle.TextFrame.TextRange.Font.Color.RGB = Color.Black.ToArgb();
            textBoxTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;

            var textBoxQuestion = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 150, 600, 50);
            textBoxQuestion.Name = "TextBoxQuestion";
            textBoxQuestion.TextFrame.TextRange.Text = "Input the description of the question here.";//设置文本框的内容
            textBoxQuestion.TextFrame.TextRange.Font.Size = 32;//设置文本字体大小
            textBoxQuestion.TextFrame.TextRange.Font.Color.RGB = Color.Black.ToArgb();

            //What if not exist?
            var textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 200, 600, 50);
            textBoxAnswer.Name = "TextBoxAnswer1";
            textBoxAnswer.TextFrame.TextRange.Text = "A. Choice 1.";
            textBoxAnswer.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小

            textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 250, 600, 50);
            textBoxAnswer.Name = "TextBoxAnswer2";
            textBoxAnswer.TextFrame.TextRange.Text = "B. Choice 2.";
            textBoxAnswer.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小

            textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 300, 600, 50);
            textBoxAnswer.Name = "TextBoxAnswer3";
            textBoxAnswer.TextFrame.TextRange.Text = "C. Choice 3.";
            textBoxAnswer.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小

            textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 350, 600, 50);
            textBoxAnswer.Name = "TextBoxAnswer4";
            textBoxAnswer.TextFrame.TextRange.Text = "D. Choice 4.";
            textBoxAnswer.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小
        }

        public void Init() {
            textboxArray.Add(textBox1);
            textboxArray.Add(textBox2);
            textboxArray.Add(textBox3);
            textboxArray.Add(textBox4);
            labelArray.Add(label1);
            labelArray.Add(label2);
            labelArray.Add(label3);
            labelArray.Add(label4);
            try {
                PowerPoint.Application Application = Globals.ThisAddIn.Application;
                PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
                if (Application.SlideShowWindows.Count > 0) {
                    slide = slides._Index(Application.SlideShowWindows[1].View.Slide.SlideIndex);
                }
                else {
                    slide = slides._Index(Application.ActiveWindow.Selection.SlideRange.SlideIndex);
                }
                if (slide.Tags["Question"] == "Yes") {
                    textBoxQuestion.Text = slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text;
                    textBoxQuestion.TextChanged += ChoiceBox_TextChanged;
                    for (int i = 1; i <= 26; i++) {
                        if (i <= 4) {
                            textboxArray[i - 1].Text = slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text;
                        }
                        else {
                            if (slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text != null) {
                                TextBox text = new TextBox();
                                this.panel1.Controls.Add(text);
                                text.Location = new Point(86, 45 + 41 * i);
                                text.Name = "textBox" + i.ToString();
                                text.Size = new System.Drawing.Size(129, 21);
                                text.Text = slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text;
                                Label label = new Label();
                                this.panel1.Controls.Add(label);
                                label.Location = new Point(20, 47 + 40 * i);
                                label.Name = "label" + i.ToString();
                                label.Text = "Choice " + i.ToString();
                                textboxArray.Add(text);
                                labelArray.Add(label);
                            }
                        }
                        count++;
                    }
                }
            }
            catch(COMException error) {
            }
            for (int i = 0; i < count; i++) {
                System.Diagnostics.Debug.WriteLine(i);
                textboxArray[i].TextChanged += ChoiceBox_TextChanged;
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e) {
            count++;
            TextBox text = new TextBox();
            this.panel1.Controls.Add(text);
            text.Location = new Point(86, 45 + 41 * (count));
            text.Name = "textBox" + count.ToString();
            text.Size = new System.Drawing.Size(129, 21);
            Label label = new Label();
            this.panel1.Controls.Add(label);
            label.Location = new Point(20, 47 + 40 * count);
            label.Name = "label" + count.ToString();
            label.Text = "Choice " + count.ToString();
            textboxArray.Add(text);
            var textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 150 + count * 50, 600, 50);
            textBoxAnswer.Name = "TextBoxAnswer" + count.ToString();
            text.Text = textBoxAnswer.TextFrame.TextRange.Text = ((char)('A'+ count - 1)).ToString() +". Choice " + count.ToString() + ".";
            textBoxAnswer.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小
            text.TextChanged += ChoiceBox_TextChanged;
            labelArray.Add(label);
        }

        private void buttonDelete_Click(object sender, EventArgs e) {
            if (count > 0) {
                count--;
                this.panel1.Controls.Remove(textboxArray[count]);
                textboxArray.RemoveAt(count);
                this.panel1.Controls.Remove(labelArray[count]);
                labelArray.RemoveAt(count);
                slide.Shapes["TextBoxAnswer" + (count + 1).ToString()].Delete();
            }
            else {
                MessageBox.Show("No choice.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ChoiceBox_TextChanged(object sender, EventArgs e) {
            try {
                if (slide.Tags["Question"] == "Yes") {
                    slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text = textBoxQuestion.Text;
                    for (int i = 1; i <= 26; i++) {
                        slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text = textboxArray[i - 1].Text;
                    }
                }
            }
            catch {
                return;
            }
        }
    }
}
