using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassPieAddin {
    public class Problem {
        public string title;
        public int count = 0;
        public List<Question> question;

        public Problem() {
            this.title = Globals.ThisAddIn.Application.ActivePresentation.FullName;
            question = new List<Question>();
        }

        public Problem(string title) {
            this.title = title;
            question = new List<Question>();
        }

        public void Add(Question question) {
            this.question.Add(question);
            count = this.question.Count;
        }
    }
}
