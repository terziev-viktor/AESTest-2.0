using System.Collections.Generic;

namespace AESTest2._0
{
    public class DataHolder
    {
        private int mark;

        private bool markCalculated;

        public DataHolder()
        {
            this.Students = new List<Student>();
            this.Exams = new List<Exam>();
            this.Questions = new List<Question>();
            this.Posts = new List<Post>();
            this.Groups = new List<string>();
            this.markCalculated = false;
        }

        public Student CurrentStudent { get; set; }

        public Post CurrentPost { get; set; }

        public Exam CurrentExam { get; set; }

        public string CurrentGroup { get; set; }

        public int ProtocolNumber { get; set; }

        public int Mark { get 
            {
                if (markCalculated)
                {
                    return this.mark;
                }
                if (this.CurrentExam == null)
                {
                    return 0;
                }
                int c = 0;
                for (int i = 0; i < this.CurrentExam.QuestionsCount; i++)
                {
                    if (this.Questions[i].RightAnswer == this.Questions[i].StudentsAnswer)
                    {
                        ++c;
                    }
                }
                this.mark = ((c * 100) / this.CurrentExam.QuestionsCount);
                this.markCalculated = true;
                return this.mark;
            } 
        }

        public List<Student> Students { get; set; }

        public List<Post> Posts { get; set; }

        public List<Exam> Exams { get; set; }

        public List<Question> Questions { get; set; }

        public List<string> Groups { get; set; }
    }
}
