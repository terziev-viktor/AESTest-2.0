using System.Collections.Generic;

namespace AESTest2._0
{
    public class DataHolder
    {
        public DataHolder()
        {
            this.Students = new List<Student>();
            this.Exams = new List<Exam>();
            this.Questions = new List<Question>();
            this.Posts = new List<Post>();
        }
        public int CurrentExamIndex { get; set; }

        public List<Student> Students { get; set; }

        public List<Post> Posts { get; set; }

        public List<Exam> Exams { get; set; }

        public List<Question> Questions { get; set; }
    }
}
