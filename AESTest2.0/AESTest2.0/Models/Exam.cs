namespace AESTest2._0
{
    public enum ExamType
    {
        Ordinance9 = 0,
        ForSafety = 1
    }

    public class Exam
    {
        public string Title { get; set; }

        public int QuestionsCount { get; set; }

        public int MinScore { get; set; }

        public string ProtocolNumberPath { get; set; }

        public ExamType Type { get; set; }
    }
}
