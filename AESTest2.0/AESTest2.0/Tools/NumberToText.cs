namespace AESTest2._0.Tools
{
    static class NumberToText
    {
        public static string AsString(int num)
        {
            switch(num)
            {
                case 1:
                    return "Първа";
                case 2:
                    return "Втора";
                case 3:
                    return "Трета";
                case 4:
                    return "Четвърта";
                case 5:
                    return "Пета";
                case 6:
                    return "Шеста";
                case 7:
                    return "Седма";
                case 8:
                    return "Осма";
                case 9:
                    return "Девета";
                default:
                    return "Нулева";
            }
        }

        public static string AsString(string num)
        {
            int parsed;
            bool success = int.TryParse(num, out parsed);
            if (success)
            {
                return AsString(parsed);
            }
            return "Нулева";
        }
    }
}
