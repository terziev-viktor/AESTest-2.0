using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AESTest2._0.Tools
{
    class Mapper
    {
        public static Dictionary<string, string> GetMap(DataHolder dataHolder)
        {
            Dictionary<string, string> map = new Dictionary<string, string>();
            map.Add("<exam>", dataHolder.CurrentExam.Title);
            map.Add("<pin>", dataHolder.CurrentStudent.PIN);
            map.Add("<post>", dataHolder.CurrentPost.Title);
            map["<group>"] = dataHolder.CurrentGroup;
            map["<protocol>"] = dataHolder.ProtocolNumber.ToString();
            map["<date>"] = DateTime.Now.ToShortDateString();
            if (dataHolder.CurrentExam.Type == ExamType.Ordinance9)
            {
                map["<dateplus>"] = DateTime.Now.AddYears(dataHolder.CurrentPost.DeltaYear).ToShortDateString();
            }
            else
            {
                map["<dateplus>"] = DateTime.Now.AddYears(1).ToShortDateString();
            }
            string[] nameSplitted = dataHolder.CurrentStudent.Fullname.Split(new char[] { ' ' });
            map["<fullname>"] = dataHolder.CurrentStudent.Fullname;
            map["<name>"] = nameSplitted[0];
            map["<sur>"] = nameSplitted[1];
            map["<famil>"] = nameSplitted[2];
            map["<mark>"] = dataHolder.Mark.ToString();
            return map;
        }
    }
}
