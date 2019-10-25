using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCView
{
    class Model
    {
        private static List<string> data = new List<string>();

        public static List<string> Data { get => data; set => data = value; }

        static public List<string> GetData()
        {
            

            Data.Add("Afzaal");
            Data.Add("Ahmad");
            Data.Add("Zeeshan");
            Data.Add("Daniyal");
            Data.Add("Rizwan");
            Data.Add("John");
            Data.Add("Doe");
            Data.Add("Johanna Doe");
            Data.Add("Pakistan");
            Data.Add("Microsoft");
            Data.Add("Programming");
            Data.Add("Visual Studio");
            Data.Add("Sofiya");
            Data.Add("Rihanna");
            Data.Add("Eminem");

            return Data;
        }
    }
}
