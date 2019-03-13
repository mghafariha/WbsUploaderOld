using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    public class test1
    {
        [sharepointFieldName("شناسه")]
        public int id { get; set; }
        public string title { get; set; }
        public int age { get; set; }

     

    }

    public class sharepointFieldName : System.Attribute
    {
        public readonly string shFieldName;

        public sharepointFieldName(string sharepointFieldName)
        {
            this.shFieldName = sharepointFieldName;
        }


    }

    //public static class aa
    //{
    //    public static string fieldName(this Object obj)
    //    {
    //        System.Attribute[] attrs = System.Attribute.GetCustomAttributes(obj);

    //        foreach (System.Attribute attr in attrs)
    //        {
    //            if (attr is sharepointFieldName)
    //            {
    //                return (sharepointFieldName)attr.fieldName();
    //            }
    //        }
    //        return "";
    //    }
    //}

}



