using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

   

            test1 myClass = new test1();
            myClass.age = 25;
            myClass.id = 8;
            myClass.title = "sdsadasdsa";
            Type myClassType = myClass.GetType();
            PropertyInfo[] properties = myClassType.GetProperties();

           // System.Reflection.PropertyInfo info = typeof(test1);
          //  object[] attributes = info.GetCustomAttributes(true);


            foreach (PropertyInfo property in properties)
            {
                System.Attribute[] attrs = System.Attribute.GetCustomAttributes(property);
               string s= ((sharepointFieldName)attrs[0]).shFieldName;
               MessageBox.Show("Name: " + property.Name + ", Value: " + property.GetValue(myClass, null));
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/ControlProject/");
            context.Credentials = new NetworkCredential("sp_admin", "nik5025085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = web.Lists.GetByTitle("ارزشیابی مدیر طرح");
            ListItem item = list.GetItemById(int.Parse(textBox2.Text));

         //   var query = new CamlQuery();
         //   query.ViewXml = "<View/>";
            //            query.ViewXml = string.Format(@"<View><Query>
            //   <Where>
            //      <Eq>
            //         <FieldRef Name='CId' />
            //         <Value Type='Text'>123</Value>
            //      </Eq>
            //   </Where>
            //</Query></View>", "123");

           // var items = list.GetItems(query);
          //  context.Load(items);
          //  context.ExecuteQuery();

          //  foreach (ListItem item in items)
            {
                item["CompanyID"] = textBox1.Text;
                item["CompanyName"] = textBox3.Text;
                item.Update();
            }
            context.ExecuteQuery();
            MessageBox.Show("Done");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/ProjectsInfo/");
            context.Credentials = new NetworkCredential("sp_admin", "nik5025085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = web.Lists.GetByTitle("عملیات اجرایی در سطح واحد زراعی");
            var query = new CamlQuery();
            //query.ViewXml = "<View/>";
            query.ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name='Contract'/><Value Type='Text'>105</Value></Eq></Where></Query></View>");

            var items = list.GetItems(query);
           context.Load(items);
            context.ExecuteQuery();

            foreach (ListItem item in items)
            {
                MessageBox.Show(item.Id.ToString());
            }
            MessageBox.Show("done!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://172.29.0.178/ProjectsInfo/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = web.Lists.GetByTitle("پیمان ها و قراردادها");
            context.Load(list,l=>l.Fields);
            context.ExecuteQuery();

            string s = "";

            foreach (Field f in list.Fields)
            {
                s += f.InternalName + "        " + f.Title+"\r\n";   
            }
            // var query = new CamlQuery();
            // query.ViewXml = "<View/>";
            // query.ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name='Contract'/><Value Type='Text'>105</Value></Eq></Where></Query></View>");

            //var items = list.GetItems(query);
            //context.Load(items);
            //context.ExecuteQuery();

            //foreach (ListItem item in items)
            //{
            //    MessageBox.Show(item.Id.ToString());
            //}


            MessageBox.Show(s);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/") { Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr") };
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();

            var list = context.Web.Lists.GetByTitle("InvoiceCMDoc");
            var q = new CamlQuery();
            var r = 0;
            q.ViewXml = @"<View  Scope='Recursive'></View>";

            var result = list.GetItems(q);
            context.Load(result);
            context.ExecuteQuery();

            foreach (ListItem item in result)
            {
                //context.Load(item.File);
                //context.ExecuteQuery();
                item["Title"] = item["FileLeafRef"].ToString().Split('.')[0];
                item.Update();
                context.ExecuteQuery();
            }
            MessageBox.Show("done");
        }
    }
}
