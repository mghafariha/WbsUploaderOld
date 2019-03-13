
using Aspose.Cells;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace WindowsFormsApplication1
{
    public partial class Form2 : System.Windows.Forms.Form
    {

        static string errorMsg = "";

        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            ClientContext context = new ClientContext("http://pmis.jnasr.com/implementation/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.Load(web.Lists);
            context.ExecuteQuery();

            string strFile = "";
            string newline = "\r\n";
            foreach (List lst in web.Lists)
            {
                if (lst.BaseType == BaseType.GenericList || lst.BaseType == BaseType.DocumentLibrary)//
                {
                    string listID = lst.Id.ToString();
                    string listName = lst.EntityTypeName.Substring(0, lst.EntityTypeName.Length - 4);
                    strFile += string.Format(@"IF OBJECT_ID ('hn_vw_{0}', 'V') IS NOT NULL DROP VIEW hn_vw_{0} {1} GO{1}", listName, newline);
                    strFile += string.Format(@"create view hn_vw_{0} as
						select {2} from AllUserData where tp_ListId='{1}' and tp_DeleteTransactionId = 0 AND (tp_IsCurrent = 1) {3} Go {3}", listName, listID, getFields(lst, context), newline);
                    strFile += newline + newline;
                    //listBox1.Items.Add(lst.Title);
                }

            }



            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0)
            {
                using (TextWriter tw = new StreamWriter(saveFileDialog1.FileName + ".sql"))
                {
                    tw.WriteLine(strFile);
                }
            }
            MessageBox.Show("save Done!");
        }

        private string getFields(List lst, ClientContext context)
        {
            context.Load(lst.Fields);
            context.ExecuteQuery();
            string fields = "tp_id as id,nvarchar1 as title, ";
            foreach (Field field in lst.Fields)
            {
                if (!field.FromBaseType && !field.Hidden)
                {
                    string s = field.SchemaXml;

                    XmlDocument doc = new XmlDocument();
                    doc.InnerXml = s;
                    XmlElement root = doc.DocumentElement;
                    var look = "";
                    if (root.Attributes["Type"].Value == "Lookup" || root.Attributes["Type"].Value == "LookupMulti")
                        look = "ID";

                    try
                    {
                        fields += string.Format(@"[{0}] as [{1}{2}], ", root.Attributes["ColName"].Value, root.Attributes["Name"].Value, look);
                    }
                    catch (Exception ex)
                    {

                    }


                }
            }
            fields = fields.Substring(0, fields.Length - 2);
            return fields;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var wb = new Workbook(@"d:\WBS Jofeir 1.xlsx");

            //ClientContext context = new ClientContext("http://net-sp:100/ProjectsInfo/");
            //context.Credentials = new NetworkCredential("spadmin", "dm!n0sp0abg", "jnasr");
            ClientContext context = new ClientContext("http://pmis.jnasr.com/ProjectsInfo/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();




            CreateOtherItems(wb, web, context);
        }
        public void CreateOtherItems(Workbook wb, Web web, ClientContext context)
        {
            List farmList = web.Lists.GetByTitle("OperationsOnFarms");
            int valueCell = 5, totalValueCell = 7, amountCell = 8, totalAmountCell = 9, activityTypeCell = 10;
            int totalWeightContractCell = 11, itemWeightActionCell = 12, totalWeightOperationCell = 13, weightActionCell = 14;
            int rowId = 0, margeRowId = 0; ;

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();


            Worksheet wsBlock = wb.Worksheets["Block"];

            for (int i = 6; i < 100; i++)
            {
                if (wsBlock.Cells[i, 2].Value == null)
                    break;
                {
                    rowId = 15;
                    var ws = wb.Worksheets[wsBlock.Cells[i, 2].Value.ToString()];
                    var cId = 103;//new SPFieldLookupValue(contract.ToString()).LookupId;
                    var blockName = ws.Cells[7, 2].Value.ToString();
                    var block = GetBlockId(cId, blockName, context);
                    var impureArea = Convert.ToDouble(ws.Cells[7, 4].Value.ToString());
                    var farmName = ws.Cells[7, 3].Value.ToString();
                    var farm = GetFarmId(cId, block, farmName, impureArea, context);

                    #region Kanal Daraje 3

                    // کانال درجه 3( درجا)
                    //اجرا
                    margeRowId = 15;
                    // مترطول
                    rowId = 15;
                    var operationId = 1;
                    var subOperationId = 1;
                    var value = ws.Cells[rowId, valueCell].Value;
                    var totalValue = ws.Cells[rowId, totalValueCell].Value;
                    var amount = ws.Cells[rowId, amountCell].Value;
                    var totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    var activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    var eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    // هکتار

                    operationId = 1;
                    subOperationId = 1;
                    rowId = 16;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    // جاده سرویس  
                    operationId = 1;
                    subOperationId = 2;
                    rowId = 17;
                    operationId = 1;
                    subOperationId = 2;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //  سازه
                    operationId = 1;
                    subOperationId = 3;
                    rowId = 18;
                    operationId = 1;
                    subOperationId = 2;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion

                    #region nasb Kanalet

                    //نصب کانالت

                    //اجرا 
                    // مترطول
                    operationId = 2;
                    subOperationId = 1;
                    margeRowId = 19;
                    rowId = 19;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    // هکتار
                    rowId = 20;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //تامین سهم پیمانکار
                    operationId = 2;
                    subOperationId = 4;
                    rowId = 21;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //جاده سرویس  
                    operationId = 2;
                    subOperationId = 2;
                    rowId = 22;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //  سازه
                    operationId = 2;
                    subOperationId = 3;
                    rowId = 23;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion

                    #region lole kamfeshar

                    //لوله کم فشار

                    //اجرا 
                    //مترطول

                    operationId = 3;
                    subOperationId = 1;
                    margeRowId = 24;
                    rowId = 24;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //هکتار
                    operationId = 3;
                    subOperationId = 1;
                    rowId = 25;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //تامین سهم پیمانکار
                    operationId = 3;
                    subOperationId = 4;
                    rowId = 26;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //جاده سرویس  
                    operationId = 3;
                    subOperationId = 2;
                    rowId = 27;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //  سازه و اتصالات
                    operationId = 3;
                    subOperationId = 3;
                    rowId = 28;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //  اجرای ایستگاه پمپاژ و نصب تجهیزات
                    operationId = 3;
                    subOperationId = 5;
                    rowId = 29;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //  تامین سهم پیمانکار (ایستگاه پمپاژ)
                    operationId = 3;
                    subOperationId = 10;
                    rowId = 30;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "درصد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);



                    #endregion

                    #region abyari tahte feshar

                    //آبیاری تحت فشار
                    //اجرا
                    //مترطول
                    operationId = 4;
                    subOperationId = 1;
                    margeRowId = 31;
                    rowId = 31;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //هکتار
                    subOperationId = 1;
                    rowId = 32;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //تامین سهم پیمانکار
                    subOperationId = 4;
                    rowId = 33;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    // جاده سرویس 
                    subOperationId = 2;
                    rowId = 34;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //سازه
                    //operationId = 4;
                    subOperationId = 3;
                    rowId = 35;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //اجرای ایستگاه پمپاژ و نصب تجهیزات

                    subOperationId = 5;
                    rowId = 36;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    // تامین سهم پیمانکار (ایستگاه پمپاژ)                
                    subOperationId = 10;
                    rowId = 35;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "درصد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion

                    #region tameen lole

                    // تامین لوله و متعلقات واجرای خط انتقال آب
                    //اجرا 
                    //مترطول

                    operationId = 4;
                    subOperationId = 1;
                    margeRowId = 38;
                    rowId = 38;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //هکتار               
                    subOperationId = 1;
                    rowId = 39;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);



                    //تامین سهم پیمانکار
                    subOperationId = 4;
                    rowId = 40;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //سازه

                    subOperationId = 3;
                    rowId = 41;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);



                    #endregion

                    #region zehkesh roobaz

                    //زهکش روباز
                    //اجرا 
                    //مترطول
                    operationId = 5;
                    subOperationId = 1;
                    margeRowId = 42;
                    rowId = 42;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //هکتار
                    subOperationId = 1;
                    rowId = 43;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    // جاده سرویس
                    subOperationId = 2;
                    rowId = 44;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //  سازه( )
                    subOperationId = 3;

                    rowId = 45;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion

                    #region collector

                    //کلکتور (زهکش جمع کننده لوله ای)
                    //اجرا 
                    //مترطول
                    operationId = 9;
                    subOperationId = 1;
                    margeRowId = 46;
                    rowId = 46;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //هکتار
                    subOperationId = 1;
                    rowId = 47;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //تامین سهم پیمانکار
                    subOperationId = 4;
                    rowId = 48;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    // سازه 
                    subOperationId = 3;
                    rowId = 49;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion

                    #region zehkesh zirzamini

                    //زهکش های زیرزمینی(لترال) 
                    //اجرا 
                    //مترطول
                    operationId = 6;
                    subOperationId = 1;
                    margeRowId = 50;
                    rowId = 50;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //هکتار
                    subOperationId = 1;
                    rowId = 51;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    //تامین سهم پیمانکار
                    subOperationId = 4;
                    rowId = 52;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //سازه
                    subOperationId = 3;
                    rowId = 53;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion

                    #region tajhiz o nosazi

                    //تجهیز و نوسازی
                    //تسطیح نسبی (خالص)

                    operationId = 7;
                    subOperationId = 12;
                    margeRowId = 54;
                    rowId = 54;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //تسطیح اساسی (خالص)

                    subOperationId = 13;
                    rowId = 55;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //تجهیز و نوسازی          
                    subOperationId = 6;
                    rowId = 65;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //هندسی سازی
                    subOperationId = 8;
                    rowId = 57;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //یکپارچه سازی

                    subOperationId = 9;
                    rowId = 58;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //تامین سهم پیمانکار

                    subOperationId = 4;
                    rowId = 59;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "متر طول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //کانال درجه 4

                    subOperationId = 11;
                    rowId = 60;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //سازه
                    subOperationId = 3;
                    rowId = 61;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "تعداد", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //آبشویی

                    subOperationId = 14;
                    rowId = 62;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "هکتار", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);

                    //جاده دسترسی بین مزارع

                    subOperationId = 7;
                    rowId = 63;
                    value = ws.Cells[rowId, valueCell].Value;
                    totalValue = ws.Cells[rowId, totalValueCell].Value;
                    amount = ws.Cells[rowId, amountCell].Value;
                    totalAmount = ws.Cells[rowId, totalAmountCell].Value;
                    activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
                    eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[rowId, 17].Value);


                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                        UpsertFarmOperation(farmList, context,
                                            cId, block, farm, operationId, subOperationId, "مترطول", value, totalValue, amount, totalAmount,
                                            ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value, eqHectar);


                    #endregion
                }
                //  catch { }
                context.ExecuteQuery();
            }

            MessageBox.Show("jfhgfhgfh");
        }
        public void CreateOtherItems2(Workbook wb, Web web, ClientContext context)
        {
            List farmList = web.Lists.GetByTitle("عملیات اجرایی در سطح واحد زراعی");


            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            //  var contract = p.AfterProperties["Contract"];
            //  int i = 2;
            // for (var i = 3; i < wb.Worksheets.Count; i++)
            {
                //try
                {
                    var ws = wb.Worksheets["PS1-LMC1"];
                    var cId = 103;//new SPFieldLookupValue(contract.ToString()).LookupId;
                    var blockName = ws.Cells[7, 2].Value.ToString();
                    var block = GetBlockId(cId, blockName, context);
                    var impureArea = Convert.ToDouble(ws.Cells[7, 4].Value.ToString());
                    var farmName = ws.Cells[7, 3].Value.ToString();
                    var farm = GetFarmId(cId, block, farmName, impureArea, context);

                    #region Kanal Daraje 3

                    // کانال درجه 3( درجا)
                    //اجرا
                    var operationId = 1;
                    var subOperationId = 2;
                    var value = ws.Cells[15, 5].Value;
                    var totalValue = ws.Cells[15, 7].Value;
                    var amount = ws.Cells[15, 8].Value;
                    var totalAmount = ws.Cells[15, 9].Value;
                    var activityType = ws.Cells[15, 10].Value.ToString();
                    var eqHectar = 0.0;
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[15, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[15, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[15, 17].Value);

                    // مترطول
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[15, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[15, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[15, 13].Value;
                        item["WeightAction"] = ws.Cells[15, 14].Value;
                        item.Update();
                    }
                    // هکتار
                    value = ws.Cells[16, 5].Value;
                    totalValue = ws.Cells[16, 7].Value;
                    amount = ws.Cells[16, 8].Value;
                    totalAmount = ws.Cells[16, 9].Value;
                    activityType = ws.Cells[16, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[16, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[16, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[16, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[16, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[16, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[15, 13].Value;
                        item["WeightAction"] = ws.Cells[15, 14].Value;
                        item.Update();
                    }
                    // جاده سرویس  
                    operationId = 1;
                    subOperationId = 3;
                    value = ws.Cells[17, 5].Value;
                    totalValue = ws.Cells[17, 7].Value;
                    amount = ws.Cells[17, 8].Value;
                    totalAmount = ws.Cells[17, 9].Value;
                    activityType = ws.Cells[17, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[17, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[17, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[17, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                         Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[17, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[17, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[15, 13].Value;
                        item["WeightAction"] = ws.Cells[15, 14].Value;
                        item.Update();
                    }
                    //  سازه
                    operationId = 1;
                    subOperationId = 4;
                    value = ws.Cells[18, 5].Value;
                    totalValue = ws.Cells[18, 7].Value;
                    amount = ws.Cells[18, 8].Value;
                    totalAmount = ws.Cells[18, 9].Value;
                    activityType = ws.Cells[18, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[18, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[18, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[18, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[18, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[18, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[15, 13].Value;
                        item["WeightAction"] = ws.Cells[15, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region nasb Kanalet

                    //نصب کانالت

                    //اجرا 
                    // مترطول
                    operationId = 2;
                    subOperationId = 2;
                    value = ws.Cells[19, 5].Value;
                    totalValue = ws.Cells[19, 7].Value;
                    amount = ws.Cells[19, 8].Value;
                    totalAmount = ws.Cells[19, 9].Value;
                    activityType = ws.Cells[19, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[19, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[19, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[19, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[19, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[19, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[19, 13].Value;
                        item["WeightAction"] = ws.Cells[19, 14].Value;
                        item.Update();
                    }
                    // هکتار
                    value = ws.Cells[20, 5].Value;
                    totalValue = ws.Cells[20, 7].Value;
                    amount = ws.Cells[20, 8].Value;
                    totalAmount = ws.Cells[20, 9].Value;
                    activityType = ws.Cells[20, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[20, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[20, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[20, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[20, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[20, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[19, 13].Value;
                        item["WeightAction"] = ws.Cells[19, 14].Value;
                        item.Update();
                    }

                    //تامین سهم پیمانکار
                    operationId = 2;
                    subOperationId = 5;
                    value = ws.Cells[21, 5].Value;
                    totalValue = ws.Cells[21, 7].Value;
                    amount = ws.Cells[21, 8].Value;
                    totalAmount = ws.Cells[21, 9].Value;
                    activityType = ws.Cells[21, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[21, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[21, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[21, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[21, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[21, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[19, 13].Value;
                        item["WeightAction"] = ws.Cells[19, 14].Value;
                        item.Update();
                    }

                    //جاده سرویس  
                    operationId = 2;
                    subOperationId = 3;
                    value = ws.Cells[22, 5].Value;
                    totalValue = ws.Cells[22, 7].Value;
                    amount = ws.Cells[22, 8].Value;
                    totalAmount = ws.Cells[22, 9].Value;
                    activityType = ws.Cells[22, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[22, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[22, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[22, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[22, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[22, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[19, 13].Value;
                        item["WeightAction"] = ws.Cells[19, 14].Value;
                        item.Update();
                    }
                    //  سازه
                    operationId = 2;
                    subOperationId = 4;
                    value = ws.Cells[23, 5].Value;
                    totalValue = ws.Cells[23, 7].Value;
                    amount = ws.Cells[23, 8].Value;
                    totalAmount = ws.Cells[23, 9].Value;
                    activityType = ws.Cells[23, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[23, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[23, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[23, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[23, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[23, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[19, 13].Value;
                        item["WeightAction"] = ws.Cells[19, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region lole kamfeshar

                    //لوله کم فشار

                    //اجرا 
                    //مترطول

                    operationId = 3;
                    subOperationId = 2;
                    value = ws.Cells[24, 5].Value;
                    totalValue = ws.Cells[24, 7].Value;
                    amount = ws.Cells[24, 8].Value;
                    totalAmount = ws.Cells[24, 9].Value;
                    activityType = ws.Cells[24, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[24, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[24, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[24, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[24, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[24, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }

                    //هکتار
                    operationId = 3;
                    subOperationId = 2;
                    value = ws.Cells[25, 5].Value;
                    totalValue = ws.Cells[25, 7].Value;
                    amount = ws.Cells[25, 8].Value;
                    totalAmount = ws.Cells[25, 9].Value;
                    activityType = ws.Cells[25, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[25, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[25, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[25, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[25, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[25, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }

                    //تامین سهم پیمانکار
                    operationId = 3;
                    subOperationId = 5;
                    value = ws.Cells[26, 5].Value;
                    totalValue = ws.Cells[26, 7].Value;
                    amount = ws.Cells[26, 8].Value;
                    totalAmount = ws.Cells[26, 9].Value;
                    activityType = ws.Cells[26, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[26, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[26, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[26, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[26, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[26, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }
                    //جاده سرویس  
                    operationId = 3;
                    subOperationId = 3;
                    value = ws.Cells[27, 5].Value;
                    totalValue = ws.Cells[27, 7].Value;
                    amount = ws.Cells[27, 8].Value;
                    totalAmount = ws.Cells[27, 9].Value;
                    activityType = ws.Cells[27, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[27, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[27, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[27, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[27, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[27, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }
                    //  سازه و اتصالات
                    operationId = 3;
                    subOperationId = 12;
                    value = ws.Cells[28, 5].Value;
                    totalValue = ws.Cells[28, 7].Value;
                    amount = ws.Cells[28, 8].Value;
                    totalAmount = ws.Cells[28, 9].Value;
                    activityType = ws.Cells[28, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[28, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[28, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[28, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[28, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[28, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }
                    //  اجرای ایستگاه پمپاژ و نصب تجهیزات
                    operationId = 3;
                    subOperationId = 15;
                    value = ws.Cells[29, 5].Value;
                    totalValue = ws.Cells[29, 7].Value;
                    amount = ws.Cells[29, 8].Value;
                    totalAmount = ws.Cells[29, 9].Value;
                    activityType = ws.Cells[29, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[29, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[29, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[29, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[29, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[29, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }
                    //  تامین سهم پیمانکار (ایستگاه پمپاژ)
                    operationId = 3;
                    subOperationId = 41;
                    value = ws.Cells[30, 5].Value;
                    totalValue = ws.Cells[30, 7].Value;
                    amount = ws.Cells[30, 8].Value;
                    totalAmount = ws.Cells[30, 9].Value;
                    activityType = ws.Cells[30, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[30, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[30, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[30, 17].Value);

                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "درصد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[29, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[29, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[24, 13].Value;
                        item["WeightAction"] = ws.Cells[24, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region abyari tahte feshar

                    //آبیاری تحت فشار
                    //اجرا
                    //مترطول
                    operationId = 4;
                    subOperationId = 2;
                    value = ws.Cells[31, 5].Value;
                    totalValue = ws.Cells[31, 7].Value;
                    amount = ws.Cells[31, 8].Value;
                    totalAmount = ws.Cells[31, 9].Value;
                    activityType = ws.Cells[31, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[31, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[31, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[31, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[31, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[31, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }
                    //هکتار
                    subOperationId = 2;
                    value = ws.Cells[32, 5].Value;
                    totalValue = ws.Cells[32, 7].Value;
                    amount = ws.Cells[32, 8].Value;
                    totalAmount = ws.Cells[32, 9].Value;
                    activityType = ws.Cells[32, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[32, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[32, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[32, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[32, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[32, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }
                    //تامین سهم پیمانکار
                    subOperationId = 5;
                    value = ws.Cells[33, 5].Value;
                    totalValue = ws.Cells[33, 7].Value;
                    amount = ws.Cells[33, 8].Value;
                    totalAmount = ws.Cells[33, 9].Value;
                    activityType = ws.Cells[33, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[33, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[33, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[33, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[33, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[33, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }
                    // جاده سرویس 
                    subOperationId = 3;
                    value = ws.Cells[34, 5].Value;
                    totalValue = ws.Cells[34, 7].Value;
                    amount = ws.Cells[34, 8].Value;
                    totalAmount = ws.Cells[34, 9].Value;
                    activityType = ws.Cells[34, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[34, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[34, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[34, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[34, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[34, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }
                    //سازه
                    operationId = 4;
                    subOperationId = 4;
                    value = ws.Cells[35, 5].Value;
                    totalValue = ws.Cells[35, 7].Value;
                    amount = ws.Cells[35, 8].Value;
                    totalAmount = ws.Cells[35, 9].Value;
                    activityType = ws.Cells[35, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[35, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[35, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[35, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[35, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[35, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }
                    //اجرای ایستگاه پمپاژ و نصب تجهیزات

                    subOperationId = 15;
                    value = ws.Cells[36, 5].Value;
                    totalValue = ws.Cells[36, 7].Value;
                    amount = ws.Cells[36, 8].Value;
                    totalAmount = ws.Cells[36, 9].Value;
                    activityType = ws.Cells[36, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[36, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[36, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[36, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[36, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[36, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }
                    // تامین سهم پیمانکار (ایستگاه پمپاژ)                
                    subOperationId = 41;
                    value = ws.Cells[37, 5].Value;
                    totalValue = ws.Cells[37, 7].Value;
                    amount = ws.Cells[37, 8].Value;
                    totalAmount = ws.Cells[37, 9].Value;
                    activityType = ws.Cells[37, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[37, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[37, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[37, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "درصد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[37, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[37, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[31, 13].Value;
                        item["WeightAction"] = ws.Cells[31, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region tameen lole

                    // تامین لوله و متعلقات واجرای خط انتقال آب
                    //اجرا 
                    //مترطول

                    operationId = 8;
                    subOperationId = 2;
                    value = ws.Cells[38, 5].Value;
                    totalValue = ws.Cells[38, 7].Value;
                    amount = ws.Cells[38, 8].Value;
                    totalAmount = ws.Cells[38, 9].Value;
                    activityType = ws.Cells[38, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[38, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[38, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[38, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[38, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[38, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[38, 13].Value;
                        item["WeightAction"] = ws.Cells[38, 14].Value;
                        item.Update();
                    }
                    //هکتار               
                    subOperationId = 2;
                    value = ws.Cells[39, 5].Value;
                    totalValue = ws.Cells[39, 7].Value;
                    amount = ws.Cells[39, 8].Value;
                    totalAmount = ws.Cells[39, 9].Value;
                    activityType = ws.Cells[39, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[39, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[39, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[39, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[39, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[39, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[39, 13].Value;
                        item["WeightAction"] = ws.Cells[39, 14].Value;
                        item.Update();
                    }

                    //تامین سهم پیمانکار
                    subOperationId = 5;
                    value = ws.Cells[40, 5].Value;
                    totalValue = ws.Cells[40, 7].Value;
                    amount = ws.Cells[40, 8].Value;
                    totalAmount = ws.Cells[40, 9].Value;
                    activityType = ws.Cells[40, 10].Value.ToString();

                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[40, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[40, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[40, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[40, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[40, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[39, 13].Value;
                        item["WeightAction"] = ws.Cells[39, 14].Value;
                        item.Update();

                    }
                    //سازه

                    subOperationId = 4;
                    value = ws.Cells[41, 5].Value;
                    totalValue = ws.Cells[41, 7].Value;
                    amount = ws.Cells[41, 8].Value;
                    totalAmount = ws.Cells[41, 9].Value;
                    activityType = ws.Cells[41, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[41, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[41, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[41, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[41, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[41, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[39, 13].Value;
                        item["WeightAction"] = ws.Cells[39, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region zehkesh roobaz

                    //زهکش روباز
                    //اجرا 
                    //مترطول
                    operationId = 5;
                    subOperationId = 2;
                    value = ws.Cells[42, 5].Value;
                    totalValue = ws.Cells[42, 7].Value;
                    amount = ws.Cells[42, 8].Value;
                    totalAmount = ws.Cells[42, 9].Value;
                    activityType = ws.Cells[42, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[42, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[42, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[42, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[42, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[42, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[42, 13].Value;
                        item["WeightAction"] = ws.Cells[42, 14].Value;
                        item.Update();
                    }
                    //هکتار
                    subOperationId = 2;
                    value = ws.Cells[43, 5].Value;
                    totalValue = ws.Cells[43, 7].Value;
                    amount = ws.Cells[43, 8].Value;
                    totalAmount = ws.Cells[43, 9].Value;
                    activityType = ws.Cells[43, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[43, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[43, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[43, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[43, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[43, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[42, 13].Value;
                        item["WeightAction"] = ws.Cells[42, 14].Value;
                        item.Update();
                    }
                    // جاده سرویس
                    subOperationId = 3;
                    value = ws.Cells[44, 5].Value;
                    totalValue = ws.Cells[44, 7].Value;
                    amount = ws.Cells[44, 8].Value;
                    totalAmount = ws.Cells[44, 9].Value;
                    activityType = ws.Cells[44, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[44, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[44, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[44, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[44, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[44, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[42, 13].Value;
                        item["WeightAction"] = ws.Cells[42, 14].Value;
                        item.Update();
                    }
                    //  سازه( )
                    subOperationId = 4;
                    value = ws.Cells[45, 5].Value;
                    totalValue = ws.Cells[45, 7].Value;
                    amount = ws.Cells[45, 8].Value;
                    totalAmount = ws.Cells[45, 9].Value;
                    activityType = ws.Cells[45, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[45, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[45, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[45, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[45, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[45, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[42, 13].Value;
                        item["WeightAction"] = ws.Cells[42, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region collector

                    //کلکتور (زهکش جمع کننده لوله ای)
                    //اجرا 
                    //مترطول
                    operationId = 12;
                    subOperationId = 2;
                    value = ws.Cells[46, 5].Value;
                    totalValue = ws.Cells[46, 7].Value;
                    amount = ws.Cells[46, 8].Value;
                    totalAmount = ws.Cells[46, 9].Value;
                    activityType = ws.Cells[46, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[46, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[46, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[46, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[46, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[46, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[46, 13].Value;
                        item["WeightAction"] = ws.Cells[46, 14].Value;
                        item.Update();
                    }
                    //هکتار
                    subOperationId = 2;
                    value = ws.Cells[47, 5].Value;
                    totalValue = ws.Cells[47, 7].Value;
                    amount = ws.Cells[47, 8].Value;
                    totalAmount = ws.Cells[47, 9].Value;
                    activityType = ws.Cells[47, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[47, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[47, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[47, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[47, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[47, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[46, 13].Value;
                        item["WeightAction"] = ws.Cells[46, 14].Value;
                        item.Update();
                    }
                    //تامین سهم پیمانکار
                    subOperationId = 5;
                    value = ws.Cells[48, 5].Value;
                    totalValue = ws.Cells[48, 7].Value;
                    amount = ws.Cells[48, 8].Value;
                    totalAmount = ws.Cells[48, 9].Value;
                    activityType = ws.Cells[48, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[48, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[48, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[48, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[48, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[48, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[46, 13].Value;
                        item["WeightAction"] = ws.Cells[46, 14].Value;
                        item.Update();
                    }

                    // سازه 
                    subOperationId = 4;
                    value = ws.Cells[49, 5].Value;
                    totalValue = ws.Cells[49, 7].Value;
                    amount = ws.Cells[49, 8].Value;
                    totalAmount = ws.Cells[49, 9].Value;
                    activityType = ws.Cells[49, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[49, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[49, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[49, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[49, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[49, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[46, 13].Value;
                        item["WeightAction"] = ws.Cells[46, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region zehkesh zirzamini

                    //زهکش های زیرزمینی(لترال) 
                    //اجرا 
                    //مترطول
                    operationId = 6;
                    subOperationId = 2;
                    value = ws.Cells[50, 5].Value;
                    totalValue = ws.Cells[50, 7].Value;
                    amount = ws.Cells[50, 8].Value;
                    totalAmount = ws.Cells[50, 9].Value;
                    activityType = ws.Cells[50, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[50, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[50, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[50, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[50, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[50, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[50, 13].Value;
                        item["WeightAction"] = ws.Cells[50, 14].Value;
                        item.Update();
                    }
                    //هکتار
                    subOperationId = 2;
                    value = ws.Cells[51, 5].Value;
                    totalValue = ws.Cells[51, 7].Value;
                    amount = ws.Cells[51, 8].Value;
                    totalAmount = ws.Cells[51, 9].Value;
                    activityType = ws.Cells[51, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[51, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[51, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[51, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[51, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[51, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[50, 13].Value;
                        item["WeightAction"] = ws.Cells[50, 14].Value;
                        item.Update();
                    }

                    //تامین سهم پیمانکار
                    subOperationId = 5;
                    value = ws.Cells[52, 5].Value;
                    totalValue = ws.Cells[52, 7].Value;
                    amount = ws.Cells[52, 8].Value;
                    totalAmount = ws.Cells[52, 9].Value;
                    activityType = ws.Cells[52, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[52, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[52, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[52, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[52, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[52, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[50, 13].Value;
                        item["WeightAction"] = ws.Cells[50, 14].Value;
                        item.Update();
                    }
                    //سازه
                    subOperationId = 4;
                    value = ws.Cells[53, 5].Value;
                    totalValue = ws.Cells[53, 7].Value;
                    amount = ws.Cells[53, 8].Value;
                    totalAmount = ws.Cells[53, 9].Value;
                    activityType = ws.Cells[53, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[53, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[53, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[53, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[53, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[53, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[50, 13].Value;
                        item["WeightAction"] = ws.Cells[50, 14].Value;
                        item.Update();
                    }

                    #endregion

                    #region tajhiz o nosazi

                    //تجهیز و نوسازی
                    //تسطیح نسبی (خالص)

                    operationId = 7;
                    subOperationId = 44;
                    value = ws.Cells[54, 5].Value;
                    totalValue = ws.Cells[54, 7].Value;
                    amount = ws.Cells[54, 8].Value;
                    totalAmount = ws.Cells[54, 9].Value;
                    activityType = ws.Cells[54, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[54, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[54, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[54, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[54, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[54, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //تسطیح اساسی (خالص)

                    subOperationId = 45;
                    value = ws.Cells[55, 5].Value;
                    totalValue = ws.Cells[55, 7].Value;
                    amount = ws.Cells[55, 8].Value;
                    totalAmount = ws.Cells[55, 9].Value;
                    activityType = ws.Cells[55, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[55, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[55, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[55, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[55, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[55, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //تجهیز و نوسازی          
                    subOperationId = 27;
                    value = ws.Cells[56, 5].Value;
                    totalValue = ws.Cells[56, 7].Value;
                    amount = ws.Cells[56, 8].Value;
                    totalAmount = ws.Cells[56, 9].Value;
                    activityType = ws.Cells[56, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[56, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[56, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[56, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[56, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[56, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //هندسی سازی
                    subOperationId = 39;
                    value = ws.Cells[57, 5].Value;
                    totalValue = ws.Cells[57, 7].Value;
                    amount = ws.Cells[57, 8].Value;
                    totalAmount = ws.Cells[57, 9].Value;
                    activityType = ws.Cells[57, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[57, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[57, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[57, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[57, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[57, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //یکپارچه سازی

                    subOperationId = 40;
                    value = ws.Cells[58, 5].Value;
                    totalValue = ws.Cells[58, 7].Value;
                    amount = ws.Cells[58, 8].Value;
                    totalAmount = ws.Cells[58, 9].Value;
                    activityType = ws.Cells[58, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[58, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[58, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[58, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[58, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[58, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //تامین سهم پیمانکار

                    subOperationId = 5;
                    value = ws.Cells[59, 5].Value;
                    totalValue = ws.Cells[59, 7].Value;
                    amount = ws.Cells[59, 8].Value;
                    totalAmount = ws.Cells[59, 9].Value;
                    activityType = ws.Cells[59, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[59, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[59, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[59, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[59, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[59, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //کانال درجه 4

                    subOperationId = 43;
                    value = ws.Cells[60, 5].Value;
                    totalValue = ws.Cells[60, 7].Value;
                    amount = ws.Cells[60, 8].Value;
                    totalAmount = ws.Cells[60, 9].Value;
                    activityType = ws.Cells[60, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[60, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[60, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[60, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[60, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[60, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //سازه
                    subOperationId = 4;
                    value = ws.Cells[61, 5].Value;
                    totalValue = ws.Cells[61, 7].Value;
                    amount = ws.Cells[61, 8].Value;
                    totalAmount = ws.Cells[61, 9].Value;
                    activityType = ws.Cells[61, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[61, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[61, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[61, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "تعداد";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[61, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[61, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //آبشویی

                    subOperationId = 47;
                    value = ws.Cells[62, 5].Value;
                    totalValue = ws.Cells[62, 7].Value;
                    amount = ws.Cells[62, 8].Value;
                    totalAmount = ws.Cells[62, 9].Value;
                    activityType = ws.Cells[62, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[62, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[62, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[62, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "هکتار";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[62, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[62, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }
                    //جاده دسترسی بین مزارع

                    subOperationId = 28;
                    value = ws.Cells[63, 5].Value;
                    totalValue = ws.Cells[63, 7].Value;
                    amount = ws.Cells[63, 8].Value;
                    totalAmount = ws.Cells[63, 9].Value;
                    activityType = ws.Cells[63, 10].Value.ToString();
                    if (activityType == "شبکه")
                        eqHectar = Convert.ToDouble(ws.Cells[63, 15].Value);
                    else if (activityType == "زهکش زیر زمینی")
                        eqHectar = Convert.ToDouble(ws.Cells[63, 16].Value);
                    else if (activityType == "تجهیز و نوسازی")
                        eqHectar = Convert.ToDouble(ws.Cells[63, 17].Value);
                    if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                        Convert.ToInt64(totalAmount) != 0)
                    {

                        ListItem item = farmList.AddItem(itemCreateInfo);
                        item["Contract"] = new FieldLookupValue() { LookupId = cId };
                        item["Block"] = new FieldLookupValue() { LookupId = block };
                        item["Farm"] = new FieldLookupValue() { LookupId = farm };
                        item["ExecutiveOperation"] = new FieldLookupValue() { LookupId = operationId };
                        item["SubExecutiveOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                        item["OrgValue"] = value;
                        item["ChangeValue"] = totalValue;
                        item["Amount"] = amount;
                        item["ChangeAmount"] = totalAmount;
                        item["Measurement"] = "مترطول";
                        item["EqHectar"] = eqHectar;
                        item["TotalWeightContract"] = ws.Cells[63, 11].Value;
                        item["ItemWeightAction"] = ws.Cells[63, 12].Value;
                        item["TotalWeightOperation"] = ws.Cells[54, 13].Value;
                        item["WeightAction"] = ws.Cells[54, 14].Value;
                        item.Update();
                    }

                    #endregion
                }
                //  catch { }
                context.ExecuteQuery();
            }

            MessageBox.Show("jfhgfhgfh");
        }

        public int GetBlockId(int cId, string blockName, ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("Blocks");
            var q = new CamlQuery();
            var r = 0;
            q.ViewXml = @"<View><Query><Where><And>
                  <Eq>
                     <FieldRef Name = 'Contract' LookupId = 'TRUE' />
                     <Value Type = 'integer'>" + cId + @"</Value>
                  </Eq>
                   <Eq>
                     <FieldRef Name = 'Title'/>
                     <Value Type = 'Text'>" + blockName + @"</Value>
                  </Eq>
                </And>
               </Where></Query></View>";

            var result = list.GetItems(q);
            context.Load(result);
            context.ExecuteQuery();
            if (result.Count > 0)
                return result[0].Id;
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem item = list.AddItem(itemCreateInfo);
            item["Title"] = blockName;
            item["Contract"] = new FieldLookupValue() { LookupId = cId };
            item.Update();
            context.ExecuteQuery();
            r = item.Id;
            return r;
        }

        public int GetFarmId(int cId, int bId, string farmName, double impureArea, ClientContext context)
        {
            var list = context.Web.Lists.GetByTitle("واحد زراعی");
            var q = new CamlQuery();
            q.ViewXml = string.Format(@"<View><Query><Where><And><Eq><FieldRef Name='Contract' LookupId = 'TRUE'/><Value Type='Lookup'>{0}</Value>
                                                      </Eq><And><Eq><FieldRef Name='Block' LookupId = 'TRUE'/><Value Type='Lookup'>{1}</Value>
                                                           </Eq><Eq><FieldRef Name='Title' /><Value Type='Text'>{2}</Value></Eq></And></And></Where></Query></View>", cId, bId, farmName);
            //q.ViewXml = @"<View> <Eq>
            //        <FieldRef Name = 'Title' />
            //         <Value Type = 'text'>" + bId + @"</Value>
            //     </Eq></View>";
            var result = list.GetItems(q);
            context.Load(result);
            context.ExecuteQuery();
            if (result.Count > 0)
                return result[0].Id;
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem item = list.AddItem(itemCreateInfo);
            item["Title"] = farmName;
            item["Block"] = new FieldLookupValue() { LookupId = bId };
            item["Contract"] = new FieldLookupValue() { LookupId = cId };
            item["ImpureArea"] = impureArea;
            item.Update();
            context.ExecuteQuery();
            return item.Id;
        }

        public ListItem UpsertFarmOperation(List list, ClientContext context,
                                         int cId, int bId, int fId, int operationId, int subOperationId, string measurement,
                                         object value, object totalValue, object amount, object totalAmount, object totalWeightContract, object itemWeightAction, object totalWeightOperation, object weightAction, object eqHectar)
        {



            var q = new CamlQuery();
            q.ViewXml = string.Format(@"<View><Query><Where>
                                                        <And><Eq><FieldRef Name='Contract' LookupId='True'/><Value Type='integer' >{0}</Value></Eq>
                                                        <And><Eq><FieldRef Name='Block' LookupId='True'/><Value Type='integer'>{1}</Value></Eq>
                                                        <And><Eq><FieldRef Name='Farm' LookupId='True'/><Value Type='integer'>{2}</Value></Eq>
                                                        <And><Eq><FieldRef Name='Operation' LookupId='True' /><Value Type='integer'>{3}</Value></Eq>
                                                        <And><Eq><FieldRef Name='SubOperation' LookupId='True'/><Value Type='integer'>{4}</Value></Eq>
                                                        <Eq><FieldRef Name='Measurement' /><Value Type='Choice'>{5}</Value></Eq>
                                    </And></And></And></And></And></Where></Query></View>", cId, bId, fId, operationId, subOperationId, measurement);

            var result = list.GetItems(q);
            context.Load(result);
            context.ExecuteQuery();

            ListItem item;
            if (result.Count > 0)
                item = result[0];
            else
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                item = list.AddItem(itemCreateInfo);
                item["Contract"] = new FieldLookupValue() { LookupId = cId };
                item["Block"] = new FieldLookupValue() { LookupId = bId };
                item["Farm"] = new FieldLookupValue() { LookupId = fId };
                item["Operation"] = new FieldLookupValue() { LookupId = operationId };
                item["SubOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                item["Measurement"] = measurement;
            }

            item["FirstVolume"] = value;
            item["FinalVolume"] = totalValue;
            item["FirstCost"] = amount;
            item["FinalCost"] = totalAmount;
            item["EqAcre"] = eqHectar;
            item["ItemWeight"] = totalWeightContract;
            item["ItemWeightOperation"] = itemWeightAction;
            item["TotalItemWeight"] = totalWeightOperation;
            item["TotalWeightOperation"] = weightAction;
            item.Update();
            context.ExecuteQuery();
            return item;
        }

        private void button3_Click(object sender, EventArgs e)
        {


            //ClientContext context = new ClientContext("http://172.29.0.162:90");
            ClientContext context = new ClientContext("http://net-sp:90");
            context.Credentials = new NetworkCredential("spadmin", "Nsr!dm$n!Sp", "nasr2");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = context.Web.Lists.GetByTitle("فرم پرداخت صورت حساب");
          //  0x01008046713216209B448C543CF3465B3CF4
                                                       
            ContentType ct = list.ContentTypes.GetById("0x01009ABAD5506DED664AB72C5C75259F2FA7");
            
            {
                ct.NewFormUrl = "/Lists/PaymentInvoice/NewForm.aspx";
              //  ct.NewFormUrl = "/lists//Lists/AgreementsAbgostrans/EditForm/index.html";
                ct.EditFormUrl = "/Lists/PaymentInvoice/EditForm.aspx";
                ct.DisplayFormUrl = "/Lists/PaymentInvoice/DispForm.aspx";

                ct.Update(false);
            }

            list.Update();
            context.ExecuteQuery();
            MessageBox.Show("done!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/implementation/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            //List list = context.Web.Lists.GetByTitle("فرم وضعیت رفع نقص");
            List list = context.Web.Lists.GetByTitle("فرم اعلام نقص");
            var query = new CamlQuery();

            query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query></View>");
            context.Load(list);
            context.Load(list.RootFolder);
            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();


            foreach (ListItem item in items)
            {
                // context.Load(item);
                context.Load(item.File);
                context.ExecuteQuery();
                item.File.MoveTo(list.RootFolder.ServerRelativeUrl + "/" + item["CId"] + "/" + item.File.Name, MoveOperations.Overwrite);
                // item.File.MoveTo(list.RootFolder.ServerRelativeUrl + "/" + item["__x007b_974d1659_5791_4ec2_8046_4c77d004973e_x007d_"] + "/" + item.File.Name, MoveOperations.Overwrite);
                item.ResetRoleInheritance();
                context.ExecuteQuery();
            }


            MessageBox.Show("done!" + items.Count);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);

            List list = context.Web.Lists.GetByTitle("صورتجلسات تحویل موقت");
            //     var query = new CamlQuery();
            //    query.ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query><Where><Eq> <FieldRef Name='FSObjType' /><Value Type = 'Integer'>0</Value></Eq></Where></Query></View>");
            //  var items = list.GetItems(query);
            // context.Load(items);
            //  context.ExecuteQuery();

            //foreach (ListItem item in items)
            //{
            //    item.ResetRoleInheritance();
            //}
            context.Load(list, ls => ls.RootFolder.Folders);

            context.ExecuteQuery();

            var folder = list.RootFolder.Folders.FirstOrDefault(x => x.Name == textBox1.Text);

            Group g1 = web.SiteGroups.GetByName("گروه مدیریت راهبری");//گروه مدیریت راهبری 
            Group g2 = web.SiteGroups.GetByName("مشاهده کنندگان اسناد");//گروه مدیریت راهبری 

            hnUser hnU = getUsers(68);
            User u1 = web.EnsureUser(hnU.advaisor), u2 = web.EnsureUser(hnU.areaManager), u3 = web.EnsureUser(hnU.contractor), u4 = web.EnsureUser(hnU.manager);


            RoleDefinitionBindingCollection collRoleDefinitionBindingRead = new RoleDefinitionBindingCollection(context);
            collRoleDefinitionBindingRead.Add(context.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
            RoleDefinitionBindingCollection collRoleDefinitionBindingCont = new RoleDefinitionBindingCollection(context);
            collRoleDefinitionBindingCont.Add(context.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type
            RoleDefinitionBindingCollection collRoleDefinitionBindingfull = new RoleDefinitionBindingCollection(context);
            collRoleDefinitionBindingfull.Add(context.Web.RoleDefinitions.GetByType(RoleType.Administrator)); //Set permission type

            folder.ListItemAllFields.RoleAssignments.Add(g1, collRoleDefinitionBindingCont);
            folder.ListItemAllFields.RoleAssignments.Add(u1, collRoleDefinitionBindingCont);
            folder.ListItemAllFields.RoleAssignments.Add(u2, collRoleDefinitionBindingCont);
            folder.ListItemAllFields.RoleAssignments.Add(g2, collRoleDefinitionBindingRead);
            folder.ListItemAllFields.RoleAssignments.Add(u4, collRoleDefinitionBindingCont);
            folder.ListItemAllFields.RoleAssignments.Add(web.EnsureUser("nasr\\sp_admin"), collRoleDefinitionBindingCont);

            //    folder.Update();
            context.ExecuteQuery();
            MessageBox.Show("done");
        }

        private hnUser getUsers(int cId)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/ProjectsInfo/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web);

            List list = context.Web.Lists.GetByTitle("پیمان ها و قراردادها");
            List list2 = context.Web.Lists.GetByTitle("حوزه مدیریتی");
            List list3 = context.Web.Lists.GetByTitle("کاربران شرکت ها");

            ListItem item = list.GetItemById(cId);
            context.Load(item);
            context.ExecuteQuery();
            ListItem item2 = list2.GetItemById((item["MangmentZone"] as FieldLookupValue).LookupId);
            context.Load(item2);
            var users = list3.GetItems(new CamlQuery() { ViewXml = "<View></View>" });

            context.Load(users);
            context.ExecuteQuery();

            //var  c = users.GetById((item["ContractorUser"] as FieldLookupValue).LookupId)["UserName"],
            //  m = (item["ProjectManagerUser"] as FieldLookupValue).LookupValue,
            //  a = (item["AdvisorUser"] as FieldLookupValue).LookupValue


            return new hnUser()
            {
                contractor = "jnasr\\" + (item["ContractorUser"] as FieldLookupValue).LookupValue,
                manager = "jnasr\\" + (item["ProjectManagerUser"] as FieldLookupValue).LookupValue,
                advaisor = "jnasr\\" + (item["AdvisorUser"] as FieldLookupValue).LookupValue,
                areaManager = (item2["AMUsername"] as FieldLookupValue).LookupValue
                //   contractor = users.GetById((item["ContractorUser"] as FieldLookupValue).LookupId)["UserName"].ToString(),
                //    manager = users.GetById((item["ProjectManagerUser"] as FieldLookupValue).LookupId)["UserName"].ToString(),
                //    advaisor = users.GetById((item["AdvisorUser"] as FieldLookupValue).LookupId)["UserName"].ToString(),
                //    areaManager = (item2["AMUsername"] as FieldLookupValue).LookupValue
            };

        }

        class hnUser
        {
            public string contractor { get; set; }
            public string advaisor { get; set; }
            public string manager { get; set; }
            public string areaManager { get; set; }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/implementation/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);

            List list = context.Web.Lists.GetByTitle("tahvil");


            context.Load(list, ls => ls.RootFolder.Folders);
            context.ExecuteQuery();

            foreach (Folder f in list.RootFolder.Folders)
            {
                if (f.Name != "127")
                {


                    string folderName = f.Name;// "69";

                    hnUser hnU = getUsers(int.Parse(folderName));
                    //User u1 = web.EnsureUser(hnU.advaisor), u2 = web.EnsureUser(hnU.areaManager), u3 = web.EnsureUser(hnU.contractor), u4 = web.EnsureUser(hnU.manager);
                    SetPermission(new string[] { hnU.advaisor, hnU.areaManager, hnU.manager }, folderName, "18EC0CC0-FE21-433F-9FEF-25928B9723D5");
                }
            }


            MessageBox.Show("done!");
        }

        public void SetPermission(string[] accountNames, string folderName, string listGuid, string roleDefinitions = "Contribute without Delete")
        {
            ClientContext ctx = new ClientContext("http://pmis.jnasr.com/implementation/");
            ctx.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            try
            {


                {
                    var list = ctx.Web.Lists.GetById(new Guid(listGuid));
                    ctx.Load(list, ls => ls.RootFolder.Folders);
                    ctx.ExecuteQuery();
                    var folder = list.RootFolder.Folders.FirstOrDefault(x => x.Name == folderName);
                    // var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);  //get Reader role
                    var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByName(roleDefinitions);
                    var roleBindings = new RoleDefinitionBindingCollection(ctx) { roleDefinition };
                    folder.ListItemAllFields.BreakRoleInheritance(false, false);  //set folder unique permissions
                    foreach (var acc in accountNames)
                    {
                        Principal user = ctx.Web.EnsureUser(acc);
                        folder.ListItemAllFields.RoleAssignments.Add(user, roleBindings);

                    }
                    ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                errorMsg += (ex.Message + "          " + folderName + "               " + "\r\n");
            }
        }
        public void SetForderPermissionByGroup(int[] groupId, string folderName, string listGuid, string roleDefinitions = "Contribute without Delete")
        {


            ClientContext ctx = new ClientContext("http://pmis.jnasr.com/implementation/");
            ctx.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            {
                var list = ctx.Web.Lists.GetById(new Guid(listGuid));
                ctx.Load(list, ls => ls.RootFolder.Folders);
                ctx.ExecuteQuery();
                var folder = list.RootFolder.Folders.FirstOrDefault(x => x.Name == folderName);

                // var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);  //get Reader role
                var roleDefinition = ctx.Site.RootWeb.RoleDefinitions.GetByName(roleDefinitions);
                var roleBindings = new RoleDefinitionBindingCollection(ctx) { roleDefinition };
                folder.ListItemAllFields.BreakRoleInheritance(false, false);
                foreach (var acc in groupId)
                {
                    var group = ctx.Web.SiteGroups.GetById(acc);

                    folder.ListItemAllFields.RoleAssignments.Add(group, roleBindings);

                }

                ctx.ExecuteQuery();
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            var i = GetFarmId(186, 1, "hghh", 5, context);
            MessageBox.Show(i.ToString());
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = web.Lists.GetByTitle("WeeklyConstructions");
            var query = new CamlQuery();
            string Iso8601Today = XmlConvert.ToString(new DateTime(2017, 10, 14), XmlDateTimeSerializationMode.Local);

            query.ViewXml = string.Format("<View><Query><Where><Lt><FieldRef Name='Period_x003a__x062a__x0627__x063' /><Value Type='Lookup' IncludeTimeValue='FALSE'>{0}</Value></Lt></Where></Query></View>", Iso8601Today);
            //query.ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name='UserName'/><Value Type='Text'>i:0#.w|nasr\\m.test4</Value></Eq></Where></Query></View>");

            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            foreach (ListItem item in items)
            {
                item["Title"] = (item["UserName"] as FieldLookupValue).LookupValue.Split('\\')[1];
                //MessageBox.Show((item["UserName"] as FieldLookupValue).LookupValue.Split('\\')[1]);
                item.Update();
                context.ExecuteQuery();
            }
            //   context.ExecuteQuery();
            MessageBox.Show(items.Count.ToString() + "      done!");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/implementation/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = web.Lists.GetByTitle("فرم اعلام نقص");

            context.Load(list, ls => ls.RootFolder.Folders);

            context.ExecuteQuery();

            foreach (Folder folder in list.RootFolder.Folders)
            {


                // var folder = list.RootFolder.Folders.FirstOrDefault(x => x.Name == textBox1.Text);
                context.Load(folder.ListItemAllFields.RoleAssignments.Groups);
                context.ExecuteQuery();
                try
                {
                    //   if (folder.ListItemAllFields.RoleAssignments.Groups.Contains(web.SiteGroups.GetByName("مشاهده کنندگان اسناد")))
                    folder.ListItemAllFields.RoleAssignments.Groups.Remove(web.SiteGroups.GetByName("مشاهده کنندگان اسناد"));
                    //   if (folder.ListItemAllFields.RoleAssignments.Groups.Contains(web.SiteGroups.GetByName("پیمانکاران")))
                    folder.ListItemAllFields.RoleAssignments.Groups.Remove(web.SiteGroups.GetByName("پیمانکاران"));
                    //  if (folder.ListItemAllFields.RoleAssignments.Groups.Contains(web.SiteGroups.GetByName("مدیران طرح")))
                    folder.ListItemAllFields.RoleAssignments.Groups.Remove(web.SiteGroups.GetByName("مدیران طرح"));
                    //  if (folder.ListItemAllFields.RoleAssignments.Groups.Contains(web.SiteGroups.GetByName("مشاوران")))
                    folder.ListItemAllFields.RoleAssignments.Groups.Remove(web.SiteGroups.GetByName("مشاوران"));

                    // . GetByPrincipalId(5).DeleteObject();

                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {

                }
            }
            MessageBox.Show("asdsadas");

        }

        private void button10_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis2.jnasr.com/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            //List list = context.Web.Lists.GetByTitle("اطلاعات مستندات و نقشه های ازبیلت");
            // List list = context.Web.Lists.GetByTitle("فرم اعلام نقص");
            List list = context.Web.Lists.GetByTitle("فرم وضعیت رفع نقص");

            var query = new CamlQuery();

            query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where></Query></View>");
            context.Load(list);
            context.Load(list.RootFolder);
            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();



            Group g1 = web.SiteGroups.GetByName("گروه مدیریت راهبری");//گروه مدیریت راهبری 
            Group g2 = web.SiteGroups.GetByName("مشاهده کنندگان اسناد");//گروه مدیریت راهبری 
            //  ListItem item = list.GetItemById(347);
            foreach (ListItem item in items)
            {
                context.Load(item.Folder);
                context.ExecuteQuery();

                hnUser hnU = getUsers(int.Parse(item.Folder.Name));
                User u1 = web.EnsureUser(hnU.advaisor), u2 = web.EnsureUser(hnU.areaManager), u3 = web.EnsureUser(hnU.contractor), u4 = web.EnsureUser(hnU.manager);

                RoleDefinitionBindingCollection collRoleDefinitionBindingRead = new RoleDefinitionBindingCollection(context);
                collRoleDefinitionBindingRead.Add(context.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
                RoleDefinitionBindingCollection collRoleDefinitionBindingCont = new RoleDefinitionBindingCollection(context);
                collRoleDefinitionBindingCont.Add(context.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type
                RoleDefinitionBindingCollection collRoleDefinitionBindingfull = new RoleDefinitionBindingCollection(context);
                collRoleDefinitionBindingfull.Add(context.Web.RoleDefinitions.GetByType(RoleType.Administrator)); //Set permission type

                item.Folder.ListItemAllFields.ResetRoleInheritance();
                item.Folder.ListItemAllFields.BreakRoleInheritance(false, true);
                item.Folder.ListItemAllFields.RoleAssignments.Add(g1, collRoleDefinitionBindingCont);
                item.Folder.ListItemAllFields.RoleAssignments.Add(u1, collRoleDefinitionBindingCont);
                item.Folder.ListItemAllFields.RoleAssignments.Add(u2, collRoleDefinitionBindingCont);
                item.Folder.ListItemAllFields.RoleAssignments.Add(g2, collRoleDefinitionBindingRead);
                item.Folder.ListItemAllFields.RoleAssignments.Add(u4, collRoleDefinitionBindingCont);
                item.Folder.ListItemAllFields.RoleAssignments.Add(web.EnsureUser("nasr\\sp_admin"), collRoleDefinitionBindingfull);
                context.ExecuteQuery();
            }

            MessageBox.Show("done");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/implementation/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = web.Lists.GetByTitle("فرم اعلام نقص");

            var query = new CamlQuery();
            // query.ViewXml = string.Format("<View><Query><Where><Neq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Neq></Where></Query></View>");
            query.ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query><Where><IsNotNull><FieldRef Name='CId' /></IsNotNull></Where></Query></View>");

            //  query.FolderServerRelativeUrl = "/implementation/AnnounceDefects/102";
            var items = list.GetItems(query);

            context.Load(list);
            context.Load(items);

            context.ExecuteQuery();

            //for (int i = 0; i < 200; i++)
            //{
            //    items[i].ResetRoleInheritance();
            //}
            //context.ExecuteQuery();
            //for (int i = 201; i < 400; i++)
            //{
            //    items[i].ResetRoleInheritance();
            //}
            //context.ExecuteQuery();
            //for (int i = 401; i < 600; i++)
            //{
            //    items[i].ResetRoleInheritance();
            //}
            //     context.ExecuteQuery();
            //for (int i = 601; i < items.Count; i++)
            //{
            //    items[i].ResetRoleInheritance();
            //}
            //context.ExecuteQuery();




            MessageBox.Show("done");
            foreach (ListItem item in items)
            {
                item.ResetRoleInheritance();
                context.ExecuteQuery();
            }


            MessageBox.Show("done");
        }

        private void button12_Click(object sender, EventArgs e)
        {


        }

        private void button13_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web, w => w.RoleDefinitions);


            List list = context.Web.Lists.GetByTitle("ارزشیابی پیمانکارانEPC");
            var query = new CamlQuery();
            query.ViewXml = string.Format("<View></View>");
            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            //RoleDefinitionBindingCollection collRoleDefinitionBindingRead = new RoleDefinitionBindingCollection(context);
            //collRoleDefinitionBindingRead.Add(context.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
            //RoleDefinitionBindingCollection collRoleDefinitionBindingCont = new RoleDefinitionBindingCollection(context);
            //collRoleDefinitionBindingCont.Add(context.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type
            // context.ExecuteQuery();
            var roleRead = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);
            var roleCont = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Contributor);
            foreach (ListItem item in items)
            {

                User advisor = web.EnsureUser(getUsers2((item["Contract"] as FieldLookupValue).LookupId, context).advaisor);
                if (item["TotalScore"] != null)
                    item.RoleAssignments.Add(advisor, new RoleDefinitionBindingCollection(context) { roleRead });
                else
                    item.RoleAssignments.Add(advisor, new RoleDefinitionBindingCollection(context) { roleCont });

            }

            context.ExecuteQuery();
            MessageBox.Show("done");


        }

        private hnUser getUsers2(int cId, ClientContext context)
        {
            List list = context.Web.Lists.GetByTitle("پیمان ها");

            ListItem item = list.GetItemById(cId);
            context.Load(item);
            context.ExecuteQuery();

            return new hnUser()
            {
                contractor = "jnasr\\" + (item["ContractorUser"] as FieldLookupValue).LookupValue,
                manager = "jnasr\\" + (item["ManagerUser"] as FieldLookupValue).LookupValue,
                advaisor = "jnasr\\" + (item["ConsultentUser"] as FieldLookupValue).LookupValue,
                areaManager = "jnasr\\" + (item["AreaManagerUser"] as FieldLookupValue).LookupValue

            };

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button14_Click(object sender, EventArgs e)
        {
            //ClientContext context = new ClientContext("http://portal.nasr.ir/sites/jmis/");
            //context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            ClientContext context = new ClientContext("http://172.29.0.162");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web, w => w.RoleDefinitions);
            List hList = context.Web.Lists.GetByTitle("حوزه ها");
            List cList = context.Web.Lists.GetByTitle("پیمان ها");
            //  List List = context.Web.Lists.GetByTitle("صورت جلسات رفع نقص");
            // List tList = context.Web.Lists.GetByTitle("صورت جلسات تحویل موقت");
            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View></View>");
            var items = cList.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            var roleRead = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);
            //var roleCont = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Contributor);
            //Group g1 = web.SiteGroups.GetByName("تیم راهبری");
            //Group g2 = web.SiteGroups.GetByName("تیم راهبری-ویرایش");
            //Group g3 = web.SiteGroups.GetByName("tahvil_Viewer");

            foreach (ListItem item in items)
            {
                ListItem hItem = hList.GetItemById((item["Area"] as FieldLookupValue).LookupId);
                context.Load(hItem);
                context.ExecuteQuery();

                User contractor = web.SiteUsers.GetById((hItem["ExperienceManager"] as FieldLookupValue).LookupId);
                item.RoleAssignments.Add(contractor, new RoleDefinitionBindingCollection(context) { roleRead });

                //ListItem tItem = tList.GetItemById((item["TemporaryDelivery"] as FieldLookupValue).LookupId);
                //context.Load(tItem);
                //context.ExecuteQuery();
                //ListItem cItem = cList.GetItemById((tItem["Contract"] as FieldLookupValue).LookupId);
                //context.Load(cItem);
                //context.ExecuteQuery();
                ////  User contractor = web.SiteUsers.GetById((cItem["ContractorUser"] as FieldLookupValue).LookupId);
                //User advaisor = web.SiteUsers.GetById((cItem["ConsultantUser"] as FieldLookupValue).LookupId);

                //User manager = web.SiteUsers.GetById((cItem["ManagerUser"] as FieldLookupValue).LookupId);
                //User areaManager = web.SiteUsers.GetById((cItem["AreaManagerUser"] as FieldLookupValue).LookupId);

                //item.BreakRoleInheritance(false, true);
                ////    item.RoleAssignments.Add(contractor, new RoleDefinitionBindingCollection(context) { roleRead });
                //item.RoleAssignments.Add(advaisor, new RoleDefinitionBindingCollection(context) { roleRead });
                //item.RoleAssignments.Add(manager, new RoleDefinitionBindingCollection(context) { roleRead });
                //item.RoleAssignments.Add(areaManager, new RoleDefinitionBindingCollection(context) { roleRead });

                //item.RoleAssignments.Add(g3, new RoleDefinitionBindingCollection(context) { roleRead });
                //item.RoleAssignments.Add(g2, new RoleDefinitionBindingCollection(context) { roleCont });
                //item.RoleAssignments.Add(g1, new RoleDefinitionBindingCollection(context) { roleRead });
                context.ExecuteQuery();
            }


            MessageBox.Show("done");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web);

            //   List cList = context.Web.Lists.GetByTitle("عملیاتهای اجرایی در سطح واحد زراعی");
            List list = context.Web.Lists.GetByTitle("پیمان ها");
            //  string cID = textBox2.Text;
            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View><Query><Where><Neq><FieldRef Name='Status' /><Value Type='Choice'>جاری</Value></Neq></Where></Query></View>");
            var items = list.GetItems(query);
            context.Load(list);
            context.Load(list.RootFolder);
            context.Load(items);
            context.ExecuteQuery();


            //   Folder f= list.RootFolder.Folders.Add(list.RootFolder.ServerRelativeUrl + "/Archive" );
            //  context.ExecuteQuery();

            foreach (ListItem item in items)
            {
                // context.Load(item.File);
                //  context.ExecuteQuery();
                item.File.MoveTo(list.RootFolder.ServerRelativeUrl + "/Archive/" + item["Title"].ToString(), MoveOperations.Overwrite);
                item.Update();
                context.ExecuteQuery();
            }
            MessageBox.Show("done");

        }

        private void button16_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web);

            List list = context.Web.Lists.GetByTitle("صورتجلسات تحویل موقت");
            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View></View>");

            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            foreach (var item in items)
            {
                RoleDefinitionBindingCollection collRoleDefinitionBindingRead = new RoleDefinitionBindingCollection(context);
                collRoleDefinitionBindingRead.Add(context.Web.RoleDefinitions.GetByType(RoleType.Reader));
                User u1 = web.EnsureUser("jnasr\\prcm");
                item.RoleAssignments.Add(u1, collRoleDefinitionBindingRead);
                context.ExecuteQuery();
            }

            MessageBox.Show("done");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://172.29.0.178/sites/jmis/");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            ClientContext context2 = new ClientContext("http://172.29.0.162/");
            context2.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web);

            Web web2 = context2.Web;
            context2.Load(web2);

            List listSource = web.Lists.GetByTitle("نوع شرکت");
            List ListDestion = web2.Lists.GetByTitle("CompanyTypes");
            context.Load(listSource);
            context.ExecuteQuery();

            context2.Load(ListDestion);
            context2.ExecuteQuery();


            CamlQuery perQuery = new CamlQuery();
            var q = new CamlQuery() { ViewXml = "<View></View>" };
            var r = listSource.GetItems(q);
            context.Load(r);

            foreach (ListItem item in r)
            {
                var newItem = ListDestion.AddItem(new ListItemCreationInformation());
                newItem["Title"] = item["Title"];
                newItem.Update();
                context2.ExecuteQuery();

            }


            MessageBox.Show("ok         " + r.Count.ToString());

        }

        public static string GetRelatedUser(int userLookupId, int contractId, bool isCompany)
        {
            string userName = "";
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            //context.Credentials = new NetworkCredential("spadmin", "dm!n0sp0abg", "jnasr");
            Web web = context.Web;
            List contractList = web.Lists.GetByTitle("پیمان ها");
            List areaList = web.Lists.GetByTitle("حوزه ها");
            List contractUserList = web.Lists.GetByTitle("کاربران پیمان");
            context.Load(areaList);
            context.Load(contractList);
            context.Load(contractUserList);
            context.ExecuteQuery();
            ListItem coItem = !isCompany ? contractList.GetItemById(contractId) : null;
            ListItem contractUserItem = contractUserList.GetItemById(userLookupId);

            context.Load(coItem);
            context.ExecuteQuery();
            context.Load(contractUserItem);
            context.ExecuteQuery();
            if (userLookupId == 1)
            {
                var contractField = coItem["ContractorUser"] as FieldLookupValue;
                int contractorId = contractField.LookupId;
                userName = contractField.LookupValue;

            }
            else if (userLookupId == 2)
            {
                var contractField = coItem["ConsultentUser"] as FieldLookupValue;
                int contractorId = contractField.LookupId;
                userName = contractField.LookupValue;

            }
            else if (userLookupId == 4)
            {
                var contractField = coItem["ManagerUser"] as FieldLookupValue;
                int contractorId = contractField.LookupId;
                userName = contractField.LookupValue;

            }
            else if (userLookupId == 5)
            {
                var areaField = coItem["Area"] as FieldLookupValue;
                int areaId = areaField.LookupId;

                ListItem areaItem = areaList.GetItemById(areaId);
                context.Load(areaItem);
                context.ExecuteQuery();
                var areaManagerField = areaItem["AreaManagerUser"] as FieldLookupValue;

                userName = areaManagerField.LookupValue;
            }
            //change
            else if (userLookupId == 9)//roo server khodemun 9
            {

                var q = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='Company' /><Value Type='Lookup'>" + contractId + "</Value></Eq></Where></Query></View>" };
                var r = areaList.GetItems(q);
                context.Load(r);
                context.ExecuteQuery();
                var areas = areaList.GetItems(q);
                context.Load(areas);
                context.ExecuteQuery();
                ListItem areaItem = areas[0];
                context.Load(areaItem);
                context.ExecuteQuery();

                var areaManagerField = areaItem["AreaManagerUser"] as FieldLookupValue;

                userName = areaManagerField.LookupValue;
                // userName = new SPFieldLookupValue(areaItem["AreaManagerUser"].ToString()).LookupValue;

            }

            //  }
            else
            {
                var userField = contractUserItem["UserName"] as FieldLookupValue;
                userName = userField.LookupValue;
            }
            return userName;
        }


        private int GetLookupId(string listname, object lookupVal)
        {

            var lookup = lookupVal as FieldLookupValue;


            string lookupValue = lookup.LookupValue;



            List lookupList = webTo.Lists.GetByTitle(listname);
            contextTo.Load(lookupList);
            //    contextTo.ExecuteQuery();
            var q = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='title3' /><Value Type='Text'>" + lookupValue + "</Value></Eq></Where></Query></View>" };

            ListItemCollection r = lookupList.GetItems(q);
            contextTo.Load(r);
            contextTo.ExecuteQuery();
            // cID = (r[0]["Contract"] as FieldLookupValue).LookupId;
            return r[0].Id;


        }
        private int GetLookupId(string listname, object lookupVal, out int cID)
        {

            var lookup = lookupVal as FieldLookupValue;


            string lookupValue = lookup.LookupValue;



            List lookupList = webTo.Lists.GetByTitle(listname);
            contextTo.Load(lookupList);
            //    contextTo.ExecuteQuery();
            var q = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + lookupValue + "</Value></Eq></Where></Query></View>" };

            ListItemCollection r = lookupList.GetItems(q);
            contextTo.Load(r);
            contextTo.ExecuteQuery();
            cID = (r[0]["Contract"] as FieldLookupValue).LookupId;
            return r[0].Id;


        }

        private int GetLookupId(string listname, object lookupVal, int cID)
        {

            var lookup = lookupVal as FieldLookupValue;
            string lookupValue = lookup.LookupValue;

            //var lookupContract = lookupContractVal as FieldLookupValue;
            //string lookupContractValue = lookupContract.LookupValue;


            List lookupList = webTo.Lists.GetByTitle(listname);
            contextTo.Load(lookupList);
            //    contextTo.ExecuteQuery();
            var q = new CamlQuery()
            {
                ViewXml = string.Format(@"<View Scope='RecursiveAll'><Query>
                                                       <Where>
                                                          <And>
                                                             <Eq>
                                                                <FieldRef Name='Contract' LookupId = 'TRUE'/>
                                                                <Value Type='Lookup'>{0}</Value>
                                                             </Eq>
                                                             <Eq>
                                                                <FieldRef Name='Title' />
                                                                <Value Type='Text'>{1}</Value>
                                                             </Eq>
                                                          </And>
                                                       </Where>
                                                    </Query></View>", cID, lookupValue)
            };

            ListItemCollection r = lookupList.GetItems(q);
            contextTo.Load(r);
            contextTo.ExecuteQuery();
            //   cID = (r[0]["Contract"] as FieldLookupValue).LookupId;
            return r[0].Id;


        }
        private ClientContext contextTo;
        private Web webTo;

        private void button18_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");


            contextTo = new ClientContext("http://pmis2.jnasr.com");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            // context.Credentials = new NetworkCredential("spadmin", "dm!n0sp0abg", "jnasr");
            Web webFrom = contextFrom.Web;
            webTo = contextTo.Web;
            contextFrom.Load(webFrom);
            List listFrom = contextFrom.Web.Lists.GetByTitle("پیمان ها");
            contextFrom.Load(listFrom);
            //   contextFrom.ExecuteQuery();
            //
            List listTo = contextTo.Web.Lists.GetByTitle("Contracts");
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();
            //

            List lookupList = webTo.Lists.GetByTitle("Companies");

            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View/>";
            ListItemCollection listItemsFrom = listFrom.GetItems(query);
            contextFrom.Load(listItemsFrom);
            contextFrom.ExecuteQuery();
            foreach (ListItem fromItem in listItemsFrom)
            {
                var newItem = listTo.AddItem(new ListItemCreationInformation());

                User ManagerUser;
                User AreaManagerUser = webTo.EnsureUser("jnasr\\" + (fromItem["AreaManagerUser"] as FieldUserValue).LookupValue);
                User ContractorUser = webTo.EnsureUser("jnasr\\" + (fromItem["ContractorUser"] as FieldUserValue).LookupValue);
                User ConsultantUser = webTo.EnsureUser("jnasr\\" + (fromItem["ConsultentUser"] as FieldUserValue).LookupValue);
                if ("jnasr\\" + (fromItem["ManagerUser"] as FieldUserValue).LookupValue == "jnasr\\m.kharkhehn")
                { ManagerUser = webTo.EnsureUser("jnasr\\m.karkhehn"); }
                else
                { ManagerUser = webTo.EnsureUser("jnasr\\" + (fromItem["ManagerUser"] as FieldUserValue).LookupValue); }
                contextTo.Load(AreaManagerUser);
                contextTo.Load(ContractorUser);
                contextTo.Load(ConsultantUser);
                contextTo.Load(ManagerUser);
                contextTo.ExecuteQuery();





                // if (fromItem["Area"] != null)
                var area = new FieldLookupValue() { LookupId = GetLookupId("Areas", fromItem["Area"]) };
                // if (fromItem["Contractor"] != null)
                var contractor = new FieldLookupValue() { LookupId = GetLookupId("Companies", fromItem["Contractor"]) };
                // if (fromItem["Consultent"] != null)
                var consultent = new FieldLookupValue() { LookupId = GetLookupId("Companies", fromItem["Consultent"]) };
                //  if (fromItem["Manager"] != null)
                var manager = new FieldLookupValue() { LookupId = GetLookupId("Companies", fromItem["Manager"]) };
                newItem["Title"] = fromItem["Title"];
                newItem["ContractType"] = fromItem["ContractType"];
                newItem["ProjectCost"] = fromItem["ProjectCost"];
                newItem["Status"] = fromItem["Status"];
                newItem["Area"] = area;
                newItem["Contractor"] = contractor;
                newItem["Consultant"] = consultent;
                newItem["Manager"] = manager;
                newItem["AreaManagerUser"] = new FieldUserValue() { LookupId = AreaManagerUser.Id };
                newItem["ContractorUser"] = new FieldUserValue() { LookupId = ContractorUser.Id };
                newItem["ConsultantUser"] = new FieldUserValue() { LookupId = ConsultantUser.Id };
                newItem["ManagerUser"] = new FieldUserValue() { LookupId = ManagerUser.Id };

                newItem.Update();
                contextTo.ExecuteQuery();
            }
            MessageBox.Show("OK");
        }

        private void button19_Click(object sender, EventArgs e)
        {
            contextTo = new ClientContext("http://pmis2.jnasr.com");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle("Test");
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();
            string folderName = "001";

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            itemCreateInfo.LeafName = folderName;

            ListItem newItem = listTo.AddItem(itemCreateInfo);
            newItem["Title"] = folderName;
            newItem.Update();
            contextTo.ExecuteQuery();

            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://pmis2.jnasr.com", "Test", folderName);
            var newItem2 = listTo.AddItem(listItemCreationInformation);
            newItem2["Title"] = "sdasdasd";
            newItem2.Update();
            contextTo.ExecuteQuery();

        }

        private void button20_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://pmis2.jnasr.com");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);
            List listFrom = contextFrom.Web.Lists.GetByTitle("بلوک ها");
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle("Blocks");
            List listContracts = contextTo.Web.Lists.GetByTitle("Contracts");
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View></View>"; ;
            ListItemCollection listItems = listContracts.GetItems(query);
            contextTo.Load(listItems);
            contextTo.ExecuteQuery();

            foreach (ListItem cItem in listItems)
            {

                CamlQuery query2 = new CamlQuery();
                query2.ViewXml = string.Format(@"<View><Query>
                                           <Where>
                                              <Eq>
                                                 <FieldRef Name='Contract' />
                                                 <Value Type='Lookup'>{0}</Value>
                                              </Eq>
                                           </Where>
                                        </Query></View>", cItem["Title"]);

                ListItemCollection listItemsFrom = listFrom.GetItems(query2);
                contextFrom.Load(listItemsFrom);
                contextFrom.ExecuteQuery();

                int folderName = cItem.Id;

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderName.ToString();

                ListItem newItem = listTo.AddItem(itemCreateInfo);
                newItem["Title"] = folderName;
                newItem.Update();
                contextTo.ExecuteQuery();

                foreach (ListItem item in listItemsFrom)
                {
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://pmis2.jnasr.com", "Blocks", folderName);
                    var newItem2 = listTo.AddItem(listItemCreationInformation);
                    newItem2["Title"] = item["Title"];
                    newItem2["Contract"] = new FieldLookupValue() { LookupId = folderName };
                    newItem2.Update();
                    contextTo.ExecuteQuery();
                }

            }
            MessageBox.Show("OK");

        }

        private void button21_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://172.29.0.178/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            //string pmisListName = "ارزشیابی نظارت عالیه جزئیات";
            //string masterListName = "ارزشیابی نظارت عالیه";
            //string detailListName = "ارزشیابی نظارت عالیه جزئیات";
            //string detailList = "GreatestSupervisionPCDetails";

            string pmisListName = "ارزشیابی نظارت عالیه جزئیات";
            string masterListName = "ارزشیابی نظارت عالیه";
            string detailListName = "ارزشیابی نظارت عالیه جزئیات";
            string detailList = "GreatestSupervisionPCDetails";
            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);
            List listFrom = contextFrom.Web.Lists.GetByTitle(pmisListName);
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle(detailListName);
            List listContracts = contextTo.Web.Lists.GetByTitle(masterListName);
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View></View>";
            ListItemCollection listItems = listContracts.GetItems(query);
            contextTo.Load(listItems);
            contextTo.ExecuteQuery();

            foreach (ListItem cItem in listItems)
            {

                CamlQuery query2 = new CamlQuery();
                query2.ViewXml = string.Format("<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='EvaluationContract'/><Value Type='Lookup'>{0}</Value></Eq></Where></Query></View>", cItem["Title"]);
                //query2.ViewXml = string.Format(@"<View Scope='Recursive'><Query>
                //                                   <Where>
                //                                      <Contains>
                //                                         <FieldRef Name='EvaluationContract' />
                //                                         <Value Type='Lookup'>مرداد 96</Value>
                //                                      </Contains>
                //                                   </Where>
                //                                </Query></View>");
                ListItemCollection listItemsFrom = listFrom.GetItems(query2);
                contextFrom.Load(listItemsFrom);
                contextFrom.ExecuteQuery();

                int folderName = (cItem["Contract"] as FieldLookupValue).LookupId;

                try
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                    itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                    itemCreateInfo.LeafName = folderName.ToString();

                    ListItem newItem = listTo.AddItem(itemCreateInfo);
                    newItem["Title"] = folderName;
                    newItem.Update();
                    contextTo.ExecuteQuery();
                }
                catch (Exception)
                {
                }

                foreach (ListItem item in listItemsFrom)
                {

                    var contractor = new FieldLookupValue() { LookupId = GetLookupId(masterListName, item["EvaluationContract"]) };

                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                    listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://pmis2.jnasr.com", detailList, folderName);
                    var newItem2 = listTo.AddItem(listItemCreationInformation);
                    newItem2["Title"] = item["Title"];
                    newItem2["Criterion"] = item["Criterion"];
                    newItem2["Index"] = item["Index"];
                    newItem2["IsRelevent"] = item["IsRelevent"];
                    newItem2["Org_Weight"] = item["Org_Weight"];
                    newItem2["Row"] = item["Row"];
                    newItem2["Score"] = item["Score"];
                    newItem2["Weight"] = item["Weight"];
                    newItem2["EvaluationContract"] = contractor;
                    newItem2.Update();
                    contextTo.ExecuteQuery();
                }

            }
            MessageBox.Show("OK");
        }

        private void button22_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://172.29.0.178/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);

            string pmisListName = "ارزشیابی نظارت عالیه جزئیات";
            string masterListName = "ارزشیابی نظارت عالیه";
            string detailListName = "ارزشیابی نظارت عالیه جزئیات";
            string detailList = "GreatestSupervisionPCDetails";

            List listFrom = contextFrom.Web.Lists.GetByTitle(pmisListName);
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle(detailListName);
            List listContracts = contextTo.Web.Lists.GetByTitle(masterListName);
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query2 = new CamlQuery();
            query2.ViewXml = string.Format(@"<View Scope='Recursive'><Query>
                                                       <Where>
                                                          <And>
                                                             <Geq>
                                                                <FieldRef Name='ID' />
                                                                <Value Type='Counter'>18143</Value>
                                                             </Geq>
                                                             <Leq>
                                                                <FieldRef Name='ID' />
                                                                <Value Type='Counter'>24268</Value>
                                                             </Leq>
                                                          </And>
                                                       </Where>
                                                    </Query></View>");
            ListItemCollection listItemsFrom = listFrom.GetItems(query2);
            contextFrom.Load(listItemsFrom);
            contextFrom.ExecuteQuery();

            foreach (ListItem item in listItemsFrom)
            {
                int folderName = 0;
                var master = new FieldLookupValue() { LookupId = GetLookupId(masterListName, item["EvaluationContract"], out folderName) };

                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "", detailList, folderName);
                var newItem2 = listTo.AddItem(listItemCreationInformation);
                newItem2["Title"] = item["Title"];
                newItem2["Criterion"] = item["Criterion"];
                newItem2["Index"] = item["Index"];
                newItem2["IsRelevent"] = item["IsRelevent"];
                newItem2["Org_Weight"] = item["Org_Weight"];
                newItem2["Row"] = item["Row"];
                newItem2["Score"] = item["Score"];
                newItem2["Weight"] = item["Weight"];
                newItem2["EvaluationContract"] = master;
                newItem2.Update();
                contextTo.ExecuteQuery();
            }
            MessageBox.Show("done!");

        }

        private void button23_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://172.29.0.162");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web, w => w.RoleDefinitions);

            List cList = context.Web.Lists.GetByTitle("پیمان ها");
            List List = context.Web.Lists.GetByTitle("بلوکهای تحویل موقت");
            var query = new CamlQuery();
            // query.ViewXml = string.Format(@"<View Scope='RecursiveAll'></View>");
            query.ViewXml = string.Format(@" < View><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where></Query></View>");
            //query.ViewXml = string.Format(@" < View><Query>
            //                           <Where>
            //                              <Eq>
            //                                 <FieldRef Name='Period' />
            //                                 <Value Type='Lookup'>مهر 96</Value>
            //                              </Eq>
            //                           </Where>
            //                        </Query></View>");
            var items = List.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();
            //foreach (ListItem item in items)
            //{
            //    item.ResetRoleInheritance();
            //    context.ExecuteQuery();
            //}

            var roleRead = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);
            var roleCont = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Contributor);
            Group g1 = web.SiteGroups.GetByName("تیم راهبری");
            Group g2 = web.SiteGroups.GetByName("تیم راهبری-ویرایش");
            // Group g3 = web.SiteGroups.GetByName("کاربران موسسه");
            //Group g3 = web.SiteGroups.GetByName("Ejra_viewer");
            foreach (ListItem item in items)
            {
                //  if (item["Contract"] != null)
                {
                    // int cId = (item["Contract"] as FieldLookupValue).LookupId;
                    int cId = int.Parse(item["Title"].ToString());
                    ListItem cItem = cList.GetItemById(cId);
                    context.Load(cItem);
                    context.ExecuteQuery();
                    User contractor = web.SiteUsers.GetById((cItem["ContractorUser"] as FieldLookupValue).LookupId);
                    User advaisor = web.SiteUsers.GetById((cItem["ConsultantUser"] as FieldLookupValue).LookupId);
                    User manager = web.SiteUsers.GetById((cItem["ManagerUser"] as FieldLookupValue).LookupId);
                    User areaManager = web.SiteUsers.GetById((cItem["AreaManagerUser"] as FieldLookupValue).LookupId);

                    item.BreakRoleInheritance(false, true);
                    item.RoleAssignments.Add(contractor, new RoleDefinitionBindingCollection(context) { roleRead });
                    item.RoleAssignments.Add(advaisor, new RoleDefinitionBindingCollection(context) { roleCont });
                    item.RoleAssignments.Add(manager, new RoleDefinitionBindingCollection(context) { roleCont });
                    item.RoleAssignments.Add(areaManager, new RoleDefinitionBindingCollection(context) { roleRead });

                    // item.RoleAssignments.Add(g3, new RoleDefinitionBindingCollection(context) { roleRead });
                    item.RoleAssignments.Add(g2, new RoleDefinitionBindingCollection(context) { roleCont });
                    item.RoleAssignments.Add(g1, new RoleDefinitionBindingCollection(context) { roleRead });
                    context.ExecuteQuery();
                }
            }


            MessageBox.Show("done" + items.Count.ToString());
        }

        private void button24_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://pmis.jnasr.com/implementation");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);


            List listFrom = contextFrom.Web.Lists.GetByTitle("لیست وضعیت اعلام نقص");
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle("اعلام نقص");
            List listTahvil = contextTo.Web.Lists.GetByTitle("صورتجلسات تحویل موقت");

            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            ListItem item = listTahvil.GetItemById(140);
            contextTo.Load(item);
            contextTo.ExecuteQuery();
            int cId = (item["Contract"] as FieldLookupValue).LookupId;
            string Iso8601Today = XmlConvert.ToString(DateTime.Parse(item["DateCommission"].ToString()), XmlDateTimeSerializationMode.Local);
            //CamlQuery query2 = new CamlQuery();
            //query2.ViewXml = string.Format(@"<View Scope='Recursive'><Query>
            //                                           <Where>
            //                                              <And>
            //                                                 <Eq>
            //                                                    <FieldRef Name='CId' />
            //                                                    <Value Type='Text'>{1}</Value>
            //                                                 </Eq>
            //                                                 <Eq>
            //                                                    <FieldRef Name='CommissionTemporaryDeliveryDate' />
            //                                                    <Value IncludeTimeValue='TRUE' Type='DateTime' IncludeTimeValue='FALSE'>{0}</Value>
            //                                                 </Eq>
            //                                              </And>
            //                                           </Where>
            //                                        </Query></View>", Iso8601Today,cId);
            //ListItemCollection listItemsFrom = listFrom.GetItems(query2);
            //contextFrom.Load(listItemsFrom);
            //contextFrom.ExecuteQuery();

            //var r=listItemsFrom.GroupBy(x=> x["DefectTitle"]).

            //foreach (ListItem oldItem in listItemsFrom)
            //{

            //}

            //        ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            //  listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://pmis2.jnasr.com", "Farms", folderName);
            //     var newItem = listTo.AddItem(listItemCreationInformation);

            //newItem["Title"] = cItem["Title"];
            //newItem["Contract"] = new FieldLookupValue() { LookupId = c };
            //newItem["Block"] = new FieldLookupValue() { LookupId = d };
            //newItem["ImpureArea"] = cItem["ImpureArea"];
            //newItem["InitialArea"] = cItem["InitialArea"];
            //newItem.Update();
            //contextTo.ExecuteQuery();

        }

        private void button25_Click(object sender, EventArgs e)
        {
            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");






            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle("بلوکهای تحویل موقت");
            List listContracts = contextTo.Web.Lists.GetByTitle("پیمان ها");
            contextTo.Load(listTo);
            contextTo.Load(listContracts);

            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View><Query><Where><Eq><FieldRef Name='Status' /><Value Type='Choice'>جاری</Value></Eq></Where></Query></View>");
            var items = listContracts.GetItems(query);
            contextTo.Load(items);
            contextTo.ExecuteQuery();



            //int[] s = { 241, 242, 244, 247, 248, 249, 250, 252, 255, 256, 257, 258, 259, 260, 262, 263, 265, 266, 268, 269, 271, 272, 273, 275, 276, 277, 281, 283, 289, 290, 291, 294, 295, 297, 298, 299, 300, 301, 302, 303, 304, 305, 306, 307, 308, 309, 311, 312, 314, 315, 316, 317, 318, 319, 320, 321, 322, 324 };
            foreach (ListItem item in items)
            {
                int folderName = item.Id;

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folderName.ToString();

                ListItem newItem = listTo.AddItem(itemCreateInfo);
                newItem["Title"] = folderName;
                newItem.Update();
                contextTo.ExecuteQuery();
            }

            MessageBox.Show("dssds");

        }

        private void button26_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://172.29.0.178/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://172.29.0.162/");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);


            List listFrom = contextFrom.Web.Lists.GetByTitle("صورت جلسات رفع نقص");
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            //List listTo = contextTo.Web.Lists.GetByTitle("اعلام نقص");
            List listTo = contextTo.Web.Lists.GetByTitle("صورت جلسات رفع نقص");

            contextTo.Load(listTo);

            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View></View>");
            var items = listTo.GetItems(query);
            contextTo.Load(items);
            contextTo.ExecuteQuery();

            foreach (var item in items)
            {
                var query2 = new CamlQuery();
                query2.ViewXml = string.Format(@"<View><Query><Where>
                                                      <Eq>
                                                         <FieldRef Name='FileLeafRef' />
                                                         <Value Type='Text'>{0}</Value>
                                                      </Eq>
                                                   </Where></Query></View>", item["FileLeafRef"].ToString());
                var items2 = listFrom.GetItems(query2);
                contextFrom.Load(items2);
                contextFrom.ExecuteQuery();

                //ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                //listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://pmis2.jnasr.com", detailList, folderName);
                //var newItem2 = listTo.AddItem(listItemCreationInformation);


                int tahvilId = GetLookupId("صورت جلسات تحویل موقت", items2[0]["TemporaryDelivery"]);
                item["Title"] = items2[0]["Title"];
                // newItem2["Filename"] = item["Filename"];
                item["TemporaryDelivery"] = new FieldLookupValue() { LookupId = tahvilId };
                item["ContractorToConsultant"] = items2[0]["ContractorToConsultant"];
                item["ConsultantToProjectManager"] = items2[0]["ConsultantToProjectManager"];
                item["ProjectManagerToJahad"] = items2[0]["ProjectManagerToJahad"];
                item["GuaranteeDate"] = items2[0]["GuaranteeDate"];
                item["ProceedingsDate"] = items2[0]["ProceedingsDate"];
                item["comment"] = items2[0]["desc"];

                item.Update();
                contextTo.ExecuteQuery();
            }

            MessageBox.Show("done!");
        }

        private void button27_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://pmis2.jnasr.com");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);
            List listFrom = webFrom.Lists.GetByTitle("پیمان ها");
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = webTo.Lists.GetByTitle("پیمان ها");
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query = new CamlQuery();
            query.ViewXml = string.Format(@"<View><Query> <Where><Neq><FieldRef Name='Status' /><Value Type='Choice'>جاری</Value></Neq></Where></Query></View>");
            ListItemCollection listItems = listFrom.GetItems(query);
            contextFrom.Load(listItems);
            contextFrom.ExecuteQuery();

            foreach (ListItem cItem in listItems)
            {


                User ManagerUser;
                if ("jnasr\\" + (cItem["ManagerUser"] as FieldUserValue).LookupValue == "jnasr\\m.kharkhehn")
                { ManagerUser = webTo.EnsureUser("jnasr\\m.karkhehn"); }
                else
                { ManagerUser = webTo.EnsureUser("jnasr\\" + (cItem["ManagerUser"] as FieldUserValue).LookupValue); }
                contextTo.Load(ManagerUser);
                contextTo.ExecuteQuery();

                var a = GetLookupId("حوزه ها", cItem["Area"]);
                var b = GetLookupId("شرکتها", cItem["Contractor"]);
                var c = GetLookupId("شرکتها", cItem["Manager"]);
                var d = GetLookupId("شرکتها", cItem["Consultent"]);

                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                var newItem = listTo.AddItem(listItemCreationInformation);

                newItem["Title"] = cItem["Title"];
                newItem["Status"] = cItem["Status"];
                newItem["ContractType"] = cItem["ContractType"];
                newItem["Area"] = new FieldLookupValue() { LookupId = a };
                newItem["Contractor"] = new FieldLookupValue() { LookupId = b };
                newItem["Manager"] = new FieldLookupValue() { LookupId = c };
                newItem["Consultant"] = new FieldLookupValue() { LookupId = d };
                newItem["ManagerUser"] = new FieldUserValue() { LookupId = ManagerUser.Id };
                newItem.Update();
                contextTo.ExecuteQuery();
            }
            MessageBox.Show("ASAsaSasa");
        }

        private void button28_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://172.29.0.178/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://172.29.0.163");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);
            List listFrom = webFrom.Lists.GetByTitle("واحد زراعی");
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = webTo.Lists.GetByTitle("واحدهای زراعی");
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View/>";
            //query.ViewXml = string.Format(@" < View><Query><Where>
            //                                          <Eq>
            //                                            <FieldRef Name='Contract' LookupId='TRUE'/>
            //                                            <Value Type='Lookup'>{0}</Value>
            //                                          </Eq>
            //                                       </Where></Query></View>", 113);
            ListItemCollection listItems = listFrom.GetItems(query);
            contextFrom.Load(listItems);
            contextFrom.ExecuteQuery();

            foreach (ListItem cItem in listItems)
            {

                var c = GetLookupId("پیمان ها", cItem["Contract"]);
                int folderName = c;
                var d = GetLookupId("بلوکها", cItem["Block"], folderName);

                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://172.29.0.163", "Farms", folderName);
                var newItem = listTo.AddItem(listItemCreationInformation);

                newItem["Title"] = cItem["Title"];
                newItem["Contract"] = new FieldLookupValue() { LookupId = c };
                newItem["Block"] = new FieldLookupValue() { LookupId = d };
                newItem["ImpureArea"] = cItem["ImpureArea"];
                newItem["InitialArea"] = cItem["InitialArea"];
                newItem.Update();
                contextTo.ExecuteQuery();
            }
            MessageBox.Show(listItems.Count.ToString());
        }

        private void button29_Click(object sender, EventArgs e)
        {
            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            webTo = contextTo.Web;
            List listTo = webTo.Lists.GetByTitle("تجربیات موضوعی");
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query = new CamlQuery();
            query.ViewXml = string.Format(@"<View></View>");
            ListItemCollection listItems = listTo.GetItems(query);
            contextTo.Load(listItems);
            contextTo.ExecuteQuery();

            Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo("en-US");
            foreach (ListItem item in listItems)
            {
                try
                {
                    Microsoft.SharePoint.Client.File f = item.File;
                    contextTo.Load(f);
                    contextTo.ExecuteQuery();
                    //  f.CheckOut();
                    f.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    // item["Title"] = item["FileLeafRef"].ToString().Split('.')[0];
                    // item.Update();
                    contextTo.ExecuteQuery();
                }
                catch (Exception ex)
                {
                }


            }

            MessageBox.Show("done!");

        }

        private void button30_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://pmis.jnasr.com/sites/jmis");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);

            List list = context.Web.Lists.GetByTitle("صورتجلسات تحویل موقت");
            context.Load(list);

            context.ExecuteQuery();

            var query = new CamlQuery();
            query.ViewXml = string.Format("<View></View>");
            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            foreach (ListItem item in items)
            {

                try
                {
                    Group g1 = web.SiteGroups.GetByName("tahvil_Viewer");//گروه مدیریت راهبری 

                    RoleDefinitionBindingCollection collRoleDefinitionBindingRead = new RoleDefinitionBindingCollection(context);
                    collRoleDefinitionBindingRead.Add(context.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
                    item.RoleAssignments.Add(g1, collRoleDefinitionBindingRead);
                    context.ExecuteQuery();
                }
                catch (Exception)
                {

                }

            }
            MessageBox.Show("done");
        }

        private void button31_Click(object sender, EventArgs e)
        {
            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle("عملیاتهای اجرایی در سطح واحد زراعی");
            List listContracts = contextTo.Web.Lists.GetByTitle("پیمان ها");
            contextTo.Load(listTo);
            contextTo.Load(listContracts);

            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where></Query></View>");
            var items = listContracts.GetItems(query);
            contextTo.Load(items);
            contextTo.ExecuteQuery();

            foreach (ListItem item in items)
            {

            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            ClientContext contextFrom = new ClientContext("http://172.29.0.178/sites/jmis");
            contextFrom.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");


            string pmisListName = "ارزشیابی نظارت کارگاهی";
            // string masterListName = "ارزشیابی نظارت عالیه";
            string detailListName = "ارزشیابی نظارت کارگاهی";
            // string detailList = "GreatestSupervisionPCDetails";


            Web webFrom = contextFrom.Web;
            contextFrom.Load(webFrom);
            List listFrom = contextFrom.Web.Lists.GetByTitle(pmisListName);
            contextFrom.Load(listFrom);

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle(detailListName);
            //  List listContracts = contextTo.Web.Lists.GetByTitle(masterListName);
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();

            CamlQuery query = new CamlQuery();
            query.ViewXml = @"<View><Query>
                                       <Where>
                                          <Eq>
                                             <FieldRef Name='Period' />
                                             <Value Type='Lookup'>مهر 96</Value>
                                          </Eq>
                                       </Where>
                                    </Query></View>";
            ListItemCollection listItemsFrom = listFrom.GetItems(query);
            contextFrom.Load(listItemsFrom);
            contextFrom.ExecuteQuery();

            foreach (ListItem item in listItemsFrom)
            {
                var contract = new FieldLookupValue() { LookupId = GetLookupId("پیمان ها", item["Contract"]) };
                User user;
                if ((item["CurrentUser"] as FieldUserValue).LookupValue.ToLower() == "m.kharkhehn")
                    user = webTo.EnsureUser("jnasr\\m.karkhehn");
                else user = webTo.EnsureUser("jnasr\\" + (item["CurrentUser"] as FieldUserValue).LookupValue);
                contextTo.Load(user);
                contextTo.ExecuteQuery();
                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();

                var newItem2 = listTo.AddItem(listItemCreationInformation);
                newItem2["Title"] = item["Title"];
                newItem2["Contract"] = contract;
                newItem2["Period"] = new FieldLookupValue() { LookupId = 8 };
                newItem2["CurrentUser"] = new FieldUserValue() { LookupId = user.Id };
                newItem2["Status"] = item["Status"];
                newItem2["TotalScore"] = item["TotalScore"];
                //newItem2["Score"] = item["Score"];
                //newItem2["Weight"] = item["Weight"];
                // newItem2["EvaluationContract"] = contractor;
                newItem2.Update();
                contextTo.ExecuteQuery();
            }


            MessageBox.Show("OK");
        }

        private void button33_Click(object sender, EventArgs e)
        {
            contextTo = new ClientContext("http://172.29.0.162");
            contextTo.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            string pmisListName = "ارزشیابی پیمانکارانEPC جزئیات";
            string masterListName = "ارزشیابی پیمانکارانEPC";
            string detailListName = "ارزشیابی پیمانکارانEPC جزئیات";
            string detailList = "EvaluationContractEPCDetails";

            webTo = contextTo.Web;
            List listTo = contextTo.Web.Lists.GetByTitle(detailListName);
            List listContracts = contextTo.Web.Lists.GetByTitle(masterListName);
            contextTo.Load(listTo);
            contextTo.ExecuteQuery();
            Folder f = listTo.RootFolder.Folders.GetByUrl(string.Format("{0}/lists/{1}/{2}", "", detailList, "237"));
            //   FolderCollection f = listTo.RootFolder.Folders;//.GetByUrl(string.Format("{0}/lists/{1}/{2}", "http://172.0.29.162", detailList, "237"));
            contextTo.Load(f);
            contextTo.ExecuteQuery();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://172.29.0.162");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web, w => w.RoleDefinitions);

            List hList = context.Web.Lists.GetByTitle("حوزه ها");
            List cList = context.Web.Lists.GetByTitle("پیمان ها");
            List List = context.Web.Lists.GetByTitle("تجربیات توصیفی");
            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View Scope='RecursiveAll'><Query><Where><Neq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Neq></Where></Query></View>");
            var items = List.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();


            var roleRead = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);
            var roleCont = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Contributor);
            Group g1 = web.SiteGroups.GetByName("تیم راهبری");
            Group g2 = web.SiteGroups.GetByName("تیم راهبری-ویرایش");

            Group g3 = web.SiteGroups.GetByName("Tajrobiat_Viewer");
            foreach (ListItem item in items)
            {
                int cId = (item["Contract"] as FieldLookupValue).LookupId;
                ListItem cItem = cList.GetItemById(cId);
                context.Load(cItem);
                context.ExecuteQuery();

                int hId = (cItem["Area"] as FieldLookupValue).LookupId;
                ListItem hItem = hList.GetItemById(hId);
                context.Load(hItem);
                context.ExecuteQuery();


                User c = web.SiteUsers.GetById((hItem["ExperienceManager"] as FieldLookupValue).LookupId);
                User a = web.SiteUsers.GetById((item["Author"] as FieldLookupValue).LookupId);


                item.BreakRoleInheritance(false, true);
                item.RoleAssignments.Add(c, new RoleDefinitionBindingCollection(context) { roleCont });
                item.RoleAssignments.Add(a, new RoleDefinitionBindingCollection(context) { roleRead });


                item.RoleAssignments.Add(g3, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(g2, new RoleDefinitionBindingCollection(context) { roleCont });
                item.RoleAssignments.Add(g1, new RoleDefinitionBindingCollection(context) { roleRead });
                context.ExecuteQuery();
            }



            MessageBox.Show("done" + items.Count.ToString());
        }

        private void button35_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://172.29.0.162");
            context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            Web web = context.Web;
            context.Load(web);
            context.Load(web, w => w.RoleDefinitions);

            List cList = context.Web.Lists.GetByTitle("گزارش هفتگی عملیات اجرایی");
            var q = new CamlQuery();
            List list = context.Web.Lists.GetByTitle("پیمان ها");
            var query = new CamlQuery();
            query.ViewXml = "<View></View>";
            var contractItems = list.GetItems(query);
            context.Load(contractItems);
            q.ViewXml = string.Format(@"<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='int'>1</Value></Eq></Where></Query></View>");
            //var q = new CamlQuery() { ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name='Status' /><Value Type='Text'>ثبت موقت</Value></Eq></Where></Query></View>" };
            var items = cList.GetItems(q);

            //var query = new CamlQuery();
            //query.ViewXml = string.Format(@"<View><Query><Where><Neq><FieldRef Name='Status' /><Value Type='Choice'>جاری</Value></Neq></Where></Query></View>");
            //var items = cList.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            var roleRead = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);
            var roleCont = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Contributor);
            var roleAdd = context.Site.RootWeb.RoleDefinitions.GetById(1073741931);
            Group g1 = web.SiteGroups.GetByName("تیم راهبری");
            Group g2 = web.SiteGroups.GetByName("تیم راهبری-ویرایش");
            Group g3 = web.SiteGroups.GetByName("کاربران موسسه");

            foreach (ListItem item in items)
            {
                int cId = int.Parse(item["Title"].ToString());
                ListItem cItem = list.GetItemById(cId);
                context.Load(cItem);
                context.ExecuteQuery();
                // User c = web.SiteUsers.GetById((item["CurrentUser"] as FieldLookupValue).LookupId);
                //  ListItem cItem = cList.GetItemById(int.Parse(item["Title"].ToString()));
                // context.Load(cItem);
                // context.ExecuteQuery();
                User contractor = web.SiteUsers.GetById((cItem["ContractorUser"] as FieldLookupValue).LookupId);
                User advaisor = web.SiteUsers.GetById((cItem["ConsultantUser"] as FieldLookupValue).LookupId);

                User manager = web.SiteUsers.GetById((cItem["ManagerUser"] as FieldLookupValue).LookupId);
                User areaManager = web.SiteUsers.GetById((cItem["AreaManagerUser"] as FieldLookupValue).LookupId);

                item.BreakRoleInheritance(true, false);
                item.RoleAssignments.Add(contractor, new RoleDefinitionBindingCollection(context) { roleAdd });
                item.RoleAssignments.Add(advaisor, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(manager, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(areaManager, new RoleDefinitionBindingCollection(context) { roleRead });

                item.RoleAssignments.Add(g3, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(g2, new RoleDefinitionBindingCollection(context) { roleCont });
                item.RoleAssignments.Add(g1, new RoleDefinitionBindingCollection(context) { roleRead });
                context.ExecuteQuery();
            }


            MessageBox.Show("done");



        }

        private void button36_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://172.29.0.162/") { Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr") };
            // ClientContext context = new ClientContext("http://sp:90/") { Credentials = new NetworkCredential("spadmin", "dm!n0sp0abg", "jnasr") };
            //ClientContext context = new ClientContext("http://pmis2.jnasr.com");
            //  context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            List list = context.Web.Lists.GetByTitle("مخزن نامه های ابلاغ تمدید تاخیرات");
            context.Load(list, l => l.DefaultNewFormUrl, l => l.DefaultEditFormUrl, l => l.DefaultDisplayFormUrl);
            context.ExecuteQuery();
            ContentType ct = list.ContentTypes.GetById("0x010100F8C11C031C12934EBB52B528FC2DC560");
            {

                ct.NewFormUrl = list.DefaultNewFormUrl;
                ct.EditFormUrl = list.DefaultEditFormUrl;
                ct.DisplayFormUrl = list.DefaultDisplayFormUrl;

                ct.Update(false);
            }

            list.Update();
            context.ExecuteQuery();
            MessageBox.Show("done!");
        }

        private void button37_Click(object sender, EventArgs e)
        {


            ClientContext context = new ClientContext("http://172.29.0.162/") { Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr") };
            Web web = context.Web;
            context.Load(web, w => w.RoleDefinitions);

            List cList = context.Web.Lists.GetByTitle("پیمان ها");
            List hList = context.Web.Lists.GetByTitle("حوزه ها");
            context.Load(cList);
            context.Load(hList);
            context.ExecuteQuery();
            var query = new CamlQuery();
            query.ViewXml = string.Format(@"<View></View>");
            //   query.ViewXml = string.Format(@"<View><Query><Where><Neq><FieldRef Name='Status' /><Value Type='Choice'>جاری</Value></Neq></Where></Query></View>");
            var items = hList.GetItems(query);
            context.Load(items);

            var roleRead = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Reader);
            //var roleCont = context.Site.RootWeb.RoleDefinitions.GetByType(RoleType.Contributor);
            Group g1 = web.SiteGroups.GetByName("تیم راهبری");
            Group g2 = web.SiteGroups.GetByName("تیم راهبری-ویرایش");
            Group g3 = web.SiteGroups.GetByName("کاربران موسسه");
            context.Load(g1);
            context.Load(g2);
            context.Load(g3);
            context.ExecuteQuery();

            foreach (ListItem item in items)
            {

                //ListItem hItem = hList.GetItemById((item["Area"] as FieldLookupValue).LookupId);


                CamlQuery cquery = new CamlQuery();
                cquery.ViewXml = string.Format(@"<View><Query><Where>
                                                      <Eq>
                                                        <FieldRef Name='Area' LookupId='TRUE'/>
                                                        <Value Type='Lookup'>{0}</Value>
                                                      </Eq>
                                                   </Where></Query></View>", item.Id);

                var contracts = cList.GetItems(cquery);
                context.Load(contracts);
                context.ExecuteQuery();
                item.BreakRoleInheritance(false, true);
                foreach (ListItem cr in contracts)
                {
                    User advisor = web.SiteUsers.GetById((cr["ConsultantUser"] as FieldLookupValue).LookupId);

                    var viewers = (FieldLookupValue[])cr["Viewers"];
                    foreach (FieldLookupValue lkp in viewers)
                    {
                        item.RoleAssignments.Add(web.SiteUsers.GetById(lkp.LookupId), new RoleDefinitionBindingCollection(context) { roleRead });
                    }
                    item.RoleAssignments.Add(advisor, new RoleDefinitionBindingCollection(context) { roleRead });
                }

                User manager = web.SiteUsers.GetById((item["CManagerUser"] as FieldLookupValue).LookupId);
                User areaManager = web.SiteUsers.GetById((item["AreaManagerUser"] as FieldLookupValue).LookupId);
                User experienceManager = web.SiteUsers.GetById((item["ExperienceManager"] as FieldLookupValue).LookupId);
                item.RoleAssignments.Add(experienceManager, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(g3, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(g2, new RoleDefinitionBindingCollection(context) { roleRead });
                item.RoleAssignments.Add(g1, new RoleDefinitionBindingCollection(context) { roleRead });
                context.ExecuteQuery();
            }


            MessageBox.Show("done");
        }

        private void button38_Click(object sender, EventArgs e)
        {
            string userName = "";
            ClientContext context = new ClientContext("http://172.29.0.162/") { Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr") };
            //context.Credentials = new NetworkCredential("spadmin", "dm!n0sp0abg", "jnasr");
            Web web = context.Web;
            List contractList = web.Lists.GetByTitle("پیمان ها");
            List cycleList = web.Lists.GetByTitle("چرخه پیمانها");

            CamlQuery cquery = new CamlQuery();
            cquery.ViewXml = string.Format(@"<View><Query><Where>
                                                      <Neq>
                                                        <FieldRef Name='Status'/>
                                                        <Value Type='Text'>جاری</Value>
                                                      </Neq>
                                                   </Where></Query></View>");
            var contracts = contractList.GetItems(cquery);
            context.Load(contracts);
            context.Load(cycleList);
            context.ExecuteQuery();

            //  List<string> statuses = new List<string>() { "شروع پروژه", "در حال اجرا", "تحویل موقت", "تحویل قطعی", "درحال اجرا-تحویل" };
            List<ItemValue> statuses = new List<ItemValue>();  // {Id=  1,Title= "شروع پروژه" }, { 2, "در حال اجرا" }, { 3 ,""}, { 4, "" }, { 8,"" } };
            statuses.Add(new ItemValue { Id = 1, Title = "شروع پروژه" });
            statuses.Add(new ItemValue { Id = 2, Title = "در حال اجرا" });
            statuses.Add(new ItemValue { Id = 3, Title = "تحویل موقت" });
            statuses.Add(new ItemValue { Id = 4, Title = "تحویل قطعی" });
            statuses.Add(new ItemValue { Id = 8, Title = "درحال اجرا-تحویل" });

            foreach (ListItem cr in contracts)
            {
                foreach (ItemValue status in statuses)
                {
                    ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();

                    var newItem2 = cycleList.AddItem(listItemCreationInformation);
                    newItem2["Title"] = cr["Title"].ToString() + " " + status.Title;
                    newItem2["Contract"] = new FieldLookupValue() { LookupId = cr.Id };
                    // newItem2["Period"] = new FieldLookupValue() { LookupId = 8 };
                    newItem2["ContractStatus"] = new FieldLookupValue() { LookupId = status.Id };

                    newItem2.Update();
                    context.ExecuteQuery();

                }
            }
            MessageBox.Show("ok");
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button39_Click(object sender, EventArgs e)
        {
            ClientContext context = new ClientContext("http://net-sp");
            context.Credentials = new NetworkCredential("spadmin", "Nsr!dm$n!Sp", "nasr2");

            Web web = context.Web;
            List list = web.Lists.GetByTitle("WeeklyPlanConstructions");
            List listdetail = web.Lists.GetByTitle("WeeklyPlanConstructionsDetails");
            context.Load(list);
            context.ExecuteQuery();
            var query = new CamlQuery();
            query.ViewXml = string.Format("<View Scope=\"RecursiveAll\"><Query><Where><And><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq><Eq><FieldRef Name='TypeCommitment'/><Value Type='Text'>تعهد اولیه</Value></Eq></And></Where></Query></View>");
            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();
            foreach (var itm in items)
            {
                //ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                //listItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://net-sp", "WeeklyPlanConstructions", (itm["Contract"] as FieldLookupValue).LookupId);
                //var newItem = list.AddItem(listItemCreationInformation);

                //newItem["Title"] = itm["Title"]+" تعهد بازنگری ";
                //newItem["Contract"] = itm["Contract"];
                //newItem["TypeCommitment"] = "تعهد بازنگری";
                //newItem["Period"] = itm["Period"];
                //newItem["Status"] = "پایان فرآیند";
                //newItem.Update();
                //context.ExecuteQuery();

                CamlQuery querydetail = new CamlQuery();
                querydetail.ViewXml = string.Format("<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='WeeklyConstruction' LookupId='TRUE'/><Value Type='Lookup'>{0}</Value></Eq></Where></Query></View>", itm.Id);
                //query2.ViewXml = string.Format(@"<View Scope='Recursive'><Query>
                //                                   <Where>
                //                                      <Contains>
                //                                         <FieldRef Name='EvaluationContract' />
                //                                         <Value Type='Lookup'>مرداد 96</Value>
                //                                      </Contains>
                //                                   </Where>
                //                                </Query></View>");
                ListItemCollection detailItems = listdetail.GetItems(querydetail);
                context.Load(detailItems);
                context.ExecuteQuery();
                foreach (var detailItem in detailItems)
                {
                    ListItemCreationInformation detaillistItemCreationInformation = new ListItemCreationInformation();
                    detaillistItemCreationInformation.FolderUrl = string.Format("{0}/lists/{1}/{2}", "http://net-sp", "WeeklyPlanConstructionsDetails", (itm["Contract"] as FieldLookupValue).LookupId);
                    var newDetailItem = listdetail.AddItem(detaillistItemCreationInformation);
                   // newDetailItem["Title"] = detailItem["Title"];
                   // newDetailItem["WeeklyConstruction"] = new FieldLookupValue() { LookupId = newItem.Id };
                   // newDetailItem["Period"] = detailItem["Period"];
                   // newDetailItem["SubOperation"] = detailItem["SubOperation"];
                   // newDetailItem["Operation"] = detailItem["Operation"];
                  //  newDetailItem["Constructed"] = detailItem["Constructed"];
                    newDetailItem["Measurement"] = detailItem["Measurement"];
                    newDetailItem.Update();
                    context.ExecuteQuery();
                }


            }
            MessageBox.Show("ok");
        }
    }
    
    public class ItemValue
    {
        public int Id { get; set; }
        public string Title { get; set; }
    }

}