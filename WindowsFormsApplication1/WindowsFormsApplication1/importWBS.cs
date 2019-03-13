using Aspose.Cells;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApplication1
{
    public partial class importWBS : System.Windows.Forms.Form
    {
        public importWBS()
        {
            InitializeComponent();
        }
        //ClientContext context = new ClientContext("http://pmis.jnasr.com/") { Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr") };
       ClientContext context = new ClientContext("http://net-sp/") { Credentials = new NetworkCredential("spadmin", "Nsr!dm$n!Sp", "nasr2") };


        private void importWBS_Load(object sender, EventArgs e)
        {
            Web web = context.Web;

          //  context.Credentials = new NetworkCredential("sp_admin", "snik#5085", "nasr");
            context.Load(web);
            context.ExecuteQuery();

            List list = web.Lists.GetByTitle("پیمان ها");
            context.Load(list);
            context.ExecuteQuery();

            ListItemCollection items = list.GetItems(new CamlQuery() { ViewXml = @"<View><Query> <Where><Eq><FieldRef Name='Status' /><Value Type='Choice'>جاری</Value></Eq></Where></Query></View>" });
            context.Load(items);
            context.ExecuteQuery();
            foreach (ListItem item in items)
            {


                comboBox1.Items.Add(new { Text = item["Title"].ToString(), Value = item.Id });
                comboBox1.DisplayMember = "Text";
                comboBox1.ValueMember = "Value";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();

            var wb = new Workbook(openFileDialog1.FileName);


            CreateOtherItems(wb, web, context);
            CreateTotalItems(wb.Worksheets["Totallll"], web);
            MessageBox.Show("Done!");
        }

        private void CreateItem(Worksheet ws, List farmList, int margeRowId, int rowId, string unit, int operationId, int subOperationId, int cId)
        {
            int valueCell = 5, totalValueCell = 7, amountCell = 8, totalAmountCell = 9, activityTypeCell = 10;
            int totalWeightContractCell = 11, itemWeightActionCell = 12, totalWeightOperationCell = 13, weightActionCell = 14;


            var value = ws.Cells[rowId, valueCell].Value;
            var totalValue = ws.Cells[rowId, totalValueCell].Value;
            var amount = ws.Cells[rowId, amountCell].Value;
            var totalAmount = ws.Cells[rowId, totalAmountCell].Value;
            var activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();

            var networkEqHectar = (ws.Cells[rowId, 15].Value);
            var drainEqHectar = (ws.Cells[rowId, 16].Value);
            var renewEqHectar = (ws.Cells[rowId, 17].Value);


            if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                Convert.ToInt64(totalAmount) != 0)
                UpsertOperation(farmList, context,
                                    cId, operationId, subOperationId, unit, value, totalValue, amount, totalAmount,
                                    ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value,
                                    networkEqHectar, drainEqHectar, renewEqHectar);


  

        }
        public void CreateItems(Workbook wb, Web web, ClientContext context)
        {

            //  MessageBox.Show("jfhgfhgfh");
        }
        public void CreateOtherItems(Workbook wb, Web web, ClientContext context)
        {
            List farmList = web.Lists.GetByTitle("ساختار شکست");
            int valueCell = 5, totalValueCell = 7, amountCell = 8, totalAmountCell = 9, activityTypeCell = 10;
            int totalWeightContractCell = 11, itemWeightActionCell = 12, totalWeightOperationCell = 13, weightActionCell = 14;
            int rowId = 0, margeRowId = 0; ;

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();


            Worksheet wsBlock = wb.Worksheets["Block"];

            for (int i = 6; i < 200; i++)
            {
                if (wsBlock.Cells[i, 2].Value == null)
                    break;
                {
                    rowId = 15;
                    var ws = wb.Worksheets[wsBlock.Cells[i, 2].Value.ToString()];
                    var cId = int.Parse((comboBox1.SelectedItem as dynamic).Value.ToString());//new SPFieldLookupValue(contract.ToString()).LookupId;
                    var blockName = ws.Cells[7, 2].Value.ToString();
                    var block = GetBlockId(cId, blockName, context);
                    var impureArea = Convert.ToDouble(ws.Cells[7, 6].Value.ToString());
                    var area = Convert.ToDouble(ws.Cells[7, 4].Value.ToString());
                    var farmName = ws.Cells[7, 3].Value.ToString();
                    var farm = 1;// GetFarmId(cId, block, farmName, impureArea, area, context);

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
                    rowId = 37;
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

                    operationId = 8;
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
                    rowId = 56;
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

            //  MessageBox.Show("jfhgfhgfh");
        }

        public void CreateTotalItems(Worksheet ws, Web web)
        {
            List farmList = web.Lists.GetByTitle("عملیاتهای اجرایی در سطح واحد پیمان");
            var cId = int.Parse((comboBox1.SelectedItem as dynamic).Value.ToString());
            int valueCell = 5, totalValueCell = 7, amountCell = 8, totalAmountCell = 9, activityTypeCell = 10;
            int totalWeightContractCell = 11, itemWeightActionCell = 12, totalWeightOperationCell = 13, weightActionCell = 14;
            int networkEqHectarId = 15, drainEqHectarId = 16, renewEqHectarId = 17;
            int rowId = 0, margeRowId = 0; ;

            // تامین جت فلاشر 	
            rowId = 64;
            var operationId = 3;
            var value = ws.Cells[rowId, valueCell].Value;
            var totalValue = ws.Cells[rowId, totalValueCell].Value;
            var amount = ws.Cells[rowId, amountCell].Value;
            var totalAmount = ws.Cells[rowId, totalAmountCell].Value;
            var activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
            var networkEqHectar = ws.Cells[rowId, networkEqHectarId].Value;
            var drainEqHectar = ws.Cells[rowId, drainEqHectarId].Value;
            var renewEqHectar = ws.Cells[rowId, renewEqHectarId].Value;

            if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                Convert.ToInt64(totalAmount) != 0)
                UpsertContractOperation(farmList, context,
                                        cId, operationId, "ریال", value, totalValue, amount, totalAmount,
                                        ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value,
                                        networkEqHectar, drainEqHectar, renewEqHectar);

            //تسهیلات کارفرمایی	

            rowId = 65;
            operationId = 2;
            value = ws.Cells[rowId, valueCell].Value;
            totalValue = ws.Cells[rowId, totalValueCell].Value;
            amount = ws.Cells[rowId, amountCell].Value;
            totalAmount = ws.Cells[rowId, totalAmountCell].Value;
            activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
            networkEqHectar = ws.Cells[rowId, networkEqHectarId].Value;
            drainEqHectar = ws.Cells[rowId, drainEqHectarId].Value;
            renewEqHectar = ws.Cells[rowId, renewEqHectarId].Value;
            if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                Convert.ToInt64(totalAmount) != 0)
                UpsertContractOperation(farmList, context,
                                        cId, operationId, "ریال", value, totalValue, amount, totalAmount,
                                        ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value,
                                        networkEqHectar, drainEqHectar, renewEqHectar);
            // سایر
            rowId = 66;
            operationId = 5;
            value = ws.Cells[rowId, valueCell].Value;
            totalValue = ws.Cells[rowId, totalValueCell].Value;
            amount = ws.Cells[rowId, amountCell].Value;
            totalAmount = ws.Cells[rowId, totalAmountCell].Value;
            activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
            networkEqHectar = ws.Cells[rowId, networkEqHectarId].Value;
            drainEqHectar = ws.Cells[rowId, drainEqHectarId].Value;
            renewEqHectar = ws.Cells[rowId, renewEqHectarId].Value;
            if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                Convert.ToInt64(totalAmount) != 0)
                UpsertContractOperation(farmList, context,
                                        cId, operationId, "ریال", value, totalValue, amount, totalAmount,
                                        ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value,
                                        networkEqHectar, drainEqHectar, renewEqHectar);
            ////تجهیز و برچیدن کارگاه	

            rowId = 67;
            operationId = 1;
            value = ws.Cells[rowId, valueCell].Value;
            totalValue = ws.Cells[rowId, totalValueCell].Value;
            amount = ws.Cells[rowId, amountCell].Value;
            totalAmount = ws.Cells[rowId, totalAmountCell].Value;
            activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
            networkEqHectar = ws.Cells[rowId, networkEqHectarId].Value;
            drainEqHectar = ws.Cells[rowId, drainEqHectarId].Value;
            renewEqHectar = ws.Cells[rowId, renewEqHectarId].Value;
            if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                Convert.ToInt64(totalAmount) != 0)
                UpsertContractOperation(farmList, context,
                                        cId, operationId, "ریال", value, totalValue, amount, totalAmount,
                                        ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value,
                                        networkEqHectar, drainEqHectar, renewEqHectar);
            ////خدمات مهندسی	

            rowId = 68;
            operationId = 4;
            value = ws.Cells[rowId, valueCell].Value;
            totalValue = ws.Cells[rowId, totalValueCell].Value;
            amount = ws.Cells[rowId, amountCell].Value;
            totalAmount = ws.Cells[rowId, totalAmountCell].Value;
            activityType = ws.Cells[rowId, activityTypeCell].Value.ToString();
            networkEqHectar = ws.Cells[rowId, networkEqHectarId].Value;
            drainEqHectar = ws.Cells[rowId, drainEqHectarId].Value;
            renewEqHectar = ws.Cells[rowId, renewEqHectarId].Value;
            if (Convert.ToInt64(value) != 0 || Convert.ToInt64(totalValue) != 0 || Convert.ToInt64(amount) != 0 ||
                Convert.ToInt64(totalAmount) != 0)
                UpsertContractOperation(farmList, context,
                                        cId, operationId, "ریال", value, totalValue, amount, totalAmount,
                                        ws.Cells[rowId, totalWeightContractCell].Value, ws.Cells[rowId, itemWeightActionCell].Value, ws.Cells[margeRowId, totalWeightOperationCell].Value, ws.Cells[margeRowId, weightActionCell].Value,
                                        networkEqHectar, drainEqHectar, renewEqHectar);
        }
        public int GetBlockId(int cId, string blockName, ClientContext context)
        {
            var list = context.Web.Lists.GetById(new Guid("62B9CA97-D7A7-48E3-A17A-79F07BFF7FC0"));
            var q = new CamlQuery();
            var r = 0;
            q.ViewXml = @"<View  Scope='RecursiveAll'><Query><Where><And>
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

        public int GetFarmId(int cId, int bId, string farmName, double impureArea, double area, ClientContext context, double NetDeliveryLevel, double DrainageDeliverylevel, double EquippedDeliverylevel, double InitialNetDeliveryLevel, double InitialDrainageDeliverylevel, double InitialEquippedDeliverylevel)
        {
            var list = context.Web.Lists.GetById(new Guid("C019ADE0-54F2-4F98-9CF9-ACCDCCCED83F"));
            var q = new CamlQuery();

            q.ViewXml = string.Format(@"<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='Contract' LookupId = 'TRUE'/><Value Type='Lookup'>{0}</Value>
                                                      </Eq><And><Eq><FieldRef Name='Block' LookupId = 'TRUE'/><Value Type='Lookup'>{1}</Value>
                                                           </Eq><Eq><FieldRef Name='Title' /><Value Type='Text'>{2}</Value></Eq></And></And></Where></Query></View>", cId, bId, farmName);

            var result = list.GetItems(q);
            context.Load(result);
            context.ExecuteQuery();
            ListItem item = null;
            if (result.Count > 0)
                //  return result[0].Id;
                item = result[0];
            else
            {

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                item = list.AddItem(itemCreateInfo);
            }
            item["Title"] = farmName;
            item["Block"] = new FieldLookupValue() { LookupId = bId };
            item["Contract"] = new FieldLookupValue() { LookupId = cId };
            item["ImpureArea"] = impureArea;
            item["InitialArea"] = area;
            item["NetDeliveryLevel"] = NetDeliveryLevel;
            item["DrainageDeliverylevel"] = DrainageDeliverylevel;
            item["EquippedDeliverylevel"] = EquippedDeliverylevel;
            item["InitialNetDeliveryLevel"] = InitialNetDeliveryLevel;
            item["InitialDrainageDeliverylevel"] = InitialDrainageDeliverylevel;
            item["InitialEquippedDeliverylevel"] = InitialEquippedDeliverylevel;
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


        public ListItem UpsertContractOperation(List list, ClientContext context,
                                        int cId, int operationId, string measurement,
                                        object value, object totalValue, object amount, object totalAmount, object totalWeightContract, object itemWeightAction, object totalWeightOperation, object weightAction, object networkEqHectar, object drainEqHectar, object renewEqHectar)
        {



            var q = new CamlQuery();
            q.ViewXml = string.Format(@"<View><Query><Where>
                                                        <And><Eq><FieldRef Name='Contract' LookupId='True'/><Value Type='integer' >{0}</Value></Eq>                                                    
                                                        <And><Eq><FieldRef Name='Operation' LookupId='True' /><Value Type='integer'>{1}</Value></Eq>
                                                        <Eq><FieldRef Name='Measurement' /><Value Type='Choice'>{2}</Value></Eq>
                                    </And></And></And></And></And></Where></Query></View>", cId, operationId, measurement);

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
                item["Operation"] = new FieldLookupValue() { LookupId = operationId };
                item["Measurement"] = measurement;
            }

            item["FirstVolume"] = value;
            item["FinalVolume"] = totalValue;
            item["FirstCost"] = amount;
            item["FinalCost"] = totalAmount;
            item["EqNetwork"] = networkEqHectar;
            item["EqEquipped"] = renewEqHectar;
            item["EqDrainage"] = drainEqHectar;
            item["ItemWeight"] = totalWeightContract;
            item["ItemWeightOperation"] = itemWeightAction;
            item["TotalItemWeight"] = totalWeightOperation;
            item["TotalWeightOperation"] = weightAction;
            item.Update();
            context.ExecuteQuery();
            return item;
        }


        public ListItem UpsertOperation(List list, ClientContext context,
                                 int cId, int operationId, int subOperationId, string measurement,
                                 object value, object totalValue, object amount, object totalAmount, object totalWeightContract, object itemWeightAction, object totalWeightOperation, object weightAction, object networkEqHectar, object drainEqHectar, object renewEqHectar)
        {



            var q = new CamlQuery();
            q.ViewXml = string.Format(@"<View Scope='Recursive'><Query><Where>
                                                        <And><Eq><FieldRef Name='Contract' LookupId='True'/><Value Type='integer' >{0}</Value></Eq>
                                                        <And><Eq><FieldRef Name='Operation' LookupId='True' /><Value Type='integer'>{1}</Value></Eq>
                                                        <And><Eq><FieldRef Name='SubOperation' LookupId='True'/><Value Type='integer'>{2}</Value></Eq>
                                                        <Eq><FieldRef Name='Measurement' /><Value Type='Choice'>{3}</Value></Eq>
                                   </And></And></And></Where></Query></View>", cId, operationId, subOperationId, measurement);

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
                item["Operation"] = new FieldLookupValue() { LookupId = operationId };
                item["SubOperation"] = new FieldLookupValue() { LookupId = subOperationId };
                item["Measurement"] = measurement;
            }

            item["FirstVolume"] = value;
            item["FinalVolume"] = totalValue;
            item["CheckValue"] = totalValue;
            item["FirstCost"] = amount;
            item["FinalCost"] = totalAmount;
            //   item["EqAcre"] = eqHectar;
            item["EqNetwork"] = networkEqHectar;
            item["EqEquipped"] = renewEqHectar;
            item["EqDrainage"] = drainEqHectar;
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
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();

            var wb = new Workbook(openFileDialog1.FileName);


            List farmList = web.Lists.GetByTitle("ساختار شکست");
            Worksheet ws = wb.Worksheets["Totallll"];

            var cId = int.Parse((comboBox1.SelectedItem as dynamic).Value.ToString());

            Worksheet wsBlock = wb.Worksheets["Block"];

            for (int i = 6; i < 300; i++)
            {
                if (wsBlock.Cells[i, 2].Value == null)
                    break;

                var wsB = wb.Worksheets[wsBlock.Cells[i, 2].Value.ToString()];

                var blockName = wsB.Cells[7, 2].Value.ToString();
                var block = GetBlockId(cId, blockName, context);
                var impureArea = Convert.ToDouble(wsB.Cells[7, 6].Value.ToString());
                var area = Convert.ToDouble(wsB.Cells[7, 4].Value.ToString());
                var farmName = wsB.Cells[7, 3].Value.ToString();
                var NetDeliveryLevel = Convert.ToDouble(wsB.Cells[7, 13].Value.ToString());
                var DrainageDeliverylevel = Convert.ToDouble(wsB.Cells[8, 13].Value.ToString());
                var EquippedDeliverylevel = Convert.ToDouble(wsB.Cells[9, 13].Value.ToString());
                var InitialNetDeliveryLevel = Convert.ToDouble(wsB.Cells[7, 11].Value.ToString());
                var InitialDrainageDeliverylevel = Convert.ToDouble(wsB.Cells[8, 11].Value.ToString());
                var InitialEquippedDeliverylevel = Convert.ToDouble(wsB.Cells[9, 11].Value.ToString());

                var farm = GetFarmId(cId, block, farmName, impureArea, area, context, NetDeliveryLevel, DrainageDeliverylevel, EquippedDeliverylevel, InitialNetDeliveryLevel, InitialDrainageDeliverylevel, InitialEquippedDeliverylevel);

            }



            // کانال درجه 3( درجا)
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 15, 15, "مترطول", 1, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 15, 16, "هکتار", 1, 1, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 15, 17, "مترطول", 1, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 15, 18, "تعداد", 1, 3, cId);


            //نصب کانالت
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 19, 19, "مترطول", 2, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 19, 20, "هکتار", 2, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 19, 21, "مترطول", 2, 4, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 19, 22, "مترطول", 2, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 19, 23, "تعداد", 2, 3, cId);


            //لوله کم فشار
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 24, 24, "مترطول", 3, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 24, 25, "هکتار", 3, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 24, 26, "مترطول", 3, 4, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 24, 27, "مترطول", 3, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 24, 28, "تعداد", 3, 3, cId);
            //  اجرای ایستگاه پمپاژ و نصب تجهیزات
            CreateItem(ws, farmList, 24, 29, "درصد", 3, 5, cId);
            //  تامین سهم پیمانکار (ایستگاه پمپاژ)
            CreateItem(ws, farmList, 24, 30, "درصد", 3, 10, cId);


            //آبیاری تحت فشار
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 31, 31, "مترطول", 4, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 31, 32, "هکتار", 4, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 31, 33, "مترطول", 4, 4, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 31, 34, "مترطول", 4, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 31, 35, "تعداد", 4, 3, cId);
            //  اجرای ایستگاه پمپاژ و نصب تجهیزات
            CreateItem(ws, farmList, 31, 36, "درصد", 4, 5, cId);
            //  تامین سهم پیمانکار (ایستگاه پمپاژ)
            CreateItem(ws, farmList, 31, 37, "درصد", 4, 10, cId);


            // تامین لوله و متعلقات واجرای خط انتقال آب
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 38, 38, "مترطول", 8, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 38, 39, "هکتار", 8, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 38, 40, "مترطول", 8, 4, cId);
            //  سازه
            CreateItem(ws, farmList, 38, 41, "تعداد", 8, 3, cId);


            //زهکش روباز
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 42, 42, "مترطول", 5, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 42, 43, "هکتار", 5, 1, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 42, 44, "مترطول", 5, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 42, 45, "تعداد", 5, 3, cId);


            //کلکتور (زهکش جمع کننده لوله ای)
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 46, 46, "مترطول", 9, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 46, 47, "هکتار", 9, 1, cId);
            // تامین سهم پیمانکار
            CreateItem(ws, farmList, 46, 48, "مترطول", 9, 4, cId);
            //  سازه
            CreateItem(ws, farmList, 46, 49, "تعداد", 9, 3, cId);

            //زهکش های زیرزمینی(لترال) 
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 50, 50, "مترطول", 6, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 50, 51, "هکتار", 6, 1, cId);
            // تامین سهم پیمانکار
            CreateItem(ws, farmList, 50, 52, "مترطول", 6, 4, cId);
            //  سازه
            CreateItem(ws, farmList, 50, 53, "تعداد", 6, 3, cId);

            //تجهیز و نوسازی
            //تسطیح نسبی (خالص)
            CreateItem(ws, farmList, 54, 54, "هکتار", 7, 12, cId);
            //تسطیح اساسی (خالص)
            CreateItem(ws, farmList, 54, 55, "هکتار", 7, 13, cId);
            //تجهیز و نوسازی    
            CreateItem(ws, farmList, 54, 56, "هکتار", 7, 6, cId);
            //هندسی سازی
            CreateItem(ws, farmList, 54, 57, "هکتار", 7, 8, cId);
            //یکپارچه سازی
            CreateItem(ws, farmList, 54, 58, "هکتار", 7, 9, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 54, 59, "مترطول", 7, 4, cId);
            //کانال درجه 4
            CreateItem(ws, farmList, 54, 60, "مترطول", 7, 11, cId);
            //سازه
            CreateItem(ws, farmList, 54, 61, "تعداد", 7, 3, cId);
            //آبشویی
            CreateItem(ws, farmList, 54, 62, "هکتار", 7, 14, cId);
            //جاده دسترسی بین مزارع
            CreateItem(ws, farmList, 54, 63, "مترطول", 7, 7, cId);


            // تامین جت فلاشر 	
            CreateItem(ws, farmList, 64, 64, "ریال", 10, 17, cId);
            //تسهیلات کارفرمایی	
            CreateItem(ws, farmList, 65, 65, "ریال", 10, 16, cId);
            // سایر
            CreateItem(ws, farmList, 66, 66, "ریال", 10, 19, cId);
            //تجهیز و برچیدن کارگاه	
            CreateItem(ws, farmList, 67, 67, "ریال", 10, 15, cId);
            //خدمات مهندسی	
            CreateItem(ws, farmList, 68, 68, "ریال", 10, 18, cId);



            MessageBox.Show("Done!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ListItemCreationInformation information;
           //string strErr= IsEmptyUserName();
           // if (strErr != "")
           // {
           //     MessageBox.Show(strErr);
           //     return;
           // }
            
         //   context.Credentials = new NetworkCredential(txtUserName.Text, txtPassWord.Text, "nasr");
            Web web = this.context.Web;
            bool flag = false;
            Folder folder = null;
            Folder docFolder = null;
            context.Load(web);
            context.ExecuteQuery();
            this.context.ExecuteQuery();
            List byTitle = web.Lists.GetByTitle("صورت وضعیت پیمانکار");
            List list2 = web.Lists.GetByTitle("مخزن صورت وضعیت پیمانکار");
            context.Load(byTitle,  l => l.RootFolder );
            context.Load(list2,  l => l.RootFolder.ServerRelativeUrl );
            context.ExecuteQuery();
            Workbook workbook = new Workbook(this.openFileDialog1.FileName);
            Worksheet worksheet = workbook.Worksheets["S1"];
            int num = int.Parse(worksheet.Cells[1, 0x18].Value.ToString());
            int contractId= int.Parse((comboBox1.SelectedItem as dynamic).Value.ToString());
            try
            {
                Folder obj3 = web.GetFolderByServerRelativeUrl((byTitle.RootFolder.ServerRelativeUrl + "/") + contractId);
                context.Load(obj3);
                context.ExecuteQuery();
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            if (!flag)
            {
                information = new ListItemCreationInformation
                {
                    UnderlyingObjectType = FileSystemObjectType.Folder,
                    LeafName = contractId.ToString()
                };
                ListItem item = byTitle.AddItem(information);
                item["Title"] = contractId.ToString();
                item.Update();
                context.ExecuteQuery();
            }
            folder = web.GetFolderByServerRelativeUrl((byTitle.RootFolder.ServerRelativeUrl + "/") + contractId);
            docFolder = web.GetFolderByServerRelativeUrl((list2.RootFolder.ServerRelativeUrl + "/") + contractId);
            context.Load(folder,  f => f.ServerRelativeUrl );
            context.Load(docFolder, d => d.ServerRelativeUrl);
            context.ExecuteQuery();
            for (int i = 5; i < (num + 5); i++)
            {
                if (worksheet.Cells[i, 2].Value == null)
                {
                    break;
                }
                CamlQuery query = new CamlQuery();
                query.ViewXml = string.Format(@"<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>{0}</Value></Eq></Where></Query></View>", worksheet.Cells[i, 2].Value.ToString() +".pdf");
                query.FolderServerRelativeUrl = docFolder.ServerRelativeUrl;
              


                ListItemCollection items = list2.GetItems(query);
                this.context.Load(items);
                this.context.ExecuteQuery();
                ListItem item2 = items[0];
                information = new ListItemCreationInformation
                {
                    FolderUrl = folder.ServerRelativeUrl
                };
                ListItem item3 = byTitle.AddItem(information);
                item3["Title"] = worksheet.Cells[i, 2].Value.ToString();
                FieldLookupValue value2 = new FieldLookupValue
                {
                    LookupId = contractId
                };
                item3["Contract"] = value2;
                item3["StartDate"] = this.ConvertToMilady(worksheet.Cells[i, 3].Value.ToString());
                item3["EndDate"] = this.ConvertToMilady(worksheet.Cells[i, 4].Value.ToString());
                item3["NetworkCM"] = decimal.Parse(worksheet.Cells[i, 5].Value.ToString());
                item3["EquippedCM"] = decimal.Parse(worksheet.Cells[i, 6].Value.ToString());
                item3["CMDate"] = this.ConvertToMilady(worksheet.Cells[i, 7].Value.ToString());
                FieldLookupValue value3 = new FieldLookupValue
                {
                    LookupId = item2.Id
                };
                item3["InvoiceCM"] = value3;
                item3["Status"] = "پایان فرآیند";
                item3.Update();
                this.context.ExecuteQuery();
            }
            MessageBox.Show("ok");
        }


        private void btnInvoiceConsultant_Click(object sender, EventArgs e)
        {
            ListItemCreationInformation information;
            //string strErr = IsEmptyUserName();
            //if (strErr != "")
            //{
            //    MessageBox.Show(strErr);
            //    return;
            //}
           // context.Credentials = new NetworkCredential(txtUserName.Text, txtPassWord.Text, "nasr");
            Web web = this.context.Web;
            bool flag = false;
            Folder folder = null;
            this.context.Load(web);
            this.context.ExecuteQuery();
            List byTitle = web.Lists.GetByTitle("صورتحساب مشاور");
            List list2 = web.Lists.GetByTitle("مخزن صورتحساب مشاور");
            this.context.Load<List>(byTitle,  l => l.RootFolder );
            this.context.Load<List>(list2,  l => l.RootFolder );
            this.context.ExecuteQuery();
            Workbook workbook = new Workbook(this.openFileDialog1.FileName);
            Worksheet worksheet = workbook.Worksheets["H1"];
            int num = int.Parse(worksheet.Cells[2, 8].Value.ToString());
            decimal consultantAmount = decimal.Parse(worksheet.Cells[0, 11].Value.ToString());
            
            int  contractId = int.Parse(((dynamic)this.comboBox1.SelectedItem).Value.ToString());
            try
            {
                Folder obj3 = web.GetFolderByServerRelativeUrl((byTitle.RootFolder.ServerRelativeUrl + "/") + contractId);
                this.context.Load(obj3);
                this.context.ExecuteQuery();
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            if (!flag)
            {
                information = new ListItemCreationInformation
                {
                    UnderlyingObjectType = FileSystemObjectType.Folder,
                    LeafName = contractId.ToString()
                };
                ListItem item = byTitle.AddItem(information);
                item["Title"] = contractId.ToString();
                item.Update();
                this.context.ExecuteQuery();
            }
            folder = web.GetFolderByServerRelativeUrl((byTitle.RootFolder.ServerRelativeUrl + "/") + contractId);
            this.context.Load(folder,  f => f.ServerRelativeUrl );
            this.context.ExecuteQuery();
            for (int i = 9; i < (num + 9); i++)
            {
                if (worksheet.Cells[i, 0].Value == null)
                {
                    break;
                }
                CamlQuery query = new CamlQuery
                {
                  ViewXml=string.Format(@"<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>{0}</Value></Eq></Where></Query></View>", worksheet.Cells[i, 0].Value.ToString() +".pdf"),
                   
                    FolderServerRelativeUrl = (list2.RootFolder.ServerRelativeUrl + "/") + contractId.ToString()
                };
                ListItemCollection items = list2.GetItems(query);
                this.context.Load(items);
                this.context.ExecuteQuery();
                ListItem item2 = items[0];
                information = new ListItemCreationInformation
                {
                    FolderUrl = folder.ServerRelativeUrl
                };
                ListItem item3 = byTitle.AddItem(information);
                item3["Title"] = worksheet.Cells[i, 0].Value.ToString();
                FieldLookupValue value2 = new FieldLookupValue
                {
                    LookupId = contractId
                };
                item3["Contract"] = value2;
                item3["StartDate"] = this.ConvertToMilady(worksheet.Cells[i, 1].Value.ToString());
                item3["EndDate"] = this.ConvertToMilady(worksheet.Cells[i, 2].Value.ToString());
                item3["CMDate"] = this.ConvertToMilady(worksheet.Cells[i, 4].Value.ToString());
                item3["CMNum"] = decimal.Parse(worksheet.Cells[i, 3].Value.ToString());
                item3["PercentConsultantNum"] = consultantAmount;
                FieldLookupValue value3 = new FieldLookupValue
                {
                    LookupId = item2.Id
                };
                item3["InvoiceCM"] = value3;
                item3["Status"] = "پایان فرآیند";
                item3.Update();
                this.context.ExecuteQuery();
            }
            MessageBox.Show("ok");
        }


        private void btnAdjustment_Click(object sender, EventArgs e)
        {
            ListItemCreationInformation information;
            //string strErr = IsEmptyUserName();
            //if (strErr != "")
            //{
            //    MessageBox.Show(strErr);
            //    return;
            //}
          //  context.Credentials = new NetworkCredential(txtUserName.Text, txtPassWord.Text, "nasr");
            Web web = this.context.Web;
            bool flag = false;
            Folder folder = null;
            this.context.Load(web);
            this.context.ExecuteQuery();
            List byTitle = web.Lists.GetByTitle("تعدیل پیمانکار");
            List list2 = web.Lists.GetByTitle("صورت وضعیت پیمانکار");
            List list3 = web.Lists.GetByTitle("مخزن تعدیل پیمانکار");
            this.context.Load<List>(byTitle, l => l.RootFolder);
            this.context.Load<List>(list3, l => l.RootFolder);
            this.context.Load<List>(list2, l => l.RootFolder);
            this.context.ExecuteQuery();
            Workbook workbook = new Workbook(this.openFileDialog1.FileName);
            Worksheet worksheet = workbook.Worksheets["T1"];
            int num = int.Parse(worksheet.Cells[1, 50].Value.ToString());
            string type = worksheet.Cells[1, 10].Value.ToString();
            dynamic obj2 = int.Parse(((dynamic)this.comboBox1.SelectedItem).Value.ToString());
            try
            {
                object obj3 = web.GetFolderByServerRelativeUrl((byTitle.RootFolder.ServerRelativeUrl + "/") + obj2);
                this.context.Load((dynamic)obj3);
                this.context.ExecuteQuery();
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            if (!flag)
            {
                information = new ListItemCreationInformation
                {
                    FolderUrl = byTitle.RootFolder.ServerRelativeUrl,
                    UnderlyingObjectType = FileSystemObjectType.Folder,
                    LeafName = (string)obj2.ToString()
                };
                ListItem item = byTitle.AddItem(information);
                item["Title"] = obj2;
                item.Update();
                this.context.ExecuteQuery();
            }
            folder = (Folder)web.GetFolderByServerRelativeUrl((byTitle.RootFolder.ServerRelativeUrl + "/") + obj2);
            this.context.Load<Folder>(folder,  f => f.ServerRelativeUrl );
            this.context.ExecuteQuery();
            for (int i = 5; i < (num + 5); i++)
            {
                if (worksheet.Cells[i, 2].Value == null)
                {
                    break;
                }

                CamlQuery query = new CamlQuery()
                {
                    ViewXml = string.Format(@"<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='Text'>{0}</Value></Eq></Where></Query></View>", worksheet.Cells[i, 2].Value.ToString() + ".pdf"),
                    FolderServerRelativeUrl = (string)((list3.RootFolder.ServerRelativeUrl + "/") + obj2.ToString())
                };

                ListItemCollection items = list3.GetItems(query);
                this.context.Load(items);
                this.context.ExecuteQuery();
                ListItem item2 = items[0];
                information = new ListItemCreationInformation
                {
                    FolderUrl = folder.ServerRelativeUrl
                };
                List<FieldLookupValue> list4 = new List<FieldLookupValue>();
                CamlQuery query2 = new CamlQuery
                {
                    ViewXml = "<View Scope='Recursive'><Query><Where><In><FieldRef Name='Title'/><Values>"
                };
                for (int j = 0; j <= 20; j++)
                {
                    if (worksheet.Cells[i, j + 11].Value == null)
                    {
                        break;
                    }
                    object viewXml = query2.ViewXml;
                    query2.ViewXml = string.Concat(new object[] { viewXml, "<Value Type='Text'>", worksheet.Cells[i, j + 11].Value, "</Value>" });
                }
                query2.ViewXml = query2.ViewXml + "</Values></In></Where></Query></View>";
                query2.FolderServerRelativeUrl = (string)((list2.RootFolder.ServerRelativeUrl + "/") + obj2.ToString());
                ListItemCollection items2 = list2.GetItems(query2);
                this.context.Load<ListItemCollection>(items2);
                this.context.ExecuteQuery();
                foreach (ListItem item3 in items2)
                {
                    FieldLookupValue value2 = new FieldLookupValue
                    {
                        LookupId = item3.Id
                    };
                    list4.Add(value2);
                }
                ListItem item4 = byTitle.AddItem(information);
                item4["Title"] = worksheet.Cells[i, 2].Value.ToString();
                FieldLookupValue value4 = new FieldLookupValue
                {
                    LookupId = (int)obj2
                };
                item4["Contract"] = value4;
                item4["Quarter"] = worksheet.Cells[i, 3].Value.ToString();
                item4["Year"] = worksheet.Cells[i, 4].Value.ToString();
                item4["NetworkCM"] = decimal.Parse(worksheet.Cells[i, 5].Value.ToString());
                item4["EquippedCM"] = decimal.Parse(worksheet.Cells[i, 6].Value.ToString());
                item4["CMDate"] = this.ConvertToMilady(worksheet.Cells[i, 7].Value.ToString());
                item4["InvoiceNumber"] = list4;
                FieldLookupValue value5 = new FieldLookupValue
                {
                    LookupId = item2.Id
                };
                item4["InvoiceCM"] = value5;
                item4["Status"] = "پایان فرآیند";
                item4["Type"] = type;
                item4.Update();
                this.context.ExecuteQuery();
            }
            MessageBox.Show("ok");
        }


        private void btnInvoiceManager_Click(object sender, EventArgs e)
        {

            //string strErr = IsEmptyUserName();
            //if (strErr != "")
            //{
            //    MessageBox.Show(strErr);
            //    return;
            //}
            
             //   context.Credentials = new NetworkCredential(txtUserName.Text, txtPassWord.Text, "nasr");

        
            Web web = this.context.Web;
            this.context.Load(web);
            this.context.ExecuteQuery();
            List byTitle = web.Lists.GetByTitle("صورتحساب مدیر طرح");
            List list2 = web.Lists.GetByTitle("مخزن صورتحساب مدیر طرح");
            this.context.Load<List>(byTitle,  l => l.RootFolder );
            this.context.Load<List>(list2, l => l.RootFolder );
            this.context.ExecuteQuery();
            Workbook workbook = new Workbook(this.openFileDialog1.FileName);
            Worksheet worksheet = workbook.Worksheets["M1"];
            int num = int.Parse(worksheet.Cells[2, 8].Value.ToString());
            int num2 = int.Parse(worksheet.Cells[2, 4].Value.ToString());
            for (int i = 9; i < (num + 9); i++)
            {
                if (worksheet.Cells[i, 0].Value == null)
                {
                    break;
                }
                ListItemCreationInformation parameters = new ListItemCreationInformation();
                ListItem item = byTitle.AddItem(parameters);
                item["Title"] = worksheet.Cells[i, 0].Value.ToString();
                FieldLookupValue value2 = new FieldLookupValue
                {
                    LookupId = num2
                };
                item["Area"] = value2;
                item["StartDate"] = this.ConvertToMilady(worksheet.Cells[i, 1].Value.ToString());
                item["EndDate"] = this.ConvertToMilady(worksheet.Cells[i, 2].Value.ToString());
                item["OwnerDate"] = this.ConvertToMilady(worksheet.Cells[i, 4].Value.ToString());
                item["OwnerNum"] = decimal.Parse(worksheet.Cells[i, 3].Value.ToString());
                item["Status"] = "پایان فرآیند";
                item.Update();
                this.context.ExecuteQuery();
            }
            MessageBox.Show("ok");
        }
        private DateTime ConvertToMilady(string faDate)
        {
            PersianCalendar calendar = new PersianCalendar();
            string[] strArray = faDate.Split(new char[] { '/' });
            string s = strArray[0];
            string str2 = strArray[1];
            string str3 = strArray[2];
            return calendar.ToDateTime(int.Parse(s), int.Parse(str2), int.Parse(str3), 0, 0, 0, 0);
        }

        private string IsEmptyUserName()
        {
            if (txtUserName.Text == "" || txtPassWord.Text == "")

                return "وارد کردن نام کاربری و پسورد الزامی است. ";
            else return "";
            
        }

        private void btnAbadanOperation_Click(object sender, EventArgs e)
        {
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();

            var wb = new Workbook(openFileDialog1.FileName);


            List farmList = web.Lists.GetByTitle("ساختار شکست");
            Worksheet ws = wb.Worksheets["Totallll"];

            var cId = int.Parse((comboBox1.SelectedItem as dynamic).Value.ToString());


            // کانال درجه 3( درجا)
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 15, 15, "مترطول", 1, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 15, 16, "هکتار", 1, 1, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 15, 17, "مترطول", 1, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 15, 18, "تعداد", 1, 3, cId);


            //نصب کانالت
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 19, 19, "مترطول", 2, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 19, 20, "هکتار", 2, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 19, 21, "مترطول", 2, 4, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 19, 22, "مترطول", 2, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 19, 23, "تعداد", 2, 3, cId);


            //لوله کم فشار
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 24, 24, "مترطول", 3, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 24, 25, "هکتار", 3, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 24, 26, "مترطول", 3, 4, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 24, 27, "مترطول", 3, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 24, 28, "تعداد", 3, 3, cId);
            //  اجرای ایستگاه پمپاژ و نصب تجهیزات
            CreateItem(ws, farmList, 24, 29, "درصد", 3, 5, cId);
            //  تامین سهم پیمانکار (ایستگاه پمپاژ)
            CreateItem(ws, farmList, 24, 30, "درصد", 3, 10, cId);
            //ترمیم و بهسازی شبکه آبیاری-اجرا  (حوزه آبادان)
            CreateItem(ws, farmList, 24, 31, "دستگاه", 3, 21, cId);
            // ترمیم و بهسازی شبکه آبیاری - اجرا(حوزه آبادان)
            CreateItem(ws, farmList, 24, 32, "هکتار", 3, 20, cId);


            //آبیاری تحت فشار
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 33, 33, "مترطول", 4, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 33, 34, "هکتار", 4, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 33, 35, "مترطول", 4, 4, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 33, 36, "مترطول", 4, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 33, 37, "تعداد", 4, 3, cId);
            //  اجرای ایستگاه پمپاژ و نصب تجهیزات
            CreateItem(ws, farmList, 33, 38, "درصد", 4, 5, cId);
            //  تامین سهم پیمانکار (ایستگاه پمپاژ)
            CreateItem(ws, farmList, 33, 39, "درصد", 4, 10, cId);




            // تامین لوله و متعلقات واجرای خط انتقال آب

            //پایپ جکینگ
            CreateItem(ws, farmList, 40, 40, "مترطول", 8, 33, cId);
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 40, 41, "مترطول", 8, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 40, 42, "هکتار", 8, 1, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 40, 43, "مترطول", 8, 4, cId);
            //  سازه
            CreateItem(ws, farmList, 40, 44, "تعداد", 8, 3, cId);


            //زهکش روباز
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 45, 45, "مترطول", 5, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 45, 46, "هکتار", 5, 1, cId);
            // جاده سرویس  
            CreateItem(ws, farmList, 45, 47, "مترطول", 5, 2, cId);
            //  سازه
            CreateItem(ws, farmList, 45, 48, "تعداد", 5, 3, cId);
            //  اجرای ایستگاه پمپاژ و نصب تجهیزات
            CreateItem(ws, farmList, 45, 49, "درصد", 5, 5, cId);
            //  تامین سهم پیمانکار (ایستگاه پمپاژ)
            CreateItem(ws, farmList, 45, 50, "درصد", 5, 10, cId);
            //لایروبی زهکش
            CreateItem(ws, farmList, 45, 51, "مترطول", 5, 22, cId);


            //کلکتور (زهکش جمع کننده لوله ای)
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 52, 52, "مترطول", 9, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 52, 53, "هکتار", 9, 1, cId);
            // تامین سهم پیمانکار
            CreateItem(ws, farmList, 52, 54, "مترطول", 9, 4, cId);
            //  سازه
            CreateItem(ws, farmList, 52, 55, "تعداد", 9, 3, cId);
            //  اجرای ایستگاه پمپاژ و نصب تجهیزات
            CreateItem(ws, farmList, 52, 56, "درصد", 9, 5, cId);
            //  تامین سهم پیمانکار (ایستگاه پمپاژ)
            CreateItem(ws, farmList, 52, 57, "درصد", 9, 10, cId);

            //زهکش های زیرزمینی(لترال) 
            //اجرا         
            // مترطول
            CreateItem(ws, farmList, 58, 58, "مترطول", 6, 1, cId);
            // هکتار
            CreateItem(ws, farmList, 58, 59, "هکتار", 6, 1, cId);
            // تامین سهم پیمانکار
            CreateItem(ws, farmList, 58, 60, "مترطول", 6, 4, cId);
            //  سازه
            CreateItem(ws, farmList, 58, 61, "تعداد", 6, 3, cId);
            // شستشوی زهکش زیرزمینی(لترال)
            CreateItem(ws, farmList, 58, 62, "مترطول", 6, 23, cId);



            //تجهیز و نوسازی
            //نقشه برداری و مشخص کردن محدوده (ناخالص)
            CreateItem(ws, farmList, 63, 63, "هکتار", 7, 34, cId);
            //مین روبی 
            CreateItem(ws, farmList, 63, 64, "هکتار", 7, 35, cId);
            //بوته کنی و تخریب و آماده سازی و پرکردن     
            CreateItem(ws, farmList, 63, 65, "هکتار", 7, 36, cId);
            //پر کردن انهار سنتی جزر و مدی
            CreateItem(ws, farmList, 63, 66, "مترطول", 7, 37, cId);
            //تسطیح نسبی (خالص)
            CreateItem(ws, farmList, 63, 67, "هکتار", 7, 12, cId);
            //تسطیح اساسی (خالص)
            CreateItem(ws, farmList, 63, 68, "هکتار", 7, 13, cId);
            //تجهیز و نوسازی    
            CreateItem(ws, farmList, 63, 69, "هکتار", 7, 6, cId);
            //هندسی سازی
            CreateItem(ws, farmList, 63, 70, "هکتار", 7, 8, cId);
            //یکپارچه سازی
            CreateItem(ws, farmList, 63, 71, "هکتار", 7, 9, cId);
            //تامین سهم پیمانکار
            CreateItem(ws, farmList, 63, 72, "مترطول", 7, 4, cId);
            //کانال درجه 4
            CreateItem(ws, farmList, 63, 73, "مترطول", 7, 11, cId);
            //سازه
            CreateItem(ws, farmList, 63, 74, "تعداد", 7, 3, cId);
            //آبشویی
            CreateItem(ws, farmList, 63, 75, "هکتار", 7, 14, cId);
            //جاده دسترسی بین مزارع
            CreateItem(ws, farmList, 63, 76, "مترطول", 7, 7, cId);
            // خود اجرایی -نی بری(ناخالص)
            CreateItem(ws, farmList, 63, 77, "هکتار", 7, 24, cId);
            //خود اجرایی -شخم و کرت بندی (ناخالص)
            CreateItem(ws, farmList, 63, 78, "هکتار", 7, 25, cId);
            //خود اجرایی -لایروبی (ناخالص)
            CreateItem(ws, farmList, 63, 79, "هکتار", 7, 26, cId);
            // خود اجرایی (ناخالص)
            CreateItem(ws, farmList, 63, 80, "هکتار", 7, 27, cId);
            //  شبکه توزیع آب درون مزارع(ناخالص) - اجرا
            CreateItem(ws, farmList, 63, 81, "مترطول", 7, 28, cId);
            //شبکه توزیع آب درون مزارع(ناخالص) - اجرا
            CreateItem(ws, farmList, 63, 82, "هکتار", 7, 29, cId);
            // شبکه توزیع آب درون مزارع(ناخالص) - تامین سهم پیمانکار
            CreateItem(ws, farmList, 63, 83, "مترطول", 7, 30, cId);

          
            // تامین جت فلاشر 	
            CreateItem(ws, farmList, 84, 84, "ریال", 10, 17, cId);
            //تسهیلات کارفرمایی	
            CreateItem(ws, farmList, 85, 85, "ریال", 10, 16, cId);
            // سایر
            CreateItem(ws, farmList, 86, 86, "ریال", 10, 19, cId);

            //تجهیز و برچیدن کارگاه	
            CreateItem(ws, farmList, 87, 87, "ریال", 10, 15, cId);

            // تسهیل گری اجتماعی
            CreateItem(ws, farmList, 88, 88, "هکتار", 10, 32, cId);
            //خدمات مهندسی	
            CreateItem(ws, farmList, 89, 89, "ریال", 10, 18, cId);
            //  حضور در زمان بهره برداری و تحویل کار به تشکل های آببران
            CreateItem(ws, farmList, 90, 90, "درصد", 10, 31, cId);




            MessageBox.Show("Done!");
        }
    }
}
