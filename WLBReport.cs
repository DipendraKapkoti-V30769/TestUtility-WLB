using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestUtility
{
    public partial class WLBReport : Form
    {

        DataTable EDWdt = new DataTable();
        DataTable HISdt = new DataTable();
        DataTable OPSdt = new DataTable();
        DataTable Finaldt = new DataTable();
        List<WLB> singOneJoin = new List<WLB>();
        List<WLB> singOneJoinFinal = new List<WLB>();
        List<OPS> OpsFinal = new List<OPS>();
        string GroupNumberInline = string.Empty;
        private OleDbConnection returnConnection(string fileName)
        {
            return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;");
        }
        public WLBReport()
        {
            InitializeComponent();
        }
        private void WLBReport_Load(object sender, EventArgs e)
        {
            textBox1.Text = @"C:\Users\V30769\TestProjects\WellNess\WLBReport.xlsx";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<string> obj = new List<string>() { "EDW", "HIS", "OPS" };
            getSearchData();
            //get data from excel...
            //EDWdt = LoadDATA("EDW");
            //HISdt = LoadDATA("HIS");
            //OPSdt = LoadDATA("OPS");

            Parallel.Invoke(() =>
                             {
                                 EDWdt = LoadDATA("EDW");

                             }, //close second Action
                             () =>
                             {
                                 OPSdt = LoadDATA("OPS");

                             }, //close second Action

                            () =>
                            {
                                HISdt = LoadDATA("HIS");

                            } //close third Action
                        ); //close parallel.invoke


            label1.Text = "Total EDW: " + EDWdt.Rows.Count;
            label4.Text = "Total OPS: " + OPSdt.Rows.Count;
            label2.Text = "Total HIS: " + HISdt.Rows.Count;

            OpsFinal = GetList();
            //DataView dv = new DataView(EDWdt);
            //dv.Sort = "dgrp_groupNumber_no,Policy, PlanCode";
            //EDWdt = dv.Table;

            //DataView dv1 = new DataView(HISdt);
            //dv1.Sort = "dgrp_groupNumber_no,Policy, PlanCode";
            //HISdt = dv1.Table;

            //get lob in sync.
            ValidateDataset();

            //Remove Paid/deny/Pend claim from the 2'nd table.
            RemoveThePaidClaims();
            Finaldt = ConvertToDataTable(singOneJoinFinal);
            label3.Text = "Total Final: " + Finaldt.Rows.Count;
            dataGridView1.DataSource = Finaldt;

        }

        private void getSearchData()
        {
            //  GroupNumberInline = textBox2.Text;

            StringBuilder sb = new StringBuilder();

            foreach (var item in textBox2.Lines)
            {
                if (!string.IsNullOrWhiteSpace(item))
                    GroupNumberInline += ",'" + item + "'";
            }

            GroupNumberInline = GroupNumberInline.Substring(1);
        }

        private List<OPS> GetList()
        {

            //OpsFinal = (from his in HISdt.AsEnumerable()
            //            join edw in OPSdt.AsEnumerable()
            //                    on new
            //                    {
            //                        Policy = his.Field<string>("Policy").Trim(),
            //                        PlanCode = his.Field<string>("PlanCode").Trim(),
            //                        Group = his.Field<string>("GroupNo").Trim()
            //                    } equals new
            //                    {
            //                        Policy = edw.Field<string>("Policy").Trim(),
            //                        PlanCode = edw.Field<string>("PlanCode").Trim(),
            //                        Group = edw.Field<string>("GroupNo").Trim()
            //                    }
            //            into tempJoin
            //            from temp in tempJoin.DefaultIfEmpty()
            //            select new OPS
            //            {
            //                Policy = his.Field<string>("Policy").Trim(),
            //                PlanCode = his.Field<string>("PlanCode").Trim(),
            //                Group = his.Field<string>("GroupNo").Trim(),
            //                Dagt_StateOp_nm = temp.Field<string>("Dagt_StateOp_nm").Trim()

            //            }).Distinct().ToList();



            for (int i = 0; i < OPSdt.Rows.Count; i++)
            {
                OpsFinal.Add(
                    new OPS()
                    {
                        Dagt_StateOp_nm = OPSdt.Rows[i].Field<string>("Dagt_StateOp_nm"),
                        PlanCode = OPSdt.Rows[i].Field<string>("PlanCode"),
                        Policy = OPSdt.Rows[i].Field<string>("Policy"),
                        Group = OPSdt.Rows[i].Field<string>("GroupNo"),

                    });

            }

            return OpsFinal;
        }

        private void RemoveThePaidClaims()
        {
            List<WLB> locallist = new List<WLB>();
            // group by Year
            var query = (from row in EDWdt.AsEnumerable()
                         group row by new
                         {
                             Year = row.Field<Int32>("Year")
                             //PlanCode = row.Field<string>("PlanCode").Trim(),
                             //Policy = row.Field<string>("Policy").Trim(),
                             //IssueState = row.Field<string>("IssueState").Trim()
                             // Year = row.Field<DateTime?>("TreatmentDate") == null ? DateTime.MinValue.Year : row.Field<DateTime>("TreatmentDate").Year
                         } into grp
                         select new
                         {
                             Year = grp.Key.Year,
                             Count = grp.Count()
                         }).ToList();

            foreach (var item in query.Where(k => k.Count > 1))
            {

                locallist = (from his in HISdt.AsEnumerable()
                             join edw in EDWdt.AsEnumerable().Where(h => h.Field<Int32>("Year") == item.Year)
                                     on new
                                     {
                                         Policy = his.Field<string>("Policy").Trim(),
                                         PlanCode = his.Field<string>("PlanCode").Trim(),
                                         Group = his.Field<string>("GroupNo").Trim()
                                     } equals new
                                     {
                                         Policy = edw.Field<string>("Policy").Trim(),
                                         PlanCode = edw.Field<string>("PlanCode").Trim(),
                                         Group = edw.Field<string>("GroupNo").Trim()
                                     }
                             into tempJoin
                             from leftJoin in tempJoin.DefaultIfEmpty()
                             select new WLB
                             {
                                 Policy = his.Field<string>("Policy").Trim(),
                                 PlanCode = his.Field<string>("PlanCode").Trim(),
                                 Group = his.Field<string>("GroupNo").Trim(),
                                 AGENT_NUMBER = his.Field<string>("AGENT_NUMBER"),
                                 Insured = his.Field<string>("Insured"),
                                 EffectiveDate = his.Field<DateTime>("EffectiveDate"),
                                 LOB = his.Field<string>("LOB"),
                                 IssueState = his.Field<string>("IssueState"),
                                 CoverageType = his.Field<string>("CoverageType"),
                                 Year = item.Year,
                                 CMAG_AGENTNO_ID = his.Field<string>("CMAG_AGENTNO_ID"),
                                 CMAG_LEVEL = his.Field<string>("CMAG_LEVEL"),
                                 CIF = his.Field<decimal>("CIF"),
                                 ClaimStatus = leftJoin == null ? "WNF" : leftJoin.Field<string>("Policy")

                             }).Where(k => k.ClaimStatus == "WNF").Distinct().ToList();

                singOneJoin.AddRange(locallist);
            }

            //for (int i = 0; i < HISdt.Rows.Count; i++)
            //{
            //    DataRow drDup = HISdt.Rows[i];

            //    if (CheckPaidClaims(drDup))
            //    {
            //        drDup.Delete();
            //        HISdt.AcceptChanges();
            //    }
            //}
            //  List<WLB> singOneJoinFinal = new List<WLB>();


            //singOneJoinFinal =
            //                   (from his in singOneJoin
            //                    join edw in OpsFinal on his.Group equals edw.Group into ps
            //                    from p in ps.DefaultIfEmpty()
            //                    where p != null
            //                    select new WLB
            //                    {
            //                        Policy = his.Policy.Trim(),
            //                        PlanCode = his.PlanCode.Trim(),
            //                        Group = his.Group.Trim(),
            //                        AGENT_NUMBER = his.AGENT_NUMBER,
            //                        Insured = his.Insured,
            //                        EffectiveDate = his.EffectiveDate,
            //                        LOB = his.LOB,
            //                        IssueState = his.IssueState,
            //                        CoverageType = his.CoverageType,
            //                        Year = his.Year,
            //                        ClaimStatus = his.ClaimStatus,
            //                        MarketOpsNumber = p == null ? string.Empty : p.Dagt_StateOp_nm

            //                    }).Distinct().ToList();




            //singOneJoinFinal = (from his in singOneJoin
            //                    join edw in OpsFinal on his.Group equals edw.Group
            //                    select new WLB
            //                    {
            //                        Policy = his.Policy.Trim(),
            //                        PlanCode = his.PlanCode.Trim(),
            //                        Group = his.Group.Trim(),
            //                        AGENT_NUMBER = his.AGENT_NUMBER,
            //                        Insured = his.Insured,
            //                        EffectiveDate = his.EffectiveDate,
            //                        LOB = his.LOB,
            //                        IssueState = his.IssueState,
            //                        CoverageType = his.CoverageType,
            //                        Year = his.Year,
            //                        ClaimStatus = his.ClaimStatus,
            //                        MarketOpsNumber = edw.Dagt_StateOp_nm

            //                    }).Distinct().ToList();

            //Parallel.ForEach(singOneJoin, data => {
            //    if (OpsFinal.Any(k => k.Group.Trim() == data.Group.Trim()))
            //        if (!singOneJoinFinal.Contains(data))
            //        {
            //            string ops = OpsFinal.Where(k => k.Group.Trim() == data.Group.Trim()).FirstOrDefault().Dagt_StateOp_nm.Trim();
            //            singOneJoinFinal.Add(new WLB()
            //            {
            //                AGENT_NUMBER = data.AGENT_NUMBER,
            //                ClaimStatus = data.ClaimStatus,
            //                CoverageType = data.CoverageType,
            //                EffectiveDate = data.EffectiveDate,
            //                Group = data.Group,
            //                Insured = data.Insured,
            //                IssueState = data.IssueState,
            //                LOB = data.LOB,
            //                MarketOpsNumber = ops,
            //                PlanCode = data.PlanCode,
            //                Policy = data.Policy,
            //                Year = data.Year
            //            });
            //        }

            //});


            singOneJoinFinal = (from his in singOneJoin
                                join edw in OpsFinal
                                        on new
                                        {
                                            Policy = his.Policy.Trim(),
                                            PlanCode = his.PlanCode.Trim(),
                                            Group = his.Group.Trim()
                                        } equals new
                                        {
                                            Policy = edw.Policy.Trim(),
                                            PlanCode = edw.PlanCode.Trim(),
                                            Group = edw.Group.Trim()
                                        }
                                into tempJoin
                                from temp in tempJoin.DefaultIfEmpty()
                                select new WLB
                                {
                                    Policy = his.Policy.Trim(),
                                    PlanCode = his.PlanCode.Trim(),
                                    Group = his.Group.Trim(),
                                    AGENT_NUMBER = his.AGENT_NUMBER,
                                    Insured = his.Insured,
                                    EffectiveDate = his.EffectiveDate,
                                    LOB = his.LOB,
                                    IssueState = his.IssueState,
                                    CoverageType = his.CoverageType,
                                    Year = his.Year,
                                    ClaimStatus = his.ClaimStatus,
                                    CMAG_LEVEL = his.CMAG_LEVEL,
                                    CMAG_AGENTNO_ID = his.CMAG_AGENTNO_ID,
                                    CIF = his.CIF,
                                    MarketOpsNumber = temp == null ? "NA" : temp.Dagt_StateOp_nm

                                }).Where(k => k.ClaimStatus != "NA").Distinct().ToList();



            //for (int i = 0; i < singOneJoin.Count; i++)
            //{
            //    if (OpsFinal.Any(k => k.Group.Trim() == singOneJoin[i].Group.Trim() && k.PlanCode.Trim() == singOneJoin[i].PlanCode.Trim()
            //            && k.Policy.Trim() == singOneJoin[i].Policy.Trim()))
            //        if (!singOneJoinFinal.Contains(singOneJoin[i]))
            //        {
            //            //  string ops = OpsFinal.Where(k => k.Group.Trim() == singOneJoin[i].Group.Trim()).FirstOrDefault().Dagt_StateOp_nm.Trim();
            //            //  singOneJoin[i].MarketOpsNumber = ops;
            //            var _res = OpsFinal.Where(k => k.Group.Trim() == singOneJoin[i].Group.Trim() && k.PlanCode.Trim() == singOneJoin[i].PlanCode.Trim()
            //            && k.Policy.Trim() == singOneJoin[i].Policy.Trim());
            //            foreach (var item in _res)
            //            {
            //                singOneJoinFinal.Add(new WLB()
            //                {
            //                    AGENT_NUMBER = singOneJoin[i].AGENT_NUMBER,
            //                    ClaimStatus = singOneJoin[i].ClaimStatus,
            //                    CoverageType = singOneJoin[i].CoverageType,
            //                    EffectiveDate = singOneJoin[i].EffectiveDate,
            //                    Group = singOneJoin[i].Group,
            //                    Insured = singOneJoin[i].Insured,
            //                    IssueState = singOneJoin[i].IssueState,
            //                    LOB = singOneJoin[i].LOB,
            //                    MarketOpsNumber = item.Dagt_StateOp_nm,
            //                    PlanCode = singOneJoin[i].PlanCode,
            //                    Policy = singOneJoin[i].Policy,
            //                    Year = singOneJoin[i].Year
            //                });
            //            }
            //        }
            //}

            singOneJoinFinal = singOneJoinFinal.Where(k => k.MarketOpsNumber != "NA").ToList();
            var qq = (from row in singOneJoinFinal
                      group row by new
                      {
                          MarketOpsNumber = row.MarketOpsNumber,
                          CIF = row.CIF
                          //PlanCode = row.Field<string>("PlanCode").Trim(),
                          //Policy = row.Field<string>("Policy").Trim(),
                          //IssueState = row.Field<string>("IssueState").Trim()
                          // Year = row.Field<DateTime?>("TreatmentDate") == null ? DateTime.MinValue.Year : row.Field<DateTime>("TreatmentDate").Year
                      } into grp
                      select new
                      {
                          MarketOpsNumber = grp.Key.MarketOpsNumber,
                          CIF = grp.Key.CIF,
                          Count = grp.Count()
                      }).ToList();

            //  listBox1.Items.Add();
            //StringBuilder sd = new StringBuilder("OpsNumber | Count");
            //sd.Append("\n");

            List<summary> ff = new List<summary>();
            foreach (var item in qq)
            {

                ff.Add(new summary()
                {

                    MarketOpsNumber = item.MarketOpsNumber,
                    CIF = item.CIF,
                    Count = item.Count
                });

            }
            dataGridView2.DataSource = ConvertToDataTable(ff);

        }



        private void SaveToExcel(List<WLB> singOneJoin)
        {

            if (Finaldt.Rows.Count < 500000)
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(Finaldt, "FinalWLBReport");
                string _filePath = textBox1.Text.Replace("WLBReport", "WNF");
                wb.SaveAs(_filePath);
            }
            else
            {
                int size = Finaldt.Rows.Count;

                List<DataTable> splitdt = SplitTable(Finaldt, size / 2);
                int i = 1;
                foreach (var item in splitdt)
                {
                    XLWorkbook wb = new XLWorkbook();
                    wb.Worksheets.Add(item, "FinalWLBReport");
                    string _filePath = textBox1.Text.Replace("WLBReport", "WNF-" + i);
                    wb.SaveAs(_filePath);
                    i++;
                }

            }

            ////Inizializar Librerias
            //var workbook = new XLWorkbook();
            //workbook.AddWorksheet("WLB");
            //var ws = workbook.Worksheet("WLB");
            ////Recorrer el objecto
            //int row = 1;
            //foreach (var c in singOneJoin.Where(k=>k.Grade== "WNF"))
            //{
            //    //Escribrie en Excel en cada celda
            //    ws.Cell("Group" + row.ToString()).Value = c.Group;
            //    ws.Cell("Policy" + row.ToString()).Value = c.Policy;
            //    ws.Cell("PlanCode" + row.ToString()).Value = c.PlanCode;              
            //    ws.Cell("Insured" + row.ToString()).Value = c.Insured;
            //    ws.Cell("LOB" + row.ToString()).Value = c.LOB;
            //    ws.Cell("AgentNumber" + row.ToString()).Value = c.AGENT_NUMBER;
            //    ws.Cell("ClaimStatus" + row.ToString()).Value = c.ClaimStatus;
            //    ws.Cell("EffectiveDate" + row.ToString()).Value = c.EffectiveDate;
            //    ws.Cell("CoverageType" + row.ToString()).Value = c.CoverageType;
            //    ws.Cell("IssueState" + row.ToString()).Value = c.IssueState;
            //    row++;

            //}
            //workbook.SaveAs(@"C:\Users\V30769\TestProjects\WellNess\Final.xlsx");
        }
        private static List<DataTable> SplitTable(DataTable originalTable, int batchSize)
        {
            List<DataTable> tables = new List<DataTable>();
            int i = 0;
            int j = 1;
            DataTable newDt = originalTable.Clone();
            newDt.TableName = "Table_" + j;
            newDt.Clear();
            foreach (DataRow row in originalTable.Rows)
            {
                DataRow newRow = newDt.NewRow();
                newRow.ItemArray = row.ItemArray;
                newDt.Rows.Add(newRow);
                i++;
                if (i == batchSize)
                {
                    tables.Add(newDt);
                    j++;
                    newDt = originalTable.Clone();
                    newDt.TableName = "Table_" + j;
                    newDt.Clear();
                    i = 0;
                }
            }
            return tables;
        }
        public DataTable ConvertToDataTable<T>(IList<T> data)
        {

            PropertyDescriptorCollection properties =

            TypeDescriptor.GetProperties(typeof(T));

            DataTable table = new DataTable();

            foreach (PropertyDescriptor prop in properties)

                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

            foreach (T item in data)

            {

                DataRow row = table.NewRow();

                foreach (PropertyDescriptor prop in properties)

                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;

                table.Rows.Add(row);

            }

            return table;

        }
        private bool CheckPaidClaims(DataRow drDup)
        {
            string policy = drDup["Policy"].ToString();
            string plan = drDup["PlanCode"].ToString();
            string Group = drDup["dgrp_groupNumber_no"].ToString();

            var groups = EDWdt.AsEnumerable().GroupBy(s => s.Field<string>("col1"));


            return true;
        }

        private void ValidateDataset()
        {
            try
            {
                foreach (DataRow dr in HISdt.Rows)
                {
                    if (Convert.ToString(dr["LOB"]) != null && Convert.ToString(dr["LOB"]).ToUpper().Trim() == "ACC")
                    {
                        dr["LOB"] = "ACCIDENT";
                    }
                    else if (Convert.ToString(dr["LOB"]) != null && Convert.ToString(dr["LOB"]).ToUpper().Trim() == "HOSP")
                    {
                        dr["LOB"] = "HOSPITAL INDEMNITY";
                    }
                    //WelnessDupsInfo.Add(new WellnessInfo()
                    //{
                    //    LOB = Convert.ToString(dr["LOB"]).Trim(),
                    //    PlanCode = Convert.ToString(dr["PlanCode"]).Trim(),
                    //    PolicyNO = Convert.ToString(dr["Policy"]).Trim(),
                    //    Year = string.Empty
                    //});
                }
                HISdt.AcceptChanges();
            }
            catch (Exception ex) { }
        }

        private DataTable LoadDATA(string strr)
        {
            DataTable sheetData = new DataTable();
            try
            {
                //*** Get from EXCEL */////****************
                //string _filePath = textBox1.Text;
                //using (OleDbConnection conn = returnConnection(_filePath))
                //{
                //    conn.Open();
                //    OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("select * from [" + strr + "$]", conn);
                //    sheetAdapter.Fill(sheetData);
                //}


                //*** Get from SQLDB */////****************     

                string Query = string.Empty;
                string connstring = string.Empty;
                switch (strr)
                {
                    case "EDW":
                        Query = GetQuery(strr);
                        connstring = GetEDWConn();
                        sheetData = GetDataFromSql(Query, connstring);
                        break;

                    case "HIS":
                        Query = GetQuery(strr);
                        connstring = GetHISConn();
                        sheetData = GetDataFromSql(Query, connstring);
                        break;

                    case "OPS":
                        Query = GetQuery(strr);
                        connstring = GetEDWConn();
                        sheetData = GetDataFromSql(Query, connstring);
                        break;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Incorrect Path!!", "Error");
            }

            return sheetData;
        }

        private DataTable GetDataFromSql(string Query, string connstring)
        {
            System.Data.DataTable dt = new DataTable("WLB");
            //using (OleDbConnection con = new OleDbConnection(connstring))
            //{
            //    con.Open();
            //    System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter();
            //    da.SelectCommand = new System.Data.OleDb.OleDbCommand(Query, con);
            //    da.Fill(dt);  
            //}

            using (SqlConnection connection = new SqlConnection(connstring))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(Query, connection);
                // Setting command timeout to 1 second  
                command.CommandTimeout = 1200;
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = command;
                da.Fill(dt);
            }

            return dt;
        }

        private string GetQuery(String queryType)
        {
            string searchtext = queryType + ".txt";
            string path = textBox1.Text.Replace("WLBReport.xlsx", searchtext);
            string edw = System.IO.File.ReadAllText(path);
            edw = edw.Replace("@GroupNumber", GroupNumberInline);
            return edw;
        }

        private string GetEDWConn()
        {

            return "Data Source=SQL-Warehouse-PROD.hq.aflac.com;" +
                   "Initial Catalog=DB_EDW;" +
                   "Integrated Security=SSPI; ";


            //return "Data Source=SQL-Warehouse-PROD.hq.aflac.com;" +
            //       "Initial Catalog=DB_EDW;" +
            //       "Integrated Security=SSPI;Provider=SQLNCLI10.1;  ";

        }

        private string GetHISConn()
        {
            return "Data Source=SQL-Warehouse-PROD.hq.aflac.com;" +
                   "Initial Catalog=DB_EnterpriseStaging;" +
                   "Integrated Security=SSPI; ";

            //return "Data Source=SQL-Warehouse-PROD.hq.aflac.com;" +
            //       "Initial Catalog=DB_EnterpriseStaging;" +
            //       "Integrated Security=SSPI; Provider=SQLNCLI10.1; ";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveToExcel(singOneJoinFinal);
        }

        private void textBox2_MouseLeave(object sender, EventArgs e)
        {
            label5.Text = "Enter GroupNo# " + textBox2.Lines.Count();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            label5.Text = "Enter GroupNo# " + textBox2.Lines.Count();
        }
    }

    public class WLB
    {
        public string Policy { get; set; }
        public string PlanCode { get; set; }
        public string Group { get; set; }

        public string Insured { get; set; }
        public DateTime EffectiveDate { get; set; }
        public string LOB { get; set; }
        public string IssueState { get; set; }
        public string CoverageType { get; set; }
        public string ClaimStatus { get; set; }
        // public string Grade { get; set; }

        public string AGENT_NUMBER { get; set; }

        public double Year { get; set; }

        public string MarketOpsNumber { get; set; }

        public string CMAG_LEVEL { get; set; }
        public string CMAG_AGENTNO_ID { get; set; }
        public decimal CIF { get; set; }

    }

    public class OPS
    {
        public string Dagt_StateOp_nm { get; set; }
        public string Policy { get; set; }
        public string PlanCode { get; set; }
        public string Group { get; set; }
    }
    public class summary
    {
        public string MarketOpsNumber { get; set; }
        public decimal CIF { get; set; }

        public int Count { get; set; }

    }
}
