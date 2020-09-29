using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Net;
using System.Windows.Forms;
using System.Xml;

namespace TestUtility
{
	public class Form1 : Form
	{
		private IContainer components = null;

		private Button button1;

		private Label label1;

		private Label label2;

		private CheckedListBox CheckBoxList;

		private Button button2;

		private Label label3;

		private Label label4;

		private Label label5;

		private ListView listView1;

		private ColumnHeader Method;

		private ColumnHeader Request;

		private ColumnHeader Response;

		private ColumnHeader Action;

		private TextBox textBox1;

		private TextBox textBox2;

		private Label label6;

		private Label label7;

		private ColumnHeader Status;

		private TextBox textBox3;
        private Button button3;
        private Label label8;

		private OleDbConnection returnConnection(string fileName)
		{
			return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;");
		}

		public Form1()
		{
			InitializeComponent();
		}

		private void button2_Click(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(textBox3.Text))
			{
				MessageBox.Show("Please Provide the test case file location!!", "Error");
				return;
			}
			Cursor.Current = Cursors.WaitCursor;
			listView1.Items.Clear();
			string strr = CheckBoxList.Items[CheckBoxList.SelectedIndex].ToString();
			LoadDATA(strr);
			Cursor.Current = Cursors.Default;
		}

		private void LoadDATA(string strr)
		{
			try
			{
				DataTable sheetData = new DataTable();
				string _filePath = textBox3.Text + "\\TestCase.xlsx";
				using (OleDbConnection conn = returnConnection(_filePath))
				{
					conn.Open();
					OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("select * from [" + strr + "$]", conn);
					sheetAdapter.Fill(sheetData);
				}
				FillUseCaseList(sheetData);
			}
			catch (Exception)
			{
				MessageBox.Show("Incorrect Path!!", "Error");
			}
		}

		private void FillUseCaseList(DataTable sheetData)
		{
			foreach (DataRow item in sheetData.Rows)
			{
				listView1.Items.Add(new ListViewItem(new string[5]
				{
					item["MethodName"].ToString(),
					item["SoAPRequest"].ToString(),
					"Not Started",
					string.Empty,
					item["SoapAction"].ToString()
				}));
			}
		}

		private void CheckBoxList_MouseClick(object sender, MouseEventArgs e)
		{
			if (string.IsNullOrEmpty(textBox3.Text))
			{
				MessageBox.Show("Please Provide the test case file location!!", "Error");
				for (int j = 0; j < CheckBoxList.Items.Count; j++)
				{
					CheckBoxList.SetItemChecked(j, value: false);
				}
				return;
			}
			listView1.Items.Clear();
			string strr = CheckBoxList.Items[CheckBoxList.SelectedIndex].ToString();
			for (int i = 0; i < CheckBoxList.Items.Count; i++)
			{
				CheckBoxList.SetItemChecked(i, value: false);
			}
			switch (strr)
			{
			case "PROD":
				CheckBoxList.SetItemChecked(CheckBoxList.SelectedIndex, value: true);
				SetEnv(strr);
				break;
			case "INT":
				CheckBoxList.SetItemChecked(CheckBoxList.SelectedIndex, value: true);
				SetEnv(strr);
				break;
			case "SYST":
				CheckBoxList.SetItemChecked(CheckBoxList.SelectedIndex, value: true);
				SetEnv(strr);
				break;
			case "DEV":
				CheckBoxList.SetItemChecked(CheckBoxList.SelectedIndex, value: true);
				SetEnv(strr);
				break;
			}
		}

		private void SetEnv(string strr)
		{
			XmlDocument doc = new XmlDocument();
			string _filePath = textBox3.Text + "\\DATA.xml";
			try
			{
				doc.Load(_filePath);
				string xmlcontents = doc.InnerXml;
				XmlNode node = doc.SelectSingleNode(strr);
				foreach (XmlNode item in doc.SelectNodes("Environments/" + strr))
				{
					foreach (XmlNode _Child in item.ChildNodes)
					{
						if (_Child.Name == "URL")
						{
							label4.Text = _Child.InnerText;
						}
					}
				}
			}
			catch (Exception)
			{
				MessageBox.Show("Incorrect Path!!", "Error");
			}
		}

		private void CheckBoxList_SelectedIndexChanged(object sender, EventArgs e)
		{
			CheckBoxList.SetItemChecked(CheckBoxList.SelectedIndex, value: true);
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Cursor.Current = Cursors.WaitCursor;
			foreach (ListViewItem item2 in listView1.Items)
			{
				item2.SubItems[2].Text = "Not Started";
			}
			foreach (ListViewItem item in listView1.Items)
			{
				try
				{
					using WebClient client = new WebClient();
					string data = item.SubItems[1].Text;
					client.Headers.Add("Content-Type", "text/xml;charset=utf-8");
					client.Headers.Add("SOAPAction", item.SubItems[4].Text);
					string response = client.UploadString(label4.Text, data);
					item.SubItems[2].Text = "Success";
					item.SubItems[3].Text = response;
					item.SubItems[2].ForeColor = Color.Green;
				}
				catch (Exception ex)
				{
					item.SubItems[2].Text = "Failed";
					item.SubItems[3].Text = ex.ToString();
					item.SubItems[2].ForeColor = Color.Red;
				}
			}
			Cursor.Current = Cursors.Default;
		}

		private void listView1_SelectedIndexChanged(object sender, EventArgs e)
		{
			textBox1.Text = string.Empty;
			textBox2.Text = string.Empty;
			foreach (ListViewItem item in listView1.SelectedItems)
			{
				textBox1.Text = item.SubItems[1].Text;
				textBox2.Text = item.SubItems[3].Text;
			}
		}

		private void textBox1_MouseDoubleClick(object sender, MouseEventArgs e)
		{
			textBox1.SelectAll();
		}

		private void textBox2_MouseDoubleClick(object sender, MouseEventArgs e)
		{
			textBox2.SelectAll();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
            System.Windows.Forms.ListViewGroup listViewGroup2 = new System.Windows.Forms.ListViewGroup("List item text", System.Windows.Forms.HorizontalAlignment.Left);
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CheckBoxList = new System.Windows.Forms.CheckedListBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.Method = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Request = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Status = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Response = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Action = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(366, 510);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 34);
            this.button1.TabIndex = 0;
            this.button1.Text = "ExecuteAll";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(362, 129);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(180, 19);
            this.label1.TabIndex = 2;
            this.label1.Text = "List Of Service methods";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(50, 180);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(182, 19);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select The Environment";
            // 
            // CheckBoxList
            // 
            this.CheckBoxList.FormattingEnabled = true;
            this.CheckBoxList.Items.AddRange(new object[] {
            "PROD",
            "SYST",
            "DEV",
            "INT"});
            this.CheckBoxList.Location = new System.Drawing.Point(54, 203);
            this.CheckBoxList.Margin = new System.Windows.Forms.Padding(4);
            this.CheckBoxList.Name = "CheckBoxList";
            this.CheckBoxList.Size = new System.Drawing.Size(188, 92);
            this.CheckBoxList.TabIndex = 5;
            this.CheckBoxList.MouseClick += new System.Windows.Forms.MouseEventHandler(this.CheckBoxList_MouseClick);
            this.CheckBoxList.SelectedIndexChanged += new System.EventHandler(this.CheckBoxList_SelectedIndexChanged);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(52, 328);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(190, 34);
            this.button2.TabIndex = 6;
            this.button2.Text = "Load The Test Case";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(362, 85);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 19);
            this.label3.TabIndex = 7;
            this.label3.Text = "URL :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(415, 85);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 19);
            this.label4.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Bahnschrift", 14F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label5.Location = new System.Drawing.Point(435, 9);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(282, 23);
            this.label5.TabIndex = 9;
            this.label5.Text = "Claims Web Service Testing Tool";
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Method,
            this.Request,
            this.Status,
            this.Response,
            this.Action});
            this.listView1.Font = new System.Drawing.Font("Bahnschrift", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listView1.FullRowSelect = true;
            listViewGroup2.Header = "List item text";
            listViewGroup2.Name = null;
            this.listView1.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup2});
            this.listView1.HideSelection = false;
            this.listView1.LabelEdit = true;
            this.listView1.Location = new System.Drawing.Point(366, 162);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(737, 318);
            this.listView1.TabIndex = 10;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            // 
            // Method
            // 
            this.Method.Text = "Method";
            this.Method.Width = 150;
            // 
            // Request
            // 
            this.Request.Text = "Request";
            this.Request.Width = 400;
            // 
            // Status
            // 
            this.Status.Text = "Status";
            this.Status.Width = 100;
            // 
            // Response
            // 
            this.Response.Text = "Response";
            this.Response.Width = 1;
            // 
            // Action
            // 
            this.Action.Text = "Action";
            this.Action.Width = 1;
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(579, 486);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(201, 216);
            this.textBox1.TabIndex = 12;
            this.textBox1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.textBox1_MouseDoubleClick);
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(903, 486);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(200, 216);
            this.textBox2.TabIndex = 13;
            this.textBox2.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.textBox2_MouseDoubleClick);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(504, 483);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 19);
            this.label6.TabIndex = 14;
            this.label6.Text = "Request";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(797, 483);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 19);
            this.label7.TabIndex = 15;
            this.label7.Text = "Response";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(54, 129);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(190, 27);
            this.textBox3.TabIndex = 16;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(50, 85);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(198, 19);
            this.label8.TabIndex = 17;
            this.label8.Text = "Provide the TestCase Path";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(54, 386);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(188, 30);
            this.button3.TabIndex = 18;
            this.button3.Text = "Get WLB Reports";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1115, 749);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.CheckBoxList);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "ClaimServiceTester";
            this.ResumeLayout(false);
            this.PerformLayout();

		}

        private void button3_Click(object sender, EventArgs e)
        {
            WLBReport obj = new WLBReport();
            obj.ShowDialog();
        }
    }
}
