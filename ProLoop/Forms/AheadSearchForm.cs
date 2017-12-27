﻿using Newtonsoft.Json;
using ProLoop.WordAddin.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProLoop.WordAddin.Forms
{
    public partial class AheadSearchForm : Form
    {
        DataTable dt = new DataTable();
        public string FilePath { get; set; }
        public AheadSearchForm()
        {
            InitializeComponent();
        }

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            //string rowFilter = string.Format("{0} Like '{1}%' OR {2} Like '{3}%' OR {4} Like '{5}%' OR {6} Like '{7}%'", "FileName", textBoxSearch.Text,
            //   "Author",textBoxSearch.Text,"Client", textBoxSearch.Text, "Matter", textBoxSearch.Text);
            ////(dataGridView1.DataSource as DataTable).DefaultView.RowFilter = rowFilter;
            //dt.DefaultView.RowFilter = rowFilter;
            ProcessAutoComplete();
        }

        private void ProcessAutoComplete()
        {
            var ObjectList = new List<MetaDataInfo>();
            using (var client = new WebClient())
            {
                if (!string.IsNullOrEmpty(textBoxPath.Text))
                {
                    string url = string.Format("{0}/api/sayt/f/{1}?keywords={2}&s={3}",AddinModule.CurrentInstance.ProLoopUrl, textBoxPath.Text, textBoxKeywords.Text, textBoxSearch.Text);
                    var jsonstring = client.DownloadString(url);
                    if (!string.IsNullOrEmpty(jsonstring))
                        ObjectList = JsonConvert.DeserializeObject<List<MetaDataInfo>>(jsonstring);
                }
                dataGridView1.AllowUserToAddRows = true;
                dataGridView1.DataSource = ObjectList;
                dataGridView1.AllowUserToAddRows = false;
            }
        }
        private void AheadSearchForm_Load(object sender, EventArgs e)
        {
           
        }

        private void DummyFilterr()
        {
            dt.Columns.Add("FileName", typeof(string));
            dt.Columns.Add("Author", typeof(string));
            dt.Columns.Add("Database", typeof(string));
            dt.Columns.Add("Client", typeof(string));
            dt.Columns.Add("Matter", typeof(string));
            dt.Columns.Add("DocNum", typeof(string));
            dt.Columns.Add("Version", typeof(string));
            dt.Columns.Add("Editdate", typeof(string));

            var list = new List<DataItem>();
            list.Add(new DataItem()
            {
                Author = "RTech",
                Client = "0110",
                Database = "DMVS",
                DocuNm = "45444",
                FileName = "Marketing.doc",
                Matter = "304",
                Version = "2",
                Editdate = DateTime.Now.AddMonths(-2).ToString("MM/dd/yyyy h:mm tt")
            });

            list.Add(new DataItem()
            {
                Author = "JRT",
                Client = "0130",
                Database = "DMVS",
                DocuNm = "45445",
                FileName = "Helper.doc",
                Matter = "345",
                Version = "1.9",
                Editdate = DateTime.Now.AddMonths(-1).ToString("MM/dd/yyyy h:mm tt")
            });

            list.Add(new DataItem()
            {
                Author = "JRT",
                Client = "0130",
                Database = "DMVS",
                DocuNm = "45445",
                FileName = "Meta.doc",
                Matter = "345",
                Version = "1.4",
                Editdate = DateTime.Now.AddMonths(-3).ToString("MM/dd/yyyy h:mm tt")
            });
            list.Add(new DataItem()
            {
                Author = "NRT",
                Client = "01102",
                Database = "KVCP",
                DocuNm = "45445",
                FileName = "Works.doc",
                Matter = "07672",
                Version = "2.8",
                Editdate = DateTime.Now.AddDays(-2).ToString("MM/dd/yyyy h:mm tt")
            });

            list.Add(new DataItem()
            {
                Author = "KMC",
                Client = "01200",
                Database = "KVCP",
                DocuNm = "23461",
                FileName = "Food.doc",
                Matter = "7878",
                Version = "3.0",
                Editdate = DateTime.Now.AddDays(-4).ToString("MM/dd/yyyy h:mm tt")
            });

            list.Add(new DataItem()
            {
                Author = "FVC",
                Client = "0130",
                Database = "DMVS",
                DocuNm = "10092",
                FileName = "Prediction.doc",
                Matter = "4763",
                Version = "3.1",
                Editdate = DateTime.Now.AddDays(-8).ToString("MM/dd/yyyy h:mm tt")
            });

            list.Add(new DataItem()
            {
                Author = "JRT",
                Client = "01100",
                Database = "DMVS",
                DocuNm = "11909",
                FileName = "Plans.doc",
                Matter = "8798",
                Version = "2.3",
                Editdate = DateTime.Now.AddDays(-3).ToString("MM/dd/yyyy h:mm tt")
            });
            foreach (var item in list)
            {

                dt.Rows.Add(new object[] { item.FileName,item.Author, item.Database , item.Client , item.Matter, item.DocuNm,
                item.Version,item.Editdate
            });
            }
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.DataSource = dt;
            dataGridView1.AllowUserToAddRows = false;

        }

        private void textBoxKeywords_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void textBoxPath_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int index = e.RowIndex;
            if(index>-1)
            {
                FilePath = dataGridView1.Rows[index].Cells[1].Value as string;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

        }
    }
}
