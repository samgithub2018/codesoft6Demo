using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using LabelManager2;
using System.IO;

namespace TestCodeSoft
{
    public partial class Main : Form
    {

        private int ilablelx;//定义列印类型

        public Main()
        {
            InitializeComponent();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            LabelManager2.Application labelapp = null;
            LabelManager2.Document labDoc = null;
            string filePathName = System.Windows.Forms.Application.StartupPath + "\\Document1.lab";
            if (!File.Exists(filePathName))
            {
                MessageBox.Show("请准备好列印模板");
                return;
            }
            try
            {
                labelapp = new LabelManager2.Application();
                labelapp.Documents.Open(filePathName, false);//调用设计好的模板
                labDoc = labelapp.ActiveDocument;
                labDoc.Variables.FormVariables.Item("Date").Value = System.DateTime.Now.ToString("yyyy-MM-dd");//往变量里面写入数据
                labDoc.PrintDocument(1);//列印份数

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                labelapp.Documents.CloseAll(true);
                labelapp.Quit();
                labelapp = null;
                labDoc = null;
            }

        }
    }
}
