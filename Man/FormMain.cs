using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Man
{
    public partial class FormMain : Form
    {
        DataHelp dataHelp;
        public FormMain()
        {
            InitializeComponent();
        }


        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(comboBox1.Text))
                {
                    return;
                }

                comboBox1.BeginUpdate();

                comboBox1.Items.Clear();

                var findUser = dataHelp.GetUser(comboBox1.Text);
                foreach (var item in findUser)
                {
                    comboBox1.Items.Add(item.Name);

                }
                if (comboBox1.Items.Count != 0)
                {

                    comboBox1.DroppedDown = true;
                }
                else
                {
                    // comboBox1.DroppedDown = false;
                }
                comboBox1.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            TreeNode root = new TreeNode("人员");
            treeView1.Nodes.Add(root);
            dataHelp = new DataHelp();
            List<BOM> rBoms = dataHelp.GetRootBom();
            foreach (var ritem in rBoms)
            {
                TreeNode tBom = new TreeNode(ritem.Name);
                tBom.Tag = ritem;
                root.Nodes.Add(tBom);
            }

            treeView1.ExpandAll();
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {


        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(comboBox1.Text))
                {
                    return;
                }
                comboBox2.Items.Clear();
                comboBox2.BeginUpdate();

                var findUser = dataHelp.GetUserForID(comboBox1.Text);
                foreach (var item in findUser)
                {
                    comboBox2.Items.Add(item.Id);
                }
                if (comboBox2.Items.Count > 0)
                {
                    comboBox2.SelectedIndex = 0;
                }

                comboBox2.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag == null) return;
            treeView2.Nodes.Clear();
            CreateTree(e.Node.Tag as BOM, null, true);
            treeView2.ExpandAll();

        }
        void CreateTree(BOM bom, TreeNode pNode, bool isRoot = false)
        {
            TreeNode sNode = GetBomNode(bom);
            if (isRoot)
            {
                treeView2.Nodes.Add(sNode);
            }
            else
            {
                pNode.Nodes.Add(sNode);
            }
            foreach (var item in bom.Son)
            {

                CreateTree(item, sNode);
            }

        }

        private static TreeNode GetBomNode(BOM bom)
        {
            TreeNode sNode = new TreeNode(bom.Name);
            sNode.Tag = bom;
            return sNode;
        }

        private void treeView2_BeforeCheck(object sender, TreeViewCancelEventArgs e)
        {

        }

        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag == null) return;
            var bom = e.Node.Tag as BOM;
            comboBox1.Text = bom.Name;
            comboBox2.Text = bom.ID;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (treeView2.SelectedNode == null) return;

            var node = GetBomNode(new BOM() { Name = "name", ID = "id" });
            var pBom = treeView2.SelectedNode.Tag as BOM;
            pBom.Son.Add(node.Tag as BOM);
            treeView2.SelectedNode.Nodes.Add(node);
            treeView2.SelectedNode = node;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (treeView2.SelectedNode == null) return;
            var bom = treeView2.SelectedNode.Tag as BOM;
            bom.ID = comboBox2.Text;
            bom.Name = comboBox1.Text;
            treeView2.SelectedNode.Text = bom.Name;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                dataHelp.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void 报表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (Report report = new Report())
            {
                var bomData = GetReportData();

                System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                //fileDialog.CreatePrompt = true;
                fileDialog.OverwritePrompt = true;

                fileDialog.FileName = string.Format("{0}.xlsx", bomData.TemplateName);
                fileDialog.DefaultExt = "xlsx";
                fileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx";
                // fileDialog.InitialDirectory = "c:\\";
                fileDialog.RestoreDirectory = true;
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        report.Create(bomData);
                        report.Save(fileDialog.FileName);
                        System.Diagnostics.Process.Start(fileDialog.FileName);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }


                }
            }
        }
        public ReportData GetReportData()
        {
            ReportData rd = new ReportData();
            rd.TemplateName = "aaaa";

            return rd;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //先在模板中设定格式

            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            application.Visible = true;
            Microsoft.Office.Interop.Word.Document t  = application.Documents.Open("d:\\aa.docx");
            t.Shapes.AddLine(0, 0, 200, 100);
            Microsoft.Office.Interop.Word.Shape  s=t.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 100, 100);
            s.TextFrame.TextRange.Text = "fdsafsadfsadfsa";
            Microsoft.Office.Interop.Word.Shape ss = t.Shapes.AddShape(9, 150, 100, 100, 50);
        
            ss.BackgroundStyle = Microsoft.Office.Core.MsoBackgroundStyleIndex.msoBackgroundStylePreset1;
            ss.TextFrame.TextRange.Text = "fdsafsadfsadfsa的的的的";
            //ss.TextFrame.TextRange. = "fdsafsadfsadfsa的的的的";

           // ss.TextFrame.TextRange.

            application.Quit();
         


        }
    }
}
