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
          


        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            BuildMainTree();

            comboBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
            comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection x = new AutoCompleteStringCollection();
            foreach (var user in DataHelp.UserAll)
            {

                x.Add(user.Name);
            }


            comboBox1.AutoCompleteCustomSource = x;
        }

        private void BuildMainTree()
        {
            treeView1.Nodes.Clear();
           TreeNode root = new TreeNode("一代");
            treeView1.Nodes.Add(root);
            dataHelp = new DataHelp();
            List<BOM> rBoms = dataHelp.GetRootBom();
            foreach (var ritem in rBoms)
            {
                TreeNode tBom = new TreeNode(ritem.ToString());
                tBom.Tag = ritem;
                root.Nodes.Add(tBom);
            }

            treeView1.ExpandAll();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            return;
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
            treeView2.SelectedNode = treeView2.Nodes[0];

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
                item.PBom = bom;
                CreateTree(item, sNode);
            }

        }

        private static TreeNode GetBomNode(BOM bom)
        {
            TreeNode sNode = new TreeNode(bom.ToString());
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
            comboBox3.Text = bom.Node;
                
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
          
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (treeView2.SelectedNode == null) return;
            var bom = treeView2.SelectedNode.Tag as BOM;
            bom.ID = comboBox2.Text;
            bom.Name = comboBox1.Text;
            bom.Node = comboBox3.Text;
            treeView2.SelectedNode.Text = bom.ToString();
            RefrenceView1();
        }

        private void toolStripButtonSave_Click(object sender, EventArgs e)
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
            Microsoft.Office.Interop.Word.Document t  = application.Documents.Open("e:\\aa.docx");
            t.Shapes.AddLine(0, 0, 200, 100);
            Microsoft.Office.Interop.Word.Shape  s=t.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 100, 100);
            s.TextFrame.TextRange.Text = "fdsafsadfsadfsa";
            Microsoft.Office.Interop.Word.Shape ss = t.Shapes.AddShape(9, 150, 100, 100, 50);
        
            ss.BackgroundStyle = Microsoft.Office.Core.MsoBackgroundStyleIndex.msoBackgroundStylePreset1;
            ss.TextFrame.TextRange.Text = "fdsafsadfsadfsa的的的的";
           
            application.Quit();
         


        }

        private void toolStripButtonCreateSon_Click(object sender, EventArgs e)
        {
            if (treeView2.SelectedNode == null) return;

            var node = GetBomNode(new BOM() { Name = "name", ID = "id" });
            var pBom = treeView2.SelectedNode.Tag as BOM;
            pBom.AddSon(node.Tag as BOM);
            treeView2.SelectedNode.Nodes.Add(node);
            treeView2.SelectedNode = node;

        }

        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (string.IsNullOrWhiteSpace(comboBox1.Text))
                {
                    return;
                }
                if (comboBox1.Text.Length <= 1)
                    return;

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

        private void 添加ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var bom = new BOM() { Name = "name", ID = "id" };
            var node = GetBomNode(bom);
            treeView1.Nodes[0].Nodes.Add(node);
            DataHelp.RootBom.Add(bom);
            treeView1.SelectedNode = node;
            treeView2.SelectedNode = treeView2.Nodes[0];


        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null) return;
            if (treeView1.SelectedNode.Tag == null) return;
            var bom = treeView1.SelectedNode.Tag as BOM;
            if (MessageBox.Show(string.Format("要删除 '{0}'", bom.Name), "", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                DataHelp.RootBom.Remove(treeView1.SelectedNode.Tag as BOM);
                BuildMainTree();
            }

        }
        private void RefrenceView1()
        {

            if (treeView1.SelectedNode == null) return;
            if (treeView1.SelectedNode.Tag == null) return;
            var bom = treeView1.SelectedNode.Tag as BOM;
            treeView1.SelectedNode.Text = bom.ToString();
        }


        private void contextMenuStrip2_Opening(object sender, CancelEventArgs e)
        {

        }

        private void 删除ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (treeView2.SelectedNode == null) return;
            if (treeView2.SelectedNode.Tag == null) return;
            var bom = treeView2.SelectedNode.Tag as BOM;
            if (MessageBox.Show(string.Format("要删除 '{0}'", bom.Name), "", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                treeView2.SelectedNode.Remove();
                bom.Remove();

            }

        }

        private void 导出EXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null) return;
            if (treeView1.SelectedNode.Tag == null) return;
            var bom = treeView1.SelectedNode.Tag as BOM;
             
            ReportData bomData = new ReportData();
            bomData.TemplateName = bom.Name;
            SetReportBOM(bomData, bom);

            using (Report report = new Report())
            {
              
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

        private void SetReportBOM(ReportData rd,BOM bom,int level=0)
        {
            BomData bd = new BomData();
            bd.Level = level;
            bd.Bom = bom;
            rd.BomData.Add(bd);
            foreach (var sitem in bom.Son)
            {
                SetReportBOM(rd, sitem, level + 1);
            }
        }
    }
}
