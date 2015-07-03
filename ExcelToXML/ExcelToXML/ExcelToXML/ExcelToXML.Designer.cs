namespace ExcelToXML
{
    partial class ExcelToXML : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelToXML()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.About = this.Factory.CreateRibbonButton();
            this.Button_XML = this.Factory.CreateRibbonButton();
            this.Key_R = this.Factory.CreateRibbonEditBox();
            this.Key_C = this.Factory.CreateRibbonEditBox();
            this.Value_R = this.Factory.CreateRibbonEditBox();
            this.Value_C = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Key_R);
            this.group1.Items.Add(this.Key_C);
            this.group1.Items.Add(this.Button_XML);
            this.group1.Items.Add(this.Value_R);
            this.group1.Items.Add(this.Value_C);
            this.group1.Items.Add(this.About);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // About
            // 
            this.About.Label = "About";
            this.About.Name = "About";
            this.About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.About_Click);
            // 
            // Button_XML
            // 
            this.Button_XML.Label = "ExcelToXML";
            this.Button_XML.Name = "Button_XML";
            this.Button_XML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_XML_Click);
            // 
            // Key_R
            // 
            this.Key_R.Label = "Key(行R)";
            this.Key_R.Name = "Key_R";
            this.Key_R.Text = null;
            // 
            // Key_C
            // 
            this.Key_C.Label = "Key(列C)";
            this.Key_C.Name = "Key_C";
            this.Key_C.Text = null;
            // 
            // Value_R
            // 
            this.Value_R.Label = "Value(行R)";
            this.Value_R.Name = "Value_R";
            this.Value_R.Text = null;
            // 
            // Value_C
            // 
            this.Value_C.Label = "Value(列C)";
            this.Value_C.Name = "Value_C";
            this.Value_C.Text = null;
            // 
            // ExcelToXML
            // 
            this.Name = "ExcelToXML";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelToXML_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton About;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_XML;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Key_R;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Key_C;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Value_R;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Value_C;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelToXML ExcelToXML
        {
            get { return this.GetRibbon<ExcelToXML>(); }
        }
    }
}
