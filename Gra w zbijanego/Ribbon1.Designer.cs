namespace Gra_w_zbijanego
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Wymagana zmienna projektanta.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Wyczyść wszystkie używane zasoby.
        /// </summary>
        /// <param name="disposing">prawda, jeżeli zarządzane zasoby powinny zostać zlikwidowane; Fałsz w przeciwnym wypadku.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kod wygenerowany przez Projektanta składników

        /// <summary>
        /// Metoda wymagana do obsługi projektanta — nie należy modyfikować
        /// jej zawartości w edytorze kodu.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.Gra = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.GameNewGame = this.Factory.CreateRibbonButton();
            this.Start = this.Factory.CreateRibbonButton();
            this.Gra.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Gra
            // 
            this.Gra.Groups.Add(this.group1);
            this.Gra.Label = "Gra w zbijanego";
            this.Gra.Name = "Gra";
            // 
            // group1
            // 
            this.group1.Items.Add(this.GameNewGame);
            this.group1.Items.Add(this.Start);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // GameNewGame
            // 
            this.GameNewGame.Image = ((System.Drawing.Image)(resources.GetObject("GameNewGame.Image")));
            this.GameNewGame.Label = "Create Game";
            this.GameNewGame.Name = "GameNewGame";
            this.GameNewGame.ShowImage = true;
            this.GameNewGame.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GameNewGame_Click);
            // 
            // Start
            // 
            this.Start.Image = ((System.Drawing.Image)(resources.GetObject("Start.Image")));
            this.Start.Label = "Start Game";
            this.Start.Name = "Start";
            this.Start.ShowImage = true;
            this.Start.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Start_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Gra);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.Gra.ResumeLayout(false);
            this.Gra.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Gra;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GameNewGame;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Start;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
