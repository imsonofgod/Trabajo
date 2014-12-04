namespace Finiquitos
{
    partial class PrincipalMDI
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PrincipalMDI));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.liquidacionesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.formularioLiquidacionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.administracionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.itemGestion = new System.Windows.Forms.ToolStripMenuItem();
            this.ayudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verLaAyudaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.liquidacionesToolStripMenuItem,
            this.administracionToolStripMenuItem,
            this.ayudaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(765, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // liquidacionesToolStripMenuItem
            // 
            this.liquidacionesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.formularioLiquidacionToolStripMenuItem});
            this.liquidacionesToolStripMenuItem.Name = "liquidacionesToolStripMenuItem";
            this.liquidacionesToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.liquidacionesToolStripMenuItem.Text = "Tareas";
            // 
            // formularioLiquidacionToolStripMenuItem
            // 
            this.formularioLiquidacionToolStripMenuItem.Name = "formularioLiquidacionToolStripMenuItem";
            this.formularioLiquidacionToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.formularioLiquidacionToolStripMenuItem.Text = "Formulario Liquidacion";
            this.formularioLiquidacionToolStripMenuItem.Click += new System.EventHandler(this.formularioLiquidacionToolStripMenuItem_Click);
            // 
            // administracionToolStripMenuItem
            // 
            this.administracionToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.itemGestion});
            this.administracionToolStripMenuItem.Name = "administracionToolStripMenuItem";
            this.administracionToolStripMenuItem.Size = new System.Drawing.Size(100, 20);
            this.administracionToolStripMenuItem.Text = "Administracion";
            this.administracionToolStripMenuItem.Visible = false;
            // 
            // itemGestion
            // 
            this.itemGestion.Name = "itemGestion";
            this.itemGestion.Size = new System.Drawing.Size(178, 22);
            this.itemGestion.Text = "Gestion de Usuarios";
            this.itemGestion.Click += new System.EventHandler(this.itemGestion_Click);
            // 
            // ayudaToolStripMenuItem
            // 
            this.ayudaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.verLaAyudaToolStripMenuItem});
            this.ayudaToolStripMenuItem.Name = "ayudaToolStripMenuItem";
            this.ayudaToolStripMenuItem.Size = new System.Drawing.Size(53, 20);
            this.ayudaToolStripMenuItem.Text = "Ayuda";
            // 
            // verLaAyudaToolStripMenuItem
            // 
            this.verLaAyudaToolStripMenuItem.Name = "verLaAyudaToolStripMenuItem";
            this.verLaAyudaToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.verLaAyudaToolStripMenuItem.Text = "Ver la Ayuda";
            this.verLaAyudaToolStripMenuItem.Click += new System.EventHandler(this.verLaAyudaToolStripMenuItem_Click);
            // 
            // PrincipalMDI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(765, 503);
            this.Controls.Add(this.menuStrip1);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "PrincipalMDI";
            this.Text = "Menu";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.PrincipalMDI_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem administracionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem itemGestion;
        private System.Windows.Forms.ToolStripMenuItem liquidacionesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem formularioLiquidacionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ayudaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verLaAyudaToolStripMenuItem;
    }
}