namespace Finiquitos
{
    partial class EditarUsuario
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EditarUsuario));
            this.label4 = new System.Windows.Forms.Label();
            this.txtApe_Mat = new System.Windows.Forms.TextBox();
            this.cbxPermisos = new System.Windows.Forms.ComboBox();
            this.btnEditar = new System.Windows.Forms.Button();
            this.txtMail = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtApe_Pat = new System.Windows.Forms.TextBox();
            this.txtNombres = new System.Windows.Forms.TextBox();
            this.txtContrasena = new System.Windows.Forms.TextBox();
            this.txtUsuario = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtRut = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 136);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(84, 13);
            this.label4.TabIndex = 105;
            this.label4.Text = "Apellido Paterno";
            // 
            // txtApe_Mat
            // 
            this.txtApe_Mat.Location = new System.Drawing.Point(103, 161);
            this.txtApe_Mat.Name = "txtApe_Mat";
            this.txtApe_Mat.Size = new System.Drawing.Size(181, 20);
            this.txtApe_Mat.TabIndex = 104;
            this.txtApe_Mat.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApe_Mat_KeyPress);
            // 
            // cbxPermisos
            // 
            this.cbxPermisos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxPermisos.FormattingEnabled = true;
            this.cbxPermisos.Location = new System.Drawing.Point(359, 165);
            this.cbxPermisos.Name = "cbxPermisos";
            this.cbxPermisos.Size = new System.Drawing.Size(181, 21);
            this.cbxPermisos.TabIndex = 103;
            // 
            // btnEditar
            // 
            this.btnEditar.Location = new System.Drawing.Point(18, 213);
            this.btnEditar.Name = "btnEditar";
            this.btnEditar.Size = new System.Drawing.Size(117, 29);
            this.btnEditar.TabIndex = 102;
            this.btnEditar.Text = "Editar Usuario";
            this.btnEditar.UseVisualStyleBackColor = true;
            this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
            // 
            // txtMail
            // 
            this.txtMail.Location = new System.Drawing.Point(359, 192);
            this.txtMail.Name = "txtMail";
            this.txtMail.Size = new System.Drawing.Size(181, 20);
            this.txtMail.TabIndex = 101;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 161);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(86, 13);
            this.label9.TabIndex = 100;
            this.label9.Text = "Apellido Materno";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(52, 105);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(49, 13);
            this.label8.TabIndex = 99;
            this.label8.Text = "Nombres";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(292, 136);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(61, 13);
            this.label5.TabIndex = 98;
            this.label5.Text = "Contraseña";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(310, 109);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 97;
            this.label2.Text = "Usuario";
            // 
            // txtApe_Pat
            // 
            this.txtApe_Pat.Location = new System.Drawing.Point(103, 133);
            this.txtApe_Pat.Name = "txtApe_Pat";
            this.txtApe_Pat.Size = new System.Drawing.Size(181, 20);
            this.txtApe_Pat.TabIndex = 96;
            this.txtApe_Pat.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApe_Pat_KeyPress);
            // 
            // txtNombres
            // 
            this.txtNombres.Location = new System.Drawing.Point(103, 102);
            this.txtNombres.Name = "txtNombres";
            this.txtNombres.Size = new System.Drawing.Size(181, 20);
            this.txtNombres.TabIndex = 95;
            this.txtNombres.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombres_KeyPress);
            // 
            // txtContrasena
            // 
            this.txtContrasena.Location = new System.Drawing.Point(359, 133);
            this.txtContrasena.Name = "txtContrasena";
            this.txtContrasena.Size = new System.Drawing.Size(181, 20);
            this.txtContrasena.TabIndex = 94;
            this.txtContrasena.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtContrasena_KeyPress);
            // 
            // txtUsuario
            // 
            this.txtUsuario.Location = new System.Drawing.Point(359, 102);
            this.txtUsuario.Name = "txtUsuario";
            this.txtUsuario.Size = new System.Drawing.Size(181, 20);
            this.txtUsuario.TabIndex = 93;
            this.txtUsuario.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUsuario_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(327, 192);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 13);
            this.label3.TabIndex = 107;
            this.label3.Text = "Mail";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(304, 168);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(49, 13);
            this.label7.TabIndex = 106;
            this.label7.Text = "Permisos";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(169, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(272, 68);
            this.pictureBox1.TabIndex = 108;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(73, 187);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(24, 13);
            this.label1.TabIndex = 110;
            this.label1.Text = "Rut";
            // 
            // txtRut
            // 
            this.txtRut.Enabled = false;
            this.txtRut.Location = new System.Drawing.Point(103, 187);
            this.txtRut.Name = "txtRut";
            this.txtRut.Size = new System.Drawing.Size(181, 20);
            this.txtRut.TabIndex = 109;
            this.txtRut.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRut_KeyPress);
            // 
            // EditarUsuario
            // 
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.ScrollBar;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(568, 254);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtRut);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtApe_Mat);
            this.Controls.Add(this.cbxPermisos);
            this.Controls.Add(this.btnEditar);
            this.Controls.Add(this.txtMail);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtApe_Pat);
            this.Controls.Add(this.txtNombres);
            this.Controls.Add(this.txtContrasena);
            this.Controls.Add(this.txtUsuario);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "EditarUsuario";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EditarUsuario";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.EditarUsuario_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtApe_Mat;
        private System.Windows.Forms.ComboBox cbxPermisos;
        private System.Windows.Forms.Button btnEditar;
        private System.Windows.Forms.TextBox txtMail;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtApe_Pat;
        private System.Windows.Forms.TextBox txtNombres;
        private System.Windows.Forms.TextBox txtContrasena;
        private System.Windows.Forms.TextBox txtUsuario;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtRut;
    }
}