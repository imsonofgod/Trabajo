namespace Finiquitos
{
    partial class GestionUsuario
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GestionUsuario));
            this.usuarioswebBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.desarrolloDataSet = new Finiquitos.desarrolloDataSet();
            this.usuarios_webTableAdapter = new Finiquitos.desarrolloDataSetTableAdapters.usuarios_webTableAdapter();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.txtUsuario = new System.Windows.Forms.TextBox();
            this.txtContrasena = new System.Windows.Forms.TextBox();
            this.txtNombres = new System.Windows.Forms.TextBox();
            this.txtApe_Pat = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMail = new System.Windows.Forms.TextBox();
            this.btnAgregar = new System.Windows.Forms.Button();
            this.cbxPermisos = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtApe_Mat = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtRut = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.usuarioswebBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.desarrolloDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // usuarioswebBindingSource
            // 
            this.usuarioswebBindingSource.DataMember = "usuarios_web";
            this.usuarioswebBindingSource.DataSource = this.desarrolloDataSet;
            // 
            // desarrolloDataSet
            // 
            this.desarrolloDataSet.DataSetName = "desarrolloDataSet";
            this.desarrolloDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // usuarios_webTableAdapter
            // 
            this.usuarios_webTableAdapter.ClearBeforeFill = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ActiveCaption;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 295);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(960, 297);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.click_boton_grid);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(8, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Agregar Usuarios";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(700, 30);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(272, 68);
            this.pictureBox1.TabIndex = 70;
            this.pictureBox1.TabStop = false;
            // 
            // txtUsuario
            // 
            this.txtUsuario.Location = new System.Drawing.Point(354, 85);
            this.txtUsuario.Name = "txtUsuario";
            this.txtUsuario.Size = new System.Drawing.Size(181, 20);
            this.txtUsuario.TabIndex = 5;
            this.txtUsuario.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUsuario_KeyPress);
            // 
            // txtContrasena
            // 
            this.txtContrasena.Location = new System.Drawing.Point(354, 116);
            this.txtContrasena.Name = "txtContrasena";
            this.txtContrasena.Size = new System.Drawing.Size(181, 20);
            this.txtContrasena.TabIndex = 6;
            this.txtContrasena.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtContrasena_KeyPress);
            // 
            // txtNombres
            // 
            this.txtNombres.Location = new System.Drawing.Point(98, 85);
            this.txtNombres.Name = "txtNombres";
            this.txtNombres.Size = new System.Drawing.Size(181, 20);
            this.txtNombres.TabIndex = 1;
            this.txtNombres.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombres_KeyPress);
            // 
            // txtApe_Pat
            // 
            this.txtApe_Pat.Location = new System.Drawing.Point(98, 116);
            this.txtApe_Pat.Name = "txtApe_Pat";
            this.txtApe_Pat.Size = new System.Drawing.Size(181, 20);
            this.txtApe_Pat.TabIndex = 2;
            this.txtApe_Pat.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApe_Pat_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(305, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 79;
            this.label2.Text = "Usuario";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(287, 119);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(61, 13);
            this.label5.TabIndex = 82;
            this.label5.Text = "Contraseña";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(299, 151);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(49, 13);
            this.label7.TabIndex = 84;
            this.label7.Text = "Permisos";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(47, 88);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(49, 13);
            this.label8.TabIndex = 85;
            this.label8.Text = "Nombres";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(10, 144);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(86, 13);
            this.label9.TabIndex = 86;
            this.label9.Text = "Apellido Materno";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(313, 178);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 13);
            this.label3.TabIndex = 88;
            this.label3.Text = "Mail";
            // 
            // txtMail
            // 
            this.txtMail.Location = new System.Drawing.Point(354, 175);
            this.txtMail.Name = "txtMail";
            this.txtMail.Size = new System.Drawing.Size(181, 20);
            this.txtMail.TabIndex = 8;
            this.txtMail.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMail_KeyPress);
            // 
            // btnAgregar
            // 
            this.btnAgregar.Location = new System.Drawing.Point(12, 209);
            this.btnAgregar.Name = "btnAgregar";
            this.btnAgregar.Size = new System.Drawing.Size(117, 23);
            this.btnAgregar.TabIndex = 89;
            this.btnAgregar.Text = "Agregar Usuario";
            this.btnAgregar.UseVisualStyleBackColor = true;
            this.btnAgregar.Click += new System.EventHandler(this.btnAgregar_Click);
            // 
            // cbxPermisos
            // 
            this.cbxPermisos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxPermisos.FormattingEnabled = true;
            this.cbxPermisos.Location = new System.Drawing.Point(354, 148);
            this.cbxPermisos.Name = "cbxPermisos";
            this.cbxPermisos.Size = new System.Drawing.Size(181, 21);
            this.cbxPermisos.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 119);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(84, 13);
            this.label4.TabIndex = 92;
            this.label4.Text = "Apellido Paterno";
            // 
            // txtApe_Mat
            // 
            this.txtApe_Mat.Location = new System.Drawing.Point(98, 144);
            this.txtApe_Mat.Name = "txtApe_Mat";
            this.txtApe_Mat.Size = new System.Drawing.Size(181, 20);
            this.txtApe_Mat.TabIndex = 3;
            this.txtApe_Mat.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApe_Mat_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(68, 178);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(24, 13);
            this.label6.TabIndex = 94;
            this.label6.Text = "Rut";
            // 
            // txtRut
            // 
            this.txtRut.Location = new System.Drawing.Point(98, 175);
            this.txtRut.MaxLength = 13;
            this.txtRut.Name = "txtRut";
            this.txtRut.Size = new System.Drawing.Size(181, 20);
            this.txtRut.TabIndex = 4;
            this.txtRut.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRut_KeyPress);
            // 
            // GestionUsuario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(984, 623);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtRut);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtApe_Mat);
            this.Controls.Add(this.cbxPermisos);
            this.Controls.Add(this.btnAgregar);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtMail);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtApe_Pat);
            this.Controls.Add(this.txtNombres);
            this.Controls.Add(this.txtContrasena);
            this.Controls.Add(this.txtUsuario);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "GestionUsuario";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestion de Usuarios";
            this.Load += new System.EventHandler(this.GestionUsuario_Load);
            ((System.ComponentModel.ISupportInitialize)(this.usuarioswebBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.desarrolloDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private desarrolloDataSet desarrolloDataSet;
        private System.Windows.Forms.BindingSource usuarioswebBindingSource;
        private desarrolloDataSetTableAdapters.usuarios_webTableAdapter usuarios_webTableAdapter;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox txtUsuario;
        private System.Windows.Forms.TextBox txtContrasena;
        private System.Windows.Forms.TextBox txtNombres;
        private System.Windows.Forms.TextBox txtApe_Pat;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtMail;
        private System.Windows.Forms.Button btnAgregar;
        private System.Windows.Forms.ComboBox cbxPermisos;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtApe_Mat;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtRut;
    }
}