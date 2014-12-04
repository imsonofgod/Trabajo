using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Finiquitos
{
    public partial class PrincipalMDI : Form
    {
        public PrincipalMDI()
        {
            InitializeComponent();
        }
      
        private void formularioLiquidacionToolStripMenuItem_Click(object sender, EventArgs e)
        {




            FormularioFiniquito.DefInstance.MdiParent = this;
            FormularioFiniquito.DefInstance.Show();
 

        }

        private void itemGestion_Click(object sender, EventArgs e)
        {
            GestionUsuario.DefInstance.MdiParent = this;
            GestionUsuario.DefInstance.Show();
        }

        private void PrincipalMDI_Load(object sender, EventArgs e)
        {
            if (frmLogin.permiso.Equals("02") || frmLogin.permiso.Equals("03"))
            {
                administracionToolStripMenuItem.Visible = true;
            }
        }

        private void verLaAyudaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ayuda.DefInstance.MdiParent = this;
            Ayuda.DefInstance.Show();
        }

 
     
    }
}
