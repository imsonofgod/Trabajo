using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Finiquitos
{
    public partial class frmLogin : Form
    {
        public static string user;
        public static string permiso;

        public frmLogin()
        {
            InitializeComponent();
        }

        private void btnIngresar_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
            {

                
           

                try {

                    SqlCommand sqlCM = new SqlCommand("sp_login_finiquito", conn);
                    sqlCM.CommandType = CommandType.StoredProcedure;
                    sqlCM.Parameters.Add("@usuario", SqlDbType.VarChar).Value = txtUsuario.Text;
                    sqlCM.Parameters.Add("@contrasena", SqlDbType.VarChar).Value = txtContrasena.Text;
                    SqlDataAdapter sqlDA = new SqlDataAdapter(sqlCM);
                    DataTable DT = new DataTable();
                    sqlDA.Fill(DT);
                    DataRow row = DT.Rows[0];
                   //MessageBox.Show(Convert.ToString(row["nombre"]));

                    user = Convert.ToString(row["usuario"]);
                    permiso = Convert.ToString(row["permiso"]);

                   this.Hide();
                   PrincipalMDI mdi = new PrincipalMDI();
                   mdi.Show();

                }catch(Exception ex ){

                    MessageBox.Show("Usuario No Registrado", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
                }

            }
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {

        }
    }
}
