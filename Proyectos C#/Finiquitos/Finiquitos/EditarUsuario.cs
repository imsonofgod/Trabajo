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
    public partial class EditarUsuario : Form
    {
        public EditarUsuario()
        {
            InitializeComponent();
        }


 

        private void EditarUsuario_Load(object sender, EventArgs e)
        {

            //Rescate de datos al load


            using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
            {



                String permiso ;

              



                try
                {

                    SqlCommand sqlCM = new SqlCommand("sp_login_finiquito", conn);
                    sqlCM.CommandType = CommandType.StoredProcedure;
                    sqlCM.Parameters.Add("@usuario", SqlDbType.VarChar).Value = GestionUsuario.persona;
                    sqlCM.Parameters.Add("@contrasena", SqlDbType.VarChar).Value = GestionUsuario.contrasena;
                    SqlDataAdapter sqlDA = new SqlDataAdapter(sqlCM);
                    DataTable DT = new DataTable();
                    sqlDA.Fill(DT);
                    DataRow row = DT.Rows[0];
                    //MessageBox.Show(Convert.ToString(row["nombre"]));

                    //user = Convert.ToString(row["usuario"]);

                    txtNombres.Text = Convert.ToString(row["nombres"]);
                    txtApe_Pat.Text = Convert.ToString(row["ape_pat"]);
                    txtApe_Mat.Text = Convert.ToString(row["ape_mat"]);
                    txtUsuario.Text = Convert.ToString(row["usuario"]);
                    txtRut.Text = Convert.ToString(row["rut"]);
                    txtContrasena.Text = Convert.ToString(row["contraseña"]);


                    if (Convert.ToString(row["permiso"]).Equals("00"))
                    {
                        permiso = "Usuario";
                    }
                    else
                    {
                        permiso = "Administrador";
                    }


                    if (permiso.Equals("Usuario"))
                    {
                        cbxPermisos.Items.Add(permiso);
                        cbxPermisos.Items.Add("Administrador");
                        cbxPermisos.SelectedIndex = (0);
                    }
                    else
                    {
                        cbxPermisos.Items.Add(permiso);
                        cbxPermisos.Items.Add("Usuario");
                        cbxPermisos.SelectedIndex = (0);
                    }


                    txtMail.Text = Convert.ToString(row["email"]);
                   
                   


                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);

                }

            }


            
        }



        




        private void btnEditar_Click(object sender, EventArgs e)
        {


            if (!txtNombres.Text.Equals("") && !txtApe_Pat.Text.Equals("") && !txtApe_Mat.Text.Equals("") &&
                !txtRut.Text.Equals("") && !txtUsuario.Text.Equals("") && !txtContrasena.Text.Equals("") && 
                !cbxPermisos.SelectedIndex.Equals(""))
            {

                // Update 

                string codigo_area = "07";
                int selectedIndex = cbxPermisos.SelectedIndex;
                Object selectedItem = cbxPermisos.SelectedItem;
                String permiso;
                try
                {
                    using (SqlConnection cn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                    {

                        SqlCommand cmd = new SqlCommand("sp_actualizar_usuario_finiquito", cn);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        {
                            cmd.CommandType = CommandType.StoredProcedure;

                            if (selectedItem.ToString().Equals("Usuario"))
                            {
                                permiso = "00";
                            }
                            else
                            {
                                permiso = "01";
                            }


                            cmd.Parameters.AddWithValue("@rut", SqlDbType.VarChar).Value = txtRut.Text;
                            cmd.Parameters.AddWithValue("@usuario", SqlDbType.Char).Value = txtUsuario.Text;
                            cmd.Parameters.AddWithValue("@contraseña", SqlDbType.VarChar).Value = txtContrasena.Text;
                            cmd.Parameters.AddWithValue("@nombres", SqlDbType.VarChar).Value = txtNombres.Text;
                            cmd.Parameters.AddWithValue("@ape_pat", SqlDbType.VarChar).Value = txtApe_Pat.Text;
                            cmd.Parameters.AddWithValue("@ape_mat", SqlDbType.VarChar).Value = txtApe_Mat.Text;
                            cmd.Parameters.AddWithValue("@cod_area", SqlDbType.VarChar).Value = codigo_area;
                            cmd.Parameters.AddWithValue("@area", SqlDbType.VarChar).Value = "RRHH";
                            cmd.Parameters.AddWithValue("@email", SqlDbType.VarChar).Value = txtMail.Text;
                            cmd.Parameters.AddWithValue("@fecha_cambio", SqlDbType.DateTime).Value = DateTime.Today;
                            cmd.Parameters.AddWithValue("@permiso", SqlDbType.VarChar).Value = permiso;


                        }
                        cn.Open();
                        cmd.ExecuteNonQuery();
                        cn.Close();

                    }



                    MessageBox.Show("Actualizado");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }



            }
            else {

                MessageBox.Show("Debe ingresar todos los campos");
            
            }

        }

        private void txtRut_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 46 && e.KeyChar != 45 && e.KeyChar != 75 && e.KeyChar != 107)
            {
                if (Char.IsDigit(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (Char.IsControl(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (Char.IsSeparator(e.KeyChar))
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            }
        }

        private void txtNombres_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back);
        }

        private void txtApe_Pat_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back);
        }

        private void txtApe_Mat_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back);
        }

        private void txtUsuario_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back);
        }

        private void txtContrasena_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar)) && (e.KeyChar != (char)Keys.Back);
        }


     


    }
}
