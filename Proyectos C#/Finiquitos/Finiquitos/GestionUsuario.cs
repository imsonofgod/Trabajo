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
    public partial class GestionUsuario : Form
    {

        // obtiene el usuario de la persona
        public static string persona;
        public static string contrasena;
        public static string rut;

        public GestionUsuario()
        {
            InitializeComponent();
        }

        private static GestionUsuario m_FormDefInstance;
        public static GestionUsuario DefInstance
        {
            get
            {
                if (m_FormDefInstance == null || m_FormDefInstance.IsDisposed)
                    m_FormDefInstance = new GestionUsuario();
                return m_FormDefInstance;
            }
            set
            {
                m_FormDefInstance = value;
            }
        }
 

        DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
        DataGridViewButtonColumn buttonColumn2 = new DataGridViewButtonColumn();
        private void GestionUsuario_Load(object sender, EventArgs e)
        {
         
            // TODO: esta línea de código carga datos en la tabla 'desarrolloDataSet.usuarios_web' Puede moverla o quitarla según sea necesario.
         
            cbxPermisos.Items.Add("Usuario");
            cbxPermisos.Items.Add("Administrador");
            cbxPermisos.SelectedIndex = (0);

                  llenar_grid();


              

                buttonColumn.Width = 50;
                buttonColumn.DisplayIndex = 0;
                buttonColumn.UseColumnTextForButtonValue = true;
                buttonColumn.Text = "Editar";

                buttonColumn2.Width = 50;
                buttonColumn2.DisplayIndex = 1;
                buttonColumn2.UseColumnTextForButtonValue = true;
                buttonColumn2.Text = "Eliminar";

                dataGridView1.Columns.Add(buttonColumn);
                dataGridView1.Columns.Add(buttonColumn2);
                dataGridView1.AllowUserToAddRows = false;
       

            
        }


        void dataGridView1_btnEditar_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ignore clicks that are not on button cells. 
            //    MessageBox.Show(dataGridView1.Rows);
            MessageBox.Show(dataGridView1.Rows[e.RowIndex].Cells["usuario"].Value.ToString());

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void click_boton_grid(object sender, DataGridViewCellEventArgs e)
        {

            //MessageBox.Show(e.RowIndex.ToString() + e.ColumnIndex.ToString());
            //especificando que reaccione solo al clic de la matriz que comienza en [0][0] 
            //osea desde los botones hacia adelante

            if ((e.RowIndex >= 0) && e.ColumnIndex >= 0)
            {
                if (dataGridView1.Columns[e.ColumnIndex] == buttonColumn )
                {

                    //MessageBox.Show(dataGridView1.Rows[e.RowIndex].Cells["usuario"].Value.ToString());
                    //editar

                    // obtiene el usuario de la persona de la grilla al precionar el boton editar
                    persona = dataGridView1.Rows[e.RowIndex].Cells["usuario"].Value.ToString();
                    contrasena = dataGridView1.Rows[e.RowIndex].Cells["contraseña"].Value.ToString();
                    rut = dataGridView1.Rows[e.RowIndex].Cells["rut"].Value.ToString();
                    EditarUsuario frmEditar = new EditarUsuario();
                    frmEditar.Show();



 

                    // verificando si el formulario 2 cerro levanta el evento closed y actualiza el datagridview
                    frmEditar.FormClosed += new FormClosedEventHandler(frmEditar_FormClosed);
      
                }


             





                if (dataGridView1.Columns[e.ColumnIndex] == buttonColumn2)
                {

                    //MessageBox.Show("eliminar"+dataGridView1.Rows[e.RowIndex].Cells["usuario"].Value.ToString());

                    DialogResult resul = MessageBox.Show("¿ Esta seguro que quiere eliminar este usuario ?", "Eliminar Usuario", MessageBoxButtons.YesNo);
                    if (resul == DialogResult.Yes)
                    {
                        try
                        {
                            using (SqlConnection cn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                             {

                                 rut = dataGridView1.Rows[e.RowIndex].Cells["rut"].Value.ToString();
                                SqlCommand cmd = new SqlCommand("sp_eliminar_usuario_finiquito", cn);
                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                {
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.AddWithValue("@rut", SqlDbType.Char).Value = rut;

                                }
                                cn.Open();
                                cmd.ExecuteNonQuery();
                                cn.Close();

                             }

                             llenar_grid();
                             MessageBox.Show("Usuario Elimado", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }

       

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!txtNombres.Text.Equals("") && !txtApe_Pat.Text.Equals("") && !txtApe_Mat.Text.Equals("") &&
               !txtRut.Text.Equals("") && !txtUsuario.Text.Equals("") && !txtContrasena.Text.Equals("") &&
               !cbxPermisos.SelectedIndex.Equals(""))
                {
                    if (!txtRut.Text.Equals(""))
                    {


                        bool estado = true;
                        rut = txtRut.Text;
                        bool validacion = false;
                        rut = rut.ToUpper();
                        rut = rut.Replace(".", "");
                        rut = rut.Replace("-", "");

                        int rutAux = int.Parse(rut.Substring(0, rut.Length - 1));

                        char dv = char.Parse(rut.Substring(rut.Length - 1, 1));

                        int m = 0, s = 1;
                        for (; rutAux != 0; rutAux /= 10)
                        {
                            s = (s + rutAux % 10 * (9 - m++ % 6)) % 11;
                        }

                        if (dv == (char)(s != 0 ? s + 47 : 75))
                        {
                            using (SqlConnection cn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                            {

                                SqlCommand sqlCM0 = new SqlCommand("sp_buscar_usuario_finiquito", cn);
                                sqlCM0.CommandType = CommandType.StoredProcedure;
                                sqlCM0.Parameters.Add("@rut", SqlDbType.VarChar).Value = rut;

                                SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                                DataTable DT0 = new DataTable();
                                sqlDA0.Fill(DT0);



                                sqlCM0.Parameters.AddWithValue("@rut", SqlDbType.VarChar).Value = rut;

                                if (!(DT0.Rows.Count > 0))
                                {
                                    string codigo_area = "07";
                                    int selectedIndex = cbxPermisos.SelectedIndex;
                                    Object selectedItem = cbxPermisos.SelectedItem;
                                    String permiso;
                                    try
                                    {
                                        using (SqlConnection cn2 = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                                        {

                                            SqlCommand cmd = new SqlCommand("sp_insertar_usuario_finiquito", cn2);
                                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                                            {
                                                cmd.CommandType = CommandType.StoredProcedure;

                                                if (selectedItem.ToString().Equals("Usuario"))
                                                {
                                                    permiso = "01";
                                                }
                                                else
                                                {
                                                    permiso = "02";
                                                }

                                                cmd.Parameters.AddWithValue("@usuario", SqlDbType.Char).Value = txtUsuario.Text;
                                                cmd.Parameters.AddWithValue("@contraseña", SqlDbType.VarChar).Value = txtContrasena.Text;
                                                cmd.Parameters.AddWithValue("@nombres", SqlDbType.VarChar).Value = txtNombres.Text;
                                                cmd.Parameters.AddWithValue("@ape_pat", SqlDbType.VarChar).Value = txtApe_Pat.Text;
                                                cmd.Parameters.AddWithValue("@ape_mat", SqlDbType.VarChar).Value = txtApe_Mat.Text;
                                                cmd.Parameters.AddWithValue("@rut", SqlDbType.Char).Value = (txtRut.Text.Replace(".","")).Replace("-","");
                                                cmd.Parameters.AddWithValue("@cod_area", SqlDbType.VarChar).Value = codigo_area;
                                                cmd.Parameters.AddWithValue("@area", SqlDbType.VarChar).Value = "RRHH";
                                                cmd.Parameters.AddWithValue("@email", SqlDbType.VarChar).Value = txtMail.Text;
                                                cmd.Parameters.AddWithValue("@fecha_creacion", SqlDbType.DateTime).Value = DateTime.Today;
                                                cmd.Parameters.AddWithValue("@fecha_cambio", SqlDbType.DateTime).Value = DateTime.Today;
                                                cmd.Parameters.AddWithValue("@permiso", SqlDbType.VarChar).Value = permiso;

                                            }
                                            cn2.Open();
                                            cmd.ExecuteNonQuery();
                                            cn2.Close();

                                            llenar_grid();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Error de Datos", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                                else
                                {

                                    MessageBox.Show("Rut ya registrado", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                }
                            }

                        }
                        else
                        {
                            MessageBox.Show("Rut no valido", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        }


                    }
                    else
                    {
                        MessageBox.Show("Rut vacio, debe ingresar todos los campos", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }


                }
                else
                {


                    MessageBox.Show("Debe ingresar todos los campos", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }catch(Exception ex){
                MessageBox.Show("Registros Invalido", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnRefrescar_Click(object sender, EventArgs e)
        {
            llenar_grid();
        }

        public void llenar_grid()
        {
            using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
            {
                SqlCommand cmd = new SqlCommand("sp_usuarios_escritorio", conn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();

                da.Fill(dt);
                this.dataGridView1.DataSource = dt;
            }                  
        }

        private void frmEditar_FormClosed(object sender, FormClosedEventArgs e)
        {
            llenar_grid();
                  
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

        private void txtMail_KeyPress(object sender, KeyPressEventArgs e)
        {
       
        }

    }
}
