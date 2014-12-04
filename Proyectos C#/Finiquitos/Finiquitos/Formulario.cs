using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;
 

namespace Finiquitos
{
    public partial class FormularioFiniquito : Form
    {
        public String fecha_in;
        public String fecha_fin;
        public String nombre_emp;
        public String rut_emp;
        public String domicilio_emp;
        public String comuna_emp;
        public String ciudad_emp;
        public String agno_serv;
        public String indep_x_agno_ser;
        public String indep_sus_x_aviso;
        public String feariado_pro;
        public String feariado_pro_dias;
        public double total;
        public String dias_p;
        public String afc;
        public String cta_cte_emp;
        public String dcto;
        public String lqdo_pago;
        public String otros_haberes;
        public String otros_dcts;
        public String haberes;
        public String dia;
        public String art;
        public String indicador;
        public String desc_art;
        public String monto_palabra_f;
        public String ape_mat;
        public String ape_pat;
        public String sb;
        public String grat;
        public String mov;
        public String col;
        public String bon_pro;
        public Boolean flag;
        public static string rut;
        public static int dias_t;
        public static float global_res = 0;
        public static int global_ndt = 0;
        public static int error = 0;
        public static int error2 = 0;
        public static int contLoad = 0;
        public static bool clickCOnsultar = false;
        public double res_prop;
        CultureInfo elGR = CultureInfo.CreateSpecificCulture("el-GR");

        private static FormularioFiniquito m_FormDefInstance;

        public FormularioFiniquito()
        {
            InitializeComponent();
        }
        private void btnCalcular_Click(object sender, EventArgs e)
        {

        }

        public static FormularioFiniquito DefInstance
                            {
                                get
                                {
                                    if (m_FormDefInstance == null || m_FormDefInstance.IsDisposed)
                                        m_FormDefInstance = new FormularioFiniquito();
                                    return m_FormDefInstance;
                                }
                                set
                                {
                                    m_FormDefInstance = value;
                                }
                            }

        private void Formulario_Load(object sender, EventArgs e)
        {
            String url = (AppDomain.CurrentDomain.BaseDirectory + "pie.jpg").ToString();
            String url2 = url.Replace("\\", "/");

            //MessageBox.Show(url2);
            //MessageBox.Show();

           // lblUsr.Text = "Bienvenido " + frmLogin.global_nombre+" "+frmLogin.global_apllido;
            dtpIngreso.Value = Convert.ToDateTime("2014/01/01");
            dtpIngreso.CustomFormat = "yyyy/MM/dd";
            dtpIngreso.Format = DateTimePickerFormat.Custom;

            dtpFiniquito.Value = Convert.ToDateTime("2014/01/01");
            dtpFiniquito.CustomFormat = "yyyy/MM/dd";
            dtpFiniquito.Format = DateTimePickerFormat.Custom;

            dtpTermino.Value = Convert.ToDateTime("2014/01/01");
            dtpTermino.CustomFormat = "yyyy/MM/dd";
            dtpTermino.Format = DateTimePickerFormat.Custom;

             using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
              {

                try
                {

                    // ME TRAE EL DIA DE TEMRINO Y EL MES DE TERMINO 
                    SqlCommand sqlCM0 = new SqlCommand("sp_ver_articulos", conn);
                    sqlCM0.CommandType = CommandType.StoredProcedure;
                
                    SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                    DataTable DT0 = new DataTable();
                    sqlDA0.Fill(DT0);
                    DataRow row0;


                    cbxCausal.Items.Add("-- Ingrese Causal --");
                                
                    for (int i = 0; i < DT0.Rows.Count; i++)
                    {
                        row0 = DT0.Rows[i];
                        cbxCausal.Items.Add(Convert.ToString("Art." + row0["articulo"] + " (" + row0["indicador"] + ") - " + row0["descripcion"].ToString()));
 
                    }

                }catch(Exception ex){
                            
                }

             }
        
             cbxCausal.Enabled = false;

            btnCalcularPRF.Enabled = false;
        }

        private int textbox_vacios() {
                int cont = 0;
                FormularioFiniquito frm = new FormularioFiniquito();
                foreach (Control oControls in frm.Controls)
                {
                    if (oControls is TextBox)
                    {
                        if (oControls.Text.Equals(""))
                        {
                            cont = cont + 1;
                        }
                    }
                }

                return cont;
        }

        private void btnCal_Click(object sender, EventArgs e)
        {
            btnCalcularPRF.Enabled = false;
            txtPrf.Text = "0";
            txtSB.Text = sb;
            txtGratificacion.Text = grat;
            txtMovilizacion.Text = mov;
            txtColacion.Text = col;
            chkMesAviso.Checked = false;
            txtPrv.Text = "0";
            cbxCausal.Enabled = false;
            txtPrf.Text = "0";
            txtPrv.Text = "0";
 
            txtFinqApagar.Text = "0";
            txtMesAviso.Text = "0";
            txtCalVd.Text = "0";
            txtCalVacP.Text = "0";
     
            cbxCausal.Enabled = false;
            cbxCausal.SelectedIndex = 0;
            txtOtrsoDcts.Enabled = false;
            txtOtrsoDcts.Text = "0";
            txtOtrosHaberes.Text = "0";
            rchOtrosDcts.Text = "";
            rchOtrosHaberes.Text = "";
 
            txtCtaCrrteEmp.Text = "0";
            txtBaseCalculo.Text = "0";
            btnBaseCalculo.Enabled = false;
            txtFeriadoProp.Text = "0";
 
            txtCtaCrrteEmp.Enabled = false;
            txtOtrosHaberes.Enabled = false;
            txtFinqApagar.Enabled = false;
            btnCalcFiniq.Enabled = false;
       
                if (txtLiq1.Text.Equals("") || txtLiq2.Text.Equals("") || txtLiq3.Text.Equals(""))
                {
                    MessageBox.Show("Debe Ingresar Las 3 Liquidaciones");
                    btnGenerarDoc.Enabled = false;
                }
                else
                {

                    btnCalcularPRF.Enabled = true;

                    try
                    {
                        double liq1 = 0, liq2 = 0, liq3 = 0;
 
                        chkSB.Checked = false;
                        chkGratificacion.Checked = false;
                        chkMovilizacion.Checked = false;
                        chkColacion.Checked = false;

                        txtPrf.Text = "0";
                        txtPrv.Text = "0";
 
                        txtFinqApagar.Text = "0";
 
                        txtCalVd.Text = "0";
                        txtCalVacP.Text = "0";


                        liq1 = double.Parse(txtLiq1.Text.Trim().ToString());
                        liq2 = double.Parse(txtLiq2.Text.Trim().ToString());
                        liq3 = double.Parse(txtLiq3.Text.Trim().ToString());

                        //MessageBox.Show(liq3.ToString());        

                        txtTotal.Text = Convert.ToString(liq1 + liq2 + liq3);

                        if (liq1 == 0 && liq2 == 0 && liq3 == 0)
                        {
                            txtPrv.Text = Convert.ToString(Math.Round((liq1 + liq2 + liq3)));
                        }

                        if (liq1 > 0 && liq2 == 0 && liq3 == 0)
                        {
                            txtPrv.Text = Convert.ToString(Math.Round((liq1 + liq2 + liq3)));
                        }

                        if (liq1 > 0 && liq2 > 0 && liq3 == 0)
                        {
                            txtPrv.Text = Convert.ToString(Math.Round((liq1 + liq2 + liq3) / 2));
                        }

                        if (liq1 > 0 && liq2 > 0 && liq3 > 0)
                        {
                            txtPrv.Text = Convert.ToString(Math.Round((liq1 + liq2 + liq3) / 3));
                        }
                
                      
               
                    }
                    catch (Exception ex)
                    {
                      //  MessageBox.Show(ex.Message);
                        MessageBox.Show("Valor mal ingresado, favor verifique bien los montos", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
    
            this.Close();
        }

        private void chkGratificacion_CheckedChanged(object sender, EventArgs e)
        {
            txtGratificacion.Text = "";
            if (chkGratificacion.Checked == true)
            {
                txtGratificacion.Enabled = true;
            }
            else {
                txtGratificacion.Enabled = false;
                txtGratificacion.Text = "0";
 
            }
        }

        public string formatearRut(string rut)
        {
            int cont = 0;
            string format;
            if (rut.Length == 0)
            {
                return "";
            }
            else
            {
                rut = rut.Replace(".", "");
                rut = rut.Replace("-", "");
                format = "-" + rut.Substring(rut.Length - 1);
                for (int i = rut.Length - 2; i >= 0; i--)
                {
                    format = rut.Substring(i, 1) + format;
                    cont++;
                    if (cont == 3 && i != 0)
                    {
                        format = "." + format;
                        cont = 0;
                    }
                }
                return format;
            }
        }

        private void chkMovilizacion_CheckedChanged(object sender, EventArgs e)
        {

 
            if (chkMovilizacion.Checked == true)
            {
                txtMovilizacion.Enabled = true;
            }
            else
            {
                txtMovilizacion.Enabled = false;           
            }
        }

        private void chkColacion_CheckedChanged(object sender, EventArgs e)
        {

            if (chkColacion.Checked == true)
            {
                txtColacion.Enabled = true;
            }
            else
            {
                txtColacion.Enabled = false;
            }
        }

        private void chkSB_CheckedChanged(object sender, EventArgs e)
        {
 
            if (chkSB.Checked == true)
            {
                txtSB.Enabled = true;
            }
            else
            {
                txtSB.Enabled = false;
            }
        }

        private void btnCalcularPRF_Click(object sender, EventArgs e)
        {
            if (btnCalcularPRF.Enabled == true)
            {
         
                txtCtaCrrteEmp.Enabled = true;
                cbxCausal.SelectedIndex = (0);
            }

            txtOtrsoDcts.Enabled = true;
            txtOtrosHaberes.Enabled = true;
 
            btnCalcFiniq.Enabled = true;
            btnBaseCalculo.Enabled = true;
            chkDiasProgresivos.Enabled = true;
          
            // verificando si algun etxt tiene un valor vacio  no sumable
            if (txtGratificacion.Text.Trim().Equals("") ||
                txtSB.Text.Trim().Equals("") ||
                txtColacion.Text.Trim().Equals("") ||
                txtMovilizacion.Text.Trim().Equals(""))
            {

                MessageBox.Show("Debe Ingresar Un Monto", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                try
                {
                   

                    txtPrf.Text = Convert.ToString(  double.Parse(txtGratificacion.Text.Trim()) +
                                                     double.Parse(txtSB.Text.Trim()) +
                                                     double.Parse(txtColacion.Text.Trim()) +
                                                     double.Parse(txtMovilizacion.Text.Trim()) +
                                                     double.Parse(txtComision.Text.Trim()));


                    // asignando la base de calculo 
               


                }
                catch (Exception )
                {

                    MessageBox.Show("Valor mal ingresado, favor verifique bien los montos", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                

            }
            
        }

        private void txtPrf_TextChanged(object sender, EventArgs e)
        {

            if (!txtPrf.Text.Equals(""))
            {
              


            if (!txtRut.Text.Equals(""))
            {
                txtCalVd.Text = Convert.ToString(Math.Round((Double.Parse(txtSB.Text) +
                                                 Double.Parse(txtPrv.Text)) / 30));

            }
            }
        }

        private void txtRut_TextChanged(object sender, EventArgs e)
        {
 
            if(txtRut.Text.Equals("")){
            txtNombres.Text = "";
            txtApePat.Text = "";
            txtApeMat.Text = "";
            txtAgnoServ.Text = "0";
            txtCargo.Text = "";
            txtCiudad.Text = "";
            txtFeriadoProp.Text = "";
            txtComuna.Text = "";
            txtDireccion.Text = "";
            txtCargo.Text = "";

            txtLiq1.Text = "";
            txtLiq2.Text = "";
            txtLiq3.Text = "";
            txtTotal.Text = "";
            btnGenerarExcel.Enabled = false;
            button1.Enabled = true;
            chkSB.Checked = false;
            chkGratificacion.Checked = false;
            chkMovilizacion.Checked = false;
            txtSB.Text = "0";
            txtGratificacion.Text = "0";
            txtMovilizacion.Text = "0";
            txtColacion.Text = "0";
            txtDiasTomados.Text = "0";
            txtDiasTomados.Enabled = false;
          
            txtNdt.Text = "0";
            txtPrf.Text = "0";
            txtPrv.Text = "0";
 
            txtFinqApagar.Text = "0";
            txtMesAviso.Text = "0";
            txtCalVd.Text = "0";
            txtCalVacP.Text = "0";
            txtMesesTrabajados.Text = "0";
            cbxCausal.Enabled = false;
            cbxCausal.SelectedIndex = 0;
            txtOtrsoDcts.Enabled = false;
            txtOtrsoDcts.Text = "0";
            txtOtrosHaberes.Text = "0";
            rchOtrosDcts.Text = "";
            rchOtrosHaberes.Text = "";
 
            txtCtaCrrteEmp.Text = "0";
            txtBaseCalculo.Text = "0";
            btnBaseCalculo.Enabled = false;
            txtFeriadoProp.Text = "0";
 
            txtCtaCrrteEmp.Enabled = false;
            txtOtrosHaberes.Enabled = false;
            txtFinqApagar.Enabled = false;
            btnCalcFiniq.Enabled = false;
            dtpFiniquito.Enabled = false;
            }    
        }

        private void txtApePat_TextChanged(object sender, EventArgs e)
        {

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

        public void calculo_mes_dias_agnos_trabajados()
        {
            using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
            {

                try
                {
                    //numero de meses trabajados dentro de un mismo año 
                    if (dtpIngreso.Value.Year == dtpFiniquito.Value.Year)
                    {
                        SqlCommand sqlCM0 = new SqlCommand("sp_detalle_tiempo_trabajo_x_fecha", conn);
                        sqlCM0.CommandType = CommandType.StoredProcedure;
                        // el mismo año
                        sqlCM0.Parameters.Add("@xfecha1_f", SqlDbType.VarChar).Value = Convert.ToString(dtpIngreso.Text.ToString());
                        sqlCM0.Parameters.Add("@xfecha2_f", SqlDbType.VarChar).Value = Convert.ToString(dtpFiniquito.Text.ToString());

                        SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                        DataTable DT0 = new DataTable();
                        sqlDA0.Fill(DT0);
                        DataRow row0 = DT0.Rows[0];

                        //asignando numero de meses trabajados dentro de un mismo año 
                        txtMesesTrabajados.Text = Convert.ToString(row0["meses"]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                
                //numero de meses trabajados dentro de un mismo año y año sioguiente (osea sin intervalos)
                    if (dtpIngreso.Value.Year + 1 == dtpFiniquito.Value.Year)
                    {
                        int Agno_ter = 0, dia_ter = 0, mes_ter = 0, num_agno = 0, num_mes = 0, dias_tra = 0, mes_tra = 0;

                        String fec_fin, fec_ini;
                        // año 1
                        SqlCommand sqlCM0 = new SqlCommand("sp_detalle_tiempo_trabajo_x_fecha", conn);
                        sqlCM0.CommandType = CommandType.StoredProcedure;
                        dtpIngreso.CustomFormat = "yyyy-MM-dd";
                        dtpFiniquito.CustomFormat = "yyyy-MM-dd";

                        fec_ini = dtpIngreso.Text.ToString();
                        fec_fin = dtpIngreso.Value.Year.ToString() + "-12-31";


                        sqlCM0.Parameters.Add("@xfecha1_f", SqlDbType.VarChar).Value = fec_ini.ToString();
                        sqlCM0.Parameters.Add("@xfecha2_f", SqlDbType.VarChar).Value = fec_fin.ToString();

                        SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                        DataTable DT0 = new DataTable();
                        sqlDA0.Fill(DT0);
                        DataRow row0 = DT0.Rows[0];

                        //meses
                        //dias_trabajados
                        //dia_termino
                        //mes_termino

                        mes_tra = int.Parse(row0["meses"].ToString());
                        dias_tra = int.Parse(row0["dias_trabajados"].ToString());
                        dia_ter = int.Parse(row0["dia_termino"].ToString());
                        mes_ter = int.Parse(row0["mes_termino"].ToString());
                        //MessageBox.Show(mes_tra.ToString() + "dt" + dias_tra + "dia_ter" + dia_ter + "MES_TER" + mes_ter);

                        // -------------------------------año2------------------------------
                        int Agno_ter2 = 0, dia_ter2 = 0, mes_ter2 = 0, num_agno2 = 0, num_mes2 = 0, dias_tra2 = 0, mes_tra2 = 0;

                        String fec_fin2, fec_ini2;

                        SqlCommand sqlCM02 = new SqlCommand("sp_detalle_tiempo_trabajo_x_fecha", conn);
                        sqlCM02.CommandType = CommandType.StoredProcedure;
                        dtpIngreso.CustomFormat = "yyyy-MM-dd";
                        dtpFiniquito.CustomFormat = "yyyy-MM-dd";

                        fec_ini2 = dtpFiniquito.Value.Year.ToString() + "-01-01";
                        fec_fin2 = dtpFiniquito.Text.ToString();


                        sqlCM02.Parameters.Add("@xfecha1_f", SqlDbType.VarChar).Value = fec_ini2.ToString();
                        sqlCM02.Parameters.Add("@xfecha2_f", SqlDbType.VarChar).Value = fec_fin2.ToString();

                        SqlDataAdapter sqlDA02 = new SqlDataAdapter(sqlCM02);
                        DataTable DT02 = new DataTable();
                        sqlDA02.Fill(DT02);
                        DataRow row02 = DT02.Rows[0];

                        //meses
                        //dias_trabajados
                        //dia_termino
                        //mes_termino

                        mes_tra2 = int.Parse(row02["meses"].ToString());
                        dias_tra2 = int.Parse(row02["dias_trabajados"].ToString());
                        dia_ter2 = int.Parse(row02["dia_termino"].ToString());
                        mes_ter2 = int.Parse(row02["mes_termino"].ToString());

                        //MessageBox.Show(mes_tra2.ToString() + "dt" + dias_tra2 + "dia_ter" + dia_ter2 + "MES_TER" + mes_ter2);


                        // dias en meses
                        int tot_dias, dias_res;
                        int dias_meses = 0;
                        tot_dias = (dias_tra + dias_tra2);
                        if (tot_dias > 30)
                        {
                            dias_res = (tot_dias % 30);
                            dias_meses = (tot_dias - dias_res) / 30;
                            tot_dias = dias_res;
                        }

                        //MessageBox.Show("dias_meses" + dias_meses + "tot_dias" + tot_dias);

                        //meses años
                        int tot_mes_agnos = 0;
                        int meses_agnos_res = 0;
                        int agnos = 0;
                        tot_mes_agnos = mes_tra + mes_tra2 + dias_meses;
                        if (tot_mes_agnos >= 12)
                        {
                            meses_agnos_res = tot_mes_agnos % 12;
                            tot_mes_agnos = (tot_mes_agnos - meses_agnos_res) / 12;

                        }


                        // MessageBox.Show("mes" + meses_agnos_res + "dias" + tot_dias + "años" + tot_mes_agnos);


                        //asignando numero de meses 
                        global_ndt = dias_tra + dias_tra2;
                        txtMesesTrabajados.Text = (mes_tra + mes_tra2 + dias_meses).ToString();
                    }

                    // si el siguiente siguiente año sigue siendo menor 
                    // quiere decir que hay intervalos

                    if (dtpIngreso.Value.Year + 1 < dtpFiniquito.Value.Year)
                    {

                        int contAgno;
                        int num_meses;
                        contAgno = (dtpFiniquito.Value.Year - dtpIngreso.Value.Year) - 1;
                        num_meses = contAgno * 12;


                        int Agno_ter = 0, dia_ter = 0, mes_ter = 0, num_agno = 0, num_mes = 0, dias_tra = 0, mes_tra = 0;

                        String fec_fin, fec_ini;
                        // año 1
                        SqlCommand sqlCM0 = new SqlCommand("sp_detalle_tiempo_trabajo_x_fecha", conn);
                        sqlCM0.CommandType = CommandType.StoredProcedure;
                        dtpIngreso.CustomFormat = "yyyy-MM-dd";
                        dtpFiniquito.CustomFormat = "yyyy-MM-dd";

                        fec_ini = dtpIngreso.Text.ToString();
                        fec_fin = dtpIngreso.Value.Year.ToString() + "-12-31";


                        sqlCM0.Parameters.Add("@xfecha1_f", SqlDbType.VarChar).Value = fec_ini.ToString();
                        sqlCM0.Parameters.Add("@xfecha2_f", SqlDbType.VarChar).Value = fec_fin.ToString();

                        SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                        DataTable DT0 = new DataTable();
                        sqlDA0.Fill(DT0);
                        DataRow row0 = DT0.Rows[0];

                        //meses
                        //dias_trabajados
                        //dia_termino
                        //mes_termino

                        mes_tra = int.Parse(row0["meses"].ToString());
                        dias_tra = int.Parse(row0["dias_trabajados"].ToString());
                        dia_ter = int.Parse(row0["dia_termino"].ToString());
                        mes_ter = int.Parse(row0["mes_termino"].ToString());
                        //MessageBox.Show(mes_tra.ToString() + "dt" + dias_tra + "dia_ter" + dia_ter + "MES_TER" + mes_ter);

                        // -------------------------------año2------------------------------
                        int Agno_ter2 = 0, dia_ter2 = 0, mes_ter2 = 0, num_agno2 = 0, num_mes2 = 0, dias_tra2 = 0, mes_tra2 = 0;

                        String fec_fin2, fec_ini2;

                        SqlCommand sqlCM02 = new SqlCommand("sp_detalle_tiempo_trabajo_x_fecha", conn);
                        sqlCM02.CommandType = CommandType.StoredProcedure;

                        dtpIngreso.CustomFormat = "yyyy-MM-dd";
                        dtpFiniquito.CustomFormat = "yyyy-MM-dd";

                        fec_ini2 = dtpFiniquito.Value.Year.ToString() + "/01/01";
                        fec_fin2 = dtpFiniquito.Text.ToString();


                        sqlCM02.Parameters.Add("@xfecha1_f", SqlDbType.VarChar).Value = fec_ini2.ToString();
                        sqlCM02.Parameters.Add("@xfecha2_f", SqlDbType.VarChar).Value = fec_fin2.ToString();

                        SqlDataAdapter sqlDA02 = new SqlDataAdapter(sqlCM02);
                        DataTable DT02 = new DataTable();
                        sqlDA02.Fill(DT02);
                        DataRow row02 = DT02.Rows[0];

                        //meses
                        //dias_trabajados
                        //dia_termino
                        //mes_termino

                        mes_tra2 = int.Parse(row02["meses"].ToString());
                        dias_tra2 = int.Parse(row02["dias_trabajados"].ToString());
                        dia_ter2 = int.Parse(row02["dia_termino"].ToString());
                        mes_ter2 = int.Parse(row02["mes_termino"].ToString());

                        //MessageBox.Show(mes_tra2.ToString() + "dt" + dias_tra2 + "dia_ter" + dia_ter2 + "MES_TER" + mes_ter2);


                        // dias en meses
                        int tot_dias, dias_res;
                        int dias_meses = 0;
                        tot_dias = (dias_tra + dias_tra2);
                        if (tot_dias > 30)
                        {
                            dias_res = (tot_dias % 30);
                            dias_meses = (tot_dias - dias_res) / 30;
                            tot_dias = dias_res;
                        }

                        //MessageBox.Show("dias_meses" + dias_meses + "tot_dias" + tot_dias);

                        //meses años
                        int tot_mes_agnos = 0;
                        int meses_agnos_res = 0;
                        int agnos = 0;
                        tot_mes_agnos = mes_tra + mes_tra2 + dias_meses;
                        if (tot_mes_agnos >= 12)
                        {
                            meses_agnos_res = tot_mes_agnos % 12;
                            tot_mes_agnos = (tot_mes_agnos - meses_agnos_res) / 12;

                        }

                        // MessageBox.Show("mes" + meses_agnos_res + "dias" + tot_dias + "años" + tot_mes_agnos);

                        //asignando numero de meses 
                        global_ndt = dias_tra + dias_tra2;
                        txtMesesTrabajados.Text = (mes_tra + mes_tra2 + dias_meses + contAgno * 12).ToString();

                    }



            

           
            }

            using (SqlConnection con = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
            {

                try
                {

                    int numerodias = (dtpFiniquito.Value - dtpIngreso.Value).Days + 1;
                    // MessageBox.Show(numerodias.ToString());

                    //  cbxCausal.SelectedIndex = (0);
                    //txtNdt.Text = (Convert.ToString(row1["Calculo"]));
                    txtNdt.Text = numerodias.ToString();
                    //MessageBox.Show(numerodias2.ToString());


                    //  año y meses de servicio 
                    SqlCommand sqlCM2 = new SqlCommand("sp_dia_mes_agno_trabajado", con);
                    sqlCM2.CommandType = CommandType.StoredProcedure;
                    sqlCM2.Parameters.Add("@diast", SqlDbType.Float).Value = float.Parse(txtNdt.Text);

                    SqlDataAdapter sqlDA2 = new SqlDataAdapter(sqlCM2);
                    DataTable DT2 = new DataTable();
                    sqlDA2.Fill(DT2);
                    DataRow row2 = DT2.Rows[0];

                    txtAgnoServ.Text = (Convert.ToString(row2["Agnos"]));
                    //   txtMesxAgno.Text = (Convert.ToString(row2["Mes_x_Agno"]));

                }
                catch (Exception ex)
                {

                    //MessageBox.Show("Empleado No Registrado");
                    // MessageBox.Show(ex.Message);
                    txtRut.Focus();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            clickCOnsultar = true;

            try
            {

                if (!(txtRut.Text.Equals("")))
                {

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

                        btnCal.Enabled = true;
                        using (SqlConnection con = new SqlConnection("Data Source=10.1.3.227;Initial Catalog=TRANSAC_GRANGUIA;Integrated Security=False;User ID=softdesa;Password=softdesa2014"))
                        {
                           

                            try
                            {
                                SqlCommand sqlCM0 = new SqlCommand("sp_softland_desde_237", con);
                                sqlCM0.CommandType = CommandType.StoredProcedure;
                                sqlCM0.Parameters.Add("@varRut", SqlDbType.VarChar).Value = formatearRut(rut);

                                SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                                DataTable DT0 = new DataTable();
                                sqlDA0.Fill(DT0);
                                DataRow row0 = DT0.Rows[0];


                                if (DT0.Rows.Count > 0)
                                {

                                    sb = Convert.ToString(row0["SueldoBase"]);
                                    grat = Convert.ToString(row0["Gratificacion"]);
                                    mov = Convert.ToString(row0["Movilizacion"]);
                                    col = Convert.ToString(row0["Colacion"]);
                                    bon_pro = Convert.ToString(row0["bono_produccion"]);
                                    dtpIngreso.Value = Convert.ToDateTime(row0["fechaIngreso"]);
                                    dtpIngreso.CustomFormat = "yyyy/MM/dd";
                                    dtpIngreso.Format = DateTimePickerFormat.Custom;

                                    dtpFiniquito.Value = Convert.ToDateTime(DateTime.Now);
                                    dtpFiniquito.CustomFormat = "yyyy/MM/dd";
                                    dtpFiniquito.Format = DateTimePickerFormat.Custom;

                                    txtNombres.Text = Convert.ToString(row0["nombres"]);
                                    txtApePat.Text = Convert.ToString(row0["appaterno"]);
                                    txtApeMat.Text = Convert.ToString(row0["apmaterno"]);
                                    txtCargo.Text = Convert.ToString(row0["cargo"]);
                                    txtDireccion.Text = Convert.ToString(row0["direccion"]);
                                    txtComuna.Text = Convert.ToString(row0["comuna"]);
                                    txtCiudad.Text = Convert.ToString(row0["ciudad"]);
                                    txtSB.Text = Convert.ToString(row0["SueldoBase"]);
                                    txtGratificacion.Text = Convert.ToString(row0["Gratificacion"]);
                                    txtMovilizacion.Text = Convert.ToString(row0["Movilizacion"]);
                                    txtColacion.Text = Convert.ToString(row0["Colacion"]);
                                    txtComision.Text = Convert.ToString(row0["bono_produccion"]);
                                    txtDiasTomados.Text = Convert.ToString(row0["diasTomados"]);
                                    txtLiq1.Enabled = true;
                                    txtLiq2.Enabled = true;
                                    txtLiq3.Enabled = true;

                                    chkSB.Enabled = true;
                                    chkComision.Enabled = true;
                                    chkMovilizacion.Enabled = true;
                                    chkColacion.Enabled = true;
                                    chkGratificacion.Enabled = true;
                                    dtpFiniquito.Enabled = true;
                                    dtpTermino.Enabled = true;
                                }
                                else {
                                    error2 = 1;
                                    MessageBox.Show("Empleado No Registrado", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                
                                }
                            
 
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Empleado No Registrado", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                //MessageBox.Show(ex.Message);
                                error2 = 1;
                                txtRut.Focus();
                            }
                        }

                    }else{

                        txtRut.Focus();
                        MessageBox.Show("Rut no valido", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    
                }
                else
                {
                    error2 = 1;
                    txtRut.Focus();
                    MessageBox.Show("Debe Ingresar un Rut Valido", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                error2 = 1;
                MessageBox.Show("Debe Ingresar un Rut Valido", "Mensage del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // MessageBox.Show(ex.Message);
                txtRut.Focus();
            }

            if(error2 < 1){
                calculo_mes_dias_agnos_trabajados();
            }
        }

        private void btnGenerarDoc_Click(object sender, EventArgs e)
        {
            btnGenerarExcel.Enabled = true;
            btnGenerarTermino.Enabled = true;
            btnCartaExcel.Enabled = true;

            if (double.Parse(txtFinqApagar.Text) > 0)
            {
                if (!cbxCausal.Text.Equals("-- Ingrese Causal --") && !txtRut.Text.Equals(""))
                {
                    //si es que esta todo guardado osea no hablitado como bandera se escoge el ultimo control
                    try
                    {
                        using (SqlConnection cn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                        {
                            int selectedIndex = cbxCausal.SelectedIndex;
                            Object selectedItem = cbxCausal.SelectedItem;

                            string strHostName = string.Empty;
                            // Getting Ip address of local machine…
                            // First get the host name of local machine.
                            strHostName = Dns.GetHostName();
                            // Then using host name, get the IP address list..
                            IPAddress[] hostIPs = Dns.GetHostAddresses(strHostName);
                            //MessageBox.Show("Direccion IP: " + hostIPs[1].ToString());

                            //MessageBox.Show("Nombre de la computadora: " + strHostName);

                            String nombrepc = hostIPs[1].ToString();
                            String ip = strHostName.ToString();

                            SqlCommand cmd = new SqlCommand("sp_log_finiquito", cn);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.AddWithValue("@rut", SqlDbType.Char).Value = txtRut.Text.Replace(".","");
                                cmd.Parameters.AddWithValue("@nombres", SqlDbType.VarChar).Value = txtNombres.Text;
                                cmd.Parameters.AddWithValue("@ape_pat", SqlDbType.VarChar).Value = txtApePat.Text;
                                cmd.Parameters.AddWithValue("@ape_mat", SqlDbType.VarChar).Value = txtApeMat.Text;
                                cmd.Parameters.AddWithValue("@direccion", SqlDbType.VarChar).Value = txtDireccion.Text;
                                cmd.Parameters.AddWithValue("@ciudad", SqlDbType.VarChar).Value = txtCiudad.Text;
                                cmd.Parameters.AddWithValue("@cargo", SqlDbType.VarChar).Value = txtCargo.Text;
                                cmd.Parameters.AddWithValue("@comuna", SqlDbType.VarChar).Value = txtComuna.Text;
                                cmd.Parameters.AddWithValue("@agnos_servicio", SqlDbType.Int).Value = int.Parse(txtAgnoServ.Text);
                                cmd.Parameters.AddWithValue("@dias_trabajados", SqlDbType.Int).Value = int.Parse(txtNdt.Text);
                                cmd.Parameters.AddWithValue("@meses_x_agnos", SqlDbType.Int).Value = int.Parse(txtMesxAgno.Text);
                                cmd.Parameters.AddWithValue("@causal_despido", SqlDbType.VarChar).Value = selectedItem.ToString();
                                cmd.Parameters.AddWithValue("@liq1", SqlDbType.Float).Value = float.Parse(txtLiq1.Text);
                                cmd.Parameters.AddWithValue("@liq2", SqlDbType.Float).Value = float.Parse(txtLiq2.Text);
                                cmd.Parameters.AddWithValue("@liq3", SqlDbType.Float).Value = float.Parse(txtLiq3.Text);
                                cmd.Parameters.AddWithValue("@total", SqlDbType.Float).Value = float.Parse(txtTotal.Text);
                                cmd.Parameters.AddWithValue("@sueldo_base", SqlDbType.Float).Value = float.Parse(txtSB.Text);
                                cmd.Parameters.AddWithValue("@gratificacion", SqlDbType.Float).Value = float.Parse(txtGratificacion.Text);
                                cmd.Parameters.AddWithValue("@movilizacion", SqlDbType.Float).Value = float.Parse(txtMovilizacion.Text);
                                cmd.Parameters.AddWithValue("@comision", SqlDbType.Float).Value = float.Parse(txtColacion.Text);
                                cmd.Parameters.AddWithValue("@colacion", SqlDbType.Float).Value = float.Parse(txtComision.Text);
                                cmd.Parameters.AddWithValue("@promedio_rem_fija", SqlDbType.Float).Value = float.Parse(txtPrf.Text);
                                cmd.Parameters.AddWithValue("@promedio_rem_var", SqlDbType.Float).Value = float.Parse(txtPrv.Text);
                                cmd.Parameters.AddWithValue("@descuento_afc", SqlDbType.Float).Value = float.Parse(txtSegCes.Text);
                                cmd.Parameters.AddWithValue("@calc_valor_dia", SqlDbType.Float).Value = float.Parse(txtCalVd.Text);
                                cmd.Parameters.AddWithValue("@calc_vacaciones_prop", SqlDbType.Float).Value = float.Parse(txtCalVacP.Text);
                                cmd.Parameters.AddWithValue("@login_usuario", SqlDbType.VarChar).Value = frmLogin.user;
                                cmd.Parameters.AddWithValue("@mes_aviso", SqlDbType.Int).Value = int.Parse(txtMesAviso.Text);
                                cmd.Parameters.AddWithValue("@dias_tomados", SqlDbType.Int).Value = int.Parse(txtDiasTomados.Text);
                                cmd.Parameters.AddWithValue("@otros_dcts", SqlDbType.Float).Value = float.Parse(txtOtrsoDcts.Text);
                                cmd.Parameters.AddWithValue("@otros_dcts_desc", SqlDbType.VarChar).Value = rchOtrosDcts.Text;
                                cmd.Parameters.AddWithValue("@otros_haberes", SqlDbType.Float).Value = float.Parse(txtOtrosHaberes.Text);
                                cmd.Parameters.AddWithValue("@otros_haberes_desc", SqlDbType.VarChar).Value = rchOtrosHaberes.Text;
                                cmd.Parameters.AddWithValue("@cuenta_corriente_emp", SqlDbType.Float).Value = float.Parse(txtCtaCrrteEmp.Text);
                                cmd.Parameters.AddWithValue("@cuenta_corriente_personal", SqlDbType.Float).Value = float.Parse(txtCtaCorrientePer.Text);
                                cmd.Parameters.AddWithValue("@dcto_fondo_fijo", SqlDbType.Float).Value = float.Parse(txtDctoFondoDijo.Text);
                                cmd.Parameters.AddWithValue("@dias_progresivos", SqlDbType.Float).Value = float.Parse(txtDiasProgresivos.Text);
                                cmd.Parameters.AddWithValue("@rem_liq_pendiente", SqlDbType.Float).Value = float.Parse(txtRemLiqPen.Text);
                                cmd.Parameters.AddWithValue("@aporte_caja_comp", SqlDbType.Float).Value = float.Parse(txtSaldoCajaCompensacion.Text);
                                cmd.Parameters.AddWithValue("@aporte_seg_cesantia", SqlDbType.Float).Value = float.Parse(txtSegCes.Text);
                                cmd.Parameters.AddWithValue("@base_calculo", SqlDbType.Float).Value = float.Parse(txtBaseCalculo.Text);
                                cmd.Parameters.AddWithValue("@fecha_ingreso", SqlDbType.DateTime).Value = dtpIngreso.Text;
                                cmd.Parameters.AddWithValue("@fecha_finiquito", SqlDbType.DateTime).Value = dtpFiniquito.Text;
                                cmd.Parameters.AddWithValue("@fecha_term_contrato", SqlDbType.DateTime).Value = dtpTermino.Text;
                                cmd.Parameters.AddWithValue("@finiquito_a_pagar", SqlDbType.Float).Value = float.Parse(txtFinqApagar.Text);
                                cmd.Parameters.AddWithValue("@ip", SqlDbType.VarChar).Value = nombrepc;
                                cmd.Parameters.AddWithValue("@nom_equipo", SqlDbType.VarChar).Value = ip;

                                desabilitarControles();
                            }
                            cn.Open();
                            cmd.ExecuteNonQuery();
                            cn.Close();
                        }


                        btnGenerarExcel.Enabled = true;
                        btnGenerarTermino.Enabled = true;
                        btnEditarFrm.Enabled = true;
                        btnCartaExcel.Enabled = true;
                        btnBaseCalculo.Enabled = false;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Debe Ingresar Todos Los Campos", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Debe seleccionar almenos una causal", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else {
                MessageBox.Show("Para guardar los cambios debe calcular el finiquito a pagar", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void calculo_feriado() {
           
                try
                {
                    if (cbxCausal.Enabled = true && txtNdt.Text != "" && btnCalcularPRF.Enabled == true)
                    {

                        using (SqlConnection con = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                        {

                            SqlCommand sqlCM2 = new SqlCommand("sp_dia_mes_agno_trabajado", con);
                            sqlCM2.CommandType = CommandType.StoredProcedure;
                            sqlCM2.Parameters.Add("@diast", SqlDbType.Float).Value = float.Parse(txtNdt.Text);

                            SqlDataAdapter sqlDA2 = new SqlDataAdapter(sqlCM2);
                            DataTable DT2 = new DataTable();
                            sqlDA2.Fill(DT2);
                            DataRow row2 = DT2.Rows[0];


                            if (cbxCausal.SelectedIndex >= 1)
                            {
                                int dias_feriado_prop = 0;
                                int diasHabiles = 0;
                                int feriados = 0;
                                double diaFraccion = 0;
                                double dia_decimal = 0;
                                //  MessageBox.Show(((float.Parse(txtMesesTrabajados.Text) * 1.25) - Math.Truncate(float.Parse(txtMesesTrabajados.Text) * 1.25)).ToString());


                            SqlCommand sqlC = new SqlCommand("sp_dias_trabajados_en_decimales_finiquitos", con);
                            sqlC.CommandType = CommandType.StoredProcedure;
                            sqlC.Parameters.Add("@fec1", SqlDbType.DateTime).Value = dtpIngreso.Value;
                            sqlC.Parameters.Add("@fec2", SqlDbType.DateTime).Value = dtpFiniquito.Value;

                            SqlDataAdapter sqlD  = new SqlDataAdapter(sqlC);
                            DataTable DTo = new DataTable();
                            sqlD.Fill(DTo);
                            DataRow row6 = DTo.Rows[0];

                            dias_feriado_prop = int.Parse(Math.Truncate(float.Parse(txtMesesTrabajados.Text) * 1.25).ToString()) - int.Parse(txtDiasTomados.Text);

                            if (Convert.ToString(row6["dia_decimal"]).Equals("0"))
                            {
               
                                diaFraccion = (float.Parse(txtMesesTrabajados.Text) * 1.25) - Math.Truncate(float.Parse(txtMesesTrabajados.Text) * 1.25);
                            }
                            else {

                                diaFraccion = double.Parse(Convert.ToString(row6["dia_decimal"])) * (1.25 / 30);
                            
                            }
                                res_prop = Math.Round(dias_feriado_prop + diaFraccion,2);

                                // MessageBox.Show("dias " + dias + "fraccion" + diaFraccion);

                                using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                                {

                                    try
                                    {

                                        // ME TRAE EL DIA DE TEMRINO Y EL MES DE TERMINO 
                                        SqlCommand sqlCM0 = new SqlCommand("sp_detalle_tiempo_trabajo_x_fecha", conn);
                                        sqlCM0.CommandType = CommandType.StoredProcedure;
                                        dtpIngreso.CustomFormat = "yyyy-MM-dd";
                                        dtpFiniquito.CustomFormat = "yyyy-MM-dd";


                                        sqlCM0.Parameters.Add("@xfecha1_f", SqlDbType.VarChar).Value = dtpFiniquito.Value.Year.ToString() + "-01-01";
                                        sqlCM0.Parameters.Add("@xfecha2_f", SqlDbType.VarChar).Value = dtpFiniquito.Text.ToString();


                                        SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                                        DataTable DT0 = new DataTable();
                                        sqlDA0.Fill(DT0);
                                        DataRow row0 = DT0.Rows[0];


                                        // EXTRACCION DE DIAS POR MES 
                                        SqlCommand sql = new SqlCommand("sp_mes_dias_x_numero_mes", conn);
                                        sql.CommandType = CommandType.StoredProcedure;
                                        sql.Parameters.Add("@numero", SqlDbType.Int).Value = int.Parse(row0["mes_termino"].ToString());
                                        sql.Parameters.Add("@agno", SqlDbType.VarChar).Value = dtpFiniquito.Value.Year.ToString();
                                        SqlDataAdapter sqlDAA = new SqlDataAdapter(sql);
                                        DataTable DTG = new DataTable();
                                        sqlDAA.Fill(DTG);
                                        DataRow roww = DTG.Rows[0];


                                        //MessageBox.Show(Convert.ToString(roww["dias"]));

                                        int avance = 0;
                                        dia = Convert.ToString(row0["dia_termino"]).ToString();
                                        String mes = Convert.ToString(row0["mes_termino"]).ToString();
                                        String agno = dtpFiniquito.Value.Year.ToString();

                                        // MessageBox.Show("dia"+dia+"mes" + mes+"agno" + agno);

                                        if (int.Parse(dia) < int.Parse(Convert.ToString(roww["dias"])))
                                        {

                                            int x = (int.Parse(dia) + 1);
                                            dia = x.ToString();

                                            do
                                            {
                                                if (x <= int.Parse(dia))
                                                {
                                                    DateTime date = DateTime.Parse((agno + "-" + mes + "-" + dia));

                                                    SqlCommand sqlCM3 = new SqlCommand("sp_dias_feriados_finiquito", conn);
                                                    sqlCM3.CommandType = CommandType.StoredProcedure;
                                                    sqlCM3.Parameters.Add("@dia_termino", SqlDbType.DateTime).Value = date;

                                                    SqlDataAdapter sqlDA3 = new SqlDataAdapter(sqlCM3);
                                                    DataTable DT3 = new DataTable();

                                                    sqlDA3.Fill(DT3);
                                                    //DataRow row = DT3.Rows[0];



                                                    if (DT3.Rows.Count > 0)
                                                    {
                                                        feriados = feriados + 1;

                                                    }
                                                    else
                                                    {
                                                        diasHabiles = diasHabiles + 1;
                                                        //dias_feriado_prop
                                                        dias_feriado_prop = dias_feriado_prop - 1;

                                                    }


                                                    dias_t = (int.Parse(dia) + 1);
                                                    dia = dias_t.ToString();

                                                }

                                                // cambiar el mes
                                                //     diasHabiles <  dias_feriado_prop && 
                                                // } while (dias_feriado_prop > 0 && int.Parse(dia) <= int.Parse(Convert.ToString(roww["dias"])));
                                                if (int.Parse(mes) == 11 && int.Parse(dia) == 31)
                                                {
                                                    flag = true;
                                                }
                                                else
                                                {
                                                    flag = false;
                                                }
                                            } while (dias_feriado_prop > 0 && int.Parse(dia) <= int.Parse(Convert.ToString(roww["dias"])));
                                        }



                                        int feriados2 = 0;
                                        int diasHabiles2 = 0;
                                        dias_t = 0;
                                        String fec_finiq = dtpFiniquito.Value.Year.ToString();
                                        if (dias_feriado_prop > 0)
                                        {
                                            int r = (int.Parse(mes) + 1);

                                            if (r >= 12 && int.Parse(dia) >= 31 && flag == false)
                                            {
                                                mes = (1).ToString();
                                                dia = "01";
                                                agno = (int.Parse(agno) + 1).ToString();
                                                fec_finiq = (int.Parse(dtpFiniquito.Value.Year.ToString()) + 1).ToString();
                                            }
                                            else
                                            {

                                                mes = r.ToString();
                                                dia = "01";
                                            }

                                            do
                                            {

                                                DateTime date = DateTime.Parse((agno + "-" + mes + "-" + dia));

                                                //// MessageBox.Show(date.ToString());

                                                SqlCommand sqlCM3 = new SqlCommand("sp_dias_feriados_finiquito", conn);
                                                sqlCM3.CommandType = CommandType.StoredProcedure;
                                                sqlCM3.Parameters.Add("@dia_termino", SqlDbType.DateTime).Value = date;

                                                SqlDataAdapter sqlDA3 = new SqlDataAdapter(sqlCM3);
                                                DataTable DT3 = new DataTable();

                                                sqlDA3.Fill(DT3);
                                                //DataRow row = DT3.Rows[0];



                                                if (DT3.Rows.Count > 0)
                                                {
                                                    feriados2 = feriados2 + 1;

                                                }
                                                else
                                                {
                                                    diasHabiles2 = diasHabiles2 + 1;

                                                }

                                                //dias_feriado_prop
                                                dias_t = (int.Parse(dia) + 1);
                                                //   dia = dias_t.ToString();

                                                // si los dias son mayor fin de mes 


                                                SqlCommand sqla = new SqlCommand("sp_mes_dias_x_numero_mes", conn);
                                                sqla.CommandType = CommandType.StoredProcedure;
                                                sqla.Parameters.Add("@numero", SqlDbType.Int).Value = int.Parse(mes);
                                                sqla.Parameters.Add("@agno", SqlDbType.VarChar).Value = fec_finiq.ToString();
                                                SqlDataAdapter sqlDAAa = new SqlDataAdapter(sqla);
                                                DataTable DTGa = new DataTable();
                                                sqlDAAa.Fill(DTGa);
                                                DataRow rowwa = DTGa.Rows[0];

                                                int h = int.Parse(Convert.ToString(rowwa["dias"]));

                                                if (dias_t > h)
                                                {
                                                    // si es el ultimo mes cambia el año 
                                                    if (dtpFiniquito.Value.Month.ToString().Equals("12"))
                                                    {
                                                        fec_finiq = (int.Parse(dtpFiniquito.Value.Year.ToString()) + 1).ToString();
                                                    }

                                                    mes = (int.Parse(mes) + 1).ToString();
                                                    dias_t = 01;

                                                }

                                                dia = dias_t.ToString();



                                            } while (diasHabiles2 < dias_feriado_prop);
 }

                                        //MessageBox.Show("existen" + Convert.ToString(diasHabiles + diasHabiles2) + "dias habiles" + (feriados + feriados2) + "dias feriados");

                                        //mas el dia de llegada 
                                       // global_res = (feriados + feriados2) + 1;
                                      global_res = (feriados + feriados2) ;


                                    }
                                    catch (Exception ex)
                                    {
                                        error = 1;
                                        //MessageBox.Show(ex.Message);
                                        MessageBox.Show("No hay feriados cargados para el calculo del año siguiente, favor comunicarse con Dpto. Informatica", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }




                            }


                            //!!!!!!!!!!!!!!!!!!!!!! rescatara txtDiasTomados de la base de softland 

                            //  calcula el numero de meses o dias del feriado proporcional * calculo de valor dia  
                            if (int.Parse(txtMesesTrabajados.Text) > 0 && error == 0)
                            {
                                txtFeriadoProp.Text = (((res_prop ) + global_res + double.Parse(txtDiasProgresivos.Text))).ToString();

                                dias_p = (double.Parse(txtFeriadoProp.Text) ).ToString();
                                if (double.Parse(txtFeriadoProp.Text) != 0)
                                {
                                    txtCalVacP.Text = Math.Round(((double.Parse(txtFeriadoProp.Text)) * double.Parse(txtCalVd.Text))).ToString();
                                }


                                if ((double.Parse(txtNdt.Text) / 365) >= 1 && cbxCausal.Text.Equals("Art.161 (1) - Necesidad de la empresa derivadas de la racionalización del servicio"))
                                {

                                    // txtMesxAgno.Text = Convert.ToString(double.Parse(Convert.ToString(row2["Mes_x_Agno"]))*(double.Parse(txtBaseCalculo.Text)));
                                    txtMesxAgno.Text = Convert.ToString(double.Parse(txtAgnoServ.Text) * (double.Parse(txtBaseCalculo.Text)));
                                }
                                else
                                {

                                    txtMesxAgno.Text = "0";
                                }
    

                            }

                            if (cbxCausal.SelectedIndex < 1 && int.Parse(txtMesesTrabajados.Text) > 0 && error == 0)
                            {
                                txtMesxAgno.Text = "0";
                                txtMesAviso.Text = "0";
                            }

                            if (cbxCausal.SelectedIndex < 1 && int.Parse(txtMesesTrabajados.Text) == 0 && error == 0)
                            {
                                txtMesxAgno.Text = "0";
                                txtMesAviso.Text = "0";
                            }

                        }

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("No hay feriados cargados para el calculo del año siguiente, favor comunicarse con Dpto. Informatica", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
 
        
        }

        private void cbxCausal_SelectedIndexChanged(object sender, EventArgs e)
        {
         
                if (cbxCausal.SelectedIndex > 0)
                {

                    try
                    {
                        if (int.Parse(txtMesesTrabajados.Text) > 0 && cbxCausal.Enabled == true)
                        {
                            calculo_feriado();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Para calcular los dias feriados es necesario que la remuneracion fija tenga un valor mayor a 0", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
        
          
        }

        private void chkMesAviso_CheckedChanged(object sender, EventArgs e)
        {
            
            if (txtPrf.Text.Equals(""))
            {


                MessageBox.Show("Debe Ingresar Valores en Remuneracion Fija", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (chkMesAviso.Checked == true)
                {
                    if(double.Parse(txtBaseCalculo.Text)>0){
                    txtMesAviso.Text = (double.Parse(txtBaseCalculo.Text)*1).ToString();
                    }else{
                        MessageBox.Show("Primero debe calcular la base da calculo", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkMesAviso.Checked = false;
                    }        
    
                }
                else
                {
                    txtMesAviso.Text = "0";
                }
            }
        }

        private void txtDiasxMes_TextChanged(object sender, EventArgs e)
        {
   
      
        

        }
 

        private void btnCalcFiniq_Click(object sender, EventArgs e)
        {
            btnGenerarDoc.Enabled = true;
            txtFinqApagar.Text = (  double.Parse(txtMesAviso.Text)  +
                                    double.Parse(txtMesxAgno.Text) +
                                    double.Parse(txtCalVacP.Text)  -
                                    (double.Parse(txtOtrsoDcts.Text)+
                                     double.Parse(txtSegCes.Text) +
                                     double.Parse(txtCtaCrrteEmp.Text))).ToString();

            


            if (!txtFinqApagar.Text.Equals(""))
            {
                btnGenerarDoc.Enabled = true;

            }
        }

        private void txtOtrsoDcts_TextChanged(object sender, EventArgs e)
        {
            if (txtOtrsoDcts.Text.Equals("0"))
            {
                rchOtrosDcts.Enabled = false;
            }

            if(txtOtrsoDcts.Text.Equals("")){
                txtOtrsoDcts.Text = "0";
            }

            if (double.Parse(txtOtrsoDcts.Text) > 0 && !(txtOtrsoDcts.Text.Equals("")))
            {

                rchOtrosDcts.Enabled = true;
            }


            if (txtOtrsoDcts.Equals("0"))
            {

                rchOtrosDcts.Text = "";
            }
          
        }

        private void txtTotal_TextChanged(object sender, EventArgs e)
        {
            chkMesAviso.Enabled = true;
        }

        private void txtOtrosHaberes_TextChanged(object sender, EventArgs e)
        {
            if (txtOtrosHaberes.Text.Equals("0"))
            {
                rchOtrosHaberes.Enabled = false;
            }

            if (txtOtrosHaberes.Text.Equals(""))
            {
                txtOtrosHaberes.Text = "0";
            }

            if (double.Parse(txtOtrosHaberes.Text) > 0 && !(txtOtrosHaberes.Text.Equals("")))
            {

                rchOtrosHaberes.Enabled = true;
            }

            if (txtOtrosHaberes.Equals("0"))
            {

                rchOtrosHaberes.Text = "";
            }
        }

        private void dtpFiniquito_ValueChanged(object sender, EventArgs e)
        {
            // ve si el formulario se ha ejecutado 1 sola vez


            if(double.Parse(txtDiasProgresivos.Text) > 0){
 
                chkDiasProgresivos.Checked = false;
                txtDiasProgresivos.Text = "0";

            }


            if (contLoad > 0 && txtRut.Text != "")
            {
            if (clickCOnsultar == true)
            {
            calculo_mes_dias_agnos_trabajados();
            clickCOnsultar = false;
            }

            try
            {
                if (int.Parse(txtMesesTrabajados.Text) > 0 && cbxCausal.Enabled == true && cbxCausal.SelectedIndex > 0)
                {
                    calculo_mes_dias_agnos_trabajados();
                    calculo_feriado();
                }
                else
                {
                    txtFeriadoProp.Text = "0";
                    calculo_mes_dias_agnos_trabajados();
                    txtCalVacP.Text = "0";
                }
            }catch(Exception ex){
                MessageBox.Show(ex.Message);
            }

           }
            contLoad = contLoad + 1;
        }

        private void FormularioFiniquito_FormClosed(object sender, FormClosedEventArgs e)
        {
            contLoad = 0;
        }

        private void txtLiq1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtLiq2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtLiq3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtSB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtGratificacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtMovilizacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtColacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtPrf_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtPrv_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtOtrsoDcts_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

 

        private void txtCtaCrrteEmp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtCalVd_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtOtrosHaberes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtCalVacP_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtFeriadoProp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtMesxAgno_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtMesAviso_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtDiasTomados_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtMesAviso_TextChanged(object sender, EventArgs e)
        {

        }

        private void desabilitarControles(){
           // txtRut.Enabled = false;
            txtNombres.Enabled = false;
            txtApePat.Enabled = false;
            txtApeMat.Enabled = false;
            txtDireccion.Enabled = false;
            txtCiudad.Enabled = false;
            txtComuna.Enabled = false;
            txtCargo.Enabled = false;
            txtAgnoServ.Enabled = false;
            txtNdt.Enabled = false;
            txtMesesTrabajados.Enabled = false;
            dtpIngreso.Enabled = false;
            dtpFiniquito.Enabled = false;
            button1.Enabled = false;

            txtLiq1.Enabled = false;
            txtLiq2.Enabled = false;
            txtLiq3.Enabled = false;
            btnCal.Enabled = false;
            chkComision.Enabled = false;
            chkSB.Enabled = false;
            chkGratificacion.Enabled = false;
            chkMovilizacion.Enabled = false;
            chkColacion.Enabled = false;
            btnCalcularPRF.Enabled = false;
            txtSB.Enabled = false;
            txtGratificacion.Enabled = false;
            txtMovilizacion.Enabled = false;
            txtColacion.Enabled = false;

            txtPrv.Enabled = false;
            txtOtrsoDcts.Enabled = false;
            txtOtrosHaberes.Enabled = false;
            rchOtrosDcts.Enabled = false;
            rchOtrosHaberes.Enabled = false;
            txtCalVacP.Enabled = false;
            txtCalVd.Enabled = false;
 
            txtCtaCrrteEmp.Enabled = false;

            cbxCausal.Enabled = false;
            chkMesAviso.Enabled = false;
            btnCalcFiniq.Enabled = false;

            txtRemLiqPen.Enabled = false;
            txtSaldoCajaCompensacion.Enabled = false;
            txtSegCes.Enabled = false;
            txtDctoFondoDijo.Enabled = false;
            txtCtaCorrientePer.Enabled = false;


        }

        private void habilitarControles()
        {
            // txtRut.Enabled = false;
            txtNombres.Enabled = true;
            txtApePat.Enabled = true;
            txtApeMat.Enabled = true;
            txtDireccion.Enabled = true;
            txtCiudad.Enabled = true;
            txtComuna.Enabled = true;
            txtCargo.Enabled = true;
            txtAgnoServ.Enabled = true;
            txtNdt.Enabled = true;
            txtMesesTrabajados.Enabled = true;
            dtpIngreso.Enabled = true;
            dtpFiniquito.Enabled = true;
            button1.Enabled = true;

            txtLiq1.Enabled = true;
            txtLiq2.Enabled = true;
            txtLiq3.Enabled = true;
            btnCal.Enabled = true;

            chkSB.Enabled = true;
            chkGratificacion.Enabled = true;
            chkMovilizacion.Enabled = true;
            chkColacion.Enabled = true;
            btnCalcularPRF.Enabled = true;
            txtSB.Enabled = true;
            txtGratificacion.Enabled = true;
            txtMovilizacion.Enabled = true;
            txtColacion.Enabled = true;

            txtPrv.Enabled = true;
            txtOtrsoDcts.Enabled = true;
            txtOtrosHaberes.Enabled = true;
            rchOtrosDcts.Enabled = true;
            rchOtrosHaberes.Enabled = true;
            txtCalVacP.Enabled = true;
            txtCalVd.Enabled = true;
 
            txtCtaCrrteEmp.Enabled = true;

            cbxCausal.Enabled = true;
            chkMesAviso.Enabled = true;
            btnCalcFiniq.Enabled = true;
            btnBaseCalculo.Enabled = true;

            txtRemLiqPen.Enabled = true;
            txtSaldoCajaCompensacion.Enabled = true;
            txtSegCes.Enabled = true;
            txtDctoFondoDijo.Enabled = true;
            txtCtaCorrientePer.Enabled = true;

        }



        private void articulo() {
            using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
            {
                // ME TRAE EL DIA DE TEMRINO Y EL MES DE TERMINO 
                SqlCommand sqlCM0 = new SqlCommand("sp_ver_articulo_x_descripcion", conn);
                sqlCM0.CommandType = CommandType.StoredProcedure;
                sqlCM0.Parameters.Add("@desc", SqlDbType.VarChar).Value = cbxCausal.Text;
                SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                DataTable DT0 = new DataTable();
                sqlDA0.Fill(DT0);
                DataRow row0 = DT0.Rows[0];


                art = Convert.ToString(row0["articulo"].ToString());
                indicador = Convert.ToString(row0["indicador"].ToString());
                desc_art = Convert.ToString(row0["descripcion"].ToString());
            }
        
        }

        private void btnGenerarFiniquito_Click(object sender, EventArgs e)
        {

            datos();

            if (!cbxCausal.Text.Equals("-- Ingrese Causal --"))
            {
                if (btnCalcFiniq.Enabled == false)
                {
                
                    //montos a palabras
                    using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;Integrated Security=False;User ID=jbustos;Password=jbustos2014"))
                    {
                        // ME TRAE EL DIA DE TEMRINO Y EL MES DE TERMINO 
                        SqlCommand sqlCM0 = new SqlCommand("sp_montos_a_palabras", conn);
                        sqlCM0.CommandType = CommandType.StoredProcedure;
                      
                        sqlCM0.Parameters.Add("@monto", SqlDbType.Float).Value = Double.Parse(lqdo_pago.Replace(".",""));
                        SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                        DataTable DT0 = new DataTable();
                        sqlDA0.Fill(DT0);
                        DataRow row0 = DT0.Rows[0];

                        monto_palabra_f = Convert.ToString(row0["Palabra_Monto"].ToString());

                    }

                    articulo();

                    String monto_palabras = monto_palabra_f;

                    // validacion numeros cuando vienen con doble cero
                   
                    if (indep_x_agno_ser.Equals("0") || indep_x_agno_ser.Equals("00")) { indep_x_agno_ser = "0"; }
                    if (indep_sus_x_aviso.Equals("0") || indep_sus_x_aviso.Equals("00")) { indep_sus_x_aviso = "0"; }
                    if (feariado_pro.Equals("0") || feariado_pro.Equals("00")) { feariado_pro = "0"; }
                    if (haberes.Equals("0") || haberes.Equals("00")) { haberes = "0"; }
                    if (dcto.Equals("0") || dcto.Equals("00")) { dcto = "0"; }
                    if (lqdo_pago.Equals("0") || lqdo_pago.Equals("00")) { lqdo_pago = "0"; }
                    if (otros_haberes.Equals("0") || otros_haberes.Equals("00")) { otros_haberes = "0"; }
                    if (otros_dcts.Equals("0") || otros_dcts.Equals("00")) { otros_dcts = "0"; }


                    // si esta normal
                    if (txtOtrsoDcts.Text.Equals("0") && txtOtrosHaberes.Text.Equals("0"))
                    {
                        MessageBox.Show("Creando Carta Finiquito", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        SaveFileDialog saveFileDialog2 = new SaveFileDialog();

                        saveFileDialog2.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                        {

                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta3(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                                agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                                afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                                art, indicador, desc_art, monto_palabra_f);

                            word.guardar(saveFileDialog2);

                            if (double.Parse(txtMesesTrabajados.Text) >= 1)
                            {
                                MessageBox.Show("Termino de Contrato de Trabajo", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }


                        }
                    }

                    // si hay descuentos sin descripcion
                    if ((double.Parse(txtOtrsoDcts.Text)) > 0 &&
                        (double.Parse(txtOtrosHaberes.Text)) > 0 &&
                        rchOtrosDcts.Text.Equals("") &&
                        rchOtrosHaberes.Text.Equals("")

                        )
                    {
                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                        saveFileDialog1.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta4(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                                agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                                afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                                art, indicador, desc_art, monto_palabra_f);

                            word.guardar(saveFileDialog1);
                        }
                    }

                    // si hay descuentos con descripcion
                    if ((double.Parse(txtOtrsoDcts.Text)) > 0 &&
                        (double.Parse(txtOtrosHaberes.Text)) > 0 &&
                        !(rchOtrosDcts.Text.Equals("") &&
                        rchOtrosHaberes.Text.Equals(""))

                        )
                    {
                        SaveFileDialog saveFileDialog3 = new SaveFileDialog();

                        saveFileDialog3.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog3.ShowDialog() == DialogResult.OK)
                        {
                            String uno = rchOtrosHaberes.Text;
                            String dos = rchOtrosDcts.Text;
                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta5(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                                agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                                afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                                art, indicador, desc_art, monto_palabras, uno, dos);
                            word.guardar(saveFileDialog3);
                        }
                    }

                    // Con Otros Descuentos sin desc
                    if ((double.Parse(txtOtrsoDcts.Text)) > 0 &&
                             rchOtrosDcts.Text.Equals("") &&

                        (double.Parse(txtOtrosHaberes.Text)) == 0 &&
                           rchOtrosHaberes.Text.Equals("")
                        )
                    {
                        SaveFileDialog saveFileDialog4 = new SaveFileDialog();

                        saveFileDialog4.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog4.ShowDialog() == DialogResult.OK)
                        {

                            String uno = rchOtrosHaberes.Text;
                            String dos = rchOtrosDcts.Text;
                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta6(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                                agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                                afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                                art, indicador, desc_art, monto_palabras, uno, dos);
                            word.guardar(saveFileDialog4);
                        }
                    }

                    // Con Otros Descuentos con desc
                    if ((double.Parse(txtOtrsoDcts.Text)) > 0 &&
                           (!rchOtrosDcts.Text.Equals("")) &&

                       (double.Parse(txtOtrosHaberes.Text)) == 0 &&
                          rchOtrosHaberes.Text.Equals("")
                       )
                    {

                        SaveFileDialog saveFileDialog4 = new SaveFileDialog();

                        saveFileDialog4.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog4.ShowDialog() == DialogResult.OK)
                        {
                            String uno = rchOtrosHaberes.Text;
                            String dos = rchOtrosDcts.Text;
                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta7(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                                agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                                afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                                art, indicador, desc_art, monto_palabras, uno, dos);
                            word.guardar(saveFileDialog4);

                        }
                    }

                    // Con Haberes  sin desc
                    if ((double.Parse(txtOtrosHaberes.Text)) > 0 &&
                           rchOtrosHaberes.Text.Equals("") &&

                      (double.Parse(txtOtrsoDcts.Text)) == 0 &&
                         rchOtrosDcts.Text.Equals("")
                      )
                    {

                        SaveFileDialog saveFileDialog4 = new SaveFileDialog();

                        saveFileDialog4.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog4.ShowDialog() == DialogResult.OK)
                        {

                            String uno = rchOtrosHaberes.Text;
                            String dos = rchOtrosDcts.Text;
                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta8(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                                agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                                afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                                art, indicador, desc_art, monto_palabras, uno, dos);
                            word.guardar(saveFileDialog4);
                        }

                    }

                    // Con Haberes  con desc
                    if ((double.Parse(txtOtrosHaberes.Text)) > 0 &&
                          (!rchOtrosHaberes.Text.Equals("")) &&

                      (double.Parse(txtOtrsoDcts.Text)) == 0 &&
                         rchOtrosDcts.Text.Equals("")
                      )
                    {
                        SaveFileDialog saveFileDialog4 = new SaveFileDialog();

                        saveFileDialog4.Filter = "Office Files|*.doc;*.docx";
                        if (saveFileDialog4.ShowDialog() == DialogResult.OK)
                        {
                            String uno = rchOtrosHaberes.Text;
                            String dos = rchOtrosDcts.Text;
                            CartaWord word = new CartaWord();
                            word.Config2();
                            word.texto_carta8(fecha_in, fecha_fin, nombre_emp, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                              agno_serv, indep_x_agno_ser, indep_sus_x_aviso, feariado_pro, feariado_pro_dias,
                                              afc, cta_cte_emp, dcto, lqdo_pago, otros_haberes, otros_dcts, total, haberes,
                                              art, indicador, desc_art, monto_palabras, uno, dos);
                            word.guardar(saveFileDialog4);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Para generar el documento primero debe guardar el finiquito", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    
                }
            }
            else {
                MessageBox.Show("Debe seleccionar almenos una causal", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnEditarFrm_Click(object sender, EventArgs e)
        {
            btnGenerarExcel.Enabled = false;
            btnGenerarTermino.Enabled = false;
            btnCartaExcel.Enabled = false;
            habilitarControles();
        }

        private void btnGenerarTermino_Click(object sender, EventArgs e)
        {
            datos();
            articulo();
            if (btnCalcFiniq.Enabled == false)
            {

                if(double.Parse(txtFeriadoProp.Text)>0){

                    SaveFileDialog saveFileDialog4 = new SaveFileDialog();

                    saveFileDialog4.Filter = "Office Files|*.doc;*.docx";
                    if (saveFileDialog4.ShowDialog() == DialogResult.OK)
                    {
                        String uno = rchOtrosHaberes.Text;
                        String dos = rchOtrosDcts.Text;
                        CartaWord word = new CartaWord();
                        word.Config();
                        word.texto_carta2(nombre_emp,ape_pat, rut_emp, domicilio_emp, comuna_emp,ciudad_emp,
                                    txtCargo.Text,fecha_fin, art,indicador,desc_art,
                                    indep_x_agno_ser , fecha_in);
                                                            word.guardar(saveFileDialog4);
                    }


                }else{
                    articulo();
                    SaveFileDialog saveFileDialog4 = new SaveFileDialog();

                    saveFileDialog4.Filter = "Office Files|*.doc;*.docx";
                    if (saveFileDialog4.ShowDialog() == DialogResult.OK)
                    {
                        String uno = rchOtrosHaberes.Text;
                        String dos = rchOtrosDcts.Text;
                        CartaWord word = new CartaWord();
                        word.Config();
                        word.texto_carta1(nombre_emp, ape_pat, rut_emp, domicilio_emp, comuna_emp, ciudad_emp,
                                    txtCargo.Text, fecha_fin, art, indicador, desc_art,
                                    indep_x_agno_ser, fecha_in);
                        word.guardar(saveFileDialog4);
                    }
                }

            }
            else {

                MessageBox.Show("Para generar el documento primero debe guardar el finiquito", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            
            
            }
        }

        private void txtCalVd_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnBaseCalculo_Click(object sender, EventArgs e)
        {
            cbxCausal.Enabled = true;
            chkMesAviso.Enabled = true;
            txtRemLiqPen.Enabled = true;
            txtSaldoCajaCompensacion.Enabled = true;
            txtSegCes.Enabled = true;
            txtDctoFondoDijo.Enabled = true;
            txtBaseCalculo.Text = Convert.ToString(double.Parse(txtPrf.Text) + double.Parse(txtPrv.Text) + double.Parse(txtOtrosHaberes.Text));
        }

        private void txtLiq1_TextChanged(object sender, EventArgs e)
        {
            btnCalcularPRF.Enabled = false;
            txtPrf.Text = "0";
            txtPrv.Text = "0";
            txtSB.Text = sb;
            txtGratificacion.Text = grat;
            txtMovilizacion.Text = mov;
            txtColacion.Text = col;
            chkMesAviso.Enabled = true;

     
    
            txtFinqApagar.Text = "0";
            txtMesAviso.Text = "0";
            txtCalVd.Text = "0";
            txtCalVacP.Text = "0";
 
            cbxCausal.Enabled = false;
            cbxCausal.SelectedIndex = 0;
            txtOtrsoDcts.Enabled = false;

            txtOtrsoDcts.Text = "0";
            txtOtrosHaberes.Text = "0";
            rchOtrosDcts.Text = "";
            rchOtrosHaberes.Text = "";

            txtCtaCrrteEmp.Text = "0";
            txtBaseCalculo.Text = "0";
            btnBaseCalculo.Enabled = false;
            txtFeriadoProp.Text = "0";
 
            txtCtaCrrteEmp.Enabled = false;
            txtOtrosHaberes.Enabled = false;
            txtFinqApagar.Enabled = false;
            btnCalcFiniq.Enabled = false;
           


            if (txtLiq1.Text.Equals(""))
            {
                txtLiq1.Text = "0";
            }
     
        }

        private void txtLiq2_TextChanged(object sender, EventArgs e)
        {
            if (txtLiq2.Text.Equals(""))
            {
                txtLiq2.Text = "0";
            }
            
            
            btnCalcularPRF.Enabled = false;
            txtPrf.Text = "0";
            txtSB.Text = sb;
            txtGratificacion.Text = grat;
            txtMovilizacion.Text = mov;
            txtColacion.Text = col;
            chkMesAviso.Enabled = false;

     
    
            txtPrv.Text = "0";
 
            txtFinqApagar.Text = "0";
            txtMesAviso.Text = "0";
            txtCalVd.Text = "0";
            txtCalVacP.Text = "0";
   
            cbxCausal.Enabled = false;
            cbxCausal.SelectedIndex = 0;
            txtOtrsoDcts.Enabled = false;
            txtOtrsoDcts.Text = "0";
            txtOtrosHaberes.Text = "0";
            rchOtrosDcts.Text = "";
            rchOtrosHaberes.Text = "";
    
            txtCtaCrrteEmp.Text = "0";
            txtBaseCalculo.Text = "0";
            btnBaseCalculo.Enabled = false;
            txtFeriadoProp.Text = "0";
 
            txtCtaCrrteEmp.Enabled = false;
            txtOtrosHaberes.Enabled = false;
            txtFinqApagar.Enabled = false;
            btnCalcFiniq.Enabled = false;
 
        }

        private void txtLiq3_TextChanged(object sender, EventArgs e)
        {
            if (txtLiq3.Text.Equals(""))
            {
                txtLiq3.Text = "0";
            }
            
            btnCalcularPRF.Enabled = false;
            txtPrf.Text = "0";
            txtSB.Text = sb;
            txtGratificacion.Text = grat;
            txtMovilizacion.Text = mov;
            txtColacion.Text = col;
            chkMesAviso.Enabled = false;

            cbxCausal.Enabled = false;
            txtPrf.Text = "0";
            txtPrv.Text = "0";
 
            txtFinqApagar.Text = "0";
            txtMesAviso.Text = "0";
            txtCalVd.Text = "0";
            txtCalVacP.Text = "0";
 
            cbxCausal.Enabled = false;
            cbxCausal.SelectedIndex = 0;
            txtOtrsoDcts.Enabled = false;
            txtOtrsoDcts.Text = "0";
            txtOtrosHaberes.Text = "0";
            rchOtrosDcts.Text = "";
            rchOtrosHaberes.Text = "";
  
            txtCtaCrrteEmp.Text = "0";
            txtBaseCalculo.Text = "0";
            btnBaseCalculo.Enabled = false;
            txtFeriadoProp.Text = "0";
    
            txtCtaCrrteEmp.Enabled = false;
            txtOtrosHaberes.Enabled = false;
            txtFinqApagar.Enabled = false;
            btnCalcFiniq.Enabled = false;
 
        }

      
        private void txtCtaCrrteEmp_TextChanged(object sender, EventArgs e)
        {
            if(txtCtaCrrteEmp.Text.Equals("")){
                txtCtaCrrteEmp.Text = "0";
            }
        }

        private void txtAgnoServ_TextChanged(object sender, EventArgs e)
        {
            if(!txtRut.Text.Equals("")){
            if (double.Parse(txtAgnoServ.Text) > 0)
            {
                chkMesAviso.Enabled = true;
            }
            else {
                chkMesAviso.Enabled = false;
            }
            }
        }
        private void datos()
        {

            //parrafo 1, 2 y 3
            DateTime dt = new DateTime(dtpIngreso.Value.Year, dtpIngreso.Value.Month, dtpIngreso.Value.Day);
            fecha_in = dt.ToLongDateString();

            DateTime dt2 = new DateTime(dtpFiniquito.Value.Year, dtpFiniquito.Value.Month, dtpFiniquito.Value.Day);
            fecha_fin = dt2.ToLongDateString();

            nombre_emp = txtNombres.Text + " " + txtApePat.Text + " " + txtApeMat.Text;
            ape_pat = txtApePat.Text;
            ape_mat = txtApeMat.Text;
            rut_emp = formatearRut(txtRut.Text);
            domicilio_emp = txtDireccion.Text;
            comuna_emp = txtComuna.Text;
            ciudad_emp = txtCiudad.Text;

            if (txtSegCes.Text.Equals("0"))
            {
            afc = txtSegCes.Text;
            }else{
                afc = double.Parse(txtSegCes.Text).ToString("0,0", elGR);
            }
            // parrafo 4
            agno_serv = txtAgnoServ.Text;

            indep_x_agno_ser = (Convert.ToDecimal((int.Parse(txtMesxAgno.Text)))).ToString("0,0", elGR);
            indep_sus_x_aviso = (double.Parse(txtMesAviso.Text)).ToString("0,0", elGR);
            feariado_pro = (double.Parse(txtCalVacP.Text)).ToString("0,0", elGR);
            feariado_pro_dias = (Math.Truncate(double.Parse(txtFeriadoProp.Text))).ToString();

            // parrafo 5
      
            cta_cte_emp = txtCtaCrrteEmp.Text;
            dcto = (double.Parse(txtOtrsoDcts.Text) + double.Parse(txtSegCes.Text) + double.Parse(cta_cte_emp)).ToString("0,0", elGR);
            lqdo_pago = (double.Parse(txtFinqApagar.Text)).ToString("0,0", elGR);
            otros_haberes = double.Parse(txtOtrosHaberes.Text).ToString("0,0", elGR);
            otros_dcts = double.Parse(txtOtrsoDcts.Text).ToString("0,0", elGR);

            total = double.Parse(txtMesxAgno.Text) + double.Parse(txtMesAviso.Text) + double.Parse(otros_haberes.Replace(".","")) + double.Parse(feariado_pro.Replace(".",""));
            //total = double.Parse(txtPrf.Text) + double.Parse(txtOtrosHaberes.Text);
            if (total == 0)
            {
                haberes = "0";
            }else{
               haberes = total.ToString("0,0", elGR);
            }
         


        }
        private void btnCartaExcel_Click(object sender, EventArgs e)
        {

            btnBaseCalculo.Enabled = false;
            txtRemLiqPen.Enabled = false;
            txtSaldoCajaCompensacion.Enabled = false;
            txtSegCes.Enabled = false;
            txtDctoFondoDijo.Enabled = false;

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Office Files|*.xls;*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                datos();
                CartasExcel c = new CartasExcel();

                using (SqlConnection conn = new SqlConnection("Data Source=10.1.3.227;Initial Catalog=TRANSAC_GRANGUIA;Integrated Security=False;User ID=softdesa;Password=softdesa2014"))
                {
                    // ME TRAE EL DIA DE TEMRINO Y EL MES DE TERMINO 
                    SqlCommand sqlCM0 = new SqlCommand("sp_ver_datos_calculo_finiquito_excel", conn);
                    sqlCM0.CommandType = CommandType.StoredProcedure;
                    sqlCM0.Parameters.Add("@ficha", SqlDbType.VarChar).Value = formatearRut(rut_emp);
                    SqlDataAdapter sqlDA0 = new SqlDataAdapter(sqlCM0);
                    DataTable DT0 = new DataTable();
                    sqlDA0.Fill(DT0);

                    String horas_ex = "0";
                    String bono_pro = "0";
                    String desc_cel = "0";
                    String seg_ces = "0";


                    if (DT0.Rows.Count > 0)
                    {
                        DataRow row0 = DT0.Rows[0];
                        horas_ex = Convert.ToString(row0["horas_extra"]).ToString();
                        bono_pro = Convert.ToString(row0["bono_produccion"]).ToString();
                        desc_cel = Convert.ToString(row0["descuentos_celular"]).ToString();
                        seg_ces = Convert.ToString(row0["seguro_cesantia"]).ToString();

                    }
                  

                String nom_empleado = nombre_emp;
                String direccion = domicilio_emp; 
                String causal = cbxCausal.Text;
                String articulo = "-"; 
                String cargo = txtCargo.Text;
                String rut = rut_emp; 
                String fono = "-";
                String fec_in = dtpIngreso.Text;
                String fec_ter = dtpFiniquito.Text;
                String fec_finiquito = dtpTermino.Text;

                String concepto = "-"; 
                String sb = txtSB.Text; 
                String gratificacion = txtGratificacion.Text;
                String comision = bono_pro; 
                String movilizacion = txtMovilizacion.Text;
                String semana_corrida = feariado_pro.Replace(".", "");
                String tot_basecal = (double.Parse(sb) + double.Parse(gratificacion) +double.Parse(txtColacion.Text) +
                                    double.Parse(comision) + double.Parse(movilizacion) + double.Parse(semana_corrida)).ToString(); 

                String ultima_rem = txtLiq3.Text; 
                String detalle = "-";
              
                String indemnizacion_vol = "0";
                String indemnizacion_agno_Serv = txtMesxAgno.Text; 
                String mes_sus = txtMesAviso.Text;
                String vac_prop = txtCalVacP.Text;
                String rem_liq_prop =txtRemLiqPen.Text;
                String horas_extra = horas_ex;
                String bono_prod = bono_pro; 
                String bono_prod_adic = "-";
                //String tot_haberes = (double.Parse(txtMesxAgno.Text) + double.Parse(txtMesAviso.Text) + double.Parse(vac_prop) +
                //                    double.Parse(gratificacion) + double.Parse(horas_extra) + double.Parse(movilizacion) + 
                //                    double.Parse(bono_prod)).ToString();
                String valor_finiquito = txtFinqApagar.Text;

                String ctaCtePersonal = "-";
                String dctoCelular = desc_cel;
                String otrosDcts = txtOtrsoDcts.Text;
                String tot_dcts = seg_ces+double.Parse(dctoCelular)+double.Parse(txtOtrsoDcts.Text);
                String colacion = txtColacion.Text;

                String pendientes_mes;
                String pendientes_dias; 
                String haberesX = (double.Parse(indemnizacion_vol) + double.Parse(indemnizacion_agno_Serv) + double.Parse(mes_sus) +
                                              double.Parse(vac_prop) + double.Parse(rem_liq_prop)).ToString();
                if ((double.Parse(dias_p) - double.Parse(txtDiasTomados.Text)) == 0)
                {
                    pendientes_mes = "0";
                    pendientes_dias = "0";
                }
                else {
                    pendientes_dias = ((double.Parse(dias_p)) - (double.Parse(txtDiasTomados.Text))).ToString();
                    pendientes_mes = ""; 
                }


           
                String proporcionales_mes = dias_p;
                String proporcionales_dias = (double.Parse(dias_p) * 1.25).ToString();
                String legal_mes = double.Parse(dias_p).ToString();
                String legal_dias = (double.Parse(dias_p) * 1.25).ToString();
                String tot_dias_co = (double.Parse(dias_p) * 1.25).ToString(); ;


                String cajaCompensacion = txtSaldoCajaCompensacion.Text;
                String seguroSec = txtSegCes.Text;
                String fondoFijo = txtDctoFondoDijo.Text;
                String ctaCorrientePersonal = txtCtaCorrientePer.Text;
                String tot_desc = (double.Parse(txtSaldoCajaCompensacion.Text) + double.Parse(txtSegCes.Text) + double.Parse(txtDctoFondoDijo.Text) +
                                              double.Parse(txtCtaCorrientePer.Text) + double.Parse(otros_dcts) + double.Parse(dctoCelular)).ToString();

                    c.crearCarta(
                        nom_empleado,  
                        direccion,  
                        causal,  
                        articulo,  
                        cargo,  
                        rut,  
                        fono,  
                        fec_in,  
                        fec_ter,  
                        fec_finiquito,  
                        concepto,  
                        sb,  
                        gratificacion, 
                        colacion,
                        comision,  
                        movilizacion,  
                        semana_corrida,
                        tot_basecal,
                        ultima_rem,  
                        detalle, 
                        haberesX,  
                        indemnizacion_vol, 
                        indemnizacion_agno_Serv,  
                        mes_sus,  
                        vac_prop,  
                        rem_liq_prop,  
                        horas_extra,  
                        bono_prod, 
                        bono_prod_adic,
                        haberesX,
                        valor_finiquito,
                        desc_cel,
                        seg_ces,
                        ctaCtePersonal,
                        dctoCelular,
                        otrosDcts,
                        tot_dcts,
                        pendientes_mes,
                        pendientes_dias,
                        proporcionales_mes,
                        proporcionales_dias,
                        legal_mes,
                        legal_dias,
                        tot_dias_co,
                        cajaCompensacion,
                        seguroSec,
                        fondoFijo,
                        ctaCorrientePersonal,
                        tot_desc);
                c.guardar(saveFileDialog);

                MessageBox.Show("Documento Excel Generado", "Mensage del sistema", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

        }

        private void txtRemLiqPen_TextChanged(object sender, EventArgs e)
        {
            if(txtRemLiqPen.Text.Equals("")){
                txtRemLiqPen.Text = "0";
            }


        }

        private void txtColacion_TextChanged(object sender, EventArgs e)
        {
            if (txtColacion.Text.Equals(""))
            {
                txtColacion.Text = "0";
            }
        }

        private void txtRemLiqPen_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtSaldoCajaCompensacion_TextChanged(object sender, EventArgs e)
        {
            if (txtSaldoCajaCompensacion.Text.Equals(""))
            {
                txtSaldoCajaCompensacion.Text = "0";
            }
        }

        private void txtSaldoCajaCompensacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtSegCes_TextChanged(object sender, EventArgs e)
        {
            if (txtSegCes.Text.Equals(""))
            {
                txtSegCes.Text = "0";
            }

 
        }

        private void txtSegCes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtDctoFondoDijo_TextChanged(object sender, EventArgs e)
        {
            if (txtDctoFondoDijo.Text.Equals(""))
            {
                txtDctoFondoDijo.Text = "0";
            }
        }

        private void txtDctoFondoDijo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtCtaCorrientePer_TextChanged(object sender, EventArgs e)
        {

            if (txtCtaCorrientePer.Text.Equals(""))
            {
                txtCtaCorrientePer.Text = "0";
            }

        }

        private void txtCtaCorrientePer_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void chkComision_CheckedChanged(object sender, EventArgs e)
        {
            if (chkComision.Checked == true)
            {
                txtComision.Enabled = true;
            }
            else
            {
                txtComision.Enabled = false;
            }
        }

        private void txtComision_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtComision_TextChanged(object sender, EventArgs e)
        {

            if (txtComision.Text.Equals(""))
            {
                txtComision.Text = "0";
            }
        }

        private void chkDiasProgresivos_CheckedChanged(object sender, EventArgs e)
        {
            if (chkDiasProgresivos.Checked == true)
            {
                txtDiasProgresivos.Enabled = true;
            }
            else
            {
                txtDiasProgresivos.Enabled = false;
            }
        }

        private void txtDiasProgresivos_TextChanged(object sender, EventArgs e)
        {
            if (txtDiasProgresivos.Text.Equals("") || txtDiasProgresivos.Text.Equals("0"))
            {
                txtDiasProgresivos.Text = "0";

                if (double.Parse(txtFeriadoProp.Text) > 0)
                {
                    calculo_feriado();
                }
            }

            if (double.Parse(txtDiasProgresivos.Text) > 0 && double.Parse(txtFeriadoProp.Text )>0)
            {
                txtFeriadoProp.Text = (double.Parse(txtFeriadoProp.Text) + double.Parse(txtDiasProgresivos.Text)).ToString();
                txtCalVacP.Text = ((double.Parse(txtFeriadoProp.Text) + double.Parse(txtDiasProgresivos.Text))*1.25).ToString();
            }





        }

        private void txtDiasProgresivos_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
            {

                e.Handled = true;

            }
        }

        private void txtSB_TextChanged(object sender, EventArgs e)
        {
            if (txtSB.Text.Equals(""))
            {
                txtSB.Text = "0";
            }
        }

        private void txtGratificacion_TextChanged(object sender, EventArgs e)
        {
            if (txtGratificacion.Text.Equals(""))
            {
                txtGratificacion.Text = "0";
            }
        }

        private void txtMovilizacion_TextChanged(object sender, EventArgs e)
        {
            if (txtMovilizacion.Text.Equals(""))
            {
                txtMovilizacion.Text = "0";
            }
        }

       


        

         
    }
}
