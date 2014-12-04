using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Finiquitos
{
    class CartasExcel
    {

        private Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        private object misValue = System.Reflection.Missing.Value;


          


        public void crearCarta(String nom_empleado, String direccion, String causal,
                                String articulo, String cargo, String rut, String fono,
                                String fec_in, String fec_ter, String fec_finiquito,
                                String concepto, String sb, String gratificacion,String colacion,
                                String comision, String movilizacion, String semana_corrida,String tot_basecal,
                                String ultima_rem, String detalle, String haberes,
                                String indemnizacion_vol, String indemnizacion_agno_Serv, String mes_sus,
                                String vac_prop, String rem_liq_prop, String horas_extra,
                                String bono_prod, String bono_prod_adic, String tot_haberes,
                                String valor_finiquito,String desc_cel,String seg_ces,String ctaCtePersonal,
                                String dctoCelular,String otrosDcts,String tot_dcts,
                                String pendientes_mes,String pendientes_dias,String proporcionales_mes,
                                String proporcionales_dias,String legal_mes,String legal_dias, String tot_dias_co,
                                String cajaCompensacion, String seguroSec, String fondoFijo, String ctaCorrientePersonal,String tot_desc){
            
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperLetter;
            // -----------------------------------------------------------------------------------
            // CALCULO DE FINIQUITO
            Range range = null;

            // ----------------------------------------------------------------------
            crearColumna(range, "A2", "J5", "CALCULO DE FINIQUITO", true, true, 1, 12);
            // -----------------------------------------------------------------------

            crearColumna(range, "A8", "B8", "Nombre del Empleado:", true, false, 1, 9);
            crearColumna(range, "A9", "B9", "Dirección:", true, false, 1, 9);
            crearColumna(range, "A10", "B10", "Causal de despido:", true, false, 1, 9);
         //   crearColumna(range, "A11", "B11", "Articulo citado:", true, false, 1, 9);
            crearColumna(range, "A11", "B11", "Cargo:", true, false, 1, 9);

            //Relleno Nombre Empleado
            crearColumna(range, "C8", "F8", nom_empleado, true, false, 0, 9);
            //Relleno Dirección
            crearColumna(range, "C9", "F9", direccion, true, false, 0, 9);
            //Relleno Causal de Despido
            crearColumna(range, "C10", "F10",causal, true, false, 1, 9);
            //Relleno Articulo citado
          //  crearColumna(range, "C11", "F11", articulo, true, false, 0, 9);
            //Relleno Cargo
            crearColumna(range, "C11", "F11", cargo, true, false, 0, 9);


            crearColumna(range, "G8", "H8", "RUT:", true, false, 1, 9);
            crearColumna(range, "G9", "H9", "Fono:", true, false, 1, 9);
            crearColumna(range, "G10", "H10", "Fecha de Ingreso:", true, false, 1, 9);
            crearColumna(range, "G11", "H11", "Fecha de Termino:", true, false, 1, 9);
            crearColumna(range, "G12", "H12", "Fecha de Finiquito:", true, false, 1, 9);

            //Relleno RUT
            crearColumna(range, "I8", "J8", rut, true, false, 0, 9);
            //Relleno Fono
            crearColumna(range, "I9", "J9", fono, true, false, 0, 9);
            //Relleno Fecha de Ingreso
            crearColumna(range, "I10", "J10", fec_in, true, false, 0, 9);
            //Relleno Fecha de Termino
            crearColumna(range, "I11", "J11", fec_ter, true, false, 0, 9);
            //Relleno Fecha de Finiquito
            crearColumna(range, "I12", "J12", fec_finiquito, true, false, 0, 9);
            // -------------------------------------------------------------------------------------
            crearColumna(range, "A14", "B14", "Concepto", true, true, 1, 9);
            // -------------------------------------------------------------------------------------
            crearColumna(range, "A15", "B15", "Sueldo Base", true, false, 1, 9);
            crearColumna(range, "A16", "B16", "Gratificación", true, false, 1, 9);
            crearColumna(range, "A17", "B17", "Colacion", true, false, 1, 9);
            crearColumna(range, "A18", "B18", "Comisión/Bonos variables", true, false, 1, 9);
            crearColumna(range, "A19", "B19", "Movilización", true, false, 1, 9);
            crearColumna(range, "A20", "B20", "Semana Corrida", true, false, 1, 9);
            crearColumna(range, "A21", "B21", "Total", true, false, 1, 9);
            // -------------------------------------------------------------------------------------
            crearColumna(range, "C14", "D14", "Ultima Remuneración", true, false, 1, 9);
            // -------------------------------------------------------------------------------------

            //Relleno sueldo base
            crearColumna(range, "C15", "D15", sb, true, false, 0, 9);
            //Relleno gratificacion
            crearColumna(range, "C16", "D16", gratificacion, true, false, 0, 9);
            //Relleno colacion  
            crearColumna(range, "C17", "D17", colacion, true, false, 0, 9);
            //Relleno comision
            crearColumna(range, "C18", "D18", comision, true, false, 0, 9);
            //Relleno movilizacion
            crearColumna(range, "C19", "D19",movilizacion, true, false, 0, 9);
            //Relleno Semana Corrida
            crearColumna(range, "C20", "D20", semana_corrida, true, false, 0, 9);
            //relleno Total (base de calculo
            crearColumna(range, "C21", "D21", tot_basecal, true, false, 0, 9);
            // -------------------------------------------------------------------------------------
            //observaciones
            crearColumna(range, "A23", "H23", "Observaciones", true, true, 1, 9);

            crearColumna(range, "A24", "H24", "", true, false, 0, 9);
            crearColumna(range, "A25", "H25", "", true, false, 0, 9);
            // -------------------------------------------------------------------------------------
            //Detalle
            crearColumna(range, "A28", "D28", "Detalle", true, true, 1, 9);
            //Haberes
            crearColumna(range, "A29", "C29", "Haberes", true, false, 1, 9);
            //Indemnizacion Voluntaria
            crearColumna(range, "A30", "C30", "Indemnización Voluntaria", true, false, 0, 9);
            //Indemnizacion Años de servicios
            crearColumna(range, "A31", "C31", "Indemnización Años de servicios", true, false, 0, 9);
            //Mes sustitutivo de aviso previo
            crearColumna(range, "A32", "C32", "Mes sustitutivo de aviso previo", true, false, 0, 9);
            //Vacaciones proporcionales
            crearColumna(range, "A33", "C33", "Vacaciones proporcionales", true, false, 0, 9);
            //Remuneracion liquida proporcional
            crearColumna(range, "A34", "C34", "Remuneracion liquida Pendiente", true, false, 0, 9);

            //TOTAL HABERES
            crearColumna(range, "A35", "C35", "TOTAL HABERES", true, false, 0, 9);
            // -------------------------------------------------------------------------------------
            //columna de enmedio
            crearColumna(range, "D29", "D35", "", false, false, 0, 9);
            // -------------------------------------------------------------------------------------
            //Monto
            crearColumna(range, "E28", "F28", "Monto", true, true, 1, 9);

            //Relleno Haber  
            crearColumna(range, "E29", "F29", "", true, false, 1, 9);
            //Relleno Indemnización Voluntaria  
            crearColumna(range, "E30", "F30", indemnizacion_vol, true, false, 1, 9);
            //Relleno Indemnización Años de Servicio  
            crearColumna(range, "E31", "F31", indemnizacion_agno_Serv, true, false, 1, 9);
            //Relleno Mes sustitutivo de aviso previo 
            crearColumna(range, "E32", "F32",mes_sus, true, false, 1, 9);
            //Relleno Vacaciones proporcionales 
            crearColumna(range, "E33", "F33", vac_prop, true, false, 1, 9);
            //Relleno Remuneracion Liquida Pendiente  
            crearColumna(range, "E34", "F34", rem_liq_prop, true, false, 1, 9);
            //Relleno Total Haberes
            crearColumna(range, "E35", "F35", haberes, true, false, 1, 9);
            // -------------------------------------------------------------------------------------
            //Total Haberes
            crearColumna(range, "G28", "J28", "Vacaciones", true, true, 1, 9);

            //encabezado - columna 1 - celda vacia
            crearColumna(range, "G29", "H29", "", true, false, 0, 9);

            //encabezado - columna 2 - meses
            crearColumna(range, "I29", "I29", "Meses", true, true, 1, 9);

            //encabezado - columna 3 - Dias
            crearColumna(range, "J29", "J29", "Días", true, true, 1, 9);
            // -------------------------------------------------------------------------------------
            //Pendientes
            crearColumna(range, "G30", "H30", "Pendientes", true, false, 1, 9);
            //Proporcionales
            crearColumna(range, "G31", "H31", "Proporcionales", true, false, 1, 9);
            //Legal
            crearColumna(range, "G32", "H32", "Legal", true, false, 1, 9);
            // Total Días Corridos
            crearColumna(range, "G33", "H33", "Total Días Corridos", true, false, 1, 9);

            //Relleno meses - pendientes
            crearColumna(range, "I30", "I30", pendientes_mes, true, false, 1, 9);
            //Relleno meses - proporcionales
            crearColumna(range, "I31", "I31",proporcionales_mes, true, false, 1, 9);
            //Relleno meses - Legal
            crearColumna(range, "I32", "I32",legal_mes, true, false, 1, 9);
            //Relleno meses - Total Dias Corridos
            crearColumna(range, "I33", "I33", "", true, false, 1, 9);

            //Relleno Dias - pendientes
            crearColumna(range, "J30", "J30", pendientes_dias, true, false, 1, 9);
            //Relleno Dias - proporcionales
            crearColumna(range, "J31", "J31",proporcionales_dias, true, false, 1, 9);
            //Relleno Dias  - Legal
            crearColumna(range, "J32", "J32",legal_dias, true, false, 1, 9);
            //Relleno Dias - Total Dias Corridos
            crearColumna(range, "J33", "J33", tot_dias_co, true, false, 1, 9);
            // -------------------------------------------------------------------------------------
            crearColumna(range, "A38", "C38", "Descuentos", true, false, 1, 9);
            crearColumna(range, "A39", "C39", "Saldo Capital Caja de Compensación", true, false, 0, 9);
            crearColumna(range, "A40", "C40", "Saldo Aporte empleador Seg Cesantia", true, false, 0, 9);
           
            crearColumna(range, "A41", "C41", "Descuento Cuenta Corriente Personal", true, false, 0, 9);
            crearColumna(range, "A42", "C42", "Descuento fondo fijo", true, false, 0, 9);
            crearColumna(range, "A43", "C43", "Descuento Telefono celular", true, false, 0, 9);
            crearColumna(range, "A44", "C44", "Otros descuentos", true, false, 0, 9);
            crearColumna(range, "A45", "C45", "TOTAL DESCUENTOS", true, false, 0, 9);

            // -------------------------------------------------------------------------------------
            //columna de enmedio
            crearColumna(range, "D38", "D45", "", false, false, 1, 9);
            // -------------------------------------------------------------------------------------
            //Descuentos
            crearColumna(range, "E38", "F38", "", true, false, 1, 9);
            //Saldo Capital Caja de Compensación
            crearColumna(range, "E39", "F39", cajaCompensacion, true, false, 1, 9);
            //Saldo Aporte empleador Seg Cesantia
            crearColumna(range, "E40", "F40", seguroSec, true, false, 1, 9);
            //Descuento Cuenta Corriente Personal
            crearColumna(range, "E41", "F41", ctaCorrientePersonal, true, false, 1, 9);
            //Decuento fondo Fijo
            crearColumna(range, "E42", "F42", fondoFijo, true, false, 1, 9);
            //Descuento Tetelefono celular
            crearColumna(range, "E43", "F43", desc_cel, true, false, 1, 9);
            //Otros descuentos 
            crearColumna(range, "E44", "F44",otrosDcts , true, false, 1, 9);
            //Total descuentos
            crearColumna(range, "E45", "F45", tot_desc, true, false, 1, 9);
        
 

            //Valor finiquito
            crearColumna(range, "C46", "D46", "Valor Finiquito", true, true, 1, 9);

            //Relleno finiquito  
            crearColumna(range, "E46", "F46", valor_finiquito, true, false, 1, 9);


            // ------------------------------------------------------------------------------------
            // Insercion de imagen firma
            // Rango donde se intertará la imagen
            Range Range088 = xlWorkSheet.get_Range("A50", "J52");
            Range088.Select(); //rango para poder insertar la imagen

            // objeto de las imagenes en excel
            Microsoft.Office.Interop.Excel.Pictures oPictures =
            (Microsoft.Office.Interop.Excel.Pictures)xlWorkSheet.Pictures(System.Reflection.Missing.Value);

            //String url = (AppDomain.CurrentDomain.BaseDirectory + "pie.jpg").ToString();
            //String url2 = url.Replace("\\", "/");

          //  xlWorkSheet.Shapes.AddPicture(@"C:\Users\juan.bustos\Desktop\pie.jpg",
            //Microsoft.Office.Core.MsoTriState.msoFalse,
            //Microsoft.Office.Core.MsoTriState.msoCTrue,
            //float.Parse(Range088.Left.ToString()), float.Parse(Range088.Top.ToString()),
            //float.Parse(Range088.Width.ToString()), float.Parse(Range088.Height.ToString()));
 
            System.Drawing.Bitmap bmp = new Bitmap(Finiquitos.Properties.Resources.pie);
            //System.Drawing.Graphics g = Graphics.FromImage(bmp);

 
 
            //g.DrawLine(new System.Drawing.Pen(System.Drawing.Color.Blue), 1, 1, 100, 100);

      
            System.Windows.Forms.Clipboard.SetDataObject(bmp, true);

 
            xlWorkSheet.Paste(Range088, bmp);

            // ------------------------------------------------------------------------------------
            //ELABORO
            crearColumna(range, "A55", "D55", "ELABORO", true, true, 1, 9);
            //APROBO
            crearColumna(range, "E55", "J55", "APROBO", true, true, 1, 9);
            //JEFE DE PERSONAS
            crearColumna(range, "A56", "D56", "JEFE DE PERSONAS", true, true, 1, 9);
            //GERENCIA ADMINISTRACION Y FINANZAS
            crearColumna(range, "E56", "J56", "GERENCIA ADMINISTRACION Y FINANZAS", true, true, 1, 9);

            // ------------------------------------------------------------------------------------

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

            }
            finally
            {
                GC.Collect();
            }
        }

        public void guardar(SaveFileDialog saveFileDialog2)
        {
            try
            {
                xlWorkBook.SaveAs(saveFileDialog2.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                                                                              Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        public void crearColumna(Range rango, String celdaA, String celdaB, String valor,
                             bool combinar, bool centrado, int bold, int size)
        {

            if (combinar == true && centrado == true)
            {
                rango = xlWorkSheet.get_Range(celdaA, celdaB);
                rango.Merge();
                rango.Font.Bold = bold;
                rango.Borders.Color = Color.Black;
                rango.Font.Size = size;
                rango.Font.Name = "Century Gothic";
                rango.Value = valor;
                rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                rango.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            }

            if (combinar == true && centrado == false)
            {
                rango = xlWorkSheet.get_Range(celdaA, celdaB);
                rango.Merge();
                rango.Font.Bold = bold;
                rango.Borders.Color = Color.Black;
                rango.Font.Size = size;
                rango.Font.Name = "Century Gothic";
                rango.Value = valor;
            }

            if (combinar == false || centrado == false)
            {
                rango = xlWorkSheet.get_Range(celdaA, celdaB);

                rango.Font.Bold = bold;
                rango.Borders.Color = Color.Black;
                rango.Font.Size = size;
                rango.Font.Name = "Century Gothic";
                rango.Value = valor;
            }



        }
    }
}
