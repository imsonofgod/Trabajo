using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Finiquitos
{
    class CartaWord
    {
        public Microsoft.Office.Interop.Word.Application objWordApplication;
        public Microsoft.Office.Interop.Word.Document objWordDocument;
        public Object oMissing = System.Reflection.Missing.Value;



public CartaWord()
{
    objWordApplication = new Microsoft.Office.Interop.Word.Application();
    objWordDocument = objWordApplication.Documents.Add(ref oMissing);
}

public void Config()
{
    objWordDocument.Activate();
    objWordApplication.ActiveDocument.PageSetup.TopMargin = (float)(28.35);
    objWordApplication.Selection.Font.Size = 10;
    objWordApplication.Selection.Font.Name = "Times New Roman";
    objWordApplication.Visible = false;
    objWordApplication.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
    objWordApplication.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = 0F;

}

public void Config2()
{

    objWordApplication.Selection.Font.Size = 11;
    objWordApplication.Selection.Font.Name = "Times New Roman";
    objWordApplication.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;

    //Microsoft.Office.Interop.Word.Selection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle;
    //Microsoft.Office.Interop.Wordndow.Selection.ParagraphFormat.SpaceAfter = 0.0F;
    objWordApplication.Selection.Paragraphs.SpaceAfter = 0;

}



public void guardar(SaveFileDialog saveFileDialog2)
{
        try { 
        objWordDocument.SaveAs(saveFileDialog2.FileName, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);

        object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdPromptToSaveChanges;
        object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
        object routeDocument = true;

        (objWordDocument).Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
}

// ************************************** Cartas Con Distinto Formato
public void texto_carta1(String nombre, String ape_pat,String rut, String direccion, 
                         String Comuna,String ciudad, String cargo, String fecha_finiquito,
                         String art, String indicador, String desc_art, String indemnizacion_valor,
                         String fecha_ingreso)
{

    Microsoft.Office.Interop.Word.Paragraph para1 = objWordDocument.Content.Paragraphs.Add();
    para1.Range.Font.Bold = 1;
    para1.Range.Text = "\t\t\t\t Santiago,  "+fecha_finiquito+".-\n";
    objWordApplication.Selection.InlineShapes.AddPicture(@"C:\\Users\\juan.bustos\\Documents\\Visual Studio 2012\\Projects\\Cartas Finiquitos\\Images\\amarillas2.png");


    Microsoft.Office.Interop.Word.Paragraph encabezado = objWordDocument.Content.Paragraphs.Add();
    encabezado.Range.Font.Bold = 1;
    encabezado.Range.Text = "\nSeñor(a)\n" +
                                "" + nombre + "\n" +
                                "RUT " + rut + "\n" +
                                "" + direccion + "\n";

    Microsoft.Office.Interop.Word.Paragraph para3 = objWordDocument.Content.Paragraphs.Add();
    para3.Range.Font.Bold = 1;
    para3.Range.Text = "Renca / Santiago\nPresente\n\n";

    Microsoft.Office.Interop.Word.Paragraph para4 = objWordDocument.Content.Paragraphs.Add();
    para4.Range.Font.Bold = 1;
    para4.Range.Text = "\t\t    REF.: NOTIFICA TÉRMINO DE CONTRATO DE TRABAJO\n\n";

    Microsoft.Office.Interop.Word.Paragraph para5 = objWordDocument.Content.Paragraphs.Add();
    para5.Range.Font.Bold = 1;
    para5.Range.Text = "Estimado Sr(a). "+ape_pat+":\n\n";

    Microsoft.Office.Interop.Word.Paragraph para6 = objWordDocument.Content.Paragraphs.Add();
    para6.Range.Font.Bold = 0;
    para6.Range.Text = "La presente carta tiene por objeto notificarle que se ha resuelto poner término a su contrato de trabajo y en " +
    "consecuencia, a sus funciones de "+cargo+", a contar del día de hoy, esto es "+fecha_finiquito+".\n\n" +
   "Esta determinación ha sido adoptada conforme a lo dispuesto en el artículo " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“," +
    " lo anterior amparado en la condiciones derivadas de la racionalización producto de" +
    " la reestructuración de la La Gran Guía S.A., (la “Empresa”), que se hizo necesaria por los bajos ingresos que" +
    " ha tenido la Empresa en los últimos 2 años.\n\n" +
    "En razón a lo anterior, y atendido el hecho que nuestra Empresa no ha dado el aviso que exige la ley con 30" +
    " días de anticipación, se le pagará una indemnización equivalente a su última remuneración mensual" +
    " devengada, con el tope legal establecido por la ley.\n\n" +
    "Del mismo modo, le informamos que como su contrato de trabajo ha estado vigente desde el "+fecha_ingreso+" hasta el "+fecha_finiquito+","+
    " por considerar que su tiempo es de menos de 1 año, no" +
    " corresponde cancelar indemnización por años de servicios.- \n\n" +
    "Recuerdo a usted que sus cotizaciones previsionales hasta el mes anterior al despido se encuentran al día y" +
    " debidamente canceladas en las entidades correspondiente. Conforme a lo dispuesto en el inc. 5 del art 162 del" +
    " Código de Trabajo se adjuntan los comprobantes de pago de las mismas.\n\n" +
    "Finalmente, le indicamos que tanto el finiquito para saldar, ajustar o cancelar las cuentas que derivan del" +
    " contrato de trabajo como el cheque mediante el cual se le pagaran las indemnizaciones y demás prestaciones" +
    " laborales a que tiene derecho, se encontrara a su disposición a contar del día 24 de Enero de 2014 en la" +
    " Notaría de don Humberto Quezada, ubicada en calle Huérfanos 835 Piso 2, Santiago.\n\n" +
    "Atentamente,\n";

    Microsoft.Office.Interop.Word.Paragraph para7 = objWordDocument.Content.Paragraphs.Add();
    para7.Range.Font.Bold = 1;
    para7.Range.Text = "\t\t\t\t\t   Carlos Zazo González\n" +
                        "\t\t\t\t\t     La Gran Guía S.A.\n" +
                        "\t\t\t\t\t     RUT 99.538.470-1\n\n";

    Microsoft.Office.Interop.Word.Paragraph para8 = objWordDocument.Content.Paragraphs.Add();
    para8.Range.Font.Bold = 0;
    para8.Range.Text = "Recibí Copia Nombre: _________________________________\n\n" +
                        "\t\t  Rut: ________________________________\n\n" +
                        "c.c.: Inspección del Trabajo\n" +
                        "Carpeta Personal";
}


// sin indemnización
public void texto_carta2(String nombre,String ape_pat, String rut, String direccion, 
                         String Comuna,String ciudad, String cargo, String fecha_finiquito, 
                         String art,String indicador,String desc_art, String indemnizacion_valor,
                         String fecha_ingreso)
{

    Microsoft.Office.Interop.Word.Paragraph para1 = objWordDocument.Content.Paragraphs.Add();
    para1.Range.Font.Bold = 1;
    para1.Range.Text = "\t\t\t\t   "+fecha_finiquito+".-\n";
    objWordApplication.Selection.InlineShapes.AddPicture(@"C:\\Users\\juan.bustos\\Documents\\Visual Studio 2012\\Projects\\Cartas Finiquitos\\Images\\amarillas.png");


    Microsoft.Office.Interop.Word.Paragraph encabezado = objWordDocument.Content.Paragraphs.Add();
    encabezado.Range.Font.Bold = 1;
    encabezado.Range.Text = "\nSeñor(a)\n" +
                                ""+nombre+"\n" +
                                "RUT "+rut+"\n" +
                                ""+direccion+"\n";

    Microsoft.Office.Interop.Word.Paragraph para3 = objWordDocument.Content.Paragraphs.Add();
    para3.Range.Font.Bold = 1;
    para3.Range.Text = ""+ciudad+"\nPresente\n\n";

    Microsoft.Office.Interop.Word.Paragraph para4 = objWordDocument.Content.Paragraphs.Add();
    para4.Range.Font.Bold = 1;
    para4.Range.Text = "\t\t    REF.: NOTIFICA TÉRMINO DE CONTRATO DE TRABAJO\n\n";

    Microsoft.Office.Interop.Word.Paragraph para5 = objWordDocument.Content.Paragraphs.Add();
    para5.Range.Font.Bold = 1;
    para5.Range.Text = "Estimado Sr(a). "+ape_pat+":\n\n";

    Microsoft.Office.Interop.Word.Paragraph para6 = objWordDocument.Content.Paragraphs.Add();
    para6.Range.Font.Bold = 0;
    para6.Range.Text = "La presente carta tiene por objeto notificarle que se ha resuelto poner término a su contrato de trabajo y en " +
    "consecuencia, a sus funciones de EJECUTIVO VENTAS, a contar del día de hoy, esto es "+fecha_finiquito+".\n\n" +
    "Esta determinación ha sido adoptada conforme a lo dispuesto en el artículo " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“,"+
    " lo anterior amparado en la condiciones derivadas de la racionalización producto de" +
    " la reestructuración de la La Gran Guía S.A., (la “Empresa”), que se hizo necesaria por los bajos ingresos que" +
    " ha tenido la Empresa en los últimos 2 años.\n\n" +
    "En razón a lo anterior, y atendido el hecho que nuestra Empresa no ha dado el aviso que exige la ley con 30" +
    " días de anticipación, se le pagará una indemnización equivalente a su última remuneración mensual" +
    " devengada, con el tope legal establecido por la ley.\n\n" +
    "Del mismo modo, le informamos que como su contrato de trabajo ha estado vigente desde el "+fecha_ingreso+" hasta el "+fecha_finiquito+","+
    " por ello se le cancelara, junto con el resto de las prestaciones y" +
    " los descuentos legales que correspondan, una indemnización por años de servicios ascendente a $"+indemnizacion_valor+".- \n\n" +
    "Recuerdo a usted que sus cotizaciones previsionales hasta el mes anterior al despido se encuentran al día y" +
    " debidamente canceladas en las entidades correspondiente. Conforme a lo dispuesto en el inc. 5 del art 162 del" +
    " Código de Trabajo se adjuntan los comprobantes de pago de las mismas.\n\n" +
    "Finalmente, le indicamos que tanto el finiquito para saldar, ajustar o cancelar las cuentas que derivan del" +
    " contrato de trabajo como el cheque mediante el cual se le pagaran las indemnizaciones y demás prestaciones" +
    " laborales a que tiene derecho, se encontrara a su disposición a contar del día 19 de Diciembre de 2013 en la" +
    " Notaría de don Humberto Quezada, ubicada en calle Huérfanos 835 Piso 2, Santiago.\n\n" +
    "Atentamente,\n";

    Microsoft.Office.Interop.Word.Paragraph para7 = objWordDocument.Content.Paragraphs.Add();
    para7.Range.Font.Bold = 1;
    para7.Range.Text = "\t\t\t\t\t   Carlos Zazo González\n" +
                        "\t\t\t\t\t     La Gran Guía S.A.\n" +
                        "\t\t\t\t\t     RUT 99.538.470-1\n\n";

    Microsoft.Office.Interop.Word.Paragraph para8 = objWordDocument.Content.Paragraphs.Add();
    para8.Range.Font.Bold = 0;
    para8.Range.Text = "Recibí Copia Nombre: _________________________________\n\n" +
                        "\t\t  Rut: ________________________________\n\n" +
                        "c.c.: Inspección del Trabajo\n" +
                        "Carpeta Personal";
}

// ************************************** Cartas

private void SearchReplace(String text)
{
    Microsoft.Office.Interop.Word.Find findObject = objWordApplication.Selection.Find;
    findObject.ClearFormatting();
    findObject.Text = text;
 


          
}


public void texto_carta3(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp, 
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras)
    {
        
    String palabra =
                        "F I N I Q U I T O\n\n\n" +
                        "En Santiago a " + fecha_in + "</color>, entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                        " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                        " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                        " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                        " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                        " se deja testimonio y se conviene lo siguiente:\n\n" +

                        "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                        " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                        " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                        " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“,"+
                        " que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                        " trabajadores. \n\n" +

                        "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                        " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                        "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                        "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                        "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                        "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                        " Descuento:\n\n" +
                        "4.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                        "5.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                        "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                        "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                        "Son: " + monto_palabras + " PESOS.-\n\n" +

                        "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                        " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                        " que la conforman.\n\n\n" +

                        "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                        " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                        " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                        " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                        " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                        " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                        " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                        " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                        " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                        " enteradas dentro del plazo legal. \n\n" +

                        "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                        " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                        " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                        " su contrato de trabajo.\n\n" +

                        "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                        " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                        " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                        " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                        " total y definitivo finiquito.\n\n" +

                        "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                        " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                        " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                        " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                        " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                        " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                        " compensaciones de cualquier orden o naturaleza.\n" +
                        "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                        " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                        "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                        "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                        "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";

                        objWordApplication.Selection.TypeText(palabra);
                       // objWordDocument.Range(1, 1).Bold = 1;
                       

    }
         

// si hay descuentos sin descripcion

public void texto_carta4(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp,
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras)
    {

    String palabra =
                    "F I N I Q U I T O\n\n\n" +
                    "En Santiago a " + fecha_in + ", entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                    " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                    " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                    " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                    " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                    " se deja testimonio y se conviene lo siguiente:\n\n" +

                    "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                    " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                    " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                    " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“, que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                    " trabajadores. \n\n" +

                    "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                    " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                    "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                    "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                    "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                    "4.- Otros Haberes   \t\t\t\t\t\t$   " + otros_haberes + "     \n" +
                    "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                    " Descuento:\n\n" +
                    "5.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                    "6.- Otros Descuentos\t\t\t\t\t\t$ ( " + otros_dcts + " .- )\n" +
                    "7.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                    "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                    "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                    "Son: " + monto_palabras + " PESOS.-\n\n" +

                    "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                    " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                    " que la conforman.\n\n\n" +

                        "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                        " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                        " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                        " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                        " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                        " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                        " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                        " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                        " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                        " enteradas dentro del plazo legal. \n\n" +

                    "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                    " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                    " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                    " su contrato de trabajo.\n\n" +

                    "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                    " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                    " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                    " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                    " total y definitivo finiquito.\n\n" +

                    "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                    " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                    " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                    " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                    " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                    " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                    " compensaciones de cualquier orden o naturaleza.\n" +
                    "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                    " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                    "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                    "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                    "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";
    objWordApplication.Selection.TypeText(palabra);
    }

// si hay descuentos con descripcion

public void texto_carta5(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp,
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras,
    String rchOtrosHaberes, String rchOtrosDcts)
    {
    String palabra =
                "F I N I Q U I T O\n\n\n" +
                "En Santiago a " + fecha_in + ", entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                " se deja testimonio y se conviene lo siguiente:\n\n" +

                "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“, que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                " trabajadores. \n\n" +

                "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                "4.- " + rchOtrosHaberes + "   \t\t\t\t\t\t$   " + otros_haberes + "     \n" +
                "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                " Descuento:\n\n" +
                "5.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                "6.- " + rchOtrosDcts + "\t\t\t\t\t\t$ ( " + otros_dcts + " .- )\n" +
                "7.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                "Son: " + monto_palabras + " PESOS.-\n\n" +

                "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                " que la conforman.\n\n\n" +

                "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                " enteradas dentro del plazo legal. \n\n" +

                "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                " su contrato de trabajo.\n\n" +

                "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                " total y definitivo finiquito.\n\n" +

                "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                " compensaciones de cualquier orden o naturaleza.\n" +
                "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";
    objWordApplication.Selection.TypeText(palabra);
    }


// Con Otros Descuentos sin desc
public void texto_carta6(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp,
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras,
    String rchOtrosHaberes, String rchOtrosDcts)
        {

            String palabra =
                            "F I N I Q U I T O\n\n\n" +
                            "En Santiago a " + fecha_in + ", entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                            " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                            " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                            " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                            " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                            " se deja testimonio y se conviene lo siguiente:\n\n" +

                            "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                            " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                            " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                            " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“, que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                            " trabajadores. \n\n" +

                            "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                            " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                            "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                            "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                            "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                            "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                            " Descuento:\n\n" +
                            "4.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                            "5.- Otros Descuentos\t\t\t\t\t\t$ ( " + otros_dcts + " .- )\n" +
                            "6.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                            "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                            "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                            "Son: " + monto_palabras + " PESOS.-\n\n" +

                            "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                            " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                            " que la conforman.\n\n\n" +

                                "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                                " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                                " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                                " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                                " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                                " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                                " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                                " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                                " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                                " enteradas dentro del plazo legal. \n\n" +

                            "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                            " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                            " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                            " su contrato de trabajo.\n\n" +

                            "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                            " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                            " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                            " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                            " total y definitivo finiquito.\n\n" +

                            "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                            " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                            " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                            " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                            " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                            " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                            " compensaciones de cualquier orden o naturaleza.\n" +
                            "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                            " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                            "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                            "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                            "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";
            objWordApplication.Selection.TypeText(palabra);
        }

// Con Otros Descuentos con desc
public void texto_carta7(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp,
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras,
    String rchOtrosHaberes, String rchOtrosDcts)
    {
    String palabra =
                            "F I N I Q U I T O\n\n\n" +
                            "En Santiago a " + fecha_in + ", entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                            " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                            " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                            " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                            " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                            " se deja testimonio y se conviene lo siguiente:\n\n" +

                            "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                            " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                        " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                            " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“, que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                            " trabajadores. \n\n" +

                            "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                            " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                            "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                            "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                            "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                            "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                            " Descuento:\n\n" +
                            "4.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                            "5.- " + rchOtrosDcts + "\t\t\t\t\t\t$ ( " + otros_dcts + " .- )\n" +
                            "6.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                            "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                            "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                            "Son: " + monto_palabras + " PESOS.-\n\n" +

                            "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                            " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                            " que la conforman.\n\n\n" +

                            "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                            " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                            " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                            " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                            " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                            " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                            " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                            " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                            " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                            " enteradas dentro del plazo legal. \n\n" +

                            "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                            " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                            " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                            " su contrato de trabajo.\n\n" +

                            "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                            " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                            " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                            " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                            " total y definitivo finiquito.\n\n" +

                            "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                            " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                            " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                            " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                            " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                            " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                            " compensaciones de cualquier orden o naturaleza.\n" +
                            "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                            " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                            "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                            "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                            "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";
    objWordApplication.Selection.TypeText(palabra);
        }


// Con Haberes  sin desc
public void texto_carta8(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp,
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras,
    String rchOtrosHaberes, String rchOtrosDcts) {

    String palabra =
                    "F I N I Q U I T O\n\n\n" +
                    "En Santiago a " + fecha_in + ", entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                    " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                    " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                    " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                    " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                    " se deja testimonio y se conviene lo siguiente:\n\n" +

                    "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                    " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                    " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                    " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“, que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                    " trabajadores. \n\n" +

                    "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                    " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                    "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                    "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                    "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                    "4.- Otros Haberes   \t\t\t\t\t\t$   " + otros_haberes + "     \n" +
                    "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                    " Descuento:\n\n" +
                    "5.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                    "6.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                    "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                    "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                    "Son: " + monto_palabras + " PESOS.-\n\n" +

                    "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                    " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                    " que la conforman.\n\n\n" +

                    "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                    " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                    " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                    " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                    " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                    " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                    " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                    " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                    " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                    " enteradas dentro del plazo legal. \n\n" +

                    "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                    " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                    " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                    " su contrato de trabajo.\n\n" +

                    "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                    " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                    " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                    " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                    " total y definitivo finiquito.\n\n" +

                    "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                    " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                    " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                    " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                    " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                    " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                    " compensaciones de cualquier orden o naturaleza.\n" +
                    "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                    " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                    "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                    "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                    "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";
    objWordApplication.Selection.TypeText(palabra);
            
}


// Con Haberes  con desc
public void texto_carta9(
    String fecha_in, String fecha_fin, String nombre_emp, String rut_emp,
    String domicilio_emp, String comuna_emp, String ciudad_emp, String agno_serv,
    String indep_x_agno_ser, String indep_sus_x_aviso, String feariado_pro,
    String feariado_pro_dias, String afc, String cta_cte_emp, String dcto,
    String lqdo_pago, String otros_haberes, String otros_dcts, double total,
    String haberes, String art, String indicador, String desc_art, String monto_palabras,
    String rchOtrosHaberes, String rchOtrosDcts) {


        String palabra =
                        "F I N I Q U I T O\n\n\n" +
                        "En Santiago a " + fecha_in + ", entre La Gran Guía S.A., persona jurídica del giro de su denominación," +
                        " Rut N° 99.538.470-1, representada por don CARLOS ZAZO GONZALEZ, cédula nacional de" +
                        " identidad Nº 22.312.429-1, ambos domiciliados calle Los Conquistadores 1700, piso 12, comuna de" +
                        " Providencia cuidad de Santiago, por una parte; y por la otra, don (ña) " + nombre_emp + ", C.I. N°" +
                        " " + rut_emp + ", domiciliado en " + domicilio_emp + ", comuna de " + comuna_emp + ", ciudad de " + ciudad_emp + ", por la otra," +
                        " se deja testimonio y se conviene lo siguiente:\n\n" +

                        "PRIMERO:\tLas partes dejan constancia que don (ña) " + nombre_emp + " prestó servicios a LA" +
                        " GRAN GUÍA S.A., bajo contrato de trabajo desde el día " + fecha_in + " hasta el día " + fecha_fin + ", fecha esta" +
                        " última en que el contrato de trabajo de don " + nombre_emp + " terminó conforme a lo dispuesto en el" +
                        " " + art + " (" + indicador + ") del Código del Trabajo esto es “" + desc_art + "“, que faculta al empleador a terminar el contrato mediante desahucio escrito a los" +
                        " trabajadores. \n\n" +

                        "SEGUNDO:\tCon motivo de la referida terminación de contrato don (ña) " + nombre_emp + " tiene" +
                        " derecho a las prestaciones que a continuación se detallan:\n\n\n" +

                        "Haberes:\n\n1.- Indemnización por años de servicios (" + agno_serv + ")\t\t\t$   " + indep_x_agno_ser + ".-\n" +
                        "2.- Indemnización Sustitutiva del aviso previo     \t\t$   " + indep_sus_x_aviso + ".-\n" +
                        "3.- Feriado Legal y Proporcional (" + feariado_pro_dias + ")                       \t\t$   " + feariado_pro + " -    \n" +
                        "4.- " + rchOtrosHaberes + "   \t\t\t\t\t\t$   " + otros_haberes + "     \n" +
                        "\t\tTotal Haberes\t\t\t\t             $   " + haberes + ".-\n\n" +

                        " Descuento:\n\n" +
                        "5.- Aporte Empleador AFC Seguro de Cesantía\t\t\t$ ( " + afc + " .- )\n" +
                        "6.- Cuenta Cte. Empresa\t\t\t\t\t$ ( " + cta_cte_emp + " .- )\n\n      \n\n" +
                        "Total Descuentos\t\t\t\t\t\t$ ( " + dcto + ".-)\n\n" +
                        "Total  Liquido a Pago\t\t\t\t\t\t$   " + lqdo_pago + ".-\n\n" +

                        "Son: " + monto_palabras + " PESOS.-\n\n" +

                        "Don (ña)  " + nombre_emp + " declara que ha revisado detenidamente la liquidación que antecede y" +
                        " deja testimonio de su total conformidad con dicha liquidación y con todas y cada una de las partidas" +
                        " que la conforman.\n\n\n" +

                        "TERCERO:\tDon (ña) " + nombre_emp + " declara que durante el tiempo que prestó servicios a LA" +
                        " GRAN GUÍA S.A., recibió oportunamente el total de las remuneraciones," +
                        " beneficios y demás  prestaciones estipuladas o que hayan derivado o deriven de" +
                        " disposiciones legales u otras normas obligatorias o de cualquier naturaleza u origen" +
                        " y que, asimismo, LA GRAN GUÍA S.A. le descontó, declaró y enteró oportuna e" +
                        " íntegramente las respectivas cotizaciones previsionales en los organismos" +
                        " pertinentes habiendo recibido, así mismo, certificados de los organismos" +
                        " previsionales de la declaración y pago de todo el período en que prestó servicios," +
                        " con excepción de las cotizaciones correspondientes al mes en curso, las que serán" +
                        " enteradas dentro del plazo legal. \n\n" +

                        "Por lo tanto, don (ña) " + nombre_emp + " deja constancia que LA GRAN GUÍA S.A." +
                        " nada le adeuda por causa o motivo alguno, legal o contractual, o de cualquier otro" +
                        " orden, sea que se relacionen con la prestación de sus servicios o la terminación de" +
                        " su contrato de trabajo.\n\n" +

                        "A mayor abundamiento, don (ña)  " + nombre_emp + " declara expresamente que no" +
                        " tiene cargo ni reclamo alguno, que formular en contra de la empleadora, razón por" +
                        " la cual libre y voluntariamente y con pleno y cabal conocimiento de sus derechos," +
                        " otorga a LA GRAN GUÍA S.A., y a sus representantes, el más amplio, completo," +
                        " total y definitivo finiquito.\n\n" +

                        "Asimismo, declara don (ña)  " + nombre_emp + " en todo caso, y a todo evento," +
                        " renuncia expresamente a cualquier derecho, acción o reclamo que pudiera o pudiere" +
                        " corresponderle en contra LA GRAN GUÍA S.A., en relación directa o indirecta con" +
                        " su contrato de trabajo, servicios prestados, o la terminación del referido contrato o" +
                        " de dichos servicios, sea que esos derechos o acciones correspondan a" +
                        " remuneraciones, imposiciones, subsidios, beneficios, indemnizaciones o" +
                        " compensaciones de cualquier orden o naturaleza.\n" +
                        "Para constancia, firman las partes sin reserva de ninguna especie en dos ejemplares del mismo" +
                        " tenor, quedando uno en poder de cada parte, previa ratificación legal de su firma.\n\n\n\n\n\n\n\n\n" +

                        "     Ex-empleador\t\t\t\t\t\tEx-trabajador\n" +
                        "     LA GRAN GUÍA S.A.\t\t\t\t\t" + nombre_emp + "\n" +
                        "     Rut.: 99.538.470-1\t\t\t\t\t\tRut.: " + rut_emp + "\n";
        objWordApplication.Selection.TypeText(palabra);
            
    }
 }      
}
