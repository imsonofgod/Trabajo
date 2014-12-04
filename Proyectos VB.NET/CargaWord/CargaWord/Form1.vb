
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Word
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO


'Autor : <Juan Pablo Bustos Saez>
'Fecha : <05-08-2014>
'Hora  : <15:57>
'Nombre: Carga de Archivos Word
'Descripcion: Software para carga de archivos word.
'Requisitos: Tener corriendo windows 2007 y una carpeta raiz como 'z:\' 
'para que pueda recorrer sus subcarpetas




Public Class Form1
    Public Ruta As String
    Public Bandera As Boolean
    Public GLobalName As String
    Public fe_cambio As DateTime
    Public co_est_aviso As String
    ' permita saber  cuando el archivo esta dañado
    Public estado As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

       
 
        Shell("net use  \\ALANO\CommonVideo /u:doclgg2014 cdsdoclgg2014 /Y")

 
        ' temrinado sin ningun problema
        For i = 0 To 10000
            'time
        Next

        listbox2.Items.Add("Cargando Archivos ....")
        lstBox.Items.Add("Cargando Archivos ....")

        Me.Show()

        Using conn As New SqlConnection("Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;User ID=jbustos;Password=jbustos2014")
            conn.Open()
            Dim consulta As String = " if exists (select * from tempdb.dbo.sysobjects where name='##avisos_a_procesar_hoy_dox2txt' and type='U') "&
"                                      begin   " &
  "                                    	drop table ##avisos_a_procesar_hoy_dox2txt   " &
   "                                   end  " &
 " select    " &
" 	av.co_aviso " &
" 	,'\\alano\CommonVideo\'+av.co_aviso+'\' as Ruta    " &
" 	,av.fe_cambio " &
" 	,av.co_est_aviso " &
 " into " &
   " 	##avisos_a_procesar_hoy_dox2txt  " &
 " from    " &
 "      	sol_item_guia sig   " &
   " 		,sol_item si   " &
    " 	,aviso av   " &
    " where   " &
    " sig.nu_solicitud=si.nu_solicitud   " &
    " 	and sig.nu_item=si.nu_item   " &
    " 	and si.co_item in (   " &
    " 							'VBASI'   " &
    " 							,'VCORP'   " &
    " 							,'VFULL'   " &
    " 								)   " &
    " 	and av.co_aviso=sig.co_aviso   " &
    " 	and av.co_est_aviso='AEN'    " &
    " 	and av.fe_cambio <getdate()   " &
    " 	and av.co_aviso<>'00000000'   " &
    " 	and av.co_aviso not in   " &
    " 					(select  " &
    " 							tp.co_aviso   " &
    " 					from  " &
    " 							texto_aviso tp  " &
    " 					where " &
    " 						tp.co_aviso=av.co_aviso " &
      "                  and tp.fe_cambio <> av.fe_cambio " &
        "                and av.co_est_aviso = 'AEN' " &
        " 				) " &
" UNION " &
 " select " &
 " av.co_aviso    " &
 " ,'\\alano\CommonVideo\'+av.co_aviso+'\' as Ruta   " &
" ,av.fe_cambio " &
" ,av.co_est_aviso " &
" from    " &
 "  	sol_item_guia sig   " &
" 	,sol_item si   " &
" 	,aviso av   " &
" where   " &
" sig.nu_solicitud=si.nu_solicitud   " &
" 	and sig.nu_item=si.nu_item   " &
" 	and si.co_item in (   " &
" 							'VBASI'   " &
" 							,'VCORP'   " &
" 							,'VFULL'   " &
" 								)   " &
" 	and av.co_aviso=sig.co_aviso   " &
" 	and av.co_est_aviso='AEN'    " &
" 	and av.fe_cambio <getdate()   " &
" 	and av.co_aviso<>'00000000'   " &
" 	and av.co_aviso in   " &
" 					(select " &
" 							tp.co_aviso   " &
" 					from  " &
" 							texto_aviso tp  " &
" 					where " &
" 						tp.co_aviso=av.co_aviso " &
" 						and tp.flag in (1,2,3)" &
" 					) " &
" select * from ##avisos_a_procesar_hoy_dox2txt "





            Dim cmd As New SqlCommand(consulta, conn)
            cmd.CommandType = CommandType.Text

            Dim da As New SqlDataAdapter(cmd)

            Dim dt As New System.Data.DataTable


            da.Fill(dt)


            For i = 0 To dt.Rows.Count - 1
                estado = False
                Bandera = False

                Dim Ruta As String = dt.Rows(i).Item("Ruta").ToString()
                'Dim Ruta As String = "Z:\00271821"
                GLobalName = dt.Rows(i).Item("co_aviso").ToString()

                fe_cambio = dt.Rows(i).Item("fe_cambio").ToString()
                'MessageBox.Show(dt.Rows(i).Item("co_est_aviso").ToString())
                co_est_aviso = dt.Rows(i).Item("co_est_aviso").ToString()
                If Not GLobalName.Equals("") Then



                    'Con Recursividad en DOCX


                    If Directory.Exists(Ruta) Then
                        'If isRecursive(Ruta) = True ThenEnd If
                        Recursive2(Ruta, fe_cambio, co_est_aviso)



                    Else
                        log(GLobalName, Ruta, 1)
                        'MsgBox("el Directorio no existe")
                        listbox2.Items.Add(Ruta + "Carpeta No Encontrada")
                    End If

                    Label4.Text = i
                    GLobalName = ""
                    Ruta = ""
                End If
            Next

        End Using


      

        listbox2.Items.Add("Proceso Terminado ....")
        lstBox.Items.Add("Proceso Terminado ....")

        Me.Close()


        Shell("net use \\ALANO\CommonVideo /delete /Y", vbHidden)

    End Sub


 

    Private Function MatarProceso(ByVal StrNombreProceso As String, _
    Optional ByVal DecirSINO As Boolean = True) As Boolean
        ' Variables para usar Wmi  
        'For i = 0 To 5000
        '    'time
        'Next

        Dim ListaProcesos As Object
        Dim ObjetoWMI As Object
        Dim ProcesoACerrar As Object

        MatarProceso = False

        ObjetoWMI = GetObject("winmgmts:")

        If ObjetoWMI Is DBNull.Value = False Then

            'instanciamos la variable  
            ListaProcesos = ObjetoWMI.InstancesOf("win32_process")

            For Each ProcesoACerrar In ListaProcesos
                If UCase(ProcesoACerrar.Name) = UCase(StrNombreProceso) Then
                    If DecirSINO Then
                        ProcesoACerrar.Terminate(0)
                        MatarProceso = True
                    End If
                End If

            Next
        End If

        'Elimina las variables  
        ListaProcesos = Nothing
        ObjetoWMI = Nothing
    End Function


    Public Sub insercion(ByVal nombre As String, ByVal ruta As String, ByVal fe_cambio As DateTime, ByVal co_est_aviso As String)
        Dim conn As New SqlClient.SqlConnection
        Dim texto As String
        conn.ConnectionString = "Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;User ID=jbustos;Password=jbustos2014"
        Try

            Dim cmd As New System.Data.SqlClient.SqlCommand
            cmd.CommandType = System.Data.CommandType.Text

            'verifico que el archivo que voy a ingresar no sea una copia del sistema

            If Trim(Mid(nombre, 1, 2)) <> "~$" Then


                Dim doc As New Word.Document
                ' Dim wordApp As Word.Application = New Application

                Dim wordApp As Word.Application = New Application

                Dim file As Object = ruta '*
                Dim Nothingobj As Object = System.Reflection.Missing.Value

                doc = wordApp.Documents.Open(file, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj, Nothingobj)
                doc.ActiveWindow.Selection.WholeStory()
                doc.ActiveWindow.Selection.Copy()

                Dim data As IDataObject = Clipboard.GetDataObject()
                doc.Close()



                texto = data.GetData(DataFormats.Text).ToString()




                cmd.CommandText = " IF EXISTS (SELECT " &
                                    "	tx_av.co_aviso " &
                                    "  FROM " &
                                    "	texto_aviso tx_av," &
                                    "	aviso av" &
                                    "  WHERE " &
                                    "		tx_av.co_aviso = '" + nombre + "' " &
                                    "  and tx_av.fe_cambio <> av.fe_cambio" &
                                    "  and av.co_est_aviso = 'AEN' " &
                                    "	) " &
                                    "BEGIN " &
                                    "UPDATE texto_aviso " &
                                    "SET " &
                                    "	tx_aviso = '" + texto + "' " &
                                    "	,flag = '10' " &
                                     "	,fe_cambio = '" + Format(fe_cambio, "yyyy/MM/dd") + "' " &
                                     "	,co_est_aviso = '" + co_est_aviso + "' " &
                                     "	where " &
                                      "	co_aviso = '" + nombre + "' " &
                                    "END " &
                                    "ELSE " &
                                    "INSERT INTO texto_aviso (co_aviso,tx_aviso,fe_modificacion,fe_cambio,co_est_aviso) values('" + nombre + "','" + texto + "',getdate(),'" + Format(fe_cambio, "yyyy/MM/dd") + "','" + co_est_aviso + "')"

                cmd.Connection = conn

                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()

                MatarProceso("WINWORD.EXE", True)

            End If


        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            listbox2.Items.Add(ruta + "   Archivo  Dañado")
            estado = True
            log(nombre, ruta, 3)
        Finally
            conn.Close()
        End Try


    End Sub

    ' FIN INSERCION 
    '  Recursiva2 ***************************

    Public Sub Recursive2(ByVal sourceDir As String, ByVal fe_cambio As DateTime, ByVal co_est_aviso As String)




        Dim cont As Integer = 1
        Dim exten As String
        Dim ex As String
        If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
            sourceDir &= System.IO.Path.DirectorySeparatorChar
        End If


        Do
            ' cambia el formato del archivo 

            Dim Files As String(), File As String
            Dim Nombre As String


            If cont = 1 Then
                exten = "*.docx"
                ex = ".docx"
            Else
                exten = "*.doc"
                ex = ".doc"
            End If


            'Leyendo los archivos de la carpeta ‘C:\Musica’
            Files = IO.Directory.GetFiles(sourceDir, exten)

            If Bandera <> True Then
                For Each File In Files
                    'Muestra los ficheros leidos, los carga al list box y omite los archivos copia de sistema (~$) 


                    If Trim(Mid(IO.Path.GetFileNameWithoutExtension(File).ToString(), 1, 2)) <> "~$" Or Trim(Mid(IO.Path.GetFileNameWithoutExtension(File).ToString(), 1, 2)) <> "" Then

                        Nombre = IO.Path.GetFileNameWithoutExtension(File)

                        If Trim(Mid(Nombre, 1, 2)) <> "~$" Then

                            Ruta = sourceDir + Nombre + ex

                            'MessageBox.Show(Ruta)
                            insercion(GLobalName, Ruta, fe_cambio, co_est_aviso)

                            If Trim(Mid(IO.Path.GetFileNameWithoutExtension(File).ToString(), 1, 2)) <> "~$" Then
                                lstBox.Items.Add(Ruta)
                                Bandera = True
                            End If

                        End If

                    End If
                Next
            End If

            cont = cont + 1
        Loop While (cont <> 3)

        ' Comienza URL 2

        Dim Url2 As String = sourceDir + Mid(sourceDir, 4, 11)
        Dim ex2 As String
        Dim cont2 As Integer = 1
        If Directory.Exists(Url2) Then

            Do
                Dim Files2 As String(), File2 As String
                Dim Nombre As String
                ' cambia el formato del archivo 


                If cont2 = 1 Then
                    exten = "*.docx"
                    ex2 = ".docx"
                Else
                    exten = "*.doc"
                    ex2 = ".doc"
                End If

                'Leyendo los archivos de la carpeta ‘C:\Musica’
                Files2 = IO.Directory.GetFiles(Url2, exten)


                If Bandera <> True Then
                    For Each File2 In Files2
                        'Muestra los ficheros leidos, los carga al list box y omite los archivos copia de sistema (~$) 


                        If Trim(Mid(IO.Path.GetFileNameWithoutExtension(File2).ToString(), 1, 2)) <> "~$" Or Trim(Mid(IO.Path.GetFileNameWithoutExtension(File2).ToString(), 1, 2)) <> "" Then

                            Nombre = IO.Path.GetFileNameWithoutExtension(File2)

                            Ruta = Url2 + Nombre + ex2


                            insercion(Nombre, Ruta, fe_cambio, co_est_aviso)

                            If Trim(Mid(IO.Path.GetFileNameWithoutExtension(File2).ToString(), 1, 2)) <> "~$" Then
                                lstBox.Items.Add(Ruta)
                                Bandera = True
                            End If

                        End If
                    Next
                End If
                cont2 = cont2 + 1
            Loop While (cont2 <> 3)

        End If

        My.Application.DoEvents() 'Deja al sistema hacer otras cosas






        If Bandera = False And estado = False Then
            log(GLobalName, sourceDir, 2)
            listbox2.Items.Add(sourceDir + "El Archivo No Existe")

        End If



    End Sub

    '  End Recursiva2 ***************************
    Public Sub log(ByVal nombre As String, ByVal ruta As String, ByVal f As Integer)


        Dim conn As New SqlClient.SqlConnection

        conn.ConnectionString = "Data Source=10.1.3.237;Initial Catalog=lgg_xdirect;User ID=jbustos;Password=jbustos2014"

        Dim cmd As New System.Data.SqlClient.SqlCommand
        cmd.CommandType = System.Data.CommandType.Text




        If f = 1 Then

            cmd.CommandText = " IF EXISTS (SELECT " &
                        "	tx_av.co_aviso " &
                        "  FROM " &
                        "	texto_aviso tx_av," &
                        "	aviso av" &
                        "  WHERE " &
                        "		tx_av.co_aviso = '" + nombre + "' " &
                        "	) " &
                        "BEGIN " &
                        "UPDATE texto_aviso " &
                        "SET " &
                        "	tx_aviso = '" + ruta + "', " &
                        "	fe_modificacion = getdate(), " &
                        "	flag =  '1', " &
                        "	fe_cambio = '" + Format(fe_cambio, "yyyy/MM/dd") + "', " &
                        "	co_est_aviso = '" + co_est_aviso + "' " &
                        "	where co_aviso = '" + nombre + "' " &
                        "END " &
                        "ELSE " & "INSERT INTO texto_aviso (co_aviso,tx_aviso,fe_modificacion,flag,fe_cambio,co_est_aviso) values('" + nombre + "','" + ruta + "',getdate(),'1','" + Format(fe_cambio, "yyyy/MM/dd") + "','" + co_est_aviso + "')"


 

            cmd.Connection = conn

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        If f = 2 Then
            Try

                cmd.CommandText = " IF EXISTS (SELECT " &
                          "	tx_av.co_aviso " &
                          "  FROM " &
                          "	texto_aviso tx_av," &
                          "	aviso av" &
                          "  WHERE " &
                          "		tx_av.co_aviso = '" + nombre + "' " &
                          "	) " &
                          "BEGIN " &
                          "UPDATE texto_aviso " &
                          "SET " &
                          "	tx_aviso = '" + ruta + "', " &
                          "	fe_modificacion = getdate(), " &
                          "	flag =  '2', " &
                          "	fe_cambio = '" + Format(fe_cambio, "yyyy/MM/dd") + "', " &
                          "	co_est_aviso = '" + co_est_aviso + "'" &
                          "	where co_aviso = '" + nombre + "' " &
                          "END " &
                          "ELSE " & "INSERT INTO texto_aviso (co_aviso,tx_aviso,fe_modificacion,flag,fe_cambio,co_est_aviso) values('" + nombre + "','" + ruta + "',getdate(),'2','" + Format(fe_cambio, "yyyy/MM/dd") + "','" + co_est_aviso + "')"



                cmd.Connection = conn

                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If


        If f = 3 Then

            cmd.CommandText = " IF EXISTS (SELECT " &
                                "	tx_av.co_aviso " &
                                "  FROM " &
                                "	texto_aviso tx_av," &
                                "	aviso av" &
                                "  WHERE " &
                                "		tx_av.co_aviso = '" + nombre + "' " &
                                "	) " &
                                "BEGIN " &
                                "UPDATE texto_aviso " &
                                "SET " &
                                "	tx_aviso = '" + ruta + "', " &
                                "	fe_modificacion = getdate(), " &
                                "	flag =  '3', " &
                                "	fe_cambio = '" + Format(fe_cambio, "yyyy/MM/dd") + "', " &
                                "	co_est_aviso = '" + co_est_aviso + "' " &
                                "	where co_aviso = '" + nombre + "' " &
                                "END " &
                                "ELSE " & "INSERT INTO texto_aviso (co_aviso,tx_aviso,fe_modificacion,flag,fe_cambio,co_est_aviso) values('" + nombre + "','" + ruta + "',getdate(),'3','" + Format(fe_cambio, "yyyy/MM/dd") + "','" + co_est_aviso + "')"






            cmd.Connection = conn

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()


        End If



    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click




        Me.Close()
        MatarProceso("CargaWord.exe", True)

    End Sub
End Class
