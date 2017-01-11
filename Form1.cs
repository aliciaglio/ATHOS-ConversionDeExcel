using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ConversionDeExcel.Controles.EntidadesAuxiliares;
using ConversionDeExcel.Models;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using EYPSCAP;
using MinimumEditDistance;
using System.Diagnostics;

//using Xamarin.Forms;



namespace ConversionDeExcel
{
    public partial class Form1 : Form
    {

        #region <<Variables de clase>>
        List<Consumos> newListaConsumos;
        List<Consumos> listaNoAlmacenados;
        public int cuenta = 0;
        public int maxlista = 0;

        public int[] indiceInicio;
        public int[] indiceInicioAltaSimple;
        public int[] indiceInicioAltaDoble;
        public int[] indiceInicioAltaTriple;
        public int[] indiceInicioAuxPequeño;
        public int[] indiceInicioAuxMediano;
        public int[] indiceInicioAuxGrande;

        string file;
        Microsoft.Office.Interop.Excel.Application xlApp;
        Workbook xlWorkBook = null;
        Sheets sheets = null;
        Worksheet xlWorkSheet = null;
        List<long> numAux;
        List<long> numBase;
        List<long> numNev;
        long max = 0;
        Thread th;
        bool configVacia;
        int noDistribuidos;

        string sizeMed;
        int filaLectura;
        List<string> numeroALetra = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
        List<Armario> newListaArmarios;
        List<DatosCajones> newListaCajones;
        List<Consumos> listRepetidos;
        struct CajonesCuenta
        {
            public string tipo;
            public int num;
        }
        CajonesCuenta[] faltan;
        CajonesCuenta[] sobran;
        #endregion


        public Form1()
        {
         
            InitializeComponent();

            this.Location = new System.Drawing.Point((971 / 2) - (this.Width / 2) + 1, (600 / 2) - (this.Height / 2) + 20);

            newListaArmarios = new List<Armario>();
            newListaCajones = new List<DatosCajones>();
            listaNoAlmacenados = new List<Consumos>();
            listRepetidos = new List<Consumos>();

            this.dgvUbic.Rows.Add("Máxima", "", "");
            this.dgvUbic.Rows.Add("Alta Triple", "", "");
            this.dgvUbic.Rows.Add("Alta Doble", "", "");
            this.dgvUbic.Rows.Add("Alta Simple", "", "");
            this.dgvUbic.Rows.Add("Media Matrix", "", "");
            this.dgvUbic.Rows.Add("Media Múltiple", "", "");
            this.dgvUbic.Rows.Add("Auxiliar Grande", "", "");
            this.dgvUbic.Rows.Add("Auxiliar Medio", "", "");
            this.dgvUbic.Rows.Add("Auxiliar Pequeño", "", "");
            this.dgvUbic.Rows.Add("Nevera", "", "");

            faltan = new CajonesCuenta[10];
            sobran = new CajonesCuenta[10];

            
        }
        public List<Consumos> noAlmacenados { get; set; }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblTitulo.Text = "PROGRAMA DE DISTRIBUCIÓN DE MEDICAMENTOS V" + Assembly.GetExecutingAssembly().GetName().Version; 
        }


            
           /* foreach (Consumos c in newListaConsumos)
            {
                listView1.Items.Add(c.codigo);
                listView1.Items.Add(c.tamano);
                listView1.Items.Add(c.consumoArmario.ToString());

            }
            Controls.Add(listView1);*/

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }



        private void btnConvExcel_Click(object sender, EventArgs e)
        {
            try
            {
                dgvMeds.Visible = false;
                lblNoDistribuidos.Visible = false;
                listaNoAlmacenados.Clear();
                for (int i = 0; i < (faltan.Count()); i++)
                {
                    dgvUbic.Rows[i].Cells["Falta"].Value = null;
                    dgvUbic.Rows[i].Cells["Libres"].Value = null;
                }
                noDistribuidos = 0;
                dgvMeds.Rows.Clear();
                Global.cierraProgressbar = false;
                th = new Thread(() => ThreadProgressBar(Color.FromArgb(20, 60, 100), Color.FromArgb(153, 180, 209), Color.FromArgb(20, 60, 100), "Generando..."));
                th.IsBackground = true;

                ObtenerExcelOrigen to = new ObtenerExcelOrigen();
                bool lista = to.FlujoObtenerExcelOrigen();
                if (!lista)
                    return;
                this.newListaConsumos = to.newListaConsumos;
                //cierraProgressBar = true;  
                if (newListaConsumos != null)
                {
                    maxlista = newListaConsumos.Count();
                    progressBar1.Maximum = maxlista;
               
                    bool continua = this.FlujoTranscribirExcelFinal(newListaConsumos);
                    
                                       
                }
                else
                {
                    frmAviso.MostrarAviso("El archivo importado no contiene medicamentos");
                }

            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.btnConvExcel_Click:" + ex.Message, 0);
            }
            return;
        }

        public bool FlujoTranscribirExcelFinal(List<Consumos> listaConsumos)
        {
            bool continuaEjecucion = false;            
            Global.cierraProgressbar = false;
            if (listaConsumos != null)
                continuaEjecucion = true;

            try
            {
                if (continuaEjecucion)
                {
                    continuaEjecucion = this.PreparacionExcelSalida();
                    
                }
                if (continuaEjecucion)
                {
                    th.Start();
                    continuaEjecucion = this.ObtenerExcelSalida();
                    if (!continuaEjecucion)
                        frmAviso.MostrarAviso("Importación de fichero erróneo");

                }
                if (!configVacia && continuaEjecucion)
                {
                    continuaEjecucion = this.VaciarConfig();
                }
                if (continuaEjecucion)
                {
                    
                    continuaEjecucion = this.ObtenerDatosDispensador();
                    if (!continuaEjecucion)
                        frmAviso.MostrarAviso("Importación del dispensador errónea");
                }
                /* if(continuaEjecucion)
                 {
                     continuaEjecucion = this.ObtenerCajones();
                     if(!continuaEjecucion)
                         MessageBox.Show("Error al obtener datos de Cajones")
                 }*/
                if (continuaEjecucion)
                {
                    continuaEjecucion = this.UbicarMedicamentos(listaConsumos);
                    if (!continuaEjecucion)
                        frmAviso.MostrarAviso("Importación de cajones errónea");
                }
                if (continuaEjecucion)
                {
                    continuaEjecucion = this.CuentaSobran();
                    if (!continuaEjecucion)
                        frmAviso.MostrarAviso("Error al contar cajones que quedan libres en el dispensador");
                }
                if (continuaEjecucion)
                {
                    continuaEjecucion = this.CuentaFaltan(listaNoAlmacenados);
                    if (!continuaEjecucion)
                        frmAviso.MostrarAviso("Error al contar medicamentos que faltan por ubicar");
                }
                if (continuaEjecucion)
                { 

                    if (listaNoAlmacenados.Count > 0)
                    {
                        noDistribuidos = listaNoAlmacenados.Count();
                        frmAviso.MostrarAviso("Hay Medicamentos que no se han conseguido ubicar por problemas de espacio");
                        //MessageBox.Show("Hay Medicamentos que no se han conseguido ubicar", "My Application", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                                dgvMeds.Visible = true;
                                lblNoDistribuidos.Text = " MEDICAMENTOS NO DISTRIBUIDOS (TOTAL: " + noDistribuidos + " )";
                                lblNoDistribuidos.Visible = dgvMeds.Visible;
                                foreach (Consumos c in listaNoAlmacenados)
                                {
                                    string[] row = new string[] { c.codigo, c.nombre, c.tamano, c.consumoArmario.ToString() };
                                    dgvMeds.Rows.Add(row);
                                }                            
                        
                    }
                    else
                    {
                        frmAviso.MostrarAviso("Se ha completado la distribución de todos los medicamentos sin problemas");
                    }
                }
                
                Global.cierraProgressbar = true;
                progressBar1.Value = progressBar1.Maximum;
                progressBar1.Value = 0;


            }
            catch (Exception ex)
            {
                Global.cierraProgressbar = true;
                Logger.Write("Ex.Form1.FlujoTranscribirExcelFinal: "+ex.Message, 0);
                frmAviso.MostrarAviso("Operación de distribución de Medicamentos No completada");
            }

            finally
            {
                if (this.xlApp != null)
                    this.xlApp.Workbooks.Close();
                    if (this.xlWorkSheet != null)
                        releaseObject(this.xlWorkSheet);
                    if (this.xlWorkBook != null)
                        releaseObject(this.xlWorkBook);
                    this.xlApp.Quit();
                    releaseObject(this.xlApp);
                    
                }
            

            return continuaEjecucion;
        }
        private bool PreparacionExcelSalida()
        {
            bool toReturn = false;
            try
            {
                //frmAviso.MostrarAviso("Elija el fichero xlms que contiene la configuración del Dispensador");

                var FD = new OpenFileDialog();
                DialogResult result = FD.ShowDialog();
                if (result == DialogResult.Cancel)
                { 
                    toReturn = false;
                }
               else if (result == DialogResult.OK && File.Exists(FD.FileName) && (Path.GetExtension(FD.FileName) == ".xlsm" || Path.GetExtension(FD.FileName) == ".xlsm" || Path.GetExtension(FD.FileName) == ".xlsm")) 
                {
                    file = FD.FileName;
                    toReturn = true;
                    DialogResult result2 = frmMessageBox.MostrarDialogo("Consulta", "¿El archivo seleccionado tiene medicamentos guardados?", true);
                    if (result2 == DialogResult.Cancel)
                        configVacia = true;
                }
                else
                {
                    frmAviso.MostrarAviso("El fichero seleccionado no es correcto");
                }
            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.PreparacionDeExcelSalida: " + ex.Message, 0);
                toReturn = false;
            }

            return toReturn;
        }

        private bool ObtenerExcelSalida()
        {
            bool toReturn;
            try
            {
                //Creo Objeto excel
                this.xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.Visible = true;
                //file = "C:\\Users\\aaglio\\Documents\\Proyectos\\DOSYS\\DispensadorExp_Empty.xlsm";
                //Objeto Workbook para crear excel
                this.xlWorkBook = xlApp.Workbooks.Open(file, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                //Objeto worksheet para hojas del documento
                this.sheets = this.xlWorkBook.Worksheets;
                int numSheets = this.sheets.Count;
                if (!configVacia)
                    this.progressBar1.Maximum = progressBar1.Maximum + (numSheets*2)+2;
                else
                { this.progressBar1.Maximum = progressBar1.Maximum + numSheets+2; }
                toReturn = true;
            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.ObtenerExcelSalida: " + ex.Message, 0);
                toReturn = false;

            }
            return toReturn;
        }

        private bool ObtenerDatosDispensador()
        {
            bool toReturn = false;
            

            try
            {
                this.xlWorkSheet = (Worksheet)this.sheets.get_Item("InformacionArmario");
                if (this.xlWorkSheet == null)
                    return toReturn;

                int indColumnaInicio = 0; //Indice del numeroALetra.

                Armario newArmario;

                while (this.xlWorkSheet.get_Range(numeroALetra[indColumnaInicio] + 1).Value2 != null)
                {
                    long armarioId = -9999;
                    int numFilas = 9999;
                    int numColumnas = 9999;
                    string numero;
                    int tipoArmario = 9999;
                    //int controladora = 9999;
                    // bool activo = false;

                    Range range = this.xlWorkSheet.get_Range(numeroALetra[indColumnaInicio] + 1).MergeArea;
                    int lastColumn = range.Columns.Count;
                    string columnaDatos;
                    if (lastColumn % 2 != 0)
                    {
                        int indColumnaAux = (int)Math.Ceiling((decimal)lastColumn / 2);
                        columnaDatos = numeroALetra[indColumnaAux + indColumnaInicio];
                    }
                    else
                        columnaDatos = numeroALetra[indColumnaInicio + (lastColumn / 2)];

                    newArmario = new Armario();

                    if (this.xlWorkSheet.get_Range(columnaDatos + "3").Value2 != null && long.TryParse(this.xlWorkSheet.get_Range(columnaDatos + "3").Value2.ToString(), out armarioId))
                        newArmario.ArmarioID = armarioId;
                    else
                        newArmario.ArmarioID = armarioId;

                    if (this.xlWorkSheet.get_Range(columnaDatos + "4").Value2 != null && Int32.TryParse(this.xlWorkSheet.get_Range(columnaDatos + "4").Value2.ToString(), out numFilas))
                        newArmario.filas = numFilas;
                    else
                        newArmario.filas = null;

                    if (this.xlWorkSheet.get_Range(columnaDatos + "5").Value2 != null && Int32.TryParse(this.xlWorkSheet.get_Range(columnaDatos + "5").Value2.ToString(), out numColumnas))
                        newArmario.columnas = numColumnas;
                    else
                        newArmario.columnas = null;

                    if (this.xlWorkSheet.get_Range(columnaDatos + "6").Value2 != null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range(columnaDatos + "6").Value2.ToString()))
                        newArmario.Nombre = this.xlWorkSheet.get_Range(columnaDatos + "6").Value2.ToString();
                    else
                        return false;

                    int iNumero;
                    if (this.xlWorkSheet.get_Range(columnaDatos + "7").Value2 != null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range(columnaDatos + "7").Value2.ToString()))
                    {
                        numero = this.xlWorkSheet.get_Range(columnaDatos + "7").Value2.ToString();
                        if (Int32.TryParse(numero, out iNumero))
                            newArmario.Numero = numero;
                        else
                            return false;
                    }
                    else
                        return false;

                    if (this.xlWorkSheet.get_Range(columnaDatos + "8").Value2 != null && Int32.TryParse(this.xlWorkSheet.get_Range(columnaDatos + "8").Value2.ToString(), out tipoArmario))
                        newArmario.tipo = tipoArmario;
                    else
                        return false;

                    /* if (this.xlWorkSheet.get_Range(columnaDatos + "9").Value2 != null && Int32.TryParse(this.xlWorkSheet.get_Range(columnaDatos + "9").Value2.ToString(), out controladora))
                         newArmario.controladora = controladora;
                     else
                         newArmario.controladora = null;

                     if (this.xlWorkSheet.get_Range(columnaDatos + "10").Value2 != null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range(columnaDatos + "10").Value2.ToString()) && ComprobacionBoolean(this.xlWorkSheet.get_Range(columnaDatos + "10").Value2.ToString()))
                     {
                         activo = ValorBoolean(this.xlWorkSheet.get_Range(columnaDatos + "10").Value2.ToString().ToUpper());
                         newArmario.Activo = activo;
                     }
                     else
                         newArmario.Activo = true;*/



                    newListaArmarios.Add(newArmario);

                    indColumnaInicio += lastColumn + 1;

                    toReturn = true;
                    progressBar1.Maximum++;
                    progressBar1.Value++;
                }
            }
            catch (Exception ex)
            {

                Logger.Write("Ex.Form1.ObtenerDatosDispensador :" + ex.Message, 0);
                toReturn = false;
            }

            return toReturn;
        }

        private bool UbicarMedicamentos(List<Consumos> listaConsumos)
        {
            indiceInicio = new int[100];
            indiceInicioAltaSimple = new int[100];
            indiceInicioAltaDoble = new int[100];
            indiceInicioAltaTriple = new int[100];
            indiceInicioAuxPequeño = new int[100];
            indiceInicioAuxMediano = new int[100];
            indiceInicioAuxGrande = new int[100];
            bool toReturn = false;
            bool continua = false;
            numAux = new List<long>();
            numBase = new List<long>();
            numNev = new List<long>();
            max = listaConsumos.Count();
            Stopwatch st = new Stopwatch();

            try
            {
                // 
                foreach (Armario a in newListaArmarios)
                {
                    if (a.tipo == 1)
                        numBase.Add(a.ArmarioID);


                    else if (a.tipo == 2)
                        numAux.Add(a.ArmarioID);



                    else
                        numNev.Add(a.ArmarioID);
                }

                //Console.WriteLine("ITERACIONES TOTAL = " + listaConsumos.Count);
                //Contemplo ubicacion por tamaño/seguridad
                foreach (Consumos c in listaConsumos)
                {
                    st.Restart();

                    sizeMed = c.tamano;
                    //Console.WriteLine("MEDICAMENTO " + c.codigo + " " + c.nombre+ " " + "CANTIDAD " + c.consumoArmario);

                    //Distribucion en armario auxiliar
                    if (sizeMed.Contains("columna"))
                    {
                        int ind = 0;
                        //Console.Write("COLUMNA ");
                        while (ind < numAux.Count)
                        {
                            string pattern = "Cajon-[A-H]-Armario-" + numAux[ind];
                            int sheetIndex = 1;
                            int numSheets = this.sheets.Count;
                            while (sheetIndex <= numSheets)
                            {

                                this.xlWorkSheet = (Worksheet)this.sheets.get_Item(sheetIndex);
                                bool isMatch = Regex.IsMatch(this.xlWorkSheet.Name, pattern, RegexOptions.IgnoreCase);
                                if (isMatch)
                                {
                                    if (sizeMed.Contains("peque"))
                                    {
                                        if (indiceInicioAuxPequeño[sheetIndex] == 0)
                                            indiceInicioAuxPequeño[sheetIndex] = 14;
                                        filaLectura = indiceInicioAuxPequeño[sheetIndex];
                                    }
                                    else if (sizeMed.Contains("doble"))
                                    {
                                        if (indiceInicioAuxMediano[sheetIndex] == 0)
                                            indiceInicioAuxMediano[sheetIndex] = 14;
                                        filaLectura = indiceInicioAuxMediano[sheetIndex];
                                    }
                                    else if (sizeMed.Contains("triple"))
                                    {
                                        if (indiceInicioAuxGrande[sheetIndex] == 0)
                                            indiceInicioAuxGrande[sheetIndex] = 14;
                                        filaLectura = indiceInicioAuxGrande[sheetIndex];
                                    }

                                    while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                    {


                                        if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                        {
                                            //Medicamento para auxiliar pequeño
                                            if (sizeMed.Contains("peque"))
                                            {
                                                if ((this.xlWorkSheet.get_Range("N" + filaLectura).Value2.ToString()).Contains("Peque"))
                                                {
                                                    this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                    this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                    continua = true;
                                                    break;
                                                }
                                                else
                                                {
                                                    filaLectura++;
                                                }
                                            }
                                            //Auxiliar Medio
                                            else if (sizeMed.Contains("doble"))
                                            {
                                                if ((this.xlWorkSheet.get_Range("N" + filaLectura).Value2.ToString()).Contains("Medio"))
                                                {
                                                    this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                    this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                    continua = true;
                                                    break;
                                                }
                                                else
                                                {
                                                    filaLectura++;
                                                }
                                            }
                                            //Auxiliar grande
                                            else if (sizeMed.Contains("triple"))
                                            {
                                                if ((this.xlWorkSheet.get_Range("N" + filaLectura).Value2.ToString()).Contains("Grande"))
                                                {
                                                    this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                    this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                    continua = true;
                                                    break;
                                                }
                                                else
                                                {
                                                    filaLectura++;
                                                }
                                            }
                                            else
                                            {
                                                frmAviso.MostrarAviso("El cajon especificado " + sizeMed + " no cumple el formato");
                                            }

                                        }

                                        else
                                        {
                                            filaLectura++;
                                        }
                                        if (sizeMed.Contains("peque"))
                                            indiceInicioAuxPequeño[sheetIndex] = filaLectura;
                                        else if (sizeMed.Contains("doble"))
                                            indiceInicioAuxMediano[sheetIndex] = filaLectura;
                                        else if (sizeMed.Contains("triple"))
                                            indiceInicioAuxGrande[sheetIndex] = filaLectura;
                                    }
                                }
                                if (continua == true)
                                    break;
                                else { sheetIndex++; }

                            }
                            if (continua == true)
                                break;
                            else { ind++; }
                        }
                    }



                    //Distribucion en nevera
                    else if (sizeMed.Contains("nevera"))
                    {
                        int ind = 0;
                        //Console.Write("NEVERA ");
                        while (ind < numNev.Count)
                        {
                            string pattern = "Cajon-[A-E]-Armario-" + numNev[ind];
                            int sheetIndex = 1;
                            int numSheets = this.sheets.Count;
                            while (sheetIndex <= numSheets)
                            {

                                this.xlWorkSheet = (Worksheet)this.sheets.get_Item(sheetIndex);
                                bool isMatch = Regex.IsMatch(this.xlWorkSheet.Name, pattern, RegexOptions.IgnoreCase);
                                if (isMatch)
                                {

                                    if (indiceInicio[sheetIndex] == 0)
                                        indiceInicio[sheetIndex] = 14;
                                    filaLectura = indiceInicio[sheetIndex];
                                    while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                    {


                                        if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                        {

                                            this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                            this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                            continua = true;

                                            break;
                                        }

                                        /* else if (this.xlWorkSheet.get_Range("F" + filaLectura).Value2 != null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range("F" + filaLectura).Value2.ToString()))
                                         {
                                             string codguardado = this.xlWorkSheet.get_Range("F" + filaLectura).Value2.ToString();
                                             Consumos guardado = listaConsumos.Find(x => x.codigo == codguardado);
                                             if (this.LevenshteinDistance(guardado.nombre, c.nombre, out porcentaje) < 6)
                                             {
                                                 break;
                                             }
                                             else
                                             {
                                                 filaLectura++;
                                             }

                                         }*/
                                        else
                                        { filaLectura++; }
                                    }
                                    indiceInicio[sheetIndex] = filaLectura;
                                }
                                if (continua == true)
                                { break; }
                                else { sheetIndex++; }

                            }
                            if (continua == true)
                                break;
                            else { ind++; }
                        }                        
                    }
                    //Distribución en base
                    else
                    {
                        int ind = 0;
                        while (ind < numBase.Count)
                        {
                            string pattern = "Cajon-[A-Z]-Armario-" + numBase[ind];
                            int sheetIndex = 1;
                            int numSheets = this.sheets.Count;
                            while (sheetIndex <= numSheets)
                            {

                                this.xlWorkSheet = (Worksheet)this.sheets.get_Item(sheetIndex);
                                bool isMax = false;
                                bool isMed = false;
                                bool isAlt = false;
                                bool isMatch = Regex.IsMatch(this.xlWorkSheet.Name, pattern, RegexOptions.IgnoreCase);
                                if (isMatch)
                                {
                                    isMax = (this.xlWorkSheet.get_Range("B3").Value2.ToString()).Contains("MAXIMA");
                                    isMed = (this.xlWorkSheet.get_Range("B3").Value2.ToString()).Contains("MATRIZ");
                                    isAlt = (this.xlWorkSheet.get_Range("B3").Value2.ToString()).Contains("ALTA");
                                }
                                //Distribucion en base maxima

                                //Console.Write("ESTUPE ");                        
                                if (sizeMed.Contains("estupe"))

                                {
                                    if (isMax)
                                    {
                                        if (indiceInicio[sheetIndex] == 0)
                                            indiceInicio[sheetIndex] = 14;
                                        filaLectura = indiceInicio[sheetIndex];
                                        while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                        {


                                            if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                            {

                                                this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                continua = true;

                                                break;
                                            }
                                            else
                                            { filaLectura++; }
                                        }
                                        indiceInicio[sheetIndex] = filaLectura;
                                    }
                                    if (continua == true)
                                    { break; }
                                    else { sheetIndex++; }
                                }


                                //Distribucion en Media Matrix
                                else if (sizeMed.Contains("matrix"))
                                {
                                    //Console.Write("MATRIX ");                               
                                    if (isMed)
                                    {
                                        int inicioFilaEstructura = 6;
                                        while (this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != null && this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                        {
                                            inicioFilaEstructura++;
                                        }
                                        int inicioFilaDatos = inicioFilaEstructura + 2;
                                        if (inicioFilaDatos == 16)
                                        {
                                            if (indiceInicio[sheetIndex] == 0)
                                                indiceInicio[sheetIndex] = inicioFilaDatos;

                                            filaLectura = indiceInicio[sheetIndex];
                                            while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != null && this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                            {
                                                if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                                {
                                                    this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                    this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                    continua = true;
                                                    break;
                                                }
                                                else
                                                { filaLectura++; }
                                            }

                                            indiceInicio[sheetIndex] = filaLectura;
                                        }
                                    }
                                    if (continua == true)
                                    { break; }
                                    else { sheetIndex++; }
                                }

                                //Distribucion en Media Matrix
                                else if (sizeMed.Contains("media"))
                                {
                                    //Console.Write("MATRIX ");                               
                                    if (isMed)
                                    {
                                        int inicioFilaEstructura = 6;
                                        while (this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != null && this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                        {
                                            inicioFilaEstructura++;
                                        }
                                        int inicioFilaDatos = inicioFilaEstructura + 2;
                                        if (inicioFilaDatos == 14)
                                        {
                                            if (indiceInicio[sheetIndex] == 0)
                                                indiceInicio[sheetIndex] = inicioFilaDatos;

                                            filaLectura = indiceInicio[sheetIndex];
                                            while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != null && this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                            {
                                                if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                                {
                                                    this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                    this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                    continua = true;
                                                    break;
                                                }
                                                else
                                                { filaLectura++; }
                                            }                                        
                                        indiceInicio[sheetIndex] = filaLectura;
                                        }
                                    }
                                    if (continua == true)
                                    { break; }
                                    else { sheetIndex++; }
                                }

                                //Distribucion en alta 
                                else
                               {
                                //Console.Write("ALTA ");                                
                                string coord = null;
                                string coordA;
                                string coordB;
                                int A;
                                int B;                               
                                    //Stopwatch stSheet = new Stopwatch();
                                    // stSheet.Restart();                                    
                                    if (isAlt)
                                    {
                                        if (sizeMed.Contains("peque"))
                                        {
                                            if (indiceInicioAltaSimple[sheetIndex] == 0)
                                                indiceInicioAltaSimple[sheetIndex] = 16;
                                            filaLectura = indiceInicioAltaSimple[sheetIndex];
                                        }
                                        else if (sizeMed.Contains("doble"))
                                        {
                                            if (indiceInicioAltaDoble[sheetIndex] == 0)
                                                indiceInicioAltaDoble[sheetIndex] = 16;
                                            filaLectura = indiceInicioAltaDoble[sheetIndex];
                                        }
                                        else if (sizeMed.Contains("triple"))
                                        {
                                            if (indiceInicioAltaTriple[sheetIndex] == 0)
                                                indiceInicioAltaTriple[sheetIndex] = 16;
                                            filaLectura = indiceInicioAltaTriple[sheetIndex];
                                        }

                                        while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                        {

                                            if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                            {
                                                Stopwatch stCoor = new Stopwatch();
                                                stCoor.Restart();
                                                coord = this.xlWorkSheet.get_Range("D" + filaLectura).Value2.ToString();

                                                coordA = coord.Substring(0, 1).ToString();
                                                A = Convert.ToInt32(coordA, 16);

                                                coordB = coord.Substring(2, 1).ToString();
                                                B = Convert.ToInt32(coordB, 16);

                                                //Alta simple
                                                if (sizeMed.Contains("peque"))
                                                {
                                                    if (coordA == coordB)
                                                    {
                                                        this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                        this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                        continua = true;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        filaLectura++;
                                                    }
                                                }
                                                //Alta doble
                                                else if (sizeMed.Contains("doble"))
                                                {
                                                    if (B - A == 1)
                                                    {
                                                        this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                        this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                        continua = true;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        filaLectura++;
                                                    }
                                                }
                                                //Alta triple
                                                else if (sizeMed.Contains("triple"))
                                                {



                                                    if (B - A == 2)
                                                    {


                                                        this.xlWorkSheet.Range["F" + filaLectura].Value2 = c.codigo;
                                                        this.xlWorkSheet.Range["H" + filaLectura].Value2 = c.consumoArmario;
                                                        continua = true;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        filaLectura++;
                                                    }
                                                }
                                                else
                                                {
                                                    frmAviso.MostrarAviso("El cajon especificado " + sizeMed + " no cumple formato");
                                                }
                                                st.Stop();
                                                //Console.WriteLine("Coordenadas: " + st.ElapsedMilliseconds);
                                            }
                                            else
                                            {
                                                filaLectura++;
                                            }

                                        }
                                        if (sizeMed.Contains("peque"))
                                        {
                                            indiceInicioAltaSimple[sheetIndex] = filaLectura;
                                        }
                                        else if (sizeMed.Contains("doble"))
                                        {
                                            indiceInicioAltaDoble[sheetIndex] = filaLectura;
                                        }
                                        else if (sizeMed.Contains("triple"))
                                        {
                                            indiceInicioAltaTriple[sheetIndex] = filaLectura;
                                        }


                                    }
                                    if (continua == true)
                                    { break; }
                                    else { sheetIndex++; }

                                    //Console.WriteLine("SHEET: " + sheetIndex + " T = " + stSheet.ElapsedMilliseconds);
                                }
                            }
                            if (continua == true)
                                break;
                            else { ind++; }
                        }
                    }

                    if (continua == false)
                        listaNoAlmacenados.Add(c);
                    else
                    { continua = false; }

                    progressBar1.Value++;
                    st.Stop();
                    //Console.WriteLine("T = " + st.ElapsedMilliseconds);
                    
                }
                noAlmacenados = listaNoAlmacenados;
                toReturn = true;
            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.UbicarMedicamentos: " + ex.Message, 0);
                toReturn = false;

            }




            return toReturn;
        }

        private bool ComprobacionBoolean(string entrada)
        {
            if (entrada.Equals("1") || entrada.Equals("0"))
                return true;
            else if (entrada.ToUpper().Equals("TRUE") || entrada.ToUpper().Equals("FALSE") || entrada.ToUpper().Equals("VERDADERO") || entrada.Equals("FALSO"))
                return true;
            else
                return false;
        }

        
        //Calculo de Cajones necesarios en el dispensador
        private bool CuentaFaltan(List<Consumos> listnoAlmacenados)
        {
            bool cuenta = false;
            try
            {            
                faltan[0].tipo = "estupe";
                faltan[0].num = 0;
                faltan[1].tipo = "triple";
                faltan[1].num = 0;
                faltan[2].tipo = "doble";
                faltan[2].num = 0;
                faltan[3].tipo = "pequeno";
                faltan[3].num = 0;
                faltan[4].tipo = "matrix";
                faltan[4].num = 0;
                faltan[5].tipo = "media";
                faltan[5].num = 0;
                faltan[6].tipo = "columna triple";
                faltan[6].num = 0;
                faltan[7].tipo = "columna doble";
                faltan[7].num = 0;
                faltan[8].tipo = "columna pequeno";
                faltan[8].num = 0;
                faltan[9].tipo = "nevera";
                faltan[9].num = 0;
                        
          
                foreach (Consumos c in listaNoAlmacenados)
                {

                    if (c.tamano.Contains("estupe"))
                        faltan[0].num++;
                    else if (c.tamano.Contains("columna triple"))
                        faltan[6].num++;
                    else if (c.tamano.Contains("columna doble"))
                        faltan[7].num++;
                    else if (c.tamano.Contains("columna peque"))
                        faltan[8].num++;
                    else if (c.tamano.Contains("triple"))
                        faltan[1].num++;
                    else if (c.tamano.Contains("doble"))
                        faltan[2].num++;
                    else if (c.tamano.Contains("peque"))
                        faltan[3].num++;
                    else if (c.tamano.Contains("matrix"))
                        faltan[4].num++;
                    else if (c.tamano.Contains("nevera"))
                        faltan[9].num++;
                    else if (c.tamano.Contains("media"))
                        faltan[5].num++;
                }
                progressBar1.Value++;
             
                for (int i=0; i<(faltan.Count()); i++)
                {
                    dgvUbic.Rows[i].Cells["Falta"].Value = faltan[i].num.ToString();
                }
                progressBar1.Value++;
                cuenta = true;
            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.CuentaFaltan :" + ex.Message, 0);
                cuenta = false;
            }
            return cuenta;
        }

        //Calculo de Cajones libres en el dispensador
        private bool CuentaSobran()
        {  
            bool cuentasob = false;
            try
            {
                sobran[0].tipo = "estupe";
                sobran[0].num = 0;
                sobran[1].tipo = "triple";
                sobran[1].num = 0;
                sobran[2].tipo = "doble";
                sobran[2].num = 0;
                sobran[3].tipo = "pequeno";
                sobran[3].num = 0;
                sobran[4].tipo = "matrix";
                sobran[4].num = 0;
                sobran[5].tipo = "media";
                sobran[5].num = 0;
                sobran[6].tipo = "columna triple";
                sobran[6].num = 0;
                sobran[7].tipo = "columna doble";
                sobran[7].num = 0;
                sobran[8].tipo = "columna pequeno";
                sobran[8].num = 0;
                sobran[9].tipo = "nevera";
                sobran[9].num++;

                string pattern = "Cajon-[A-Z]-Armario+";
                int sheetIndex = 1;
                int numSheets = this.sheets.Count;
                bool isAux = false;
                bool isNev = false;
                bool isMax = false;
                bool isAlt = false;
                bool isMed = false;
                bool isBase = false;
                string coordA=null;
                string coordB=null;
                int A=-999;
                int B= -999;

                //Cuenta Cajones libres
                while (sheetIndex <=numSheets)
                {
                    this.xlWorkSheet = (Worksheet)this.sheets.get_Item(sheetIndex);
                    if (this.xlWorkSheet == null)
                        return cuentasob;

                    bool isMatch = Regex.IsMatch(this.xlWorkSheet.Name, pattern, RegexOptions.IgnoreCase);
                    if (isMatch)
                    {
                        isAux = (this.xlWorkSheet.get_Range("B4").Value2.ToString()).Contains("AUXILIAR");
                        isNev = (this.xlWorkSheet.get_Range("B4").Value2.ToString()).Contains("NEVERA");
                        isBase = (this.xlWorkSheet.get_Range("B4").Value2.ToString()).Contains("BASE");
                        //La Hoja es del Armario Base
                        if (isBase)
                        {
                            isMax = (this.xlWorkSheet.get_Range("B3").Value2.ToString()).Contains("MAXIMA");
                            isAlt = (this.xlWorkSheet.get_Range("B3").Value2.ToString()).Contains("ALTA");
                            isMed = (this.xlWorkSheet.get_Range("B3").Value2.ToString()).Contains("MEDIA");
                        
                            // Cajon tipo Maxima
                            if (isMax)
                            {
                                filaLectura = 14;
                                while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                {
                                    if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)                                        
                                        sobran[0].num++;

                                 filaLectura++;
                                }
                            }

                            //Cajon tipo ALta
                            else if (isAlt)
                            {
                                filaLectura = 16;
                                while (this.xlWorkSheet.get_Range("A"+filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                {
                                    if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 !=null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range("C" + filaLectura).Value2.ToString()) && this.xlWorkSheet.get_Range("F"+filaLectura).Value2 == null)
                                    {
                                        coordA = (this.xlWorkSheet.get_Range("D" + filaLectura).Value2.ToString()).Substring(0, 1).ToString();
                                        A = Convert.ToInt32(coordA, 16);
                                        coordB = (this.xlWorkSheet.get_Range("D" + filaLectura).Value2.ToString()).Substring(2, 1).ToString();
                                        B = Convert.ToInt32(coordB, 16);

                                        //Cajetin Alta Triple
                                        if (B - A == 2)
                                            sobran[1].num++;
                                        //Cajetin ALta Doble
                                        else if (B - A == 1)
                                            sobran[2].num++;
                                        //Cajetin Alta Simple
                                        else if (B - A == 0)
                                            sobran[3].num++;
                                        else
                                        {
                                            frmAviso.MostrarAviso("El Cajetin en " + filaLectura + " de " + this.xlWorkSheet.Name + " no cumple las condiciones de configuración requeridas");
                                        }
                                       
                                    }
                                    filaLectura++;
                                }
                            }

                            //Cajon tipo Media
                            else if(isMed)
                            {
                                int inicioFilaEstructura = 6;
                                int columnas = 0;
                                while (this.xlWorkSheet.get_Range("A"+inicioFilaEstructura).Interior.Color !=null && this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != ColorTranslator.ToOle(Color.White))
                                {
                                    columnas++;
                                    inicioFilaEstructura++;
                                }
                                int inicioFilaDatos = inicioFilaEstructura + 2;
                                //Cajon tipo Media Matrix
                                if (inicioFilaDatos == 16)
                                {
                                    while (this.xlWorkSheet.get_Range("A" + inicioFilaDatos).Interior.Color != null && this.xlWorkSheet.get_Range("A" + inicioFilaDatos).Interior.Color != ColorTranslator.ToOle(Color.White))
                                    {
                                        if (this.xlWorkSheet.get_Range("C" + inicioFilaDatos).Value2 != null && this.xlWorkSheet.get_Range("F" + inicioFilaDatos).Value2 == null)
                                        {
                                            sobran[4].num++;
                                        }
                                        inicioFilaDatos++;

                                    }
                                }
                                //Cajon tipo Media Multiple
                                if (inicioFilaDatos == 14)
                                {
                                    while (this.xlWorkSheet.get_Range("A" + inicioFilaDatos).Interior.Color != null && this.xlWorkSheet.get_Range("A" + inicioFilaDatos).Interior.Color != ColorTranslator.ToOle(Color.White))
                                    {
                                        if (this.xlWorkSheet.get_Range("C" + inicioFilaDatos).Value2 != null && this.xlWorkSheet.get_Range("F" + inicioFilaDatos).Value2 == null)
                                        {
                                            sobran[5].num++;
                                        }
                                        inicioFilaDatos++;

                                    }
                                }
                            }

                            else
                            {
                                frmAviso.MostrarAviso("El Cajón" + this.xlWorkSheet.Name + " no cumple las condiciones de configuración requeridas");
                            }
                        }

                        //Armario Auxiliar
                        else if(isAux)
                        {
                            filaLectura = 14;
                            while (this.xlWorkSheet.get_Range("A" + filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                            {
                                if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                {
                                    //Cajon Grande
                                    if ((this.xlWorkSheet.get_Range("N" + filaLectura).Value2.ToString()).Contains("Grande"))
                                        sobran[6].num++;
                                    else if ((this.xlWorkSheet.get_Range("N" + filaLectura).Value2.ToString()).Contains("Medio"))
                                        sobran[7].num++;
                                    else if ((this.xlWorkSheet.get_Range("N" + filaLectura).Value2.ToString()).Contains("Peque"))
                                        sobran[8].num++;
                                    else { frmAviso.MostrarAviso("El Cajón" + this.xlWorkSheet.Name + " no cumple las condiciones de configuración requeridas"); }
                                }
                                filaLectura++;
                            }
                        }

                        //Nevera
                        else if (isNev)
                        {
                            filaLectura = 14;
                            while (this.xlWorkSheet.get_Range("A"+filaLectura).Interior.Color != ColorTranslator.ToOle(Color.White))
                            {
                                if (this.xlWorkSheet.get_Range("C" + filaLectura).Value2 != null && this.xlWorkSheet.get_Range("F" + filaLectura).Value2 == null)
                                    sobran[9].num++;

                                filaLectura++;
                            }
                        }

                        else
                        {
                            frmAviso.MostrarAviso("La Hoja " + this.xlWorkSheet.Name + " no cumple las condiciones de configuración requeridas");
                        }
                      }
                    progressBar1.Value++;
                       sheetIndex++;                        
                    }

                //Muestra Cuenta en tabla
                for (int i = 0; i < (sobran.Count()); i++)
                {
                    dgvUbic.Rows[i].Cells["Libres"].Value = sobran[i].num.ToString();
                }

                cuentasob = true;
            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.CuentaSobran " + ex.Message, 0);
                cuentasob = false;
            }
            return cuentasob;
        }

        private bool VaciarConfig()
        {
           
            bool vacio = false;
            try
            {
                string pattern = "Cajon-[A-Z]-Armario+";
                int sheetIndex = 1;
                int numSheets = this.sheets.Count;

                int minimo = -9999;
                int maximo = -9999;
                decimal cantidad = -9999;
                DateTime caducidad = new DateTime();
                

                //Cuenta Cajones libres
                while (sheetIndex <= numSheets)
                {
                    this.xlWorkSheet = (Worksheet)this.sheets.get_Item(sheetIndex);
                    if (this.xlWorkSheet == null)
                        return vacio;

                    bool isMatch = Regex.IsMatch(this.xlWorkSheet.Name, pattern, RegexOptions.IgnoreCase);
                    if (isMatch)
                    {
                        int inicioFilaEstructura = 6;
                        while (this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != null && this.xlWorkSheet.get_Range("A" + inicioFilaEstructura).Interior.Color != ColorTranslator.ToOle(Color.White))
                        {
                            inicioFilaEstructura++;
                        }
                           int inicioFilaDatos = inicioFilaEstructura + 2;
                           while (this.xlWorkSheet.get_Range("A" + inicioFilaDatos).Interior.Color != null && this.xlWorkSheet.get_Range("A" + inicioFilaDatos).Interior.Color != ColorTranslator.ToOle(Color.White))
                           {
                            if (this.xlWorkSheet.get_Range("F" + inicioFilaDatos).Value2 != null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range("F" + inicioFilaDatos).Value2.ToString()))
                            {
                                this.xlWorkSheet.get_Range("F" + inicioFilaDatos).Value2 = null;
                                if (this.xlWorkSheet.get_Range("H" + inicioFilaDatos).Value2 != null && Int32.TryParse(this.xlWorkSheet.get_Range("H" + inicioFilaDatos).Value2.ToString(), out maximo))
                                    this.xlWorkSheet.get_Range("H" + inicioFilaDatos).ClearContents();
                                if (this.xlWorkSheet.get_Range("G" + inicioFilaDatos).Value2 != null && Int32.TryParse(this.xlWorkSheet.get_Range("G" + inicioFilaDatos).Value2.ToString(), out minimo))
                                    this.xlWorkSheet.get_Range("G" + inicioFilaDatos).ClearContents();
                                if (this.xlWorkSheet.get_Range("I" + inicioFilaDatos).Value2 != null && decimal.TryParse(this.xlWorkSheet.get_Range("I" + inicioFilaDatos).Value2.ToString(), out cantidad))
                                    this.xlWorkSheet.get_Range("I" + inicioFilaDatos).ClearContents();
                                if (this.xlWorkSheet.get_Range("J" + inicioFilaDatos).Value2 != null && DateTime.TryParse(this.xlWorkSheet.get_Range("J" + inicioFilaDatos).Value2.ToString(), out caducidad))
                                    this.xlWorkSheet.get_Range("J" + inicioFilaDatos).ClearContents();
                                if (this.xlWorkSheet.get_Range("K" + inicioFilaDatos).Value2 != null && !string.IsNullOrEmpty(this.xlWorkSheet.get_Range("K" + inicioFilaDatos).Value2.ToString()) && ComprobacionBoolean(this.xlWorkSheet.get_Range("K" + inicioFilaDatos).Value2.ToString()))
                                    this.xlWorkSheet.get_Range("K" + inicioFilaDatos).ClearContents();
                            }
                            inicioFilaDatos++;
                        }
                           

                    }
                    
                    sheetIndex++;
                    progressBar1.Value++;
                }
                vacio = true;
            }
            catch (Exception ex)
            {
                Logger.Write("Ex.Form1.VaciarConfig :" + ex.Message, 0);
                vacio = false;
            }
            return vacio;
        }
     

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dGVLleno_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        /* public int LevenshteinDistance(string s, string t, out double porcentaje)
         {
             porcentaje = 0;

             // d es una tabla con m+1 renglones y n+1 columnas
             int costo = 0;
             int m = s.Length;
             int n = t.Length;
             int[,] d = new int[m + 1, n + 1];

             // Verifica que exista algo que comparar
             if (n == 0) return m;
             if (m == 0) return n;

             // Llena la primera columna y la primera fila.
             for (int i = 0; i <= m; d[i, 0] = i++) ;
             for (int j = 0; j <= n; d[0, j] = j++) ;


             /// recorre la matriz llenando cada unos de los pesos.
             /// i columnas, j renglones
             for (int i = 1; i <= m; i++)
             {
                 // recorre para j
                 for (int j = 1; j <= n; j++)
                 {
                     /// si son iguales en posiciones equidistantes el peso es 0
                     /// de lo contrario el peso suma a uno.
                     costo = (s[i - 1] == t[j - 1]) ? 0 : 1;
                     d[i, j] = System.Math.Min(System.Math.Min(d[i - 1, j] + 1,  //Eliminacion
                                   d[i, j - 1] + 1),                             //Inserccion 
                                   d[i - 1, j - 1] + costo);                     //Sustitucion
                 }
             }

             /// Calculamos el porcentaje de cambios en la palabra.
             if (s.Length > t.Length)
                 porcentaje = ((double)d[m, n] / (double)s.Length);
             else
                 porcentaje = ((double)d[m, n] / (double)t.Length);
             return d[m, n];
         }*/

        public void ThreadProgressBar(Color colorFondo, Color colorBolas, Color colorTexto, String texto)
        {
            frmProgressBar fpb = new frmProgressBar(colorFondo, colorBolas, colorTexto, texto);
            fpb.Size = new Size(102, 97);
            fpb.Location = new System.Drawing.Point((1920 / 2) - (fpb.Width / 2) + 1, (1080 / 2) - (fpb.Height / 2) + 20);
            fpb.Show();
            fpb.Start();
            fpb.Refresh();
            while (!Global.cierraProgressbar) 
            {
                System.Windows.Forms.Application.DoEvents();
            } 
           // fpb.Stop();
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
                frmAviso.MostrarAviso("AccionNoPosible");
                Logger.Write("Ex.Form1.releaseObject" + ex.Message, 0);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
