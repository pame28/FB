using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Telerik.Reporting;
using Telerik.Reporting.ExcelRendering;
using Telerik.Reporting.Processing;
using System.Data;
using System.IO;
using System.Text;
using System.Drawing;
using System.Data.SqlClient;
using System.Configuration;
using DevExpress.Web;
using DevExpress.Web.ASPxUploadControl;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace BasePerdidasOperativas.FormIngresoInfo
{

    public partial class CargaArchivoDemandasJuiciosEjecucion : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["usuario"] == null)
            {
                Response.Redirect("~/Transacciones/Login.aspx");
            }

            if (!IsPostBack)
            {
                gvInconsistencias.Visible = false;
            }
        }

        protected void uploadArchivoExcel_FileUploadComplete(object sender, DevExpress.Web.ASPxUploadControl.FileUploadCompleteEventArgs e)
        {
            gvInconsistencias.Visible = false;
            LblError.Visible = false;
        }

        #region Estructura de Archivo

        protected bool compararEstructura(DataTable datos)
        {
            bool iguales = true;
            string rutaPlantilla = HttpContext.Current.Request.MapPath("~/Transacciones/Plantillas/PlantillaDemandasJuiciosEjecucion.xlsx");

            try
            {
                DataTable dtPlantilla;
                using (ExcelPackage pk = new ExcelPackage(new FileInfo(rutaPlantilla)))
                {
                    dtPlantilla = ExcelPackageToDataTable(pk);
                }

                //Se valida si tienen el mismo número de columnas
                if (dtPlantilla.Columns.Count != datos.Columns.Count)
                {
                    return false;
                }

                //Se valida que las columnas se llamen igual
                for (int i = 0; i < dtPlantilla.Columns.Count; i++)
                {
                    if (dtPlantilla.Columns[i].ColumnName != datos.Columns[i].ColumnName)
                    {
                        iguales = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MostrarError("Ocurrió un error al comparar la estructura del archivo. " + ex.Message);
                iguales = false;
            }

            return iguales;
        }

        /// <summary>
        /// Guarda el archivo en el servidor y lo convierte a DataTable (sin modificar datos).
        /// </summary>
        public DataTable CargaArchivoExcelSpreadSheet(string RutaArchivo)
        {
            if (File.Exists(RutaArchivo)) File.Delete(RutaArchivo);
            uploadArchivoExcel.SaveAs(RutaArchivo);

            using (ExcelPackage pk = new ExcelPackage(new FileInfo(RutaArchivo)))
            {
                return ExcelPackageToDataTable(pk);
            }
        }

        //Funcion que convierte excel a un DataTable
        protected static DataTable ExcelPackageToDataTable(ExcelPackage excelPackage)
        {
            DataTable dt = new DataTable();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];
            //check if the worksheet is completely empty
            if (worksheet.Dimension == null)
            {
                return dt;
            }
            //create a list to hold the column names
            List<string> columnNames = new List<string>();
            //needed to keep track of empty column headers
            int currentColumn = 1;
            //loop all columns in the sheet and add them to the datatable
            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();
                //check if the previous header was empty and add it if it was
                if (cell.Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                    currentColumn++;
                }
                //add the column name to the list to count the duplicates
                columnNames.Add(columnName);
                //count the duplicate column names and make them unique to avoid the exception
                int occurrences = columnNames.Count(x => x.Equals(columnName));
                if (occurrences > 1)
                {
                    columnName = columnName + "_" + occurrences;
                }
                //add the column to the datatable
                dt.Columns.Add(columnName);
                currentColumn++;
            }
            //start adding the contents of the excel file to the datatable
            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();
                //loop all cells in the row
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(newRow);
            }
            return dt;
        }

        /// <summary>
        /// Quita filas vacías y guarda el número de fila real de Excel de cada registro.
        /// </summary>
        private DataTable QuitarFilasVacias(DataTable dt, out List<int> filasExcel)
        {
            filasExcel = new List<int>();
            DataTable limpio = dt.Clone();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool vacia = dt.Rows[i].ItemArray.All(v => string.IsNullOrWhiteSpace(Convert.ToString(v)));
                if (!vacia)
                {
                    limpio.ImportRow(dt.Rows[i]);
                    filasExcel.Add(i + 2); // fila 1 = encabezados
                }
            }
            return limpio;
        }

        #endregion

        #region Validaciones

        // Campos que pueden venir vacíos (por requerimiento, ninguno)
        private static readonly HashSet<string> CamposOpcionales = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
        };

        // Nombres legibles para los mensajes de error
        private static readonly Dictionary<string, string> NombresCampos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            {"empresa", "Empresa"},
            {"Zona", "Zona"},
            {"NumeroDemanda", "Número de Demanda"},
            {"NumeroExpediente", "Número de Expediente"},
            {"FechaPresentacion", "Fecha de Presentación"},
            {"NaturalezaDemanda", "Naturaleza de la Demanda"},
            {"NombreDemandado", "Nombre del Demandado"},
            {"NombreDemandante", "Nombre del Demandante"},
            {"TipoDemandante", "Tipo de Demandante"},
            {"TipoDemanda", "Tipo de Demanda"},
            {"MotivoDemanda", "Motivo de la Demanda"},
            {"AccionInstitucion", "Acción de la Institución"},
            {"JuzgadoDemanda", "Juzgado de la Demanda"},
            {"NombreJuzgado", "Nombre del Juzgado"},
            {"MontoDemandado", "Monto Demandado"},
            {"moneda", "Moneda"},
            {"TipoCambio", "Tipo de Cambio"},
            {"ResultadoPrimera", "Resultado Primera Instancia"},
            {"FechaActuacion", "Fecha de Actuación"},
            {"ResultadoSegunda", "Resultado Segunda Instancia"},
            {"MontoGastosLegalesHonorarios", "Monto Gastos Legales u Honorarios"},
            {"DemandaAsociadaRiesgoCredito", "Demanda Asociada a Riesgo de Crédito"},
            {"TipoGarantia", "Tipo de Garantía"},
            {"ProbabilidadPerdida", "Probabilidad de Pérdida"},
            {"ProcuradorAsignado", "Procurador Asignado"}
        };

        /// <summary>
        /// Valida todo el archivo. Devuelve una tabla Fila | Campo | Error.
        /// Si no hay errores, deja montos y fechas listos para insertar.
        /// </summary>
        private DataTable ValidarArchivo(DataTable dt, List<int> filasExcel)
        {
            DataTable errores = new DataTable();
            errores.Columns.Add("Fila", typeof(int));
            errores.Columns.Add("Campo", typeof(string));
            errores.Columns.Add("Error", typeof(string));

            HashSet<string> existentes = ObtenerDemandasExistentes();
            HashSet<string> enArchivo = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            DateTime fechaRegistro = DateTime.Today; // fecha de registro de la información = fecha de carga

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
                int fila = filasExcel[i];

                // 1. Campos obligatorios
                foreach (DataColumn col in dt.Columns)
                {
                    if (CamposOpcionales.Contains(col.ColumnName)) continue;
                    if (string.IsNullOrWhiteSpace(Convert.ToString(row[col])))
                    {
                        string mensaje = col.ColumnName.Equals("MontoDemandado", StringComparison.OrdinalIgnoreCase)
                            ? "Error en monto demandado, no puede quedar vacío, se debe colocar un valor estimado."
                            : "El campo " + Nombre(col.ColumnName) + " no puede quedar en blanco.";
                        Agregar(errores, fila, col.ColumnName, mensaje);
                    }
                }

                // 2. Número de demanda: existente en el sistema o repetido en el archivo
                string numero = Txt(row, "NumeroDemanda");
                if (numero != "")
                {
                    if (existentes.Contains(numero))
                    {
                        Agregar(errores, fila, "NumeroDemanda", "Número de demanda duplicado o existente.");
                    }
                    if (!enArchivo.Add(numero))
                    {
                        Agregar(errores, fila, "NumeroDemanda", "Número de demanda duplicado o existente en el archivo.");
                    }
                }

                // 3. Zona: I, II o III
                string zona = Txt(row, "Zona").ToUpper();
                if (zona != "" && zona != "I" && zona != "II" && zona != "III")
                {
                    Agregar(errores, fila, "Zona", "Error en campo zona, debe registrarse como: I, II o III.");
                }

                // 4. Fechas (formato DD/MM/AAAA)
                DateTime fPresentacion, fActuacion;
                bool okPresentacion = ValidarFecha(Txt(row, "FechaPresentacion"), out fPresentacion);
                bool okActuacion = ValidarFecha(Txt(row, "FechaActuacion"), out fActuacion);

                if (Txt(row, "FechaPresentacion") != "" && !okPresentacion)
                {
                    Agregar(errores, fila, "FechaPresentacion", "Error en formato fecha de presentación, formato correcto: \"DD/MM/AAAA\".");
                }
                if (Txt(row, "FechaActuacion") != "" && !okActuacion)
                {
                    Agregar(errores, fila, "FechaActuacion", "Error en formato de fecha de actuación, formato correcto: \"DD/MM/AAAA\".");
                }
                if (okPresentacion && fPresentacion > fechaRegistro)
                {
                    Agregar(errores, fila, "FechaPresentacion", "Error en fecha de presentación, es mayor a la fecha de registro de información.");
                }
                if (okPresentacion && okActuacion && fActuacion < fPresentacion)
                {
                    Agregar(errores, fila, "FechaActuacion", "Error en fecha de actuación, es menor a la fecha de presentación de la demanda.");
                }

                // 5. Tipo de cambio: 4 decimales
                string tipoCambio = LimpiarMonto(Txt(row, "TipoCambio"), 4);
                decimal valor;
                if (tipoCambio != "")
                {
                    if (!ValidarDecimales(tipoCambio, 4, out valor))
                        Agregar(errores, fila, "TipoCambio", "Error en tipo de cambio, debe estar en formato lempiras con cuatro cifras decimales después del punto.");
                    else
                        row["TipoCambio"] = tipoCambio;
                }

                // 6. Monto demandado: 2 decimales y distinto de 0.00
                string montoDemandado = LimpiarMonto(Txt(row, "MontoDemandado"), 2);
                if (montoDemandado != "")
                {
                    if (!ValidarDecimales(montoDemandado, 2, out valor))
                        Agregar(errores, fila, "MontoDemandado", "Error en monto demandado, debe ir en formato lempiras con dos cifras decimales después del punto.");
                    else if (valor == 0)
                        Agregar(errores, fila, "MontoDemandado", "Error en monto demandado, no puede quedar vacío, se debe colocar un valor estimado.");
                    else
                        row["MontoDemandado"] = montoDemandado;
                }

                // 7. Gastos legales u honorarios: 2 decimales (0.00 sí se permite)
                string gastos = LimpiarMonto(Txt(row, "MontoGastosLegalesHonorarios"), 2);
                if (gastos != "")
                {
                    if (!ValidarDecimales(gastos, 2, out valor))
                        Agregar(errores, fila, "MontoGastosLegalesHonorarios", "Error en monto gastos legales u honorarios, debe ir en formato lempiras con dos cifras decimales después del punto.");
                    else
                        row["MontoGastosLegalesHonorarios"] = gastos;
                }

                // Fechas en formato ISO para que SQL Server no las confunda (día/mes)
                if (okPresentacion) row["FechaPresentacion"] = fPresentacion.ToString("yyyy-MM-dd");
                if (okActuacion) row["FechaActuacion"] = fActuacion.ToString("yyyy-MM-dd");
            }

            return errores;
        }

        private HashSet<string> ObtenerDemandasExistentes()
        {
            HashSet<string> existentes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string cadena = ConfigurationManager.ConnectionStrings["BasePerdidasOperativas.Properties.Settings.SettingConnectionString"].ConnectionString;
            const string sql = "SELECT LTRIM(RTRIM(NumeroDemanda)) FROM tblRegistroDemandasJuiciosEjecucion WHERE NumeroDemanda IS NOT NULL";

            using (SqlConnection cn = new SqlConnection(cadena))
            using (SqlCommand cmd = new SqlCommand(sql, cn))
            {
                cn.Open();
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        existentes.Add(dr.GetString(0));
                    }
                }
            }
            return existentes;
        }

        private static bool ValidarFecha(string texto, out DateTime fecha)
        {
            return DateTime.TryParseExact(texto, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha);
        }

        /// <summary>
        /// Quita L, Lps, $, comas y espacios. Un "-" solo (cero contable) se convierte en 0.
        /// </summary>
        private static string LimpiarMonto(string texto, int decimales)
        {
            string limpio = texto.ToUpper().Replace("LPS", "").Replace("L", "").Replace("$", "").Replace(",", "").Replace(" ", "").Trim();
            if (limpio == "-")
            {
                limpio = 0m.ToString("F" + decimales, CultureInfo.InvariantCulture);
            }
            return limpio;
        }

        /// <summary>
        /// Valida número positivo con exactamente N decimales después del punto.
        /// </summary>
        private static bool ValidarDecimales(string texto, int decimales, out decimal valor)
        {
            valor = 0;
            if (!Regex.IsMatch(texto, @"^\d+\.\d{" + decimales + "}$")) return false;
            return decimal.TryParse(texto, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out valor);
        }

        private static string Txt(DataRow row, string columna)
        {
            if (!row.Table.Columns.Contains(columna)) return "";
            return Convert.ToString(row[columna]).Trim();
        }

        private static string Nombre(string columna)
        {
            string nombre;
            return NombresCampos.TryGetValue(columna, out nombre) ? nombre : columna;
        }

        private static void Agregar(DataTable errores, int fila, string columna, string mensaje)
        {
            errores.Rows.Add(fila, Nombre(columna), mensaje);
        }

        private void MostrarError(string mensaje)
        {
            LblError.Visible = true;
            LblError.ForeColor = System.Drawing.Color.Red;
            LblError.Text = mensaje;
        }

        #endregion

        protected void cbCargarArchivo_Callback(object sender, DevExpress.Web.ASPxClasses.CallbackEventArgsBase e)
        {
            LblError.Visible = false;
            gvInconsistencias.Visible = false;
        }

        protected void btnCargarArchivo_Click(object sender, EventArgs e)
        {
            gvInconsistencias.Visible = false;
            gvDemandasNoCargadas.Visible = false;
            gvDemandasCargadas.Visible = false;
            gvErroresCarga.Visible = false;
            LblError.Visible = false;

            if (!uploadArchivoExcel.HasFile)
            {
                MostrarError("Seleccione el archivo a cargar.");
                return;
            }

            string FinalPath = "";
            try
            {
                String TempoPath = Server.MapPath("~/Temporal_Files/");
                if (!Directory.Exists(TempoPath))
                {
                    Directory.CreateDirectory(TempoPath);
                }
                FinalPath = TempoPath + Path.GetFileName(uploadArchivoExcel.FileName);

                // 1. Leer el archivo
                DataTable dtArchivoExcel = CargaArchivoExcelSpreadSheet(FinalPath);

                // 2. Quitar filas vacías guardando la fila real de Excel
                List<int> filasExcel;
                dtArchivoExcel = QuitarFilasVacias(dtArchivoExcel, out filasExcel);

                if (dtArchivoExcel.Rows.Count == 0)
                {
                    MostrarError("El archivo no contiene registros.");
                    return;
                }

                // 3. Validar que sea la plantilla correcta
                if (!compararEstructura(dtArchivoExcel))
                {
                    MostrarError("Ocurrió un error validando estructura del archivo. Favor utilizar la plantilla correcta.");
                    return;
                }

                // 4. Validar los datos. Si hay un solo error, NO se carga nada.
                DataTable errores = ValidarArchivo(dtArchivoExcel, filasExcel);
                if (errores.Rows.Count > 0)
                {
                    MostrarError("El archivo presenta inconsistencias, no se cargó ningún registro. Corrija los errores indicados e intente nuevamente.");
                    gvErroresCarga.DataSource = errores;
                    gvErroresCarga.DataBind();
                    gvErroresCarga.Visible = true;
                    return;
                }

                // 5. Sin errores: insertar
                CargarEnBaseDeDatos(dtArchivoExcel);
            }
            catch (Exception ex)
            {
                MostrarError("Ocurrió un error: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (FinalPath != "" && File.Exists(FinalPath)) File.Delete(FinalPath); //Se elimina el archivo temporal
                }
                catch { }
            }
        }

        /// <summary>
        /// Proceso original: tabla temporal + SP de validación/inserción.
        /// </summary>
        private void CargarEnBaseDeDatos(DataTable dtArchivoExcel)
        {
            string connection_String = ConfigurationManager.ConnectionStrings["BasePerdidasOperativas.Properties.Settings.SettingConnectionString"].ToString();
            using (SqlConnection sqlcone = new SqlConnection(connection_String))
            {
                sqlcone.Open();

                // Truncar la tabla antes de insertar
                using (SqlCommand cmd = new SqlCommand("TRUNCATE TABLE dbo.tblExcelDemandasJuiciosEjecucion", sqlcone))
                {
                    cmd.ExecuteNonQuery();
                }

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlcone))
                {
                    bulkCopy.DestinationTableName = "dbo.tblExcelDemandasJuiciosEjecucion";
                    bulkCopy.WriteToServer(dtArchivoExcel);
                }

                using (SqlCommand cmd = new SqlCommand("dbo.usp_validacionDemandasJuiciosEjecucion", sqlcone))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@usuario", Session["usuarioid"]);
                    SqlDataAdapter sqladapter = new SqlDataAdapter(cmd);
                    DataSet dataTableNuevo = new DataSet();
                    sqladapter.Fill(dataTableNuevo);

                    if (dataTableNuevo.Tables[0].Rows.Count > 0)
                    {
                        MostrarError("Se encontraron inconsistencias en el archivo, favor revisar validaciones no superadas.");
                        Session["gvInconsistencias"] = dataTableNuevo.Tables[0];
                        gvInconsistencias.DataBind();
                        gvInconsistencias.Visible = true;
                    }

                    if (dataTableNuevo.Tables[1].Rows.Count > 0)
                    {
                        Session["gvDemandasNoCargadas"] = dataTableNuevo.Tables[1];
                        gvDemandasNoCargadas.DataBind();
                        gvDemandasNoCargadas.Visible = true;
                    }

                    if (dataTableNuevo.Tables[2].Rows.Count > 0)
                    {
                        if (dataTableNuevo.Tables[0].Rows.Count == 0)
                        {
                            LblError.Visible = true;
                            LblError.ForeColor = System.Drawing.Color.Green;
                            LblError.Text = "Se insertaron Registros Exitosamente";
                        }

                        Session["gvDemandasCargadas"] = dataTableNuevo.Tables[2];
                        gvDemandasCargadas.DataBind();
                        gvDemandasCargadas.Visible = true;
                    }
                }
            }
        }

        protected void gvInconsistencias_DataBinding(object sender, EventArgs e)
        {
            if (Session["gvInconsistencias"] != null)
                gvInconsistencias.DataSource = (DataTable)Session["gvInconsistencias"];
            else
                gvInconsistencias.DataSource = null;
        }

        protected void gvDemandasNoCargadas_DataBinding(object sender, EventArgs e)
        {
            if (Session["gvDemandasNoCargadas"] != null)
                gvDemandasNoCargadas.DataSource = (DataTable)Session["gvDemandasNoCargadas"];
            else
                gvDemandasNoCargadas.DataSource = null;
        }

        protected void gvDemandasCargadas_DataBinding(object sender, EventArgs e)
        {
            if (Session["gvDemandasCargadas"] != null)
                gvDemandasCargadas.DataSource = (DataTable)Session["gvDemandasCargadas"];
            else
                gvDemandasCargadas.DataSource = null;
        }
    }
}
