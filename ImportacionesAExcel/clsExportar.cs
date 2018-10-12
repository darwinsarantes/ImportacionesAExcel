using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;

namespace ImportacionesAExcel
{
    public class clsExportar
    {
        private string Filename;
        private bool chkexcel;
        private Application oexcel;
        private Workbook obook;
        private Worksheet osheet;
        private System.Data.DataTable DT;
        private string prtTitulo;
        private string prtSubtitulo;
        private string prtNombreHoja;
        private string prtNombreDelSistema;
        private bool prtNombreSistema;
        private bool prtFecha;
        private bool prtHora;
        private bool prtLogo;
        private bool prtGraficoLinea;
        private string prtTituloGraficoLinea;
        private string prtTituloEjeX;
        private string prtTituloEjeY;
        private bool prtGraficoBarra;
        private string prtTituloGraficoBarra;
        private string prtTituloEjeX1;
        private string prtTituloEjeY1;
        private int[] prtEjes;
        private bool prtGraficoDeBarrasEjes;
        private string StartupPath;
        private string prtEmpresa_Nombre;
        private string prtEmpresa_Datos1;
        private string prtEmpresa_Datos2;
        private string prtRutaParaExportacion;
        private char PrimerColumna;
        private char PrimerColumnaGrafico;
        private short PrimerFila;
        private short PrimerFilaRegistros;
        private int Ncm;
        private int Nc;
        private string Rango;
        private string RangoCategorias;
        private string RangoGrafico;
        private int i;
        private int NFilas;
        private int NFilas_Ocultas;

        public clsExportar()
        {
            this.prtRutaParaExportacion = "";
            this.Ncm = 0;
            this.Rango = "";
            this.RangoCategorias = "";
            this.RangoGrafico = "";
            this.i = 0;
            this.NFilas = 0;
            this.NFilas_Ocultas = 0;
        }

        public System.Data.DataTable pty_DT
        {
            get
            {
                return this.DT;
            }
            set
            {
                this.DT = value;
            }
        }

        public string pty_prtTitulo
        {
            get
            {
                return this.prtTitulo;
            }
            set
            {
                this.prtTitulo = value;
            }
        }

        public string pty_prtSubtitulo
        {
            get
            {
                return this.prtSubtitulo;
            }
            set
            {
                this.prtSubtitulo = value;
            }
        }

        public string pty_prtNombreHoja
        {
            get
            {
                return this.prtNombreHoja;
            }
            set
            {
                this.prtNombreHoja = value;
            }
        }

        public string pty_prtNombreDelSistema
        {
            get
            {
                return this.prtNombreDelSistema;
            }
            set
            {
                this.prtNombreDelSistema = value;
            }
        }

        public bool pty_prtNombreSistema
        {
            get
            {
                return this.prtNombreSistema;
            }
            set
            {
                this.prtNombreSistema = value;
            }
        }

        public bool pty_prtFecha
        {
            get
            {
                return this.prtFecha;
            }
            set
            {
                this.prtFecha = value;
            }
        }

        public bool pty_prtHora
        {
            get
            {
                return this.prtHora;
            }
            set
            {
                this.prtHora = value;
            }
        }

        public bool pty_prtLogo
        {
            get
            {
                return this.prtLogo;
            }
            set
            {
                this.prtLogo = value;
            }
        }

        public bool pty_prtGraficoLinea
        {
            get
            {
                return this.prtGraficoLinea;
            }
            set
            {
                this.prtGraficoLinea = value;
            }
        }

        public string pty_prtTituloGraficoLinea
        {
            get
            {
                return this.prtTituloGraficoLinea;
            }
            set
            {
                this.prtTituloGraficoLinea = value;
            }
        }

        public string pty_prtTituloEjeX
        {
            get
            {
                return this.prtTituloEjeX;
            }
            set
            {
                this.prtTituloEjeX = value;
            }
        }

        public string pty_prtTituloEjeY
        {
            get
            {
                return this.prtTituloEjeY;
            }
            set
            {
                this.prtTituloEjeY = value;
            }
        }

        public bool pty_prtGraficoBarra
        {
            get
            {
                return this.prtGraficoBarra;
            }
            set
            {
                this.prtGraficoBarra = value;
            }
        }

        public string pty_prtTituloGraficoBarra
        {
            get
            {
                return this.prtTituloGraficoBarra;
            }
            set
            {
                this.prtTituloGraficoBarra = value;
            }
        }

        public string pty_prtTituloEjeX1
        {
            get
            {
                return this.prtTituloEjeX1;
            }
            set
            {
                this.prtTituloEjeX1 = value;
            }
        }

        public string pty_prtTituloEjeY1
        {
            get
            {
                return this.prtTituloEjeY1;
            }
            set
            {
                this.prtTituloEjeY1 = value;
            }
        }

        public int[] pty_prtEjes
        {
            get
            {
                return this.prtEjes;
            }
            set
            {
                this.prtEjes = value;
            }
        }

        public bool pty_prtGraficoDeBarrasEjes
        {
            get
            {
                return this.prtGraficoDeBarrasEjes;
            }
            set
            {
                this.prtGraficoDeBarrasEjes = value;
            }
        }

        public string pty_StartupPath
        {
            get
            {
                return this.StartupPath;
            }
            set
            {
                this.StartupPath = value;
            }
        }

        public string pty_Empresa_Nombre
        {
            get
            {
                return this.prtEmpresa_Nombre;
            }
            set
            {
                this.prtEmpresa_Nombre = value;
            }
        }

        public string pty_Empresa_Datos1
        {
            get
            {
                return this.prtEmpresa_Datos1;
            }
            set
            {
                this.prtEmpresa_Datos1 = value;
            }
        }

        public string pty_Empresa_Datos2
        {
            get
            {
                return this.prtEmpresa_Datos2;
            }
            set
            {
                this.prtEmpresa_Datos2 = value;
            }
        }

        public string pty_RutaParaExportacion
        {
            get
            {
                return this.prtRutaParaExportacion;
            }
            set
            {
                this.prtRutaParaExportacion = value;
            }
        }

        public void CrearLibro()
        {
            Random random = new Random();
            this.Filename = this.prtRutaParaExportacion.Trim().Length > 0 ? this.prtRutaParaExportacion : "C:\\Exportacion_Excel";
            if (!Directory.Exists(this.Filename))
                Directory.CreateDirectory(this.Filename);
            this.Filename = this.Filename + "\\" + this.prtTitulo + " " + DateAndTime.DatePart(DateInterval.Day, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1).ToString() + DateAndTime.Month(DateTime.Now).ToString() + DateAndTime.Year(DateTime.Now).ToString() + DateAndTime.Hour(DateTime.Now).ToString() + DateAndTime.Minute(DateTime.Now).ToString() + DateAndTime.Second(DateTime.Now).ToString() + ".xlsx";
            if (File.Exists(this.Filename))
                File.Delete(this.Filename);
            if (File.Exists(this.Filename))
                return;
            this.chkexcel = false;
            this.oexcel = (Application)Interaction.CreateObject("Excel.Application", "");
            // ISSUE: reference to a compiler-generated method
            this.obook = this.oexcel.Workbooks.Add((object)Missing.Value);
            this.oexcel.Application.DisplayAlerts = true;
            this.chkexcel = true;
        }

        public void Configuracion_inicial_Excel()
        {
            this.osheet = (Worksheet)this.oexcel.Worksheets[(object)1];
            //this.osheet.Name = this.prtNombreHoja;
            this.osheet.Name = "Hoja1";
            this.osheet.PageSetup.Zoom = (object)100;
            this.PrimerColumna = 'A';
            this.PrimerColumnaGrafico = 'B';
            this.Nc = this.DT.Columns.Count;
            if (this.prtLogo)
            {
                this.PrimerFila = (short)6;
                this.PrimerFilaRegistros = (short)10;
            }
            else
            {
                this.PrimerFila = (short)1;
                this.PrimerFilaRegistros = (short)5;
            }
            try
            {
                foreach (System.Data.DataColumn column in (System.Data.InternalDataCollectionBase)this.DT.Columns)
                {
                    if (!column.ColumnName.ToString().Contains("EECIIS"))
                    {
                        this.Rango = this.nombreColumna(checked(this.Ncm + 1)) + checked((int)this.PrimerFila + 4).ToString();
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)column.ColumnName.ToString());
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).EntireColumn.AutoFit();
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).BorderAround((object)8, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, (object)Missing.Value);
                        checked { ++this.Ncm; }
                    }
                }
            }
            finally
            {
                IEnumerator enumerator = null;
                if (enumerator is IDisposable)
                    (enumerator as IDisposable).Dispose();
            }

            this.Rango = Conversions.ToString(this.PrimerColumna) + this.PrimerFila.ToString() + ":" + this.nombreColumna(this.Ncm) + this.PrimerFila.ToString();
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)14;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)true;
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Merge((object)Missing.Value);
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)this.prtTitulo);
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignCenter;
            checked { ++this.PrimerFila; }
            this.Rango = Conversions.ToString(this.PrimerColumna) + this.PrimerFila.ToString() + ":" + this.nombreColumna(this.Ncm) + this.PrimerFila.ToString();
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)12;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)true;
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Merge((object)Missing.Value);
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)this.prtSubtitulo);
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignCenter;
            checked { this.PrimerFila += (short)3; }
            this.Rango = Conversions.ToString(this.PrimerColumna) + this.PrimerFila.ToString() + ":" + this.nombreColumna(this.Ncm) + this.PrimerFila.ToString();
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.ColorIndex = (object)0;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)true;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)10;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignCenter;
            if (!this.prtLogo)
                return;
            if (Directory.Exists(this.StartupPath + "\\Logos") && File.Exists(this.StartupPath + "\\Logos\\Logo_Empresa.png"))
            {
                // ISSUE: reference to a compiler-generated method
                this.osheet.Shapes.AddPicture(this.StartupPath + "\\Logos\\Logo_Empresa.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0.0f, 0.0f, 60f, 69f);
            }
            this.Rango = Conversions.ToString(this.PrimerColumna) + "2";
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)15;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)true;
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)("                    " + this.prtEmpresa_Nombre));
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignLeft;
            this.Rango = Conversions.ToString(this.PrimerColumna) + "3";
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)10;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)false;
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)("                                  " + this.prtEmpresa_Datos1));
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignLeft;
            this.Rango = Conversions.ToString(this.PrimerColumna) + "4";
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)9;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)false;
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)("                                  " + this.prtEmpresa_Datos2));
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignLeft;
        }

        public void Exportar_a_Excel(System.Data.DataRow DR)
        {
            checked { ++this.NFilas; }
            this.i = 0;
            int num = 0;
            checked { ++this.PrimerFila; }
            this.NFilas_Ocultas = 0;
            try
            {
                foreach (System.Data.DataColumn column in ( System.Data.InternalDataCollectionBase)this.DT.Columns)
                {
                    if (column.ColumnName.ToString().Contains("AgruparTitulo") && Microsoft.VisualBasic.CompilerServices.Operators.CompareString(DR[this.i].ToString().Trim(), "1", false) == 0)
                    {
                        this.Rango = this.nombreColumna(checked(this.i + 1 - this.NFilas_Ocultas)) + checked((int)this.PrimerFilaRegistros - 1).ToString() + ":" + this.nombreColumna(checked(this.i + 1 - this.NFilas_Ocultas + Conversions.ToInteger(DR[this.i + 1]) - 1)) + checked((int)this.PrimerFilaRegistros - 1).ToString();
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.ColorIndex = (object)0;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)true;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)10;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignCenter;
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Merge((object)Missing.Value);
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)DR[checked(this.i + 2)].ToString());
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).BorderAround((object)8, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, (object)Missing.Value);
                        column.ColumnName = Strings.Replace(column.ColumnName, "AgruparTitulo", "AgruparYATitulo", 1, -1, CompareMethod.Binary);
                    }
                    if (column.ColumnName.ToString().Contains("EECIIS"))
                        checked { ++this.NFilas_Ocultas; }
                    this.Rango = this.nombreColumna(checked(num + 1)) + this.PrimerFila.ToString();
                    // ISSUE: reference to a compiler-generated method
                    // ISSUE: reference to a compiler-generated method
                    if (!column.ColumnName.ToString().Contains("EECIIS") && Information.IsNothing(RuntimeHelpers.GetObjectValue(this.osheet.get_Range((object)this.Rango, (object)Missing.Value).get_Value((object)Missing.Value))))
                    {
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)DR[this.i].ToString().Trim());
                    }
                    if (this.i == 0 & (Microsoft.VisualBasic.CompilerServices.Operators.CompareString(DR[this.i].ToString(), "1", false) == 0 | Microsoft.VisualBasic.CompilerServices.Operators.CompareString(DR[this.i].ToString(), "2", false) == 0) & column.ColumnName.ToString().Contains("EECIIS"))
                    {
                        // ISSUE: reference to a compiler-generated method
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)DR[checked(this.i + 1)].ToString());
                    }
                    if (((num >= this.Ncm || Microsoft.VisualBasic.CompilerServices.Operators.CompareString(DR[this.i].ToString(), "1", false) != 0 || this.i != 0 ? 0 : 1) & (column.ColumnName.ToString().Contains("EECIIS") ? 1 : 0)) != 0)
                    {
                        this.RangoCategorias = this.nombreColumna(checked(this.i + 1)) + this.PrimerFila.ToString() + ":" + this.nombreColumna(this.Ncm).ToString() + this.PrimerFila.ToString();
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.ColorIndex = (object)11;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.Bold = (object)true;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.Size = (object)10;
                    }
                    if (((num >= this.Ncm || Microsoft.VisualBasic.CompilerServices.Operators.CompareString(DR[this.i].ToString(), "2", false) != 0 || this.i != 0 ? 0 : 1) & (column.ColumnName.ToString().Contains("EECIIS") ? 1 : 0)) != 0)
                    {
                        this.RangoCategorias = this.nombreColumna(checked(this.i + 1)) + this.PrimerFila.ToString() + ":" + this.nombreColumna(this.Ncm).ToString() + this.PrimerFila.ToString();
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.ColorIndex = (object)30;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.Bold = (object)true;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.Size = (object)10;
                        // ISSUE: reference to a compiler-generated method
                        this.osheet.get_Range((object)this.RangoCategorias, (object)Missing.Value).Font.Italic = (object)true;
                    }
                    if (!column.ColumnName.ToString().Contains("EECIIS"))
                        checked { ++num; }
                    checked { ++this.i; }
                }
            }
            finally
            {
                System.Collections.IEnumerator enumerator = null;
                if (enumerator is IDisposable)
                    (enumerator as IDisposable).Dispose();
            }
        }

        public void Configuracion_final_Excel()
        {
            this.Rango = Conversions.ToString(this.PrimerColumna) + this.PrimerFilaRegistros.ToString() + ":" + this.nombreColumna(this.Ncm) + this.PrimerFila.ToString();
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Columns.BorderAround((object)1, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, (object)Missing.Value);
            checked { ++this.PrimerFila; }
            this.Rango = "B" + this.PrimerFila.ToString() + ":" + this.nombreColumna(this.Ncm) + this.PrimerFila.ToString();
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Size = (object)9;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.Bold = (object)false;
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Merge((object)Missing.Value);
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).set_Value((object)Missing.Value, (object)(Interaction.IIf(this.prtFecha, (object)Strings.Format((object)DateAndTime.Now.Date, "dd/MM/yyyy"), (object)"").ToString() + "   " + Interaction.IIf(this.prtHora, (object)DateTime.Now.ToLongTimeString().ToString(), (object)"").ToString() + Interaction.IIf(this.prtNombreSistema, (object)("   " + this.prtNombreDelSistema), (object)"").ToString()));
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).EntireColumn.AutoFit();
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).HorizontalAlignment = (object)XlHAlign.xlHAlignRight;
            // ISSUE: reference to a compiler-generated method
            this.osheet.get_Range((object)this.Rango, (object)Missing.Value).Font.ColorIndex = (object)0;
            this.oexcel.WindowState = XlWindowState.xlMaximized;
            this.oexcel.Visible = true;
            // ISSUE: reference to a compiler-generated method
            this.osheet.Activate();
            // ISSUE: reference to a compiler-generated method
            this.obook.SaveAs((object)this.Filename, (object)Missing.Value, (object)Missing.Value, (object)Missing.Value, (object)Missing.Value, (object)Missing.Value, XlSaveAsAccessMode.xlNoChange, (object)Missing.Value, (object)Missing.Value, (object)Missing.Value, (object)Missing.Value, (object)Missing.Value);
            this.osheet = (Worksheet)null;
            this.oexcel.Visible = true;
        }

        private void LiberarDatos()
        {
            if (!this.chkexcel)
                return;
            this.osheet = (Worksheet)null;
            this.oexcel.Application.DisplayAlerts = false;
            // ISSUE: reference to a compiler-generated method
            this.obook.Close((object)Missing.Value, (object)Missing.Value, (object)Missing.Value);
            this.oexcel.Application.DisplayAlerts = true;
            this.obook = (Workbook)null;
            // ISSUE: reference to a compiler-generated method
            this.oexcel.Quit();
            this.oexcel = (Application)null;
        }

        private string nombreColumna(int numero)
        {
            return new string[257]
            {
        null,
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z",
        "AA",
        "AB",
        "AC",
        "AD",
        "AE",
        "AF",
        "AG",
        "AH",
        "AI",
        "AJ",
        "AK",
        "AL",
        "AM",
        "AN",
        "AO",
        "AP",
        "AQ",
        "AR",
        "AS",
        "AT",
        "AU",
        "AV",
        "AW",
        "AX",
        "AY",
        "AZ",
        "BA",
        "BB",
        "BC",
        "BD",
        "BE",
        "BF",
        "BG",
        "BH",
        "BI",
        "BJ",
        "BK",
        "BL",
        "BM",
        "BN",
        "BO",
        "BP",
        "BQ",
        "BR",
        "BS",
        "BT",
        "BU",
        "BV",
        "BW",
        "BX",
        "BY",
        "BZ",
        "CA",
        "CB",
        "CC",
        "CD",
        "CE",
        "CF",
        "CG",
        "CH",
        "CI",
        "CJ",
        "CK",
        "CL",
        "CM",
        "CN",
        "CO",
        "CP",
        "CQ",
        "CR",
        "CS",
        "CT",
        "CU",
        "CV",
        "CW",
        "CX",
        "CY",
        "CZ",
        "DA",
        "DB",
        "DC",
        "DD",
        "DE",
        "DF",
        "DG",
        "DH",
        "DI",
        "DJ",
        "DK",
        "DL",
        "DM",
        "DN",
        "DO",
        "DP",
        "DQ",
        "DR",
        "DS",
        "DT",
        "DU",
        "DV",
        "DW",
        "DX",
        "DY",
        "DZ",
        "EA",
        "EB",
        "EC",
        "ED",
        "EE",
        "EF",
        "EG",
        "EH",
        "EI",
        "EJ",
        "EK",
        "EL",
        "EM",
        "EN",
        "EO",
        "EP",
        "EQ",
        "ER",
        "ES",
        "ET",
        "EU",
        "EV",
        "EW",
        "EX",
        "EY",
        "EZ",
        "FA",
        "FB",
        "FC",
        "FD",
        "FE",
        "FF",
        "FG",
        "FH",
        "FI",
        "FJ",
        "FK",
        "FL",
        "FM",
        "FN",
        "FO",
        "FP",
        "FQ",
        "FR",
        "FS",
        "FT",
        "FU",
        "FV",
        "FW",
        "FX",
        "FY",
        "FZ",
        "GA",
        "GB",
        "GC",
        "GD",
        "GE",
        "GF",
        "GG",
        "GH",
        "GI",
        "GJ",
        "GK",
        "GL",
        "GM",
        "GN",
        "GO",
        "GP",
        "GQ",
        "GR",
        "GS",
        "GT",
        "GU",
        "GV",
        "GW",
        "GX",
        "GY",
        "GZ",
        "HA",
        "HB",
        "HC",
        "HD",
        "HE",
        "HF",
        "HG",
        "HH",
        "HI",
        "HJ",
        "HK",
        "HL",
        "HM",
        "HN",
        "HO",
        "HP",
        "HQ",
        "HR",
        "HS",
        "HT",
        "HU",
        "HV",
        "HW",
        "HX",
        "HY",
        "HZ",
        "IA",
        "IB",
        "IC",
        "ID",
        "IE",
        "IF",
        "IG",
        "IH",
        "II",
        "IJ",
        "IK",
        "IL",
        "IM",
        "IN",
        "IO",
        "IP",
        "IQ",
        "IR",
        "IS",
        "IT",
        "IU",
        "IV"
            }[numero];
        }
    }
}
