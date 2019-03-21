using CapaNegocio;
using CarlosAg.ExcelXmlWriter;
using Entidades;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace CapaPresentacion
{
    public partial class frmIndicadorProducto : Form
    {
        public frmIndicadorProducto()
        {
            InitializeComponent();
        }
        negProducto oProducto = new negProducto();
        private void CrearGrid()
        {
            try
            {
                dgvProducto.Columns.Add("ColumnId", "Id");
                dgvProducto.Columns.Add("ColumnCodigo", "Código");
                dgvProducto.Columns.Add("ColumnNombre", "Nombre");
                dgvProducto.Columns.Add("ColumnPrecioCompra", "P Compra");
                dgvProducto.Columns.Add("ColumnPrecio", "P Venta");
                dgvProducto.Columns.Add("ColumnStock", "Stock");
                dgvProducto.Columns.Add("ColumnUniMedida", "Categoria");
                dgvProducto.Columns.Add("ColumnMaterila", "Uni. Medida");
                dgvProducto.Columns.Add("ColumnUniMedida", "Material");
                DataGridViewImageColumn dgvImagenColumn = new DataGridViewImageColumn();
                dgvImagenColumn.HeaderText = "Estado";
                dgvImagenColumn.Name = "ColumnEstado";
                dgvProducto.Columns.Add(dgvImagenColumn);

                dgvProducto.Columns[0].Visible = false;
                dgvProducto.Columns[2].Width = 300;
                dgvProducto.Columns[3].Width = 88;
                dgvProducto.Columns[4].Width =88;
                dgvProducto.Columns[5].Width = 88;
                dgvProducto.Columns[6].Width = 100;
                dgvProducto.Columns[7].Width = 100;
                dgvProducto.Columns[8].Width = 110;
                dgvProducto.Columns[9].Width = 55;
               


                DataGridViewCellStyle cssCabecera = new DataGridViewCellStyle();
                cssCabecera.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvProducto.ColumnHeadersDefaultCellStyle = cssCabecera;

                dgvProducto.AllowUserToAddRows = false;
                dgvProducto.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgvProducto.AllowUserToResizeColumns = false;
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void LlenarCombo() {
            try
            {
                cboCategoria.ValueMember = "Id_Cat";
                cboCategoria.DisplayMember = "Nombre_Cat";
                cboCategoria.DataSource = negProducto.Instancia.ListarCategoria();
                // cargar estado rbstock
                rbStock.Checked = true;
            }
            catch (Exception)
            {

                throw;
            }
        }
        
        private void LlenarGrid(String name) {
            try
            {
                int cat = 0;
                if (name==null) {
                    if (cboCategoria.SelectedValue == null) throw new ApplicationException("Debe seleccionar una categoria");
                }
                if (!String.IsNullOrEmpty(Convert.ToString(cboCategoria.SelectedValue))) cat = (int)cboCategoria.SelectedValue;
                dgvProducto.Rows.Clear();
                int rango = 0; Image img = null;
                if (rbStock.Checked == true) rango = 3;
                else if (rbStockPromedio.Checked == true) rango = 2;
                else if (rbStcokMin.Checked == true) rango = 1;
                else if (rbstockcero.Checked == true) rango = 0;
                List<entProducto> Lista = null;
                if (name == null) { Lista = negProducto.Instancia.ListarProductoIndicador(txtCodigo.Text,(int)cat , rango); }
                else { Lista = negProducto.Instancia.BuscarprodAvanzadaIndicador(name); }
                for (int i = 0; i < Lista.Count; i++)
                {
                    if (Lista[i].Stock_Prod >= 0 && Lista[i].Stock_Prod <= Lista[i].StockMin_Prod)
                    {
                        img = Properties.Resources.circulorojo_24x24;
                    }//Image.FromFile(Path.Combine(Application.StartupPath, "ImgAplicacion\\circulorojo_24x24.png")); }
                    else if (Lista[i].Stock_Prod > Lista[i].StockMin_Prod && Lista[i].Stock_Prod <= Lista[i].StockProm_Prod)
                    {
                        img = Properties.Resources.CirculoNaranja24x24;  //Image.FromFile(Path.Combine(Application.StartupPath, "ImgAplicacion\\CirculoNaranja24x24.png"));
                    }
                    else if (Lista[i].Stock_Prod > Lista[i].StockProm_Prod)
                    {
                        img = Properties.Resources.circulo_verde24x24; //Image.FromFile(Path.Combine(Application.StartupPath, "ImgAplicacion\\circulo_verde24x24.png"));
                    }
                    String[] fila = new String[] { Lista[i].Id_Prod.ToString(),Lista[i].Codigo_Prod,Lista[i].Nombre_Prod,Lista[i].PrecioCompra_Prod.ToString(),Lista[i].Precio_Prod.ToString(),
                    Lista[i].Stock_Prod.ToString(),Lista[i].categoria.Nombre_Cat,Lista[i].unidmedida.Abreviatura_Umed,Lista[i].material.Nombre};
                    dgvProducto.Rows.Add(fila);
                    dgvProducto.Rows[i].Cells[9].Value = img;

                }
               
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void frmIndicadorProducto_Load(object sender, EventArgs e)
        {
            try
            {
                CrearGrid();
                LlenarCombo();
            }
            catch (ApplicationException ae) { MessageBox.Show(ae.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnListar_Click(object sender, EventArgs e)
        {
            try
            {
                LlenarGrid(null);
            }
            catch (ApplicationException ae) { MessageBox.Show(ae.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtNombreProd_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                LlenarGrid(txtNombreProd.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnImprimirExcel_Click(object sender, EventArgs e)
        {
            try
            {

                DataSet table = new DataSet();
                table = oProducto.ListarProductoExcel();
          
                ExportToSpreadsheet(table);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void ExportToSpreadsheet(DataSet dsExcel)
        {
            try
            {
                   CarlosAg.ExcelXmlWriter.Workbook book = new CarlosAg.ExcelXmlWriter.Workbook();

                //Crear propiedades del excel
                CrearPropiedades(book);
               //Add styles to the workbook
                GenerarEstilos(book.Styles);

                // Add a Worksheet with some data
                GenerarHojas(book.Worksheets, dsExcel);

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel|*.xls";
                saveFileDialog1.Title = "Save an Excel File";
                saveFileDialog1.FileName = DateTime.Now.ToString("yyyy-MM-dd") + " PRODUCTO";
                saveFileDialog1.ShowDialog();

                if (saveFileDialog1.FileName != string.Empty)
                {
                    book.Save(saveFileDialog1.FileName);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }           
        }
        private void GenerarHojas(WorksheetCollection sheets, DataSet data)
        {
         
            for (int i=0; i< data.Tables.Count; i++)
            {
               
                DataTable rtRedumen = data.Tables[i];
               GenerarHoja(sheets, data.Tables[i].TableName, rtRedumen);
            }
        }
        private void GenerarHoja(WorksheetCollection sheets, string NombHoja,DataTable rtRedumen)
        {
            switch(NombHoja)
            {
                case "Table":
                    NombHoja = "PRODUCTO PRECIO MAXIMO";
                    break;
                case "Table1":
                    NombHoja = "PRODUCTO PRECIO MINIMO";
                    break;
           }
            Worksheet sheet = sheets.Add(NombHoja);
            sheet.Table.DefaultRowHeight = 15;
            sheet.Table.DefaultColumnWidth = 66;
            sheet.Table.FullColumns = 1;
            sheet.Table.FullRows = 1;
            sheet.Table.StyleID = "s62";
            sheet.Table.Columns.Add(50);
            sheet.Table.Columns.Add(65);
            sheet.Table.Columns.Add(250);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            sheet.Table.Columns.Add(100);
            GenerarEncabezadoDatosEnHoja(sheet);
             GenerarDatosEnHoja(sheet, rtRedumen);
                  
        }
        private void CrearPropiedades(Workbook book)
        {
            book.Properties.Author = "Aplicativo 911";
            book.Properties.LastAuthor = "Aplicativo 911";
            book.Properties.Created = DateTime.Now;
            book.Properties.LastSaved = DateTime.Now;
            book.Properties.Company = "Banco de Crédito";
            book.Properties.Version = "12.00";

            book.ExcelWorkbook.WindowHeight = 9150;
            book.ExcelWorkbook.WindowWidth = 16620;
            book.ExcelWorkbook.WindowTopX = 360;
            book.ExcelWorkbook.WindowTopY = 45;
            book.ExcelWorkbook.ProtectWindows = true;
            book.ExcelWorkbook.ProtectStructure = true;
        }
        private void GenerarEncabezadoDatosEnHoja(Worksheet sheet )
        {
            WorksheetColumn column0 = sheet.Table.Columns.Add();
            column0.Width = 75;
            column0.StyleID = "s62";
            WorksheetColumn column1 = sheet.Table.Columns.Add();
            column1.Width = 75;
            column1.StyleID = "s62";
            WorksheetColumn column2 = sheet.Table.Columns.Add();
            column2.Width = 75;
            column2.StyleID = "s62";
            WorksheetColumn column3 = sheet.Table.Columns.Add();
            column3.Width = 75;
            column3.StyleID = "s62";
            WorksheetColumn column4 = sheet.Table.Columns.Add();
            column4.Width = 75;
            column4.StyleID = "s62";

            WorksheetRow Row0 = sheet.Table.Rows.Add();
            Row0.AutoFitHeight = false;
            WorksheetRow Row1 = sheet.Table.Rows.Add();
            Row1.Height = 27;
            Row1.AutoFitHeight = false;
            Row1.Cells.Add("FARMACIA AGRICOLA >> TIERRA SANA >>", DataType.String, "s65");

            WorksheetRow Row2 = sheet.Table.Rows.Add();
            Row2.AutoFitHeight = true;
            Row2.Cells.Add("SUCURSAL", DataType.String, "s63");
            WorksheetCell cell;
            cell = Row2.Cells.Add();
            cell.Data.Type = DataType.String;
            cell.Data.Text = "SUCURSAL NUMERO N° 1";

            string date = Convert.ToString(DateTime.Now);
            WorksheetRow Row3 = sheet.Table.Rows.Add();
            Row3.AutoFitHeight = true;
            Row3.Cells.Add("FECHA", DataType.String, "s63");
            cell = Row3.Cells.Add();
            cell.Data.Type = DataType.String;
            cell.Data.Text = date.ToString();

            WorksheetRow Row4 = sheet.Table.Rows.Add();
            Row4.AutoFitHeight = true;
            Row4.Cells.Add("MONEDA", DataType.String, "s63");
            cell = Row4.Cells.Add();
            cell.Data.Type = DataType.String;
            cell.Data.Text = "SOLES";
            WorksheetRow Row8 = sheet.Table.Rows.Add();
            Row8.AutoFitHeight = true;
            Row8.Height = 20;
            Row8.Cells.Add("N°", DataType.String, "s84");
            Row8.Cells.Add("CODIGO", DataType.String, "s84");          
            Row8.Cells.Add("NOMBRE PRODUCTO", DataType.String, "s84");
            Row8.Cells.Add("PRECIO COMPRA", DataType.String, "s84");
            Row8.Cells.Add("PRECIO PRODUCTO", DataType.String, "s84");
            Row8.Cells.Add("STOCK", DataType.String, "s84");
            Row8.Cells.Add("STOCK PROMEDIO", DataType.String, "s84");
            Row8.Cells.Add("STOCK MINIMO", DataType.String, "s84");
            Row8.Cells.Add("CATEGORIA", DataType.String, "s84");
            Row8.Cells.Add("UNIDAD", DataType.String, "s84");
            Row8.Cells.Add("ALMACEN", DataType.String, "s84");

        

            //WorksheetRow Row10 = sheet.Table.Rows.Add();
            //Row10.AutoFitHeight = true;
            //Row10.Cells.Add("MONEDA s87", DataType.String, "s87");
            //cell = Row10.Cells.Add();
            //cell.Data.Type = DataType.String;
            //cell.Data.Text = "SOLES s87";

            //WorksheetRow Row11 = sheet.Table.Rows.Add();
            //Row11.AutoFitHeight = true;
            //Row11.Cells.Add("MONEDA s88", DataType.String, "s88");
            //cell = Row11.Cells.Add();
            //cell.Data.Type = DataType.String;
            //cell.Data.Text = "SOLES s88";

            //WorksheetRow Row12 = sheet.Table.Rows.Add();
            //Row12.AutoFitHeight = true;
            //Row12.Cells.Add("MONEDA s89", DataType.String, "s89");
            //cell = Row12.Cells.Add();
            //cell.Data.Type = DataType.String;
            //cell.Data.Text = "2";

            //WorksheetRow Row9 = sheet.Table.Rows.Add();
            //Row9.AutoFitHeight = true;
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";
            //cell = Row9.Cells.Add();
            //cell.StyleID = "s85";

        }
        private void GenerarEstilos(WorksheetStyleCollection styles)
        {
            //[Default]
            WorksheetStyle Default = styles.Add("Default");
            Default.Name = "Normal";
            Default.Font.FontName = "Formata Regular";
            Default.Font.Color = "#000000";
            Default.Alignment.Vertical = StyleVerticalAlignment.Bottom;

            WorksheetStyle s16 = styles.Add("s16");
            s16.Name = "Millares";
            s16.NumberFormat = "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * " + "-" + "??_ ;_ @_ ";

            WorksheetStyle s31 = styles.Add("s31");
            s31.Font.FontName = "Tahoma";
            s31.Font.Size = 8;
            s31.Font.Color = "#000000";
            s31.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s31.Alignment.Vertical = StyleVerticalAlignment.Center;
            s31.Alignment.WrapText = true;
           
            s31.NumberFormat = "@";

            WorksheetStyle s32 = styles.Add("s32");
            s32.Font.FontName = "Formata Regular";
            s32.NumberFormat = "#####";
            s32.Font.Color = "#000000";
            s32.Alignment.Vertical = StyleVerticalAlignment.Center;
           

            WorksheetStyle s62 = styles.Add("s62");
            s62.Alignment.Vertical = StyleVerticalAlignment.Center;

            WorksheetStyle s63 = styles.Add("s63");
            s63.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s63.Alignment.Vertical = StyleVerticalAlignment.Center;

            WorksheetStyle s65 = styles.Add("s65");
            s65.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s65.Alignment.Vertical = StyleVerticalAlignment.Center;

            WorksheetStyle s84 = styles.Add("s84");
            s84.Interior.Color = "#BFBFBF";
            s84.Interior.Pattern = StyleInteriorPattern.Solid;
            s84.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s84.Alignment.Vertical = StyleVerticalAlignment.Center;
            s84.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s84.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s84.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s84.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);

            WorksheetStyle s85 = styles.Add("s85");
            s85.Alignment.Vertical = StyleVerticalAlignment.Center;
            s85.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s85.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s85.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s85.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);

            WorksheetStyle s86 = styles.Add("s86");
            s86.Parent = "s16";
            s86.Font.FontName = "Formata Regular";
            s86.Font.Color = "#000000";
            s86.Alignment.Vertical = StyleVerticalAlignment.Center;
;


            WorksheetStyle s87 = styles.Add("s87");
            s87.Font.FontName = "Formata Regular";
            s87.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            s87.Alignment.Vertical = StyleVerticalAlignment.Center;
            s87.Alignment.WrapText = true;



            //WorksheetStyle s88 = styles.Add("s88");
            //s88.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            //s88.Alignment.Vertical = StyleVerticalAlignment.Center;
            //s88.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            //s88.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            //s88.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            //s88.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            //s88.NumberFormat = "@";

            //WorksheetStyle s89 = styles.Add("s89");
            //s89.Parent = "s16";
            //s89.Font.FontName = "Formata Regular";
            //s89.Font.Color = "#000000";
            //s89.Alignment.Vertical = StyleVerticalAlignment.Center;
            //s89.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            //s89.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            //s89.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            //s89.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
            //s89.NumberFormat = "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * " + "-" + "??_ ;_ @_ ";


            WorksheetStyle s94 = styles.Add("s94");
            s94.Alignment.Horizontal = StyleHorizontalAlignment.Left;
            s94.Alignment.Vertical = StyleVerticalAlignment.Center;
            s94.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            s94.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            s94.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            s94.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);
        }
        private void GenerarDatosEnHoja(Worksheet sheet, DataTable dtResumen)
        {

            decimal tTotal = 0;
            //Loop through each table in dataset
            foreach (DataRow row in dtResumen.Rows)
            {
                WorksheetRow Row9 = sheet.Table.Rows.Add();               
                Row9.AutoFitHeight = true;
                Row9.Cells.Add(row["Id_Prod"].ToString(), DataType.Number, "s32");
                Row9.Cells.Add(row["CODIGO_PRODUCTO"].ToString(), DataType.String, "s87");         
                Row9.Cells.Add(row["NOMBRE_PRODUCTO"].ToString(), DataType.String, "s87");
                Row9.Cells.Add(row["PRECIO_COMPRA"].ToString(), DataType.Number, "s86");
                Row9.Cells.Add(row["PRECIO_PRODUCTO"].ToString(), DataType.Number, "s86");
                Row9.Cells.Add(row["STOCK"].ToString(), DataType.Number, "s32");
                Row9.Cells.Add(row["STOCK_PROMEDIO"].ToString(), DataType.Number, "s32");
                Row9.Cells.Add(row["STOCK_MONIMO"].ToString(), DataType.Number, "s32");
                Row9.Cells.Add(row["CATEGORIA"].ToString(), DataType.String, "s87");
                Row9.Cells.Add(row["UNIDAD"].ToString(), DataType.String, "s31");
                Row9.Cells.Add(row["ALMACEN"].ToString(), DataType.String, "s31");
                tTotal += (decimal)row["PRECIO_COMPRA"];
            }
            if (tTotal > 0)
            {
                WorksheetRow Row9 = sheet.Table.Rows.Add();
                Row9.Cells.Add("", DataType.Number, "s32");
                Row9.Cells.Add("", DataType.String, "s87");
                Row9.Cells.Add("TOTAL", DataType.String, "s87");
                Row9.Cells.Add(tTotal.ToString(), DataType.Number, "s86");
                Row9.Cells.Add("", DataType.Number, "s86");
                Row9.Cells.Add("", DataType.Number, "s32");
                Row9.Cells.Add("", DataType.Number, "s32");
                Row9.Cells.Add("", DataType.Number, "s32");
                Row9.Cells.Add("", DataType.String, "s87");
                Row9.Cells.Add("", DataType.String, "s31");
                Row9.Cells.Add("", DataType.String, "s31");
            }
                    
        }
    }
}

