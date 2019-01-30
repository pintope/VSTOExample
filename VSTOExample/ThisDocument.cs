namespace VSTOExample
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Configuration;
    using System.Collections.Generic;
    using Microsoft.SqlServer.Dac.Model;
    using Word = Microsoft.Office.Interop.Word;


    public partial class ThisDocument
    {
        /// <summary>
        /// Manejador del evento de apertura del documento.
        /// </summary>
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            this.ShowSpellingErrors    = false;
            this.ShowGrammaticalErrors = false;

            // La base de datos se va a leer desde un archivo ".dacpac".
            string dacpacFile = ConfigurationManager.AppSettings["dacpacFile"];

            DirectoryInfo appDir = new DirectoryInfo(Directory.GetCurrentDirectory());
            string solutionDir   = appDir.Parent.Parent.Parent.FullName;

            using (TSqlModel model = new TSqlModel(System.IO.Path.Combine(solutionDir, dacpacFile)))
            {
                // Obtener todas las propiedades de extensión de la base de datos.
                var allExtProperties = model.GetObjects(DacQueryScopes.All, ModelSchema.ExtendedProperty);

                InsertTables(model, allExtProperties);
                InsertFunctions(model, allExtProperties);
                InsertStoredProcedures(model, allExtProperties);
            }

            TablesOfContents[1].Update();
        }

        /// <summary>
        /// Inserta en el documento de Word las tablas de la base de datos.
        /// </summary>
        /// <param name="model">Base de datos como TSqlModel.</param>
        /// <param name="allExtProperties">Todas las propiedades de extensión de la base de datos.</param>
        private void InsertTables(TSqlModel model, IEnumerable<TSqlObject> allExtProperties)
        {
            // Busca en el documento la sección "Tablas", que tiene un marcador también así llamado.
            // Es donde se deben añadir todas las tablas correspondientes al modelo de datos.
            Word.Range range = this.Bookmarks["Tablas"].Range.Next();

            var records = new List<DocObject>();

            // Obtiene todas las tablas de la base de datos.
            var allDbTables = model.GetObjects(DacQueryScopes.All, ModelSchema.Table);

            foreach (var table in allDbTables) // Leer todas las tablas del modelo de datos.
            {
                string tableSchema = table.Name.Parts.First();

                // Incluye solo aquellas tablas que pertenecen al esquema "dbo".
                if (tableSchema != "dbo")
                {
                    continue;
                }

                // Nuevo registro de la tabla.
                var acDocObject = new DocObject
                {
                    Name = table.Name.Parts.Last(),
                    Type = DBObject.Table
                };

                // Propiedades de extensión para la tabla.
                var tableExtProperty = allExtProperties.Where(f => f.Name.ToString().Contains(acDocObject.Name));
                acDocObject.Description = GetExtPropertyDescription(new List<string>() { "SqlTableBase", tableSchema, acDocObject.Name }, tableExtProperty);

                foreach (var column in table.GetChildren().Where(child => child.ObjectType.Name == "Column"))
                {
                    // Restricción de clave primaria.
                    var pkConstraint = column.GetReferencing(PrimaryKeyConstraint.Columns, DacQueryScopes.UserDefined);

                    // Tipo del campo.
                    var fieldType = column.GetReferenced(Column.DataType).Any() ? column.GetReferenced(Column.DataType).First() : null;
                    var fieldName = column.Name.Parts.Last();

                    // Guarda la información del campo.
                    acDocObject.Columns.Add(new Tuple<bool, string, string, string>(
                        pkConstraint.Any(),                       // Es clave primaria o no.
                        fieldName,                                // Nombre del campo.
                        fieldType?.Name.Parts[0].ToUpper() ?? "", // Tipo del campo.
                        GetExtPropertyDescription(new List<string>() { "SqlColumn", tableSchema, acDocObject.Name, fieldName }, tableExtProperty)));
                }

                records.Add(acDocObject);
            }

            // Crea la tabla en base a la información obtenida desde la base de datos.
            range = PrintRecords(records, range, true);
            CommitRange(range);
            InsertSpace(range);
        }

        /// <summary>
        /// Inserta en el documento de Word las funciones.
        /// </summary>
        /// <param name="model">Base de datos como TSqlModel.</param>
        /// <param name="allExtProperties">Todas las propiedades de extensión de la base de datos.</param>
        private void InsertFunctions(TSqlModel model, IEnumerable<TSqlObject> allExtProperties)
        {
            // Busca en el documento la sección "Funciones", que tiene un marcador también así llamado.
            // Es donde se deben añadir todas las funciones correspondientes al modelo de datos.
            Word.Range range = this.Bookmarks["FuncionesTabulares"].Range.Next();

            // Funciones de tipo tabla-valor.
            var functions = model.GetObjects(DacQueryScopes.All, ModelSchema.TableValuedFunction).Where(f => !f.Name.ToString().StartsWith("[master]"));
            InsertFunctionType(range, functions, allExtProperties);

            // Busca en el documento la sección "Funciones", que tiene un marcador también así llamado.
            // Es donde se deben añadir todas las funciones correspondientes al modelo de datos.
            range = this.Bookmarks["FuncionesEscalares"].Range.Next();

            // Funciones de tipo escalar.
            functions = model.GetObjects(DacQueryScopes.All, ModelSchema.ScalarFunction).Where(f => !f.Name.ToString().StartsWith("[master]"));
            InsertFunctionType(range, functions, allExtProperties);
        }

        /// <summary>
        /// Inserta en el documento de Word las funciones indicadas.
        /// </summary>
        /// <param name="range">Rango del documento.</param>
        /// <param name="model">Base de datos como TSqlModel.</param>
        /// <param name="functions">Todas las funciones como enumerable de objetos TSqlObject.</param>
        /// <param name="allExtProperties">Todas las propiedades de extensión de la base de datos.</param>
        public void InsertFunctionType(Word.Range range, IEnumerable<TSqlObject> functions, IEnumerable<TSqlObject> allExtProperties)
        {
            var records = new List<DocObject>();

            foreach (var function in functions) // Leer todas las funciones del modelo de datos.
            {
                string functionSchema = function.Name.Parts.First();

                // Nuevo registro de la tabla.
                var acDocObject = new DocObject
                {
                    Name = function.Name.Parts.Last(),
                    Type = DBObject.Function
                };

                var allFunctions = allExtProperties.Where(f => f.Name.ToString().Contains(acDocObject.Name));

                // Propiedades de extensión para la función.
                var functionExtProperty = allExtProperties.Where(f => f.Name.ToString().Contains(acDocObject.Name));
                acDocObject.Description = GetExtPropertyDescription(new List<string>() { "SqlSubroutineParameter", functionSchema, acDocObject.Name }, functionExtProperty);

                foreach (var parameter in function.GetChildren().Where(child => child.ObjectType.Name == "Parameter"))
                {
                    // Tipo del campo.
                    var parameterType = parameter.GetReferenced(Parameter.DataType).Any() ? parameter.GetReferenced(Parameter.DataType).First() : null;
                    var parameterName = parameter.Name.Parts.Last();

                    // Guarda la información del parámetro.
                    acDocObject.Columns.Add(new Tuple<bool, string, string, string>(
                        false,
                        parameterName,                                // Nombre del campo.
                        parameterType?.Name.Parts[0].ToUpper() ?? "", // Tipo del campo.
                        GetExtPropertyDescription(new List<string>() { "SqlSubroutineParameter", functionSchema, acDocObject.Name, parameterName }, allExtProperties)));
                }

                records.Add(acDocObject);
            }

            // Crea la tabla en base a la información obtenida desde la base de datos.
            range = PrintRecords(records, range);
            CommitRange(range);
            InsertSpace(range);
        }

        /// <summary>
        /// Inserta en el documento de Word los procedimientos almacesados.
        /// </summary>
        /// <param name="model">Base de datos como TSqlModel.</param>
        /// <param name="allExtProperties">Todas las propiedades de extensión de la base de datos.</param>
        private void InsertStoredProcedures(TSqlModel model, IEnumerable<TSqlObject> allExtProperties)
        {
            // Busca en el documento la sección "Procedimientos", que tiene un marcador también así llamado.
            // Es donde se deben añadir todos los  correspondientes al modelo de datos.
            Word.Range range = this.Bookmarks["Procedimientos"].Range.Next();

            var records = new List<DocObject>();

            // Obtener todas las tablas de la base de datos.
            var allProcedures = model.GetObjects(DacQueryScopes.All, ModelSchema.Procedure);

            foreach (var proc in allProcedures) // Leer todas las tablas del modelo de datos.
            {
                string procSchema = proc.Name.Parts.First();

                // Nuevo registro de la tabla.
                var acDocObject = new DocObject
                {
                    Name = proc.Name.Parts.Last(),
                    Type = DBObject.Table
                };

                var procExtProperties = allExtProperties.Where(f => f.Name.ToString().Contains(acDocObject.Name));
                acDocObject.Description = GetExtPropertyDescription(new List<string>() { "SqlProcedure", procSchema, acDocObject.Name }, procExtProperties);

                foreach (var parameter in proc.GetChildren().Where(child => child.ObjectType.Name == "Parameter"))
                {
                    var dataType = parameter.GetReferenced(Parameter.DataType).First(); // Obtiene el tipo del parametro.
                    var parameterName = parameter.Name.Parts.Last();

                    // Guarda la información del parámetro.
                    acDocObject.Columns.Add(new Tuple<bool, string, string, string>(
                        false,
                        parameterName,                           // Nombre del parámetro.
                        dataType?.Name.Parts[0].ToUpper() ?? "", // Tipo del parámetro.
                        GetExtPropertyDescription(new List<string>() { "SqlSubroutineParameter", procSchema, acDocObject.Name, parameterName }, allExtProperties)));
                }

                records.Add(acDocObject);
            }

            // Crea la tabla en base a la información obtenida desde la base de datos.
            range = PrintRecords(records, range);
            CommitRange(range);
            InsertSpace(range);
        }

        /// <summary>
        /// Crea la tabla en Word para los registros indicados.
        /// </summary>
        /// <param name="records">Registros a imprimir.</param>
        /// <param name="range">Rango donde se ubica la tabla.</param>
        /// <returns></returns>
        private Word.Range PrintRecords(List<DocObject> records, Word.Range range, bool isTable = false)
        {
            InsertSpace(range);

            // Calcular las dimensiones de la tabla.
            int numCols = 5;
            int numRows = 2; // Número total de columnas de la tabla. El valor inicial es 2 por la cabecera.
            foreach (var record in records)
            {
                numRows += Math.Max(2, record.Columns.Count);
            }

            // Añade la tabla y le da formato general.
            Word.Table table = range.Tables.Add(range, numRows, numCols);
            table.Range.Font.Size = 9;
            table.Rows.SetHeight(15, Word.WdRowHeightRule.wdRowHeightExactly);

            table.Columns[1].Width = 120; // Nombre de la tabla.
            table.Columns[2].Width = 20;  // Icono de clave primaria.
            table.Columns[3].Width = 70;  // Nombre del campo/parámetro.
            table.Columns[4].Width = 75;  // Tipo del campo/parámetro.
            table.Columns[5].Width = 200; // Descripción del campo/parámetro.

            // Inserta la cabecera de las columnas.
            const int verySmallFontSize = 8;
            const int smallFontSize = 11;
            const int bigFontSize = 14;

            // Campos/parámetros.
            table.Cell(1, 2).Range.Text = isTable ? "CAMPOS" : "PARÁMETROS";
            table.Cell(1, 2).Range.Font.Size = bigFontSize;
            table.Cell(1, 2).Merge(table.Cell(1, 5));
            table.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            SetCellBorder(table.Cell(1, 2), borderType: Word.WdBorderType.wdBorderBottom);
            SetCellBorder(table.Cell(1, 2), borderType: Word.WdBorderType.wdBorderTop);
            SetCellBorder(table.Cell(1, 2), borderType: Word.WdBorderType.wdBorderLeft);
            SetCellBorder(table.Cell(1, 2), borderType: Word.WdBorderType.wdBorderRight);

            // Nombre.
            table.Cell(2, 2).Range.Text = "Nombre";
            table.Cell(2, 2).Range.Font.Size = smallFontSize;
            table.Cell(2, 2).Merge(table.Cell(2, 3));
            table.Cell(2, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Cell(2, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray65;
            table.Cell(2, 2).Range.Font.Color = Word.WdColor.wdColorWhite;
            SetCellBorder(table.Cell(2, 2));
            SetCellBorder(table.Cell(2, 2), borderType: Word.WdBorderType.wdBorderLeft);

            // Tipo.
            table.Cell(2, 3).Range.Text = "Tipo";
            table.Cell(2, 3).Range.Font.Size = smallFontSize;
            table.Cell(2, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Cell(2, 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray65;
            table.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorWhite;
            SetCellBorder(table.Cell(2, 3));

            // Descipción.
            table.Cell(2, 4).Range.Text = "Descripción";
            table.Cell(2, 4).Range.Font.Size = smallFontSize;
            table.Cell(2, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Cell(2, 4).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray65;
            table.Cell(2, 4).Range.Font.Color = Word.WdColor.wdColorWhite;
            SetCellBorder(table.Cell(2, 4));
            SetCellBorder(table.Cell(2, 4), borderType: Word.WdBorderType.wdBorderRight);

            int row = 3;
            foreach (var record in records)
            {
                bool isLastRecord = record == records.Last();
                int cellHeight = Math.Max(2, record.Columns.Count);

                // Nombre de la tabla.
                table.Cell(row, 1).Range.Text = record.Name;
                table.Cell(row, 1).Range.Font.Size = bigFontSize;
                table.Cell(row, 1).HeightRule = Word.WdRowHeightRule.wdRowHeightAuto;
                table.Cell(row, 1).FitText = record.Name.Length > 15;
                table.Cell(row, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                SetCellBorder(table.Cell(row, 1), Word.WdLineStyle.wdLineStyleSingle, Word.WdBorderType.wdBorderRight);

                // Descripción de la tabla.
                table.Cell(row + 1, 1).Range.Text = record.Description;
                table.Cell(row + 1, 1).HeightRule = Word.WdRowHeightRule.wdRowHeightAuto;
                if (cellHeight > 2)
                {
                    table.Cell(row + 1, 1).Merge(table.Cell(row + cellHeight - 1, 1));
                }

                table.Cell(row + 1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                SetCellBorder(table.Cell(row + 1, 1), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);

                // Campos.
                int hOffset = 0;
                int vOffset = 0;
                for (; vOffset < record.Columns.Count; vOffset++)
                {
                    bool isLastField = vOffset == cellHeight - 1;

                    if (isTable)
                    {
                        if (record.Columns[vOffset].Item1) // El campo es clave primaria.
                        {
                            table.Cell(row + vOffset, 2).Range.InlineShapes.AddPicture("..\\..\\Resources\\key_16xLG.png");
                        }

                        table.Cell(row + vOffset, 2).Range.Font.Size = verySmallFontSize;
                        table.Cell(row + vOffset, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        table.Cell(row + vOffset, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        SetCellBorder(table.Cell(row + vOffset, 2), borderType: Word.WdBorderType.wdBorderLeft);
                        hOffset = 1;

                        if (isLastField)
                        {
                            SetCellBorder(table.Cell(row + vOffset, 2), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                        }
                    }
                    else // No hay necesidad de indicar la clave primaria, por lo que se fusionan las celdas.
                    {
                        table.Cell(row + vOffset, 2).Merge(table.Cell(row + vOffset, 3));
                    }

                    // Nombre del campo.
                    table.Cell(row + vOffset, 2 + hOffset).Range.Text = record.Columns[vOffset].Item2;
                    table.Cell(row + vOffset, 2 + hOffset).FitText = record.Columns[vOffset].Item2.Length > 15;
                    table.Cell(row + vOffset, 2 + hOffset).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    table.Cell(row + vOffset, 2 + hOffset).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;

                    if (!isTable)
                    {
                        SetCellBorder(table.Cell(row + vOffset, 2), borderType: Word.WdBorderType.wdBorderLeft);
                    }

                    if (isLastField)
                    {
                        SetCellBorder(table.Cell(row + vOffset, 2 + hOffset), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                    }

                    // Tipo del campo.
                    table.Cell(row + vOffset, 3 + hOffset).Range.Text = record.Columns[vOffset].Item3;
                    table.Cell(row + vOffset, 3 + hOffset).FitText = record.Columns[vOffset].Item3.Length > 10;
                    table.Cell(row + vOffset, 3 + hOffset).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    if (isLastField)
                    {
                        SetCellBorder(table.Cell(row + vOffset, 3 + hOffset), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                    }

                    // Descripción del campo.
                    table.Cell(row + vOffset, 4 + hOffset).Range.Text = record.Columns[vOffset].Item4;
                    table.Cell(row + vOffset, 4 + hOffset).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    SetCellBorder(table.Cell(row + vOffset, 4 + hOffset), borderType: Word.WdBorderType.wdBorderRight);
                    if (isLastField)
                    {
                        SetCellBorder(table.Cell(row + vOffset, 4 + hOffset), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                    }
                }

                // En caso de que haya un solo campo/parámetro o ninguno, hay que formatear la/s fila/s que 
                // corresponden al nombre y la descripción de la tabla/función/procedimiento.
                if (record.Columns.Count < 2)
                {
                    int numRowsToFormat = 2 - record.Columns.Count;
                    for (int i = 0; i < numRowsToFormat; i++)
                    {
                        if (!isTable)
                        {
                            table.Cell(row + i + vOffset, 2).Merge(table.Cell(row + i + vOffset, 3));
                        }

                        table.Cell(row + i + vOffset, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        SetCellBorder(table.Cell(row + i + vOffset, 2), borderType: Word.WdBorderType.wdBorderLeft);
                        table.Cell(row + i + vOffset, 2 + hOffset).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        SetCellBorder(table.Cell(row + i + vOffset, 4 + hOffset), Word.WdLineStyle.wdLineStyleSingle, Word.WdBorderType.wdBorderRight);

                        if (numRowsToFormat == 1 || i == 1) // Última iteración.
                        {
                            // Imprimir la línea divisoria entre registros.
                            SetCellBorder(table.Cell(row + i + vOffset, 2), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                            SetCellBorder(table.Cell(row + i + vOffset, 2 + hOffset), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                            SetCellBorder(table.Cell(row + i + vOffset, 3 + hOffset), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                            SetCellBorder(table.Cell(row + i + vOffset, 4 + hOffset), isLastRecord ? Word.WdLineStyle.wdLineStyleSingle : Word.WdLineStyle.wdLineStyleDot);
                        }
                    }
                }

                row += cellHeight;
            }

            return table.Range;
        }

        /// <summary>
        /// Consolida el último cambio en el rango. De no hacerlo, cambios ulteriores sobreescriben el rango.
        /// </summary>
        /// <param name="range">Rango.</param>
        private void CommitRange(Word.Range range)
        {
            range.SetRange(range.End, range.End);
        }

        /// <summary>
        /// Inserta espacio en el rango especificado.
        /// </summary>
        /// <param name="range">Rango.</param>
        private void InsertSpace(Word.Range range)
        {
            range.InsertParagraph();
            CommitRange(range);
            range.InsertParagraph();
            CommitRange(range);
        }

        /// <summary>
        /// Obtiene la descripción de una propiedad extendida de un elemento de la base de datos.
        /// </summary>
        /// <param name="subBranches">Elemento de la base de datos expresado como lista de sub ramas.</param>
        /// <param name="allExtProperties">Lista con todas las propiedades extendidas.</param>
        /// <returns></returns>
        private string GetExtPropertyDescription(List<string> subBranches, IEnumerable<TSqlObject> allExtProperties)
        {
            // Encuentra la propiedad cuyo nombre en formato [rama].[subrama]... coincide con las sub ramas indicadas.
            var property = allExtProperties.FirstOrDefault(
                p => {
                    List<string> branches = new List<string>();
                    for (int i = 0; i < subBranches.Count; i++)
                    {
                        branches.Add(p.Name.Parts[i]);
                    }

                    return subBranches.SequenceEqual(branches);
                });

            string description = property?.GetProperty(ExtendedProperty.Value).ToString() ?? string.Empty;
            return description.StartsWith("N'") && description.EndsWith("'") ? description.Substring(2, description.Length - 3) : description;
        }

        /// <summary>
        /// Establece el estilo de la arista de una celda.
        /// </summary>
        /// <param name="cell">Celda a formatear.</param>
        /// <param name="style">Estilo de la línea.</param>
        /// <param name="borderType">Arista a colorear.</param>
        /// <param name="color">Color de la línea.</param>
        private void SetCellBorder(Word.Cell cell, Word.WdLineStyle style = Word.WdLineStyle.wdLineStyleSingle, Word.WdBorderType borderType = Word.WdBorderType.wdBorderBottom, Word.WdColor color = Word.WdColor.wdColorBlack)
        {
            Word.Border border = cell.Borders[borderType];
            border.Visible     = true;
            border.LineStyle   = style;
            border.LineWidth   = Word.WdLineWidth.wdLineWidth050pt;
            border.Color       = color;
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup  += new EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new EventHandler(this.ThisDocument_Shutdown);
        }

        /// <summary>
        /// Manejador del evento de cierre del documento.
        /// </summary>
        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }
    }
}