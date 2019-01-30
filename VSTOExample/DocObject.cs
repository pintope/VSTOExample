namespace VSTOExample
{
    using System;
    using System.Collections.Generic;


    /// <summary>
    /// Objeto que representa una fila de cada una de las tablas.
    /// </summary>
    public class DocObject
    {
        public DocObject()
        {
            Columns = new List<Tuple<bool, string, string, string>>();
        }


        /// <summary>
        /// Nombre del objeto.
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Descripción del objeto.
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Descripción del objeto de base de datos a presentar en forma tabular.
        /// </summary>
        public DBObject Type { get; set; }
        /// <summary>
        /// Columnas que corresponden a cada uno de los objetos de la base de datos, ya sean parámetros de una función
        /// o procedimiento almacenado, ya sea un campo de una tabla. Cada uno de los elementos de la tupla representan,
        /// respectivamente:
        ///     - Si el campo de la tabla es clave primaria o no, en caso de una tabla.
        ///     - Nombre del objeto.
        ///     - Tipo del objeto.
        ///     - Descripción del objeto.
        /// </summary>
        public List<Tuple<bool, string, string, string>> Columns { get; set; }
    }
}