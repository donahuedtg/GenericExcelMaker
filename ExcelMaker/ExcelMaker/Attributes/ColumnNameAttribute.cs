namespace ExcelMaker.Attributes
{
    using System;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true, Inherited = true)]
    public sealed class ColumnNameAttribute : Attribute
    {
        private readonly string colName;

        public string ColName
        {
            get
            {
                return colName;
            }
        }

        public ColumnNameAttribute(string colName)
        {
            this.colName = colName;
        }
    }
}

