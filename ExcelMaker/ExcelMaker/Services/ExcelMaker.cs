namespace ExcelMaker.Services
{
    using global::ExcelMaker.Attributes;
    using global::ExcelMaker.Model;
    using OfficeOpenXml;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    public class ExcelMaker : IExcelMaker
    {
        public byte[] CreateExcel<T>(string sheetName, T model) where T : class
        {
            using (ExcelPackage xlPackage = new ExcelPackage())
            {
                Dictionary<string, List<string>> data = this.GetDataReflection(model);

                string fileName = string.Format($"{sheetName}");
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add(fileName);


                int row = 1;
                int col = 1;
                int countFullCol = 1;

                IEnumerable<List<string>> list = data.Where(x => x.Value.Count > 1).Select(x => x.Value);

                if (!list.Any())
                {
                    throw new ApplicationException("List of arg is empty");
                }

                foreach (var label in list)
                {
                    for (int i = 0; i < label.Count; i++)
                    {
                        string value = label[i].Split(new char[] { '|' }).FirstOrDefault(); //A1 ...... R1
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            throw new ArgumentNullException($"Empty column name for col index {i}");
                        }
                        worksheet.Cells[row, col].Value = value;
                            
                        worksheet.Cells[row, col].Style.Font.Bold = true;
                        col++;
                        countFullCol++;
                    }

                    break;
                }

                

                row++;
                col = 1;
                foreach (var item in list)
                {
                    for (int i = 0; i < item.Count; i++)
                    {
                        string value = item[i].Split(new char[] { '|' }).LastOrDefault();
                        worksheet.Cells[row, col].Value = value;
                        col++;
                    }

                    row++;
                    col = 1;
                }

                for (int i = 1; i <= countFullCol; i++)
                {
                    worksheet.Column(i).AutoFit();
                }


                return xlPackage.GetAsByteArray();
            }
        }


        private Dictionary<string, List<string>> GetDataReflection<T>(T model) where T : class
        {
            Dictionary<string, List<string>> data = new Dictionary<string, List<string>>();

            Type type = typeof(T);

            PropertyInfo[] currTypeProps = type.GetProperties();

            foreach (var prop in currTypeProps)
            {
                if (prop.PropertyType.IsPrimitive || prop.PropertyType == typeof(Decimal) || prop.PropertyType == typeof(String))
                {
                    var value = prop.GetValue(model);
                    //string propType = prop.PropertyType.FullName;
                    string val = value == null ? "Null" : value.ToString();

                    data.Add(prop.Name, new List<string> { val });
                }
                else
                {
                    var innerType = prop.PropertyType.GetGenericArguments().First();

                    var list = (IList)prop.GetValue(model, null);
                    int counter = 0;

                    foreach (var item in list)
                    {
                        bool isCustomAttrExist = innerType.GetProperties().Any(x => x.CustomAttributes.Count() > 0);

                        List<ColNameModel> listClassProps = new List<ColNameModel>();

                        if (isCustomAttrExist)
                        {
                            listClassProps = innerType.GetProperties()
                                    .Where(x => x.CustomAttributes
                                                .Select(t => t.AttributeType)
                                                .Contains(typeof(ColumnNameAttribute)))
                                    .Select(x => new ColNameModel
                                    {
                                        Properties = x,
                                        Name = x.CustomAttributes
                                                .SelectMany(t => t.ConstructorArguments.Select(v => v.Value))
                                                .FirstOrDefault()
                                    })
                                    .ToList();
                        }
                        else
                        {
                            listClassProps = innerType.GetProperties()
                                                .Select(x => new ColNameModel
                                                {
                                                    Properties = x,
                                                    Name = x.Name
                                                })
                                                .ToList();
                        }

                        List<string> listProps = new List<string>();

                        foreach (var classProp in listClassProps)
                        {
                            //object val = classProp.GetValue(item, null);
                            object val = classProp.Properties.GetValue(item, null);

                            listProps.Add($"{classProp.Name}|{val}");
                        }

                        data.Add($"{innerType.Name}{counter}", listProps);
                        counter++;
                    }
                }
            }

            return data;
        }
    }
}
