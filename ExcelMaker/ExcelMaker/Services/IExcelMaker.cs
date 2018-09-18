namespace ExcelMaker.Services
{
    public interface IExcelMaker
    {
        byte[] CreateExcel<T>(string sheetName, T model) where T : class;
    }
}
