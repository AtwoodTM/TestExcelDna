using ExcelDna.Integration;

namespace TestExcelDna
{
    public static class ExcelFunctions
    {
        [ExcelFunction(Name = "HelloWorld", Description = "Returns 'Hello World'")]
        public static string HelloWorld() {
            return "Hello World";
        }
    }
}
