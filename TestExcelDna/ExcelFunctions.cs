using ExcelDna.Integration;

namespace TestExcelDna
{
    public static class ExcelFunctions
    {
        [ExcelFunction(Name = "HelloWorld", Description = "Returns 'Hello World'")]
        public static string HelloWorld() {
            return "Hello World";
        }

        [ExcelFunction(Name = "TestStaticAdd", Description = "Returns addition of two integers.")]
        public static int TestStaticAdd(
            [ExcelArgument(Name="i1",Description = "The first integer.")]
            int i1,
            [ExcelArgument(Name="i2",Description = "The second integer.")]
            int i2)
        {
            return (i1 + i2);
        }
    }
}
