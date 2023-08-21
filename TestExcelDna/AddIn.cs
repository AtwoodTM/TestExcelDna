using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Reflection;
using System.Runtime.InteropServices;
using Westwind.Scripting;

namespace TestExcelDna
{
    [ComVisible(false)]
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {
                // Rosyln warmup
                // at app startup - runs a background task, but don't await
                _ = RoslynLifetimeManager.WarmupRoslyn();
                IntelliSenseServer.Install();
                RegisterFunctionsWorks();
                RegisterFunctionsDoesNotWork();
                IntelliSenseServer.Refresh();
            }
            catch (Exception ex)
            {
                var error = ex.StackTrace;
                Console.WriteLine(error);
            }
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public void RegisterFunctionsWorks()
        {
            var script = new CSharpScriptExecution() { SaveGeneratedCode = true };
            script.AddDefaultReferencesAndNamespaces();

            var code = $@"
                using System;

                namespace MyApp
                {{
	                public class MathWorks
	                {{

                        public MathWorks() {{}}

                        public static string TestAddWorks(int num1, int num2)
		                {{
			                // string templates
			                var result = num1 + "" + "" + num2 + "" = "" + (num1 + num2);
			                Console.WriteLine(result);
		                
			                return result;
		                }}
		                
		                public static string TestMultiplyWorks(int num1, int num2)
		                {{
			                // string templates
			                var result = $""{{num1}}  *  {{num2}} = {{ num1 * num2 }}"";
			                Console.WriteLine(result);
			                
			                result = $""Take two: {{ result ?? ""No Result"" }}"";
			                Console.WriteLine(result);
			                
			                return result;
		                }}
	                }}
                }}";

            // need dynamic since current app doesn't know about type
            dynamic math = script.CompileClass(code);

            Console.WriteLine(script.GeneratedClassCodeWithLineNumbers);

            // Grabbing assembly for below registration
            dynamic mathClass = script.CompileClassToType(code);
            var assembly = mathClass.Assembly;

            if (!script.Error)
            {
                Assembly asm = assembly;
                Type[] types = asm.GetTypes();
                List<MethodInfo> methods = new();

                // Get list of MethodInfo's from assembly for each method with ExcelFunction attribute
                foreach (Type type in types)
                {
                    foreach (MethodInfo info in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                    {
                        methods.Add(info);
                    }
                }

                Integration.RegisterMethods(methods);
            }
            else
            {
                //MessageBox.Show("Errors during compile!");
            }
        }

        public void RegisterFunctionsDoesNotWork()
        {
            var script = new CSharpScriptExecution() { SaveGeneratedCode = true };
            script.AddDefaultReferencesAndNamespaces();

            // No ExcelDna.Integration.dll, so the below commented line does not work
            // script.AddAssembly("ExcelDna.Integration.dll");
            script.AddAssembly(typeof(ExcelIntegration));
            script.AddAssembly(typeof(ExcelArgumentAttribute));
            script.AddAssembly(typeof(ExcelFunctionAttribute));

            var code = $@"
                using System;
                using ExcelDna.Integration;

                namespace MyApp
                {{
	                public class MathDoesNotWork
	                {{

                        public MathDoesNotWork() {{}}

                        [ExcelFunction(Name = ""TestAddDoesNotWork"", Description = ""Returns 'TestAddDoesNotWork'"")]
                        public static string TestAddDoesNotWork(int num1, int num2)
		                {{
			                // string templates
			                var result = num1 + "" + "" + num2 + "" = "" + (num1 + num2);
			                Console.WriteLine(result);
		                
			                return result;
		                }}
		                
                        [ExcelFunction(Name = ""TestMultiplyDoesNotWork"", Description = ""Returns 'TestMultiplyAddDoesNotWork'"")]
		                public static string TestMultiplyDoesNotWork(int num1, int num2)
		                {{
			                // string templates
			                var result = $""{{num1}}  *  {{num2}} = {{ num1 * num2 }}"";
			                Console.WriteLine(result);
			                
			                result = $""Take two: {{ result ?? ""No Result"" }}"";
			                Console.WriteLine(result);
			                
			                return result;
		                }}
	                }}
                }}";

            // need dynamic since current app doesn't know about type
            dynamic math = script.CompileClass(code);

            Console.WriteLine(script.GeneratedClassCodeWithLineNumbers);

            // Grabbing assembly for below registration
            dynamic mathClass = script.CompileClassToType(code);
            var assembly = mathClass.Assembly;

            if (!script.Error)
            {
                Assembly asm = assembly;
                Type[] types = asm.GetTypes();
                List<MethodInfo> methods = new();

                // Get list of MethodInfo's from assembly for each method with ExcelFunction attribute
                foreach (Type type in types)
                {
                    foreach (MethodInfo info in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                    {
                        methods.Add(info);
                    }
                }

                Integration.RegisterMethods(methods);
            }
            else
            {
                //MessageBox.Show("Errors during compile!");
            }
        }
    }
}