using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Reflection;
using System.Runtime.Loader;
using Westwind.Scripting;

namespace TestExcelDna
{
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
                RegisterFunctions();
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
            try
            {
                IntelliSenseServer.Uninstall();
            }
            catch (Exception ex)
            {
                var error = ex.StackTrace;
                Console.WriteLine(error);
            }
        }

        public void RegisterFunctions()
        {
            var script = new CSharpScriptExecution
            {
                SaveGeneratedCode = true,
                AlternateAssemblyLoadContext = AssemblyLoadContext.GetLoadContext(this.GetType().Assembly)
            };

            using (var ctx = System.Runtime.Loader.AssemblyLoadContext.EnterContextualReflection(this.GetType().Assembly))
            {
                // ... run code that loads extra assemblies here
                script.AddNetCoreDefaultReferences();
                script.AddAssembly(typeof(ExcelIntegration));
            }

            var code = $@"
                using System;
                using ExcelDna.Integration;

                namespace MyApp
                {{
	                public class Math
	                {{

                        public Math() {{}}

                        [ExcelFunction(Name = ""TestAdd"", Description = ""Returns 'TestAdd'"")]
                        public static string TestAdd(
                            [ExcelArgument(Name=""num1"",Description = ""The first integer."")]                            
                            int num1,
                            [ExcelArgument(Name=""num2"",Description = ""The second integer."")]                            
                            int num2)
		                {{
			                // string templates
			                var result = num1 + "" + "" + num2 + "" = "" + (num1 + num2);
			                Console.WriteLine(result);
		                
			                return result;
		                }}
		                
                        [ExcelFunction(Name = ""TestMultiply"", Description = ""Returns 'TestMultiply'"")]
		                public static string TestMultiply(
                            [ExcelArgument(Name=""num1"",Description = ""The first integer."")]                            
                            int num1,
                            [ExcelArgument(Name=""num2"",Description = ""The second integer."")]                            
                            int num2)
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