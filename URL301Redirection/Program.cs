using System;


namespace URL301Redirection
{
    class Program
    {
        public static ValidateURLredirection objValidateURL = new ValidateURLredirection();


        public Program()
        {
            // kill the excel process before script execution 
            try
            {
                var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (var p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        p.Kill();
                    }
                }
            }
            catch (Exception e)
            {
                System.Console.WriteLine("Exception related to Excel process in the Task Manager .... " + e);
            }
        }

        public static void Main(string[] args)
        {
            Execute();
        }

        public static void Execute()
        {
            objValidateURL.ValidateUrlsMethod();
        }
    }
}


