using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleSync sync;
            
            if (args.Length > 0)
            {
                if (args.GetValue(0).ToString() == "--ExchangeSyncOnly")
                {
                    if (args.GetValue(1).ToString() == "true")
                    {
                        Console.WriteLine("Executing Exchange Sync Only");
                        sync = new ConsoleSync(true); 
                    }
                    else if (args.GetValue(1).ToString() == "false")
                    {
                        Console.WriteLine("Executing Exchange Sync All");
                        sync = new ConsoleSync(false);
                    }
                    else
                    {
                        Console.WriteLine("Executing Exchange Sync - Invalid Parameter Value Specified");
                        sync = new ConsoleSync();
                    }
                }
                else if (args.GetValue(0).ToString() == "--ReSyncFailed")
                {
                    if (args.GetValue(1).ToString() == "true")
                    {
                        Console.WriteLine("Executing Resync Failed");
                        sync = new ConsoleSync(true, true);
                    }
                    else
                    {
                        Console.WriteLine("Executing Exchange Sync - Invalid Parameter Value for Resync Failed");
                        sync = new ConsoleSync();
                    }

                }
                else
                {
                    Console.WriteLine("Executing Exchange Sync - Invalid Arguments Specified");
                    sync = new ConsoleSync();
                }
            }
            else
            {
                sync = new ConsoleSync();
                Console.WriteLine("Executing Exchange Sync - No Arguments Specified");
            }
        }
    }
}
