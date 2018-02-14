using System;
using System.IO;

namespace WJRK
{

    public class Logger {
        private static readonly string LOG_FILE = "wjrk-update-addresses.log";

        private static void Write(string type, string message){
            using (System.IO.StreamWriter file = 
            new System.IO.StreamWriter(LOG_FILE, true))
            {
                file.WriteLine(String.Format("[{0}] {1}: {2}", DateTime.Now.ToString(), type, message));
            }
            Console.WriteLine(message);
        }

        public static void Log(string message){
            Console.ForegroundColor = ConsoleColor.Cyan;
            Write("LOG    ", message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        public static void Info(string message){
            Console.ForegroundColor = ConsoleColor.Blue;
            Write("INFO   ", message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        public static void Error(string message){
            Console.ForegroundColor = ConsoleColor.Red;
            Write("ERROR  ", message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        public static void Success(string message){
            Console.ForegroundColor = ConsoleColor.Green;
            Write("SUCCESS", message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        public static void Warn(string message){
            Console.ForegroundColor = ConsoleColor.Yellow;
            Write("WARNING", message);
            Console.ForegroundColor = ConsoleColor.White;
        }

    }
}