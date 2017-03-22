using System;
using System.IO;
using Serilog;
using Serilog.Events;

namespace ProLoop.WordAddin
{
    public static class Logger
    {
        public static string LogFolderPath;

        public static void SetupLogging()
        {
            //Set MyDocuments folder as the defaulth path for logs
            if (string.IsNullOrEmpty(LogFolderPath))
                throw new ArgumentNullException("logFolderPath");

            if (!Directory.Exists(LogFolderPath))
                Directory.CreateDirectory(LogFolderPath);

#if (DEBUG)
            var outputTemplateContext =
                "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {SourceContext} {MethodInvoked} {Message}{NewLine}{Exception}";
#else
            var outputTemplateContext =
                "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message}{NewLine}{Exception}";
#endif
            //Start a new logging session
            Log.Logger = new LoggerConfiguration()

#if (DEBUG)
                .MinimumLevel.Verbose()
#else
                .MinimumLevel.Information()
#endif

                //.Enrich.WithThreadId()
                .Enrich.FromLogContext()

#if (DEBUG)
                //.WriteTo.LiterateConsole()
                .WriteTo.RollingFile(
                    string.Concat(LogFolderPath, Path.DirectorySeparatorChar, "Log-{Hour}-Debug.log"),
                    buffered: true,
                    restrictedToMinimumLevel: LogEventLevel.Verbose,
                    flushToDiskInterval: TimeSpan.FromSeconds(1),
                    outputTemplate: outputTemplateContext,
                    retainedFileCountLimit: 1)

                //.WriteTo.RollingFile(
                //    string.Concat(LogFolderPath, Path.DirectorySeparatorChar, "Log-{Hour}.log"),
                //    buffered: true,
                //    restrictedToMinimumLevel: LogEventLevel.Information,
                //    flushToDiskInterval: TimeSpan.FromSeconds(1),
                //    outputTemplate: outputTemplateContext,
                //    retainedFileCountLimit: 24)
#else
                .WriteTo.RollingFile(
                    string.Concat(LogFolderPath, Path.DirectorySeparatorChar, "Log-{Hour}.log"),
                    buffered: true,
                    restrictedToMinimumLevel: LogEventLevel.Information,
                    flushToDiskInterval: TimeSpan.FromSeconds(1),
                    outputTemplate: outputTemplateContext,
                    retainedFileCountLimit: 2)
#endif
                .CreateLogger();
        }

        public static void CloseLogging()
        {
            Log.CloseAndFlush();
        }

        private static void DeleteLogFiles(string folderPath)
        {
            //Delete any previous log files in the PDF Folder
            var logFiles = Directory.GetFiles(folderPath, "*.log");
            foreach (var logFile in logFiles)
                try
                {
                    File.Delete(logFile);
                }
                catch (Exception exception)
                {
                    Log.Verbose("Exception while deleting log file {0}: {1}", logFile, exception.Message);
                }
        }
    }
}