
using ConsoleExcel;
using log4net;
using log4net.Core;
using log4net.Repository.Hierarchy;
using System.CommandLine;
using System.Security;

// Configure log4net using the .config file
[assembly: log4net.Config.XmlConfigurator(Watch = true)]
// This will cause log4net to look for a configuration file
// called ConsoleApp.exe.config in the application base
// directory (i.e. the directory containing ConsoleApp.exe)

internal class Program
{
    private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(typeof(Program));
    
    


    // delete the old log files in the logging subfolder every month
    private static void DeleteOldLogs()
    {
        string logsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logging");
        if (Directory.Exists(logsFolder))
        {
            try
            {
                var files = Directory.GetFiles(logsFolder, "*.log", SearchOption.TopDirectoryOnly);
                foreach (var file in files)
                {
                    FileInfo fileInfo = new(file);
                    // Check if the file is older than 30 days
                    if (fileInfo.CreationTime < DateTime.Now.AddMonths(-1))
                    {
                        _log.Info($"Deleting old log file: {fileInfo.Name}");
                        File.Delete(file);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Error($"Error deleting old log files: {ex.Message}");
            }
        }
    }

    private static void ConfigureLoggingFromConfig()
    {
        // Get the repository
        Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();

        // Remove all existing appenders
        hierarchy.Root.RemoveAllAppenders();

        // Create logs subfolder in application directory
        string logsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logging");
        try
        {
            if (!Directory.Exists(logsFolder))
            {
                Directory.CreateDirectory(logsFolder);
            }
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException || ex is PathTooLongException || ex is IOException || ex is SecurityException)
        {
            Console.Error.WriteLine($"Error creating log directory: {ex.Message}");
            logsFolder = AppDomain.CurrentDomain.BaseDirectory; // Fall back to base directory
        }

        // Get appenders from config
        var consoleAppender = LogManager.GetRepository()
            .GetAppenders()
            .FirstOrDefault(a => a.Name == "ConsoleAppender");

        var fileAppender = LogManager.GetRepository()
            .GetAppenders()
            .FirstOrDefault(a => a.Name == "LogFileAppender") as log4net.Appender.FileAppender;

        // Configure file appender to use subfolder if found
        if (fileAppender != null)
        {
            try
            {
                // Create a timestamped log filename
                string logFileName = $"application_{DateTime.Now:yyyyMMdd_HHmmss}.log";
                string fullPath = Path.Combine(logsFolder, logFileName);
                // Update the file path
                fileAppender.File = fullPath;
                // Need to activate options for changes to take effect
                fileAppender.ActivateOptions();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error configuring file appender: {ex.Message}");
                // Continue without the file appender
                fileAppender = null;
            }
        }

        // Add only console and file appenders back
        if (consoleAppender != null)
            hierarchy.Root.AddAppender(consoleAppender);

        if (fileAppender != null)
            hierarchy.Root.AddAppender(fileAppender);

        // If neither appender was found, create defaults with subfolder
        if (consoleAppender == null && fileAppender == null)
        {
            _log.Warn("No appenders found in config, creating defaults with subfolder");

            // Create a default file appender to the subfolder
            var defaultFileAppender = new log4net.Appender.FileAppender
            {
                Name = "DefaultFileAppender",
                File = Path.Combine(logsFolder, $"default_{DateTime.Now:yyyyMMdd_HHmmss}.log"),
                AppendToFile = true,
                Layout = new log4net.Layout.PatternLayout("%date [%thread] %-5level %logger - %message%newline")
            };
            defaultFileAppender.ActivateOptions();

            // Create a default console appender
            var defaultConsoleAppender = new log4net.Appender.ConsoleAppender
            {
                Name = "DefaultConsoleAppender",
                Layout = new log4net.Layout.PatternLayout("%date [%thread] %-5level %logger - %message%newline")
            };
            defaultConsoleAppender.ActivateOptions();

            // Add both default appenders
            hierarchy.Root.AddAppender(defaultFileAppender);
            hierarchy.Root.AddAppender(defaultConsoleAppender);
        }

        // Make sure root level is configured
        #if DEBUG
                // Use Debug level in Debug builds
                hierarchy.Root.Level = Level.Debug;
                Console.WriteLine("Debug logging is enabled");
        #else
            // Use Error level in Release builds
            hierarchy.Root.Level = Level.Error;
        #endif
        hierarchy.Configured = true;
        // Add a log observer to detect when entries are written
        log4net.Repository.Hierarchy.Logger rootLogger = hierarchy.Root;

        
        
    }

    private static async Task<int> Main(string[] args)
    {
        DeleteOldLogs();
        ConfigureLoggingFromConfig();
        // Log an info level message
        // Log an info level message
        if (_log.IsInfoEnabled)
        {
            _log.Info("Application [ConsoleApp] Start");            
        }



        var fileOption = new Option<FileInfo?>(
            name: "--file",
            description: "The Excel file to read");

        var customOption = new Option<string?>(
            name: "--option",
            description: "A custom option for additional functionality.");

        var rootCommand = new RootCommand("Sample app for CloseXML");
        rootCommand.AddOption(fileOption);
        rootCommand.AddOption(customOption);

        rootCommand.SetHandler((file, option) =>
        {
            ParseXL(file!, option);
        },
            fileOption, customOption);

        var ret =  await rootCommand.InvokeAsync(args);       
        return ret;
    }

    static void ParseXL(FileInfo file, string? option)
    {
        if (file == null || !file.Exists)
        {
            _log.Error("File not found or invalid file path provided.");
            return;
        }
        // Check file size limits
        const long maxFileSizeBytes = 10 * 1024 * 1024; // 10MB limit
        if (file.Length > maxFileSizeBytes)
        {
            _log.Error($"File exceeds maximum allowed size: {file.Length} bytes");
            return;
        }
        // Check file extension
        if (!file.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            _log.Error($"Invalid file type: {file.Extension}");
            return;
        }

        if (!string.IsNullOrEmpty(option))
        {
            XLParser xlParser = new();
            if (option.Equals("test1", StringComparison.OrdinalIgnoreCase))
            {
                _log.Info("test1 option selected.");
                xlParser.OpenFilewithOption(file.FullName, option);

            }
            else if (option.Equals("test2", StringComparison.OrdinalIgnoreCase))
            {
                _log.Info("test2 option selected.");
                xlParser.OpenFilewithOption(file.FullName, option);
            }
            else
            {
                _log.Error($"Unknown option: {option}");
            }
        }
        else
        {
            _log.Error("No option provided.");
        }
    }
}