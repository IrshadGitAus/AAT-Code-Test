using AAT.File.Processor3.Helpers;
using AAT.File.Processor3.Services;
using AAT.File.Processor3.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using System;
using System.IO;

namespace AAT.File.Processor3
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var host = AppStartup();
            var serviceProvider = host.Services;
            var zipService = serviceProvider.GetService<IZipService>();
            System.AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
            bool confirmed = false;

            do
            {
                ConsoleKey response;
                do
                {
                    Console.Write("Are you sure you want to extract the zip files? [Y/N] ");
                    response = Console.ReadKey(false).Key;   // true is intercept key (dont show), false is show
                    if (response != ConsoleKey.Enter)
                        Console.WriteLine();

                } while (response != ConsoleKey.Y && response != ConsoleKey.N);
                if (response == ConsoleKey.N) Environment.Exit(1);
                confirmed = response == ConsoleKey.Y;
            } while (!confirmed);

            if (confirmed)
            {
                Log.Logger.Information("-----------******** START: Processing zip files*******----------");
                bool success = zipService.ProcessFiles();
                Log.Logger.Information("-----------******** END: Processing zip files*******----------");
            }
        }

        static void BuildConfig(IConfigurationBuilder builder)
        {
            builder.SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables().Build();
        }

        static IHost AppStartup()
        {
            var builder = new ConfigurationBuilder();

            BuildConfig(builder);

            var Configuration = builder.Build();

            // Specifying the configuration for serilog
            Log.Logger = new LoggerConfiguration() // initiate the logger configuration
                                                   //.ReadFrom.Configuration(builder.Build()) // connect serilog to our configuration folder
                            .WriteTo.Console()
                            .WriteTo.RollingFile(ApplicationHelper.GetApplicationRoot() + "\\logs\\log-{Date}.txt", outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss zzz} [{Level:u3}] {Properties}{Message}{NewLine}{Exception}")
                            .CreateLogger(); //initialise the logger

            var host = Host.CreateDefaultBuilder() // Initialising the Host 
                        .ConfigureServices((context, services) =>
                        { // Adding the DI container for configuration

                            services.Configure<ZipConfiguration>(Configuration.GetSection("ZipSettings"));
                            services.AddTransient<IZipService, ZipService>();

                        })
                        .UseSerilog() // Add Serilog
                        .Build(); // Build the Host

            return host;
        }

        static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {
            Log.Logger.Error(e.ExceptionObject.ToString());
            Console.WriteLine("An error occured. See the log-{date}.txt file in the logs folder for more information. Press Enter to continue");
            Console.ReadLine();
            Environment.Exit(1);
        }
    }
}
