// Copyright (c) 2020 Cognex Corporation. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.IO;
using log4net;
using Newtonsoft.Json;

namespace WindowsSnap
{
    internal static class Logger
    {
        private static string AppDataFolder = "WindowsSnap";
        private static string JsonFile = "WindowsSnap.ws";

        private static ILog _logger => LogManager.GetLogger(typeof(WindowsSnapProgram));

        public static void Log(string msg)
        {
            try
            {
                _logger?.Info(msg); // Info logs to the file
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex);
            }
        }

        public static void Console(string msg)
        {
            try
            {
                _logger?.Debug(msg); // Debug only prints to console
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex);
            }
        }

        public static List<Snapshot> ReadWindowsSnapJson()
        {
            try
            {
                var path = GetJsonPath();
                using (var file = new FileStream(path,
                                                 FileMode.Open,
                                                 FileAccess.Read,
                                                 FileShare.Read,
                                                 64 * 1024,
                                                 true))
                using (var reader = new StreamReader(file))
                {
                    var json = reader.ReadToEnd();
                    return JsonConvert.DeserializeObject<List<Snapshot>>(json);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error reading json: " + ex);
            }

            return new List<Snapshot>();
        }

        public static void WriteWindowsSnapJson(List<Snapshot> snapshots)
        {
            try
            {
                var path = GetJsonPath();
                using (var sw = new StreamWriter(path))
                {
                    var json = JsonConvert.SerializeObject(snapshots);
                    sw.Write(json);
                }
            }
            catch (Exception ex)
            {
                Logger.Log("Error writing json: " + ex);
            }
        }

        private static string GetJsonPath()
        {
            var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                      AppDataFolder);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            var path = Path.Combine(folder, JsonFile);
            return path;
        }
    }
}
