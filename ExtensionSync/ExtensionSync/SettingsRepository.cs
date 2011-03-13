using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace LatishSehgal.ExtensionSync
{
    public class SettingsRepository
    {
        public SettingsRepository(ExtensionManagerFacade visualStudioExtensionManager, string path)
        {
            extensionManager = visualStudioExtensionManager;
            settingsFilePath = path;
        }

        public event Action<string> Log;

        public void PersistExtensionSettings()
        {
            var installedExtensionsInformation = extensionManager.GetInstalledExtensionsInformation();

            using (var fileStream = new FileStream(settingsFilePath, FileMode.OpenOrCreate))
            {
                var serializer = new XmlSerializer(typeof(List<ExtensionInformation>));
                serializer.Serialize(fileStream, installedExtensionsInformation);
            }
        }

        public List<ExtensionInformation> GetPersistedExtensionSettings()
        {
            var settings = new List<ExtensionInformation>();
            try
            {
                if (File.Exists(settingsFilePath))
                {
                    using (var fileStream = new FileStream(settingsFilePath, FileMode.Open))
                    {
                        var serializer = new XmlSerializer(typeof(List<ExtensionInformation>));
                        settings = (List<ExtensionInformation>)serializer.Deserialize(fileStream);
                        return settings;
                    }
                }
            }
            catch (Exception exception)
            {
                LogMessage(string.Format("Error while retrieving persisted Extension Settings: {0}",exception.Message));
            }

            return settings;
        }

        void LogMessage(string message)
        {
            if (Log != null)
                Log.Invoke(message);
        }

        ExtensionManagerFacade extensionManager;
        readonly string settingsFilePath;
    }
}
