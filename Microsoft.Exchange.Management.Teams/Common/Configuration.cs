//-----------------------------------------------------------------------
// <copyright file="Configuration.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.Management.Teams.Common
{
    using System;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Linq;

    /// <summary>
    /// Defines the <see cref="Configuration"/> class.
    /// </summary>
    public class Configuration
    {
        /// <summary>
        /// Gets the value for a given key in the app.config
        /// </summary>
        /// <param name="key">Name of the configuration key</param>
        /// <returns>Returns the value for a given key</returns>
        public string GetConfigKeyValue(string key)
        {
            string appConfigPath = $"{Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path)}.config";
            XDocument appConfig = XDocument.Load(appConfigPath);
            
            return appConfig.Descendants("add")
                            .First(node => (string)node.Attribute("key") == key)
                            .Attribute("value").Value;
        }
    }
}
