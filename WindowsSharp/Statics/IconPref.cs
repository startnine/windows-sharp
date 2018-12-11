using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace WindowsSharp
{
    public static class IconPref
    {
        static XmlDocument prefDocument = new XmlDocument();
        static String prefPath = Environment.ExpandEnvironmentVariables(@"%appdata%\Start9\IconPref.xml");
        public static String iconsPath = Environment.ExpandEnvironmentVariables(@"%appdata%\Start9\IconOverrides");

        static IconPref()
        {
            if (!File.Exists(prefPath))
            {
                //File.Create(prefPath);
                File.WriteAllLines(prefPath, new List<String>()
                {
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>",
                    "<icon>",
                    "	<file>",
                    "	</file>",
                    "	<process>",
                    "	</process>",
                    "</icon>"
                });
            }

            if (!Directory.Exists(iconsPath))
            {
                Directory.CreateDirectory(iconsPath);
            }

            prefDocument.Load(prefPath);
        }

        public static void Save()
        {
            IconPref.Save();
        }

        public static List<IconOverride> FileIconOverrides
        {
            get
            {
                List<IconOverride> overrides = new List<IconOverride>();
                foreach (XmlNode node in prefDocument.SelectSingleNode(@"/icon/file").ChildNodes)
                {
                    overrides.Add(new IconOverride(node));
                }
                return overrides;
            }
        }

        internal static IconOverride IconOverrideFromPath(string path)
        {
            IconOverride returnVal = null;
            foreach (IconOverride i in FileIconOverrides)
            {
                if (Environment.ExpandEnvironmentVariables(i.TargetName).ToLowerInvariant() == (Environment.ExpandEnvironmentVariables(path).ToLowerInvariant()))
                {
                    returnVal = i;
                    break;
                }
            }
            return returnVal;
        }
    }

    public class IconOverride// : INotifyPropertyChanged
    {
        XmlNode target;

        public String TargetName
        {
            get
            {
                var result = "";
                if (target.Attributes["targetName"] != null)
                {
                    result = target.Attributes["targetName"].Value;
                }
                return result;
            }
            set
            {
                target.Attributes["targetName"].Value = value.ToString();
                IconPref.Save();
            }
        }

        public Boolean IsFullPath
        {
            get
            {
                var result = false;
                if (target.Attributes["fullPath"] != null)
                {
                    result = bool.Parse(target.Attributes["fullPath"].Value);
                }
                return result;
            }
            set
            {
                target.Attributes["fullPath"].Value = value.ToString();
                IconPref.Save();
            }
        }

        public String ReplacementName
        {
            get
            {
                var result = "";
                if (target.Attributes["replacement"] != null)
                {
                    result = target.Attributes["replacement"].Value;
                }
                return result;
            }
            set
            {
                target.Attributes["replacement"].Value = value.ToString();
                IconPref.Save();
            }
        }

        public IconOverride(XmlNode targetNode)
        {
            target = targetNode;
        }
    }
}
