using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace Style_Manager
{
    static class StyleData
    {
        const string FileName = @"Styles.bin";

        internal static SortedDictionary<string, Style> GetStyles()
        {
            SortedDictionary<string, Style> styles = null;

            string path = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                FileName);

            if (File.Exists(path))
            {
                Stream fileStream = null;
                try
                {
                    fileStream = File.OpenRead(path);
                    BinaryFormatter deserializer = new BinaryFormatter();
                    styles = (SortedDictionary<string, Style>)deserializer.Deserialize(fileStream);
                }
                catch (Exception)
                {
                    System.Windows.Forms.MessageBox.Show("Error getting styles from disk. Contact support.");
                }
                finally
                { 
                    if (fileStream != null)
                        fileStream.Close(); 
                }

            }

            return styles;
        }

        internal static void SaveStyles(SortedDictionary<string, Style> styles)
        {
            string path = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                FileName);
            // save to disk
            Stream fileStream = File.Create(path);
            BinaryFormatter serializer = new BinaryFormatter();
            try
            {
                serializer.Serialize(fileStream, styles);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Error saving styles to disk. Contact support.");
            }
            finally
            {
                fileStream.Close();
            }
        }
    }
}
