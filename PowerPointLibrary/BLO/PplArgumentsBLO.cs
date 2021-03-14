using PowerPointLibrary.Entities;
using PowerPointLibrary.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.BLO
{
    public class PplArgumentsBLO
    {
        /// <summary>
        /// Read arguments from main params
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public PplArguments Read(string[] args)
        {
            PplArguments pplArguments = new PplArguments();

            // CurrentDirectoryPath
            pplArguments.CurrentDirectoryPath = Environment.CurrentDirectory;
            pplArguments.TemplatePath = "template.potx";

            // First params is the md document
            if (args != null)

                for (int i = 0; i < args.Length; i++)
                {
                    if (i == 0)
                    {
                        pplArguments.MdDocumentPath = args[i];
                        pplArguments.OutPutFile = Path.GetFileNameWithoutExtension(pplArguments.MdDocumentPath) + ".pptx";
                        pplArguments.UseSlideOutPutFile = Path.GetFileNameWithoutExtension(pplArguments.MdDocumentPath) + ".slides.pptx";


                    }



                    if (args[i] == "-d")
                    {
                        pplArguments.CurrentDirectoryPath = args[i + 1];
                        var IsPathRooted = Path.IsPathRooted(pplArguments.CurrentDirectoryPath);
                        if (!IsPathRooted)
                            pplArguments.CurrentDirectoryPath = Environment.CurrentDirectory + "\\" + pplArguments.CurrentDirectoryPath;
     
                        i++;
                    }
                    if (args[i] == "-t")
                    {
                        pplArguments.TemplatePath = args[i + 1];

                        i++;
                    }

                }


            // MdDocumentPath
            pplArguments.MdDocumentPath = Path.Combine(pplArguments.CurrentDirectoryPath, pplArguments.MdDocumentPath);
            if (!File.Exists(pplArguments.MdDocumentPath))
            {
                string msg = $"The file '{pplArguments.MdDocumentPath}' doesn't exist";
                throw new PplException(msg);
            }

            // Template Path
            pplArguments.TemplateConfigurationPath = Path.Combine(pplArguments.CurrentDirectoryPath, pplArguments.TemplatePath.Replace(".potx", ".json"));
            pplArguments.TemplatePath = Path.Combine(pplArguments.CurrentDirectoryPath, pplArguments.TemplatePath);

            if (!File.Exists(pplArguments.TemplateConfigurationPath))
            {
                string msg = $"The configuration file '{pplArguments.TemplateConfigurationPath}' doesn't exist";
                throw new PplException(msg);
            }
            if (!File.Exists(pplArguments.TemplatePath))
            {
                string msg = $"The template file '{pplArguments.TemplatePath}' doesn't exist";
                throw new PplException(msg);
            }
      


            // MdDocumentDirectoryPath
            pplArguments.MdDocumentDirectoryPath = Path.GetDirectoryName(pplArguments.MdDocumentPath);


            // OutPutPath
            pplArguments.OutPutPath = Path.Combine(pplArguments.CurrentDirectoryPath, "PowerPointFiles");
            this.CreatePathIfNotExist(pplArguments.OutPutPath);

            // OutPutFile
            pplArguments.OutPutFile = Path.Combine(pplArguments.OutPutPath, pplArguments.OutPutFile);
            pplArguments.UseSlideOutPutFile = Path.Combine(pplArguments.OutPutPath, pplArguments.UseSlideOutPutFile);

            return pplArguments;
        }

        private void CreatePathIfNotExist(string outPutPath)
        {
            if (!Directory.Exists(outPutPath))
                Directory.CreateDirectory(outPutPath);
        }
    }
}
