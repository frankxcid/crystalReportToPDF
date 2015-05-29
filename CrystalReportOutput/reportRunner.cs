using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;

namespace CrystalReportOutput
{
    public class ReportRunner
    {
        /// <summary>
        /// The password to the database
        /// </summary>
        public static String defaultPassword = "";
        /// <summary>
        /// Stores any error messages
        /// </summary>
        public static String ErrorMessage = "";
        /// <summary>
        /// Select the output type
        /// </summary>
        public enum outputType
        {
            excel,
            pdf,
            word
        };
      
        /// <summary>
        /// Run the report with the output as a memory stream for download on webpage
        /// </summary>
        /// <param name="reportPath">The full path to the report</param>
        /// <param name="outputStream">The Memory Stream containing the output</param>
        /// <param name="databasePassword">The password to use when connecting to the database</param>
        /// <param name="outputFileName">The name of the file to download</param>
        /// <param name="Parameters"(Optional) >Ordered list of parameter values for the report</param>
        /// <param name="_outputType">(Optional) Type of output. Defaults to PDF</param>
        /// <returns>True if report ran and created output successfully</returns>
        public static Boolean execute(String reportPath, System.Web.HttpResponse resp, String databasePassword, String outputFileName, List<Object> Parameters = null, outputType _outputType = outputType.pdf)
        {
            defaultPassword = databasePassword;
            var rd = execReport(reportPath, Parameters, _outputType);
            String t = "";
            try
            {    
                if (rd == null) { return false; }
                var paramName = new List<string>();
                for (int i = 0; i < rd.ParameterFields.Count; i++)
                {
                    var thisParam = rd.ParameterFields[i];
                    var val = thisParam.CurrentValues;
                    var pn = thisParam.Name;
                    paramName.Add(thisParam.Name);
                }
                rd.ExportToHttpResponse(getEType(_outputType), resp, true, outputFileName);
                return true;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.ToString();
                return false;
            }

        }
        /// <summary>
        /// Run the report so the output will be a file on the machine running this application
        /// </summary>
        /// <param name="reportPath">The full path to the report file</param>
        /// <param name="outputPath">The full path including file name of the output</param>
        /// <param name="databasePassword">Password used when report connects to the database</param>
        /// <param name="Parameters">(Optional) Ordered list of parameter values for the report</param>
        /// <param name="_outputType">(Optional) Type of output. Defaults to PDF</param>
        /// <returns>True if report ran and created output successfully</returns>
        public static Boolean executeSave(String reportPath, String outputPath, String databasePassword, List<Object> Parameters = null, outputType _outputType = outputType.pdf)
        {
            defaultPassword = databasePassword;
            var rd = execReport(reportPath, Parameters, _outputType);
            try
            {
                if (rd == null) { return false; }
                var paramName = new List<string>();
                for (int i = 0; i < rd.ParameterFields.Count; i++)
                {
                    var thisParam = rd.ParameterFields[i];
                    var val = thisParam.CurrentValues;
                    var pn = thisParam.Name;
                    paramName.Add(thisParam.Name);
                }
                var fileOptions = new DiskFileDestinationOptions();
                var ExportPDF = rd.ExportOptions;
                fileOptions.DiskFileName = outputPath;
                ExportPDF.ExportDestinationOptions = fileOptions;
                ExportPDF.ExportDestinationType = ExportDestinationType.DiskFile;
                ExportPDF.ExportFormatType = getEType(_outputType);
                rd.Export();
                rd.Close();
                rd.Dispose();
                ErrorMessage = "";
                return true;
            }
            catch(Exception ex)
            {
                ErrorMessage = ex.ToString();
                return false;
            }
        }
        /// <summary>
        /// Sets up the output type
        /// </summary>
        /// <param name="thisOutputType">The output type selected by users</param>
        /// <returns>ExportFormatType</returns>
        private static ExportFormatType getEType(outputType thisOutputType)
        {   
            switch (thisOutputType)
            {
                case outputType.excel:
                    return ExportFormatType.Excel;
                case outputType.pdf:
                    return ExportFormatType.PortableDocFormat;
                case outputType.word:
                    return ExportFormatType.WordForWindows;
            }
            return ExportFormatType.PortableDocFormat;
        }
        /// <summary>
        /// Executes the report by loading the report, applying security credentials
        /// </summary>
        /// <param name="reportPath">Full path to the report file</param>
        /// <param name="Parameters">Ordered list of paramaters for the report</param>
        /// <param name="_outputType">The Output type</param>
        /// <returns>The report document if successful, otherwise, null</returns>
        private static ReportDocument execReport(String reportPath, List<Object> Parameters, outputType _outputType){
             try
            {
                var rd = new ReportDocument();
                rd.Load(reportPath, OpenReportMethod.OpenReportByTempCopy);
                var logonInfo = new TableLogOnInfo();
                if (Parameters != null)
                {
                    for (int i = 0; i < Parameters.Count; i++)
                    {
                        
                        rd.SetParameterValue(i, (Parameters[i].ToString().ToUpper() == "NULL" ? null : Parameters[i]));
                    }
                }
                applyCredentials(ref rd);
                if (rd.Subreports != null)
                {
                    foreach (Section thisSection in rd.ReportDefinition.Sections)
                    {
                        foreach (ReportObject subRO in thisSection.ReportObjects)
                        {
                            if (subRO.Kind == ReportObjectKind.SubreportObject)
                            {
                                var sro = (SubreportObject)subRO;
                                var sr = new ReportDocument();
                                sr = sro.OpenSubreport(sro.SubreportName);
                                applyCredentials(ref sr);
                            }
                        }
                    }
                }
                return rd;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.ToString();
                return null;
            }
        }
        /// <summary>
        /// Applies credentials to the report so it connects to the database if a password is supplied
        /// </summary>
        /// <param name="thisDoc">The report document</param>
        private static void applyCredentials(ref ReportDocument thisDoc)
        {
            if (defaultPassword == "") { return; }
            foreach (Table thisTable in thisDoc.Database.Tables)
            {
                var logonInfo = thisTable.LogOnInfo;
                if (logonInfo.ConnectionInfo.Password == null || logonInfo.ConnectionInfo.Password == "")
                {
                    logonInfo.ConnectionInfo.Password = defaultPassword;
                    thisTable.ApplyLogOnInfo(logonInfo);
                }
            }
        }
    }
}
