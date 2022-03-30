using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPITraining42
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
     public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
           
            string pipesInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

           
                foreach (Pipe pipe in pipes)
                {
                
                double pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                double outerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                double innerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                double len = Math.Round(UnitUtils.ConvertFromInternalUnits(pipeLength, UnitTypeId.Millimeters), 2);
                double oDim = Math.Round(UnitUtils.ConvertFromInternalUnits(outerDiameter, UnitTypeId.Millimeters), 2);
                double iDim = Math.Round(UnitUtils.ConvertFromInternalUnits(innerDiameter, UnitTypeId.Millimeters), 2);
                pipesInfo += $"Имя типа трубы: {pipe.Name}\tНаружный диаметр: {oDim}мм\tВнутренний диаметр:{iDim}{Environment.NewLine}";
                
            }
            var saveDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "pipesInfo.csv",
                DefaultExt = ".csv"
            };

            string selectedFilePath = string.Empty;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            File.WriteAllText(selectedFilePath, pipesInfo);

            
           
            return Result.Succeeded;
        }
    }

}
