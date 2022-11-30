using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        { }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {   
            if (context.ExternalPlanSetup != null)
            {
                if (context.ExternalPlanSetup.Beams.First().Blocks.First().Outline[0].Count() > 0)
                {
                    Point[][] points = context.ExternalPlanSetup.Beams.First().Blocks.First().Outline;
                    double viewBox = SetViewBoxFromApplicator(context.ExternalPlanSetup.Beams.First().Applicator.Id);
                    
                    StringBuilder sbBlock = new StringBuilder();
                    StringBuilder sbBlockPath = new StringBuilder();
                                                            
                    sbBlockPath.Append("<path d =\"");

                    //Flip Y across the X-axis since SVG coordinate system has +Y going down
                    //Scale all points by 0.95 since the outline coordinates of the block are at 100cm and the block needs to print at 95cm size
                    for (int i = 0; i < points[0].Length; i++)
                    {
                        points[0][i] = new Point(points[0][i].X * 0.95, points[0][i].Y * -0.95);
                                                
                        if (i == 0)
                        { sbBlockPath.Append("M " + points[0][i].X + " " + points[0][i].Y); }
                        else
                        { sbBlockPath.Append(" L " + points[0][i].X + " " + points[0][i].Y); }
                    }
                    sbBlockPath.Append(" z\" stroke=\"black\" stroke-width=\"1\" />\n");

                    //Start the XML and set the viewbox
                    sbBlock.AppendLine("<?xml version=\"1.0\" standalone=\"no\"?>");
                    sbBlock.AppendLine("<svg xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" " +
                        "width=\"" + viewBox.ToString() + "mm\" height=\"" + viewBox.ToString() + "mm\" " +
                        "viewBox=\"" + (-0.5 * viewBox).ToString() + " " + (-0.5 * viewBox).ToString() + " " + 
                        viewBox.ToString() + " " + viewBox.ToString() + "\" preserveAspectRatio=\"minXminY meet\" >") ;

                    //Add the block path
                    sbBlock.Append(sbBlockPath);

                    //Add locating L and Pin
                    sbBlock.Append(AddLocatingJigFromApplicator(context.ExternalPlanSetup.Beams.First().Applicator.Id));
                    
                    //end XML
                    sbBlock.AppendLine("</svg>");

                    string filePath = @"\\thimage-16\va_transfer\blockexports\block_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".svg";

                    try
                    { File.AppendAllText(filePath, sbBlock.ToString()); }
                    catch (Exception err)
                    { MessageBox.Show("Error writing data to file.\n" + err.Message); }

                    if (viewBox < 150)
                    {
                        //SaveSVGWithoutHoles(sbBlockPath.ToString(), context.ExternalPlanSetup.Beams.First().Applicator.Id);                    
                    }
                }
                else
                { MessageBox.Show("Please open a plan with an electron field that has a block outline."); }
            }
        }

        private double SetViewBoxFromApplicator(string id)
        {
            // viewbox is the dimension of the baseplate for the applicator
            switch (id)
            {
                case "A06":
                    return 82;
                case "A10":
                    return 120;
                case "A15":
                    return 168;
                case "A20":
                    return 215;
                case "A25":
                    return 263;
                default:
                    return 263;
            }
        }

        private string AddLocatingJigFromApplicator(string id)
        {
            if (id == "A20" | id == "A25")
                return AddCAXMarkOnly();

            StringBuilder sbJig = new StringBuilder();
            // Adding the 'L-hole'
            sbJig.AppendLine("<path d=\"M -6.3 2.7 L 2.7 2.7 L 2.7 -3.3 L 6.3 -3.3 L 6.3 6.3 L -6.3 6.3 z\" stroke=\"black\" stroke-width=\"1\" />");

            // Adding the Pin-hole
            switch (id)
            {
                case "A06":
                    sbJig.AppendLine("<circle cx=\"0\" cy=\"0\" r=\"1.8\" stroke=\"blue\" stroke-width=\"1\" />");
                    return sbJig.ToString();
                case "A10":
                    sbJig.AppendLine("<circle cx=\"-1\" cy=\"0\" r=\"1.8\" stroke=\"blue\" stroke-width=\"1\" />");
                    return sbJig.ToString();
                case "A15":
                    sbJig.AppendLine("<circle cx=\"-1\" cy=\"-1\" r=\"1.8\" stroke=\"blue\" stroke-width=\"1\" />");
                    return sbJig.ToString();
                case "A20":
                    sbJig.AppendLine("<circle cx=\"-2\" cy=\"-2\" r=\"1.8\" stroke=\"blue\" stroke-width=\"1\" />");
                    return sbJig.ToString();
                case "A25":
                    sbJig.AppendLine("<circle cx=\"-3\" cy=\"-3\" r=\"1.8\" stroke=\"blue\" stroke-width=\"1\" />");
                    return sbJig.ToString();
                default:
                    return "";
            }
        }

        private string AddCAXMarkOnly()
        {
            StringBuilder sbCAX = new StringBuilder();
            sbCAX.AppendLine("<path d=\"M -10 -1 L -1 -1 L -1 -10 L 1 -10 L 1 -1 L 10 -1 L 10 1 L 1 1 L 1 10 L -1 10 L -1 1 L -10 1 z\" stroke=\"black\" stroke-width=\"1\" />");
            return sbCAX.ToString();
        }

        private void SaveSVGWithoutHoles(string blockPath, string coneID)
        {
            double viewBox = SetViewBoxFromApplicator(coneID);

            StringBuilder sbBlock = new StringBuilder();
            sbBlock.AppendLine("<?xml version=\"1.0\" standalone=\"no\"?>");
            sbBlock.AppendLine("<svg xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" " +
                "width=\"" + viewBox.ToString() + "mm\" height=\"" + viewBox.ToString() + "mm\" " +
                "viewBox=\"" + (-0.5 * viewBox).ToString() + " " + (-0.5 * viewBox).ToString() + " " +
                viewBox.ToString() + " " + viewBox.ToString() + "\" preserveAspectRatio=\"minXminY meet\" >");
            sbBlock.Append(blockPath);
            sbBlock.AppendLine("</svg>");

            string filePath = @"\\thimage-16\va_transfer\blockexports\block_noHoles_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".svg";

            try
            { File.AppendAllText(filePath, sbBlock.ToString()); }
            catch (Exception err)
            { MessageBox.Show("Error writing data to file.\n" + err.Message); }
        }
    }
}
