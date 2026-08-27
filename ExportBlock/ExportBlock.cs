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
                    double windSum = 0.0;


                    StringBuilder sbBlock = new StringBuilder();
                    StringBuilder sbBlockPath = new StringBuilder();
                    string appId = context.ExternalPlanSetup.Beams.First().Applicator.Id;
                    if (appId == "A06")
                    { appId = "A"; }
                    else if (appId == "A10")
                    { appId = "B"; }
                    else if (appId == "A15")
                    { appId = "C"; }
                    else if (appId == "A20")
                    { appId = "D"; }
                    else if (appId == "A25")
                    { appId = "E"; }
                    else
                    { appId = "UnknownAppID"; }

                    sbBlockPath.Append("<path d =\"");

                    //Flip Y across the X-axis, since SVG coordinate system has +Y going down, by using -1 multiplier on Y coordinates
                    //Flip X across the Y-axis, since a 'flipped' block prints with the smoothest side (print bed side) down against the tray, by using -1 multiplier on X coordinates
                    //windSum is used to determine if clockwise or CCW winding order negative is CCW positive is Clockwise
                    //Scale all points by 0.95 since the outline coordinates of the block are at 100cm and the block needs to print at 95cm size
                    for (int i = 0; i < points[0].Length; i++)
                    {

                        points[0][i] = new Point(points[0][i].X * -0.95, points[0][i].Y * -0.95);
                        if (i < points[0].Length - 1)
                        { windSum += (points[0][i].X * (points[0][i + 1].Y * -0.95)) - ((points[0][i + 1].X * -0.95) * points[0][i].Y); }
                        else // wrapping around to the first point when we reach the end
                        { windSum += (points[0][i].X * (points[0][0].Y * -0.95)) - ((points[0][0].X * -0.95) * points[0][i].Y); }
                        if (i == 0)
                        { sbBlockPath.Append("M " + points[0][i].X + " " + points[0][i].Y); }
                        else
                        { sbBlockPath.Append(" L " + points[0][i].X + " " + points[0][i].Y); }
                    }

                    //Add locating L and Pin
                    if (windSum < 0)
                    { sbBlockPath.Append(AddLocatingJigFromApplicatorCCW(context.ExternalPlanSetup.Beams.First().Applicator.Id)); }
                    else
                    { sbBlockPath.Append(AddLocatingJigFromApplicatorCW(context.ExternalPlanSetup.Beams.First().Applicator.Id)); }
                    sbBlockPath.Append(" z\" />\n");

                    //Start the XML and set the viewbox
                    sbBlock.AppendLine("<?xml version=\"1.0\" standalone=\"no\"?>");
                    sbBlock.AppendLine("<svg xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" " +
                        "width=\"" + viewBox.ToString() + "mm\" height=\"" + viewBox.ToString() + "mm\" " +
                        "viewBox=\"" + (-0.5 * viewBox).ToString() + " " + (-0.5 * viewBox).ToString() + " " +
                        viewBox.ToString() + " " + viewBox.ToString() + "\" preserveAspectRatio=\"minXminY meet\" >");

                    //Add the block path
                    sbBlock.Append(sbBlockPath);



                    //end XML
                    sbBlock.AppendLine("</svg>");

                    string filePath = @"\\C0103AE1-SF\VA_Transfer\GRTC Transfer\blockexports\block_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + appId + ".svg";

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

        private string AddLocatingJigFromApplicatorCW(string id)
        {
            if (id == "A20" | id == "A25")
                return AddCAXMarkOnly();

            StringBuilder sbJig = new StringBuilder();
            // Adding the 'L-hole' ** remember that the block is printing upside down -- so these X coordinates are reflected from what you will see on the tray
            // starting the 'L' here
            sbJig.AppendLine(" M 6.1 2.9 L -2.9 2.9 L -2.9 -3.1 L -6.1 -3.1 L -6.1 6.1");

            // Adding the arc and finishing the L

            switch (id)
            {
                case "A06":
                    sbJig.AppendLine(" L -5.5,6.1 A 1.5,1.5 0 1,0, -2.5,6.1 L 6.1 6.1");
                    return sbJig.ToString();
                case "A10":
                    sbJig.AppendLine(" L -3.5,6.1 A 1.5,1.5 0 1,0, -0.5,6.1 L 6.1 6.1");
                    return sbJig.ToString();
                case "A15":
                    sbJig.AppendLine(" L -1.5,6.1 A 1.5,1.5 0 1,0, 1.5,6.1 L 6.1 6.1");
                    return sbJig.ToString();
                case "A20":
                    sbJig.AppendLine(" L 0.5,6.1 A 1.5,1.5 0 1,0, 3.5,6.1 L 6.1 6.1");
                    return sbJig.ToString();
                case "A25":
                    sbJig.AppendLine(" L 2.5,6.1 A 1.5,1.5 0 1,0, 5.5,6.1 L 6.1 6.1");
                    return sbJig.ToString();
                default:
                    return "";
            }
        }

        private string AddLocatingJigFromApplicatorCCW(string id)
        {
            if (id == "A20" | id == "A25")
                return AddCAXMarkOnly();

            StringBuilder sbJig = new StringBuilder();
            // Adding the 'L-hole' ** remember that the block is printing upside down -- so these X coordinates are reflected from what you will see on the tray
            // starting the 'L' here
            sbJig.AppendLine(" M 6.1 2.9 L 6.1 6.1");

            // Adding the arc and finishing the L

            switch (id)
            {
                case "A06":
                    sbJig.AppendLine(" L -2.5,6.1 A 1.5,1.5 0 1,1, -5.5,6.1 L -6.1 6.1 L -6.1 -3.1 L -2.9 -3.1 L -2.9 2.9 L 6.1 2.9");
                    return sbJig.ToString();
                case "A10":
                    sbJig.AppendLine(" L -0.5,6.1 A 1.5,1.5 0 1,1, -3.5,6.1 L -6.1 6.1 L -6.1 -3.1 L -2.9 -3.1 L -2.9 2.9 L 6.1 2.9");
                    return sbJig.ToString();
                case "A15":
                    sbJig.AppendLine(" L 1.5,6.1 A 1.5,1.5 0 1,1, -1.5,6.1 L -6.1 6.1 L -6.1 -3.1 L -2.9 -3.1 L -2.9 2.9 L 6.1 2.9");
                    return sbJig.ToString();
                case "A20":
                    sbJig.AppendLine(" L 3.5,6.1 A 1.5,1.5 0 1,1, 0.5,6.1 L -6.1 6.1 L -6.1 -3.1 L -2.9 -3.1 L -2.9 2.9 L 6.1 2.9");
                    return sbJig.ToString();
                case "A25":
                    sbJig.AppendLine(" L 5.5,6.1 A 1.5,1.5 0 1,1, 2.5,6.1 L -6.1 6.1 L -6.1 -3.1 L -2.9 -3.1 L -2.9 2.9 L 6.1 2.9");
                    return sbJig.ToString();
                default:
                    return "";
            }
        }

        private string AddCAXMarkOnly()
        {
            StringBuilder sbCAX = new StringBuilder();
            sbCAX.AppendLine(" M -10 -1 L -1 -1 L -1 -10 L 1 -10 L 1 -1 L 10 -1 L 10 1 L 1 1 L 1 10 L -1 10 L -1 1 L -10 1");
            return sbCAX.ToString();
        }

        private void SaveSVGWithoutHoles(string blockPath, string coneID)
        {
            double viewBox = SetViewBoxFromApplicator(coneID);

            string appId = coneID;
            if (appId == "A06")
            { appId = "A"; }
            else if (appId == "A10")
            { appId = "B"; }
            else if (appId == "A15")
            { appId = "C"; }
            else if (appId == "A20")
            { appId = "D"; }
            else if (appId == "A25")
            { appId = "E"; }
            else
            { appId = ""; }

            StringBuilder sbBlock = new StringBuilder();
            sbBlock.AppendLine("<?xml version=\"1.0\" standalone=\"no\"?>");
            sbBlock.AppendLine("<svg xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" " +
                "width=\"" + viewBox.ToString() + "mm\" height=\"" + viewBox.ToString() + "mm\" " +
                "viewBox=\"" + (-0.5 * viewBox).ToString() + " " + (-0.5 * viewBox).ToString() + " " +
                viewBox.ToString() + " " + viewBox.ToString() + "\" preserveAspectRatio=\"minXminY meet\" >");
            sbBlock.Append(blockPath);
            sbBlock.AppendLine("</svg>");

            string filePath = @"\\C0103AE1-SF\VA_Transfer\GRTC Transfer\blockexports\block_noHoles_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + appId + ".svg";

            try
            { File.AppendAllText(filePath, sbBlock.ToString()); }
            catch (Exception err)
            { MessageBox.Show("Error writing data to file.\n" + err.Message); }
        }
    }
}
