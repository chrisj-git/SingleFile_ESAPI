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
using TargetOverlap;

[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        { }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            if (context.ExternalPlanSetup == null)
            { 
                MessageBox.Show("Please open a plan before running this script."); 
                return; 
            }

            Structure targetStruct = ShowSelectionDialog(context.ExternalPlanSetup.StructureSet.Structures);
            
            if (targetStruct != null)
            { CalculateOverlaps(context, targetStruct); }
        }

        private Structure ShowSelectionDialog(IEnumerable<Structure> structures)
        {
            SelectionDialog dialog = new SelectionDialog();
            dialog.Structures_comboBox.ItemsSource = structures.Select(x => x.Id);

            bool? result = dialog.ShowDialog();
            if (result == true)
            {
                string selectedOption = dialog.SelectedOption;
                return structures.FirstOrDefault(x => x.Id == selectedOption);
            }
            else
            { MessageBox.Show("Selection was canceled."); return null; }
        }

        private void CalculateOverlaps(ScriptContext context, Structure targetStruct)
        {
            try
            {
                context.Patient.BeginModifications();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Modifications not allowed.\n\n" + ex.Message);
            }

            Structure targetHighRes = context.ExternalPlanSetup.StructureSet.Structures.FirstOrDefault(x => x.Id == "zTarget_HighRes");
            if (targetHighRes == null)
            {
                if (context.StructureSet.CanAddStructure("CONTROL", "zTarget_HighRes"))
                { targetHighRes = context.StructureSet.AddStructure("CONTROL", "zTarget_HighRes"); }
                else
                { MessageBox.Show("Can't add temp structure for boolean operations! (zTarget_HighRes)"); return; }
            }

            targetHighRes.SegmentVolume = targetStruct.SegmentVolume;

            if (!targetHighRes.IsHighResolution)
            {
                if (targetHighRes.CanConvertToHighResolution())
                { targetHighRes.ConvertToHighResolution(); }
                else
                { MessageBox.Show("Can't convert " + targetHighRes.Id + " to high resolution for boolean operations!"); return; }
            }

            double targetVolume = targetStruct.Volume;

            Structure overlapStruct = context.ExternalPlanSetup.StructureSet.Structures.FirstOrDefault(x => x.Id == "zTemp_Boolean");
            if (overlapStruct == null)
            {
                if (context.StructureSet.CanAddStructure("CONTROL", "zTemp_Boolean"))
                { overlapStruct = context.StructureSet.AddStructure("CONTROL", "zTemp_Boolean"); }
                else
                { MessageBox.Show("Can't add temp structure for boolean operations! (zTemp_Boolean)"); return; }
            }

            if (!overlapStruct.IsHighResolution)
            {
                if (overlapStruct.CanConvertToHighResolution())
                { overlapStruct.ConvertToHighResolution(); }
                else
                { MessageBox.Show("Can't convert " + overlapStruct.Id + " to high resolution for boolean operations!"); return; }
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("Target: {0}\tVolume: {1}\n", targetStruct.Id, targetVolume.ToString("F3")));

            foreach (Structure s in context.ExternalPlanSetup.StructureSet.Structures)
            {
                string ID = s.Id.ToLower();
                if (ID != "body" && s.Id != targetStruct.Id && !ID.StartsWith("z") && !ID.StartsWith("BB") && !ID.StartsWith("wire"))
                {
                    SegmentVolume overlapSegmentVolume;
                    if (s.IsHighResolution)
                    {
                        overlapSegmentVolume = targetHighRes.SegmentVolume.And(s.SegmentVolume);
                    }
                    else // s is not high resolution
                    {
                        overlapStruct.SegmentVolume = s.SegmentVolume;
                        if (overlapStruct.CanConvertToHighResolution())
                        {
                            overlapStruct.ConvertToHighResolution();
                            overlapSegmentVolume = targetHighRes.SegmentVolume.And(overlapStruct.SegmentVolume);
                        }
                        else
                        { MessageBox.Show("Can't convert " + s.Id + " to high resolution for boolean operations!"); continue; }
                    }

                    overlapStruct.SegmentVolume = overlapSegmentVolume;
                    if (overlapStruct.Volume != 0)
                    {
                        double overlapPercent = overlapStruct.Volume / targetVolume;

                        sb.AppendLine(string.Format("Struct: {0}\tOverlap: {1}",
                            s.Id, overlapPercent.ToString("P2")));
                    }
                }
            }
            MessageBox.Show(sb.ToString(), "Percent of target structure overlapping other structures.", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
