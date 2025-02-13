using System;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
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
            if (context.StructureSet == null)
            { 
                MessageBox.Show("Please open a structure set before running this script."); 
                return; 
            }

            try
            { context.Patient.BeginModifications(); }
            catch (Exception ex)
            { MessageBox.Show("Modifications not allowed.\n\n" + ex.Message); }

            ShowWindow(context.StructureSet);
        }

        private void ShowWindow(StructureSet structureSet)
        {
            OverlapView overlapView = new OverlapView();
            overlapView._StructureSet = structureSet;
            overlapView.Structures_comboBox.ItemsSource = structureSet.Structures.Select(x => x.Id);
            overlapView.ShowDialog();
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

        private List<Overlap> CalculateOverlaps(ScriptContext context, Structure targetStruct)
        {
            

            Structure targetHighRes = context.ExternalPlanSetup.StructureSet.Structures.FirstOrDefault(x => x.Id == "zTarget_HighRes");
            if (targetHighRes == null)
            {
                if (context.StructureSet.CanAddStructure("CONTROL", "zTarget_HighRes"))
                { targetHighRes = context.StructureSet.AddStructure("CONTROL", "zTarget_HighRes"); }
                else
                { MessageBox.Show("Can't add temp structure for boolean operations! (zTarget_HighRes)"); return null; }
            }

            targetHighRes.SegmentVolume = targetStruct.SegmentVolume;

            if (!targetHighRes.IsHighResolution)
            {
                if (targetHighRes.CanConvertToHighResolution())
                { targetHighRes.ConvertToHighResolution(); }
                else
                { MessageBox.Show("Can't convert " + targetHighRes.Id + " to high resolution for boolean operations!"); return null; }
            }

            double targetVolume = targetStruct.Volume;

            Structure overlapStruct = context.ExternalPlanSetup.StructureSet.Structures.FirstOrDefault(x => x.Id == "zTemp_Boolean");
            if (overlapStruct == null)
            {
                if (context.StructureSet.CanAddStructure("CONTROL", "zTemp_Boolean"))
                { overlapStruct = context.StructureSet.AddStructure("CONTROL", "zTemp_Boolean"); }
                else
                { MessageBox.Show("Can't add temp structure for boolean operations! (zTemp_Boolean)"); return null; }
            }

            if (!overlapStruct.IsHighResolution)
            {
                if (overlapStruct.CanConvertToHighResolution())
                { overlapStruct.ConvertToHighResolution(); }
                else
                { MessageBox.Show("Can't convert " + overlapStruct.Id + " to high resolution for boolean operations!"); return null; }
            }

            List<Overlap> overlapList = new List<Overlap>();

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
                        double overlapFraction = overlapStruct.Volume / targetVolume;
                        overlapList.Add(new Overlap(s.Id, Math.Round(overlapFraction, 3)));
                    }
                }
            }
            return overlapList;
        }

        private void ShowResultView(string target_ID, List<Overlap> overlapList)
        {
            ResultView resultView = new ResultView();
            resultView.Target_Label.Content = target_ID;
            resultView.Result_DataGrid.ItemsSource = overlapList;
            resultView.Result_DataGrid.IsReadOnly = true;
            resultView.ShowDialog();
        }
    }

    public class Overlap
    {
        public string Structure { get; set; }
        public double OverlapFraction { get; set; }
        public Overlap(string name, double overlapFraction)
        {
            Structure = name;
            OverlapFraction = overlapFraction;
        }
    }
}
