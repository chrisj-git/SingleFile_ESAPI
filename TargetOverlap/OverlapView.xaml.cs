using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using System.Threading.Tasks;

namespace TargetOverlap
{
    public class Overlap
    {
        public string Structure { get; set; }
        
        private double _overlapFraction;
        public double OverlapFraction
        {
            get => _overlapFraction;
            
            set
            { _overlapFraction = value; OverlapPercentage = $"{_overlapFraction:P1}"; }
        }
        // Computed Property
        public string OverlapPercentage { get; set; }
        
        public Overlap(string name, double overlapFraction)
        {
            Structure = name;
            OverlapFraction = overlapFraction;
        }
    }

    /// <summary>
    /// Interaction logic for OverlapView.xaml
    /// </summary>
    public partial class OverlapView : Window
    {
        public string SelectedStructure { get; private set; }
        public StructureSet _StructureSet { get; set; }
        public ObservableCollection<Overlap> Overlaps { get; } = new ObservableCollection<Overlap>();

        public OverlapView()
        {
            InitializeComponent();
            DataContext = this;
        }

        private async void Calculate_Button_Click(object sender, RoutedEventArgs e)
        {
            Calculate_Button.IsEnabled = false;
            try
            {
                if (Structures_comboBox.SelectedItem is string selectedItem)
                {
                    SelectedStructure = selectedItem;
                    Overlaps.Clear();

                    var progress = new Progress<Overlap>(item =>
                    {
                        Overlaps.Add(item);
                    });

                    await CalculateOverlapsAsync(progress);
                }
                else
                { MessageBox.Show("Please select a structure."); }
            }
            finally { Calculate_Button.IsEnabled = true; }
        }

        private void CleanUp_TempStructs(List<string> tempIDs)
        {
            foreach (string tempID in tempIDs)
            {
                Structure structToDelete = _StructureSet.Structures.FirstOrDefault(x => x.Id == tempID);
                if (structToDelete != null)
                {
                    if (_StructureSet.CanRemoveStructure(structToDelete))
                    { _StructureSet.RemoveStructure(structToDelete); }
                }
            }
        }

        private async Task CalculateOverlapsAsync(IProgress<Overlap> progress)
        {
            List<string> structsCreated = new List<string>();

            Structure targetStruct = _StructureSet.Structures.FirstOrDefault(x => x.Id == SelectedStructure);

            Structure targetHighRes = _StructureSet.Structures.FirstOrDefault(x => x.Id == "zTarget_HighRes");
            if (targetHighRes == null)
            {
                string newID = "zTarget_HighRes";
                if (_StructureSet.CanAddStructure("CONTROL", newID))
                { 
                    targetHighRes = _StructureSet.AddStructure("CONTROL", newID); 
                    structsCreated.Add(targetHighRes.Id); 
                }
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

            Structure overlapStruct = _StructureSet.Structures.FirstOrDefault(x => x.Id == "zTemp_Boolean");
            if (overlapStruct == null)
            {
                string newID = "zTemp_Boolean";
                if (_StructureSet.CanAddStructure("CONTROL", newID))
                {
                    overlapStruct = _StructureSet.AddStructure("CONTROL", newID);
                    structsCreated.Add(overlapStruct.Id);
                }
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

            foreach (Structure s in _StructureSet.Structures)
            {
                string ID = s.Id.ToLower();

                // ---- Skip structures we don't need to compute overlap for -- BODY, Target itself, Z's, Wires, BBs
                if (ID == "body" || s.Id == targetStruct.Id || ID.StartsWith("z") || ID.StartsWith("BB") || ID.StartsWith("wire"))
                    continue;

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
                if (overlapStruct.Volume >= 0.0005) // rounds to 0.1% or more
                {
                    double overlapFraction = overlapStruct.Volume / targetVolume;
                    progress.Report(new Overlap(s.Id, overlapFraction));
                }

            }
            CleanUp_TempStructs(structsCreated);
            return;
        }
    }
}
