using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using VMS.TPS.Common.Model.API;

namespace TargetOverlap
{
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

    /// <summary>
    /// Interaction logic for OverlapView.xaml
    /// </summary>
    public partial class OverlapView : Window
    {
        public string SelectedOption { get; private set; }
        public StructureSet _StructureSet { get; set; }
        public ObservableCollection<Overlap> Overlaps { get; set; }

        public OverlapView()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Structures_comboBox.SelectedItem is string selectedItem)
            {
                SelectedOption = selectedItem;
                CalculateOverlaps();
            }
            else
            { MessageBox.Show("Please select an option."); }
        }

        private void CalculateOverlaps()
        {
            Structure targetStruct = _StructureSet.Structures.FirstOrDefault(x => x.Id == SelectedOption);

            Structure targetHighRes = _StructureSet.Structures.FirstOrDefault(x => x.Id == "zTarget_HighRes");
            if (targetHighRes == null)
            {
                if (_StructureSet.CanAddStructure("CONTROL", "zTarget_HighRes"))
                { targetHighRes = _StructureSet.AddStructure("CONTROL", "zTarget_HighRes"); }
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
                if (_StructureSet.CanAddStructure("CONTROL", "zTemp_Boolean"))
                { overlapStruct = _StructureSet.AddStructure("CONTROL", "zTemp_Boolean"); }
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
                        Overlaps.Add(new Overlap(s.Id, Math.Round(overlapFraction, 3)));
                    }
                }
            }
            return;
        }
    }
}
