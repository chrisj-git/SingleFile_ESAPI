using System;
using System.Linq;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: ESAPIScript(IsWriteable = true)]
namespace VMS.TPS
{
    public class Script
    {
        public Script()
        { }

        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            if (context.ExternalPlanSetup == null)
            { MessageBox.Show("Please open a plan before running this script."); return; }

            Beam txBeam = context.ExternalPlanSetup.Beams.FirstOrDefault();
            if (txBeam == null)
            { MessageBox.Show("There must be at least one field to match isocenter with."); return; }

            ExternalBeamMachineParameters machineParamters = SetMachineParameters();
            DRRCalculationParameters boneDRRCalculationParameters = createBoneDRRParameters();
            
            string ap = "AP Setup";
            string rt = "Rt Setup";
            string cbct = "CBCT";

            try
            {
                context.Patient.BeginModifications();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Modifications not allowed.\n\n" + ex.Message);
            }

            Beam apSetup = context.ExternalPlanSetup.AddSetupBeam(machineParamters, new VRect<double>(-100, -100, 100, 100), 0, 0, 0, txBeam.IsocenterPosition);
            Beam rtSetup = context.ExternalPlanSetup.AddSetupBeam(machineParamters, new VRect<double>(-100, -100, 100, 100), 0, 270, 0, txBeam.IsocenterPosition);
            Beam cbctSetup = context.ExternalPlanSetup.AddSetupBeam(machineParamters, new VRect<double>(-100, -100, 100, 100), 0, 0, 0, txBeam.IsocenterPosition);
            
            if (context.ExternalPlanSetup.Beams.Any(x => 
                x.Id.ToLower() == ap.ToLower() || 
                x.Id.ToLower() == rt.ToLower() || 
                x.Id.ToLower() == cbct.ToLower()
                ))
            {
                apSetup.Id += "_AP Setup";
                rtSetup.Id += "_Rt Setup";
                cbctSetup.Id += "_CBCT";
            }
            else
            {
                apSetup.Id = "AP Setup";
                rtSetup.Id = "Rt Setup";
                cbctSetup.Id = "CBCT";
            }

            apSetup.CreateOrReplaceDRR(boneDRRCalculationParameters);
            rtSetup.CreateOrReplaceDRR(boneDRRCalculationParameters);
            MessageBox.Show($"Please move the three newly added setup fields to the top of the Field Order list:\n\n" +
                "New Fields:\n" +
                $"{apSetup.Id}\n{rtSetup.Id}\n{cbctSetup.Id}", 
                "Setup fields added", MessageBoxButton.OK, MessageBoxImage.Information);


            ExternalBeamMachineParameters SetMachineParameters()
            {
                return new ExternalBeamMachineParameters(
                     txBeam.TreatmentUnit.Id, "6X", 600, "STATIC", "");
            }

            DRRCalculationParameters createBoneDRRParameters()
            {
                return new DRRCalculationParameters(500, 1.0, 100, 1000);
            }
        }
    }
}