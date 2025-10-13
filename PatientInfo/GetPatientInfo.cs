using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

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
                string patientID = context.Patient.Id;
                string patientLastName = context.Patient.LastName;
                string patientFirstName = context.Patient.FirstName;
                string oncID = context.Patient.PrimaryOncologistId;
                string planID = context.ExternalPlanSetup.Id;
                string rxTotalDose = context.ExternalPlanSetup.TotalDose.ValueAsString;
                string numberFractions = context.ExternalPlanSetup.NumberOfFractions.Value.ToString();
                if (oncID == "1821059338")
                { oncID = "RZ"; }
                else if (oncID == "1437331733")
                { oncID = "NA"; }

                string data = string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}", patientLastName, patientFirstName, patientID, oncID, planID, rxTotalDose, numberFractions);

                MessageBox.Show(data);
                Clipboard.SetText(data);
            }
        }
    }
}
