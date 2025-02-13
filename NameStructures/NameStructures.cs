using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
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

            if (context.StructureSet != null)
            {

                try
                {
                    context.Patient.BeginModifications();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Modifications not allowed.\n\n" + ex.Message);
                }
                
                foreach ( var item in context.StructureSet.Structures )
                {
                    string errorMessage = "";
                    MessageBox.Show(string.Format("Struct: {0} -- Code: {1} -- Color: {2}", item.Id, item.StructureCode.Code, item.Color.ToString()));
                    if (item.CanSetAssignedHU(out errorMessage))
                    {
                        if (item.Id.ToLower().Contains("ptv"))
                        {
                            item.Color = Colors.YellowGreen;
                        }
                    }
                    else
                    { MessageBox.Show(errorMessage); }
                }
            }
        }
    }
}