//**********************
//Code developed by Ichchha Gupta
//last commit : 19-04-24, 5:31PM
//**********************

// Importing required namespaces
using addin.Properties;
using CodeStack.SwEx.AddIn;
using CodeStack.SwEx.AddIn.Attributes;
using CodeStack.SwEx.AddIn.Enums;
using CodeStack.SwEx.Common.Attributes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace addin
{
    // SolidWorks add-in class definition
    // Indicates that the class is visible to COM components
    // Specifies a unique identifier (GUID) for the class. This GUID is used to uniquely identify the add-in in the COM registry
    [ComVisible(true), Guid("611C5832-CB81-4122-AC78-954B26ABDA26")]
    [AutoRegister("Slicing", "Generates Slices")] //automatically registers solidworks addin in the registry, it could be recognized and loaded automatically
    public class AddIn : SwAddInEx
    {
        // Enum to define add-in commands
        [Title("Slicing")]
        [Description("Generates Slices")]
        [Icon(typeof(Resources), nameof(Resources.slices))]
        private enum Commands_e
        {
            [Title("Slicing")]
            [Description("Generates Slices")]
            [Icon(typeof(Resources), nameof(Resources.slices))]
            [CommandItemInfo(true, true, swWorkspaceTypes_e.Part, true)]
            GenerateSlices
        }

        // Method called when the add-in is connected to SolidWorks
        public override bool OnConnect()
        {
            // Add command group to SolidWorks command manager
            AddCommandGroup<Commands_e>(OnButtonClick);
            return true;
        }

        // Method to handle button click events
        private void OnButtonClick(Commands_e cmd)
        {
            try
            {
                switch (cmd)
                {
                    case Commands_e.GenerateSlices:
                        RunSlicingMacro();
                        break;
                }
            }
            catch (Exception ex)
            {
                // Display error message to the user
                App.SendMsgToUser2(ex.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        private void RunSlicingMacro()
        {
            if (App is ISldWorks swApp)
            {
                var macroPath = @"E:\Desktop\BTP\macros\SlicingmacroOG.swp"; // Update with your macro path

                if (System.IO.File.Exists(macroPath))
                {
                    int macroError;
                    // Run the macro
                    swApp.RunMacro2(macroPath, "", "main", (int)swRunMacroOption_e.swRunMacroDefault, out macroError);

                    if (macroError != 0)
                    {
                        // Throw exception if macro execution fails
                        throw new Exception($"Macro execution failed with error code: {macroError}");
                    }
                }
                else
                {
                    // Throw exception if macro file not found
                    throw new Exception($"Macro file not found at {macroPath}");
                }
            }
            else
            {
                // Throw exception if SolidWorks application object retrieval fails
                throw new Exception("Failed to retrieve SolidWorks application object");
            }
        }

    }
}
