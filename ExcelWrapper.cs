// Only these assemblies are allowed in Eplan scripts:
using System;
using System.Xml;
using System.Drawing;
using System.Windows.Forms;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Gui;

// However, "Late Binding" using Reflection allows to connect to other assemblies.
using System.Reflection;

public class ExcelDemo
{
    [Start]
    public void StartExcel()
    {
        // Create new Excel instance and add a workbook.
        // Using the Excel Wrapper for EPLAN Scripts this looks like COM Interop:
        Excel.Application xlsApp = new Excel.Application();
        xlsApp.Visible = true;
        xlsApp.Workbooks.Add();
        
        // Your code here...
    }
}


// Excel Wrapper for EPLAN Scripts
//
// To be continued...
// The whole code for an Eplan script must be in one file.
namespace Excel
{
    public class Application
    {
        private object _application;
        
        public Application()
        {
            Type type = Type.GetTypeFromProgID("Excel.Application");
            _application = Activator.CreateInstance(type);
        }
        
        public bool Visible
        {
            get
            {
                return (bool)_application.GetType().InvokeMember("Visible", BindingFlags.GetProperty , null, _application, null);
            }
            set
            {
                _application.GetType().InvokeMember("Visible", BindingFlags.SetProperty, Type.DefaultBinder, _application, new object[] {value});
            }
        }
        
        public Excel.Workbooks Workbooks
        {
            get
            {
                object _workbooks = _application.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty , null, _application, null);
                Excel.Workbooks workbooks = new Excel.Workbooks(_workbooks);
                return workbooks;
            }
        }
    }
    
    public class Workbooks
    {
        private object _workbooks;
        
        public Workbooks(object workbooks)
        {
            _workbooks = workbooks;
        }
        
        public void Add()
        {
            _workbooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, _workbooks, null);
        }
    }
}
