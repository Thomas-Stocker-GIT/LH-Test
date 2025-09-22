using _360DegreeTestSuiteDemo.ObjectRepository;
using System;
using System.Collections.Generic;
using System.Data;
using UiPath.CodedWorkflows;
using UiPath.Core;
using UiPath.Core.Activities.Storage;
using UiPath.Excel;
using UiPath.Excel.Activities;
using UiPath.Excel.Activities.API;
using UiPath.Excel.Activities.API.Models;
using UiPath.Orchestrator.Client.Models;
using UiPath.Testing;
using UiPath.Testing.Activities.TestData;
using UiPath.Testing.Activities.TestDataQueues.Enums;
using UiPath.Testing.Enums;
using UiPath.UIAutomationNext.API.Contracts;
using UiPath.UIAutomationNext.API.Models;
using UiPath.UIAutomationNext.Enums;

namespace _360DegreeTestSuiteDemo
{
    public class Workflow : CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
           var screen = uiAutomation.Open(Descriptors.__Chrome__UiBank_Welcome.Chrome__UiBank_Welcome);
            screen.Click("Products");
            screen.Click("Loans", new ClickOptions().WithVariable("aaname", "Loans"));
        }
    }
}