using _360DegreeTestSuiteDemo.CODEDTESTCASES.Selectors;
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

namespace _360DegreeTestSuiteDemo.CODEDTESTCASES
{
    public class Demo : CodedWorkflow
    {
        [TestCase]
        public void Execute()
        {
             UiBank_Web selectors = new UiBank_Web();
            ClickOptions click = new ClickOptions();
            


            var homeScreen = uiAutomation.Open("Home");
            homeScreen.Click("Products",NClickType.Down);
            System.Threading.Thread.Sleep(3000);
            homeScreen.Click("Products", NClickType.Up);
        }
    }
}