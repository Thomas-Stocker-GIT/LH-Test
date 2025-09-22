using System;
using System.Collections.Generic;
using System.Data;
using _360DegreeTestSuiteDemo.ObjectRepository;
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

namespace _360DegreeTestSuiteDemo.CODEDTESTCASES.Selectors
{
    public class Type_Options : CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
            TypeIntoOptions type = new TypeIntoOptions();
            type.Text = "john@doe.com";

            type.ActivateBefore = true;
            type.AlterIfDisabled = true;
            type.ClickBeforeMode = NClickMode.Single;
            type.ContinueOnError = true;
            type.DelayAfter = 1000;
            type.DelayBetweenKeys = 100;
            type.EmptyFieldMode = NEmptyFieldMode.None;


            var formScreen = uiAutomation.Attach("Form");
            formScreen.TypeInto("Email Address", type);


        }
    }
}