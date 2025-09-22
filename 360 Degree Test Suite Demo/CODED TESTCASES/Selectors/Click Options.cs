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
    public class Click_Options : CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
           
            ClickOptions click = new ClickOptions();
            

            var homeScreen = uiAutomation.Open("Home");
            
            //Single Click
            homeScreen.Click("mySelectorReference");
            
            //Double Click
            homeScreen.Click("mySelectorReference",NClickType.Double);
            
            //Long Click
            homeScreen.Click("Products",NClickType.Down);
            System.Threading.Thread.Sleep(3000);
            homeScreen.Click("Products", NClickType.Up);
            
            
            //cursor motion
            click.CursorMotionType = CursorMotionType.Instant;
            click.CursorMotionType = CursorMotionType.Smooth;
            
            
            //Interaction Mode           
            click.InteractionMode = NChildInteractionMode.Simulate;
            click.InteractionMode = NChildInteractionMode.WindowMessages;     
            click.InteractionMode = NChildInteractionMode.HardwareEvents;
            
            // Misc Options
            click.ActivateBefore = true;
            click.ContinueOnError = true;
            click.DelayBefore = 1000;
            click.DelayAfter = 1000;
            
            
             
            
            
        }
    }
} 