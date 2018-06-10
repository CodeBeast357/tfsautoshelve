﻿using EnvDTE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VsSDK.IntegrationTestLibrary;
using Microsoft.VSSDK.Tools.VsIdeTesting;
using System;

namespace VsExt.AutoShelve_IntegrationTests
{
    [TestClass]
    public class VisualBasicProjectTests
    {
        #region fields

        private delegate void ThreadInvoker();

        #endregion

        #region properties

        /// <summary>
        ///     Gets or sets the test context which provides
        ///     information about and functionality for the current test run.
        /// </summary>
        public TestContext TestContext { get; set; }

        #endregion

        #region ctors

        #endregion

        [HostType("VS IDE")]
        [TestMethod]
        public void VbWinformsApplication()
        {
            UIThreadInvoker.Invoke((ThreadInvoker)delegate
           {
               //Solution and project creation parameters
               const string solutionName = "VBWinApp";
               const string projectName = "VBWinApp";

               //Template parameters
               const string language = "VisualBasic";
               const string projectTemplateName = "WindowsApplication.Zip";
               const string itemTemplateName = "CodeFile.zip";
               const string newFileName = "Test.vb";

               var dte = (DTE)VsIdeTestHostContext.ServiceProvider.GetService(typeof(DTE));

               var testUtils = new TestUtils();

               testUtils.CreateEmptySolution(TestContext.TestDir, solutionName);
               Assert.AreEqual(0, testUtils.ProjectCount());

               //Add new  Windows application project to existing solution
               testUtils.CreateProjectFromTemplate(projectName, projectTemplateName, language, false);

               //Verify that the new project has been added to the solution
               Assert.AreEqual(1, testUtils.ProjectCount());

               //Get the project
               Project project = dte.Solution.Item(1);
               Assert.IsNotNull(project);
               Assert.IsTrue(string.Compare(project.Name, projectName, StringComparison.InvariantCultureIgnoreCase)
                             == 0);

               //Verify Adding new code file to project
               ProjectItem newCodeFileItem = testUtils.AddNewItemFromVsTemplate(
                   project.ProjectItems,
                   itemTemplateName,
                   language,
                   newFileName);
               Assert.IsNotNull(newCodeFileItem, "Could not create new project item");
           });
        }
    }
}