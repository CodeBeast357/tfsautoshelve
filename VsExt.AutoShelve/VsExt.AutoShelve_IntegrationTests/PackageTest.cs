using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Windows.Threading;
using ThreadingTask = System.Threading.Tasks.Task;

namespace VsExt.AutoShelve_IntegrationTests
{
    /// <summary>
    /// Integration test for package validation
    /// </summary>
    [TestClass]
    public class PackageTest
    {
        protected static Dispatcher UIThreadDispatcher { get; private set; }

        [AssemblyInitialize]
        public static void AssemblyInitialize()
        {
            ThreadHelper.JoinableTaskFactory.Run(() => ThreadingTask.Run(() => UIThreadDispatcher = Dispatcher.CurrentDispatcher));
        }

        [TestMethod]
        [HostType("VS IDE")]
        public void PackageLoadTest()
        {
            //Get the Shell Service
            var serviceProvider = ServiceProvider.GlobalProvider;

            var shellService = (IVsShell)serviceProvider.GetService(typeof(SVsShell));
            Assert.IsNotNull(shellService);

            //Validate package load
            var packageGuid = new Guid(AutoShelve.Packaging.GuidList.GuidAutoShelvePkgString);
            Assert.IsTrue(0 == shellService.LoadPackage(ref packageGuid, out IVsPackage package));
            Assert.IsNotNull(package, "Package failed to load");
        }
    }
}
