using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(VSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string)]
    public sealed class VSPackage : AsyncPackage
    {
        /// <summary>
        /// VSPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "685129d8-626a-4201-b4d2-51796691600d";

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage"/> class.
        /// </summary>
        public VSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var solution = await this.GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            await this.WriteProjectInfoAsync(solution);
        }

        private async System.Threading.Tasks.Task WriteProjectInfoAsync(IVsSolution solution)
        {
            solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_ALLINSOLUTION, Guid.Empty, out IEnumHierarchies enumHierarchies);

            IVsHierarchy[] hierarchies = new IVsHierarchy[1];
            uint hierarchiesRetrieved = 0;

            while (enumHierarchies.Next(1, hierarchies, out hierarchiesRetrieved) == 0)
            {
                if (hierarchies[0] is IVsHierarchy2)
                {
                    var properties = new int[] { (int)__VSHPROPID.VSHPROPID_Name, (int)__VSHPROPID.VSHPROPID_ProjectDir };
                    var values = new object[properties.Length];
                    var results = new int[properties.Length];

                    try
                    {
                        ((IVsHierarchy2)hierarchies[0]).GetProperties((uint)VSConstants.VSITEMID.Root, 2, properties, values, results);

                        if (ErrorHandler.Succeeded(results[0]))
                        {
                            string name = values[0] as string;
                            await this.WriteToOutputWindowAsync("\tProject name: " + name + Environment.NewLine);
                        }

                        if (ErrorHandler.Succeeded(results[1]))
                        {
                            string projectDir = values[1] as string;
                            await this.WriteToOutputWindowAsync("\tProject dir : " + projectDir + Environment.NewLine);
                        }

                        var guidProperties = new int[] { (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, (int)__VSHPROPID.VSHPROPID_TypeGuid };
                        var guidValues = new Guid[properties.Length];

                        ((IVsHierarchy2)hierarchies[0]).GetGuidProperties((uint)VSConstants.VSITEMID.Root, 2, guidProperties, guidValues, results);

                        if (ErrorHandler.Succeeded(results[0]))
                        {
                            var projectId = (Guid)guidValues[0];
                            await this.WriteToOutputWindowAsync("\tProject id  : " + projectId + Environment.NewLine);
                        }

                        if (ErrorHandler.Succeeded(results[1]))
                        {
                            var projectType = (Guid)guidValues[1];
                            await this.WriteToOutputWindowAsync("\tProject type: " + projectType + Environment.NewLine);
                        }

                        await this.WriteToOutputWindowAsync(Environment.NewLine);
                    }
                    catch (Exception e)
                    {
                        await this.WriteToOutputWindowAsync("\tException: " + e.Message + Environment.NewLine);
                        continue;
                    }
                }
            }
        }

        private async Task WriteToOutputWindowAsync(string output)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync();

            var outputWindow = await this.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;

            Guid paneGuid = VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            int hr = outputWindow.GetPane(ref paneGuid, out IVsOutputWindowPane pane);

            if (ErrorHandler.Failed(hr) || (pane == null))
            {
                if (ErrorHandler.Succeeded(outputWindow.CreatePane(ref paneGuid, "General", fInitVisible: 1, fClearWithSolution: 1)))
                {
                    hr = outputWindow.GetPane(ref paneGuid, out pane);
                }
            }

            if (ErrorHandler.Succeeded(hr))
            {
                pane?.Activate();
                pane?.OutputString(output);
            }
        }
    }

    internal class IVsHierarchy2Impl : IVsHierarchy2
    {
        void IVsHierarchy2.GetGuidProperties(uint itemid, int count, int[] propids, Guid[] rgGuids, int[] results)
        {
            throw new NotImplementedException();
        }

        void IVsHierarchy2.GetProperties(uint itemid, int count, int[] propids, object[] vars, int[] results)
        {
            throw new NotImplementedException();
        }
    }
}
