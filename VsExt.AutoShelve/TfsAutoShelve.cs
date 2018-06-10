using EnvDTE;
using EnvDTE80;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.VisualStudio.TeamFoundation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using VsExt.AutoShelve.EventArgs;

// Ref http://visualstudiogallery.msdn.microsoft.com/080540cb-e35f-4651-b71c-86c73e4a633d
namespace VsExt.AutoShelve
{
    public class TfsAutoShelve : ISAutoShelve, IAutoShelve, IDisposable
    {
        private static string _extensionName => Resources.ExtensionName;

        private readonly Timer _timer;
        private readonly IServiceProvider _serviceProvider;

        private TimerCallback AutoShelveCallback => _ => CreateShelveset();

        public TfsAutoShelve(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
            var autoResetEvent = new AutoResetEvent(false);
            _timer = new Timer(AutoShelveCallback, autoResetEvent, Timeout.Infinite, Timeout.Infinite);
        }

        #region IAutoShelveService

        public bool IsRunning { get; private set; }

        private ushort _maximumShelvesets;

        public ushort MaximumShelvesets
        {
            get
            {
                return ShelvesetName.IsNameSpecificToDate() ? _maximumShelvesets : (ushort)0;
            }
            set
            {
                _maximumShelvesets = value;
            }
        }

        public string ShelvesetName { get; set; }

        public double TimerInterval { get; set; }

        public void CreateShelveset(bool force = false)
        {
            try
            {
                if (ProjectCollectionUri == null) return;
                var teamProjectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(ProjectCollectionUri);
                teamProjectCollection.Credentials = CredentialCache.DefaultNetworkCredentials;
                teamProjectCollection.EnsureAuthenticated();

                var service = (VersionControlServer)teamProjectCollection.GetService(typeof(VersionControlServer));
                var infos = new AutoShelveInfos(this);
                foreach (var workspaceInfo in Workstation.Current.GetAllLocalWorkspaceInfo())
                {
                    if (workspaceInfo.MappedPaths.Length <= 0 || !workspaceInfo.ServerUri.Equals(ProjectCollectionUri))
                    {
                        continue;
                    }

                    Task.Run(() =>
                    {
                        var result = CreateShelvesetInternal(service, workspaceInfo, infos, force);
                        if (!result.IsSuccess)
                        {
                            InvalidateConnection(); // Force re-init on next attempt
                        }
                        ShelvesetCreated?.Invoke(this, result);
                    });
                }
            }
            catch (Exception ex)
            {
                InvalidateConnection(); // Force re-init on next attempt
                TfsShelvesetErrorReceived?.Invoke(this, new TfsShelvesetErrorEventArgs(error: ex));
            }
        }

        #region Events

        public event EventHandler<ShelvesetCreatedEventArgs> ShelvesetCreated;

        public event EventHandler<TfsShelvesetErrorEventArgs> TfsShelvesetErrorReceived;

        public event EventHandler<TfsConnectionErrorEventArgs> TfsConnectionErrorReceived;

        public event EventHandler Started;

        public event EventHandler Stopped;

        #endregion

        #endregion

        #region Private Members

        private class AutoShelveInfos
        {
            public event EventHandler<ShelvesetCreatedEventArgs> ShelvesetCreated
            {
                add { service.ShelvesetCreated += value; }
                remove { service.ShelvesetCreated -= value; }
            }

            public ushort MaximumShelvesets { get; }

            public string ShelvesetName { get; }

            private readonly IAutoShelve service;

            public AutoShelveInfos(IAutoShelve service)
            {
                this.service = service;
                MaximumShelvesets = service.MaximumShelvesets;
                ShelvesetName = service.ShelvesetName;
            }
        }

        private Uri _projectCollectionUri;

        public Uri ProjectCollectionUri
        {
            get
            {
                if (_projectCollectionUri != null) return _projectCollectionUri;
                try
                {
                    var dte = (DTE2)_serviceProvider.GetService(typeof(DTE));
                    var obj = (TeamFoundationServerExt)dte.GetObject("Microsoft.VisualStudio.TeamFoundation.TeamFoundationServerExt");
                    _projectCollectionUri = new Uri(WebUtility.UrlDecode(obj.ActiveProjectContext.DomainUri));
                }
                catch (Exception ex)
                {
                    Stop();  // Disable timer to prevent Ref: Q&A "endless error dialogs" @ http://visualstudiogallery.msdn.microsoft.com/080540cb-e35f-4651-b71c-86c73e4a633d 
                    TfsConnectionErrorReceived?.Invoke(this, new TfsConnectionErrorEventArgs(error: ex));
                }
                return _projectCollectionUri;
            }
        }

        private void InvalidateConnection()
        {
            _projectCollectionUri = null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="service"></param>
        /// <param name="infos"></param>
        /// <param name="workspaceInfo"></param>
        /// <param name="force">True when the user manually initiates a ShelveSet via the Team menu or mapped shortcut key.</param>
        private static ShelvesetCreatedEventArgs CreateShelvesetInternal(VersionControlServer service, WorkspaceInfo workspaceInfo, AutoShelveInfos infos, bool force)
        {
            var shelvesetName = default(string);
            var shelvesetChangeCount = 0;
            var shelvesetPurgeCount = 0;
            var error = default(Exception);

            try
            {
                var workspace = service.GetWorkspace(workspaceInfo);

                // Build a new, valid shelve set name
                shelvesetName = string.Format(infos.ShelvesetName, workspace.Name, workspace.OwnerName, DateTime.Now, workspace.OwnerName.GetDomain(), workspace.OwnerName.GetLogin());
                shelvesetName = shelvesetName.CleanShelvesetName();

                // If there are no pending changes that have changed since the last shelveset then there is nothing to do
                var hasChanges = force;
                var pendingChanges = workspace.GetPendingChanges();
                var numPending = pendingChanges.Length;

                if (numPending > 0)
                {
                    var pastShelvesets = GetPastShelvesets(service, workspace, infos.ShelvesetName);

                    if (!force)
                    {
                        var lastShelveset = pastShelvesets.FirstOrDefault();
                        if (lastShelveset == null)
                        {
                            // If there are pending changes and no shelveset yet exists, then create shelveset.
                            hasChanges = true;
                        }
                        else
                        {
                            // Compare numPending to shelvedChanges.Count();  Force shelveset if they differ
                            // Otherwise, resort to comparing file HashValues
                            var shelvedChanges = service.QueryShelvedChanges(lastShelveset).FirstOrDefault();
                            hasChanges = (shelvedChanges == null
                                || numPending != shelvedChanges.PendingChanges.Length)
                                || pendingChanges.DifferFrom(shelvedChanges.PendingChanges);
                        }
                    }

                    if (hasChanges)
                    {
                        shelvesetChangeCount = numPending;

                        // Actually create a new Shelveset 
                        var shelveset = new Shelveset(service, shelvesetName, workspace.OwnerName)
                        {
                            Comment = string.Format("Shelved by {0}. {1} items", _extensionName, numPending)
                        };
                        workspace.Shelve(shelveset, pendingChanges, ShelvingOptions.Replace);

                        // Clean up past Shelvesets
                        if (infos.MaximumShelvesets > 0)
                        {
                            foreach (var set in pastShelvesets.Skip(infos.MaximumShelvesets))
                            {
                                service.DeleteShelveset(set);
                                shelvesetPurgeCount++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex;
            }

            // Build event args for notification create shelveset result
            return new ShelvesetCreatedEventArgs(
                shelvesetName: shelvesetName,
                shelvesetChangeCount: shelvesetChangeCount,
                shelvesetPurgeCount: shelvesetPurgeCount,
                error: error);
        }

        private static IEnumerable<Shelveset> GetPastShelvesets(VersionControlServer service, Workspace workspace, string shelvesetName)
        {
            var pastShelvesets = service.QueryShelvesets(null, workspace.OwnerName).Where(s => s.Comment?.Contains(_extensionName) == true);
            if (pastShelvesets?.Any() == true)
            {
                if (shelvesetName.IsNameSpecificToWorkspace())
                {
                    pastShelvesets = pastShelvesets.Where(s => s.Name.Contains(workspace.Name));
                }
                return pastShelvesets.OrderByDescending(s => s.CreationDate);
            }
            else
            {
                return pastShelvesets;
            }
        }

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // The bulk of the clean-up code is implemented in Dispose(bool)
        protected virtual void Dispose(bool disposeManaged)
        {
            _timer?.Dispose();
        }

        #endregion

        public void Start()
        {
            _timer.Change(TimeSpan.FromMinutes(TimerInterval), TimeSpan.FromMinutes(TimerInterval));
            IsRunning = true;
            Started?.Invoke(this, System.EventArgs.Empty);
        }

        public void Stop()
        {
            if (!IsRunning) return;
            _timer.Change(Timeout.Infinite, Timeout.Infinite);
            IsRunning = false;
            Stopped?.Invoke(this, System.EventArgs.Empty);
        }

        #endregion

    }
}