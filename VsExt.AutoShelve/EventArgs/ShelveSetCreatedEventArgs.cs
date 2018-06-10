using System;

namespace VsExt.AutoShelve.EventArgs
{
    public class ShelvesetCreatedEventArgs : System.EventArgs
    {
        public string ShelvesetName { get; }

        public int ShelvesetChangeCount { get; }

        public int ShelvesetsPurgeCount { get; }

        public Exception Error { get; }

        public bool IsSuccess => Error == null;

        public ShelvesetCreatedEventArgs(string shelvesetName, int shelvesetChangeCount, int shelvesetPurgeCount, Exception error)
        {
            ShelvesetName = shelvesetName;
            ShelvesetChangeCount = shelvesetChangeCount;
            ShelvesetsPurgeCount = shelvesetPurgeCount;
            Error = error;
        }
    }
}