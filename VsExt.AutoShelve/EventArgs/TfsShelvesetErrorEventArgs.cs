using System;

namespace VsExt.AutoShelve.EventArgs
{
    public class TfsShelvesetErrorEventArgs : System.EventArgs
    {
        public Exception Error { get; }

        public TfsShelvesetErrorEventArgs(Exception error)
        {
            Error = error;
        }
    }
}
