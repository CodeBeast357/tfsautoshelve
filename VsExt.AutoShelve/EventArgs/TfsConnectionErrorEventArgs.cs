using System;

namespace VsExt.AutoShelve.EventArgs
{
    public class TfsConnectionErrorEventArgs : System.EventArgs
    {
        public Exception Error { get; }

        public TfsConnectionErrorEventArgs(Exception error)
        {
            Error = error;
        }
    }
}