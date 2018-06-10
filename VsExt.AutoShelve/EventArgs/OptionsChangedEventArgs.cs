namespace VsExt.AutoShelve.EventArgs
{
    public class OptionsChangedEventArgs : System.EventArgs
    {
        public bool PauseWhileDebugging { get; }

        public double Interval { get; }

        public ushort MaximumShelvesets { get; }

        public string OutputPane { get; }

        public string ShelvesetName { get; }

        public OptionsChangedEventArgs(bool pauseWhileDebugging, double interval, ushort maximumShelvesets, string outputPane, string shelvesetName)
        {
            PauseWhileDebugging = pauseWhileDebugging;
            Interval = interval;
            MaximumShelvesets = maximumShelvesets;
            OutputPane = outputPane;
            ShelvesetName = shelvesetName;
        }
    }
}