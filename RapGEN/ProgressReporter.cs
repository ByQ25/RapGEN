namespace RapGEN
{
    class ProgressReporter : IProgressInfo
    {
        // Pola:
        protected bool stopRequired;

        // Konstruktory:
        public ProgressReporter(int progValue, int maxValue)
        {
            this.Progress = progValue;
            this.ProgressMax = maxValue;
            this.stopRequired = false;
            RapGEN_MainWin.ProcessCanceled += OnStop;
        }
        public ProgressReporter(int maxValue) : this(0, maxValue) { }
        public ProgressReporter() : this(0, 1) { }
        
        // Metody:
        protected void OnStop(object sender, System.EventArgs e)
        {
            stopRequired = true;
        }

        // Właściwości:
        public int Progress { get; protected set; }
        public int ProgressMax { get; protected set; }
        public byte PercentageProgress {
            get { return System.Convert.ToByte((double)Progress * 100 / ProgressMax); }
        }
    }
}
