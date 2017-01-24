namespace RapGEN
{
    internal interface IProgressInfo
    {
        int Progress { get; }
        int ProgressMax { get; }
        byte PercentageProgress { get; }
    }
}