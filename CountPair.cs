class CountPair
{
    public int SourceCount { get; set; }
    public int TargetCount { get; set; }

    public CountPair(int sourceCount, int targetCount)
    {
        this.SourceCount = sourceCount;
        this.TargetCount = targetCount;
    }

    public bool IsSatisfied()
    {
        if (SourceCount <= TargetCount) return true;
        return false;
    }
}


