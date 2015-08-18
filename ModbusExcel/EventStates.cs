namespace ModbusExcel
{
    /// <summary>
    ///   Socket and Topic States enums
    /// </summary>
    public enum EventStates
    {
        None,
        New,
        Opening,
        Idle,
        Reading,
        Exception
    }
}