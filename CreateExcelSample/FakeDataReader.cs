namespace CreateExcelSample;

/// <summary>
/// Used to simulate calling something like an OracleDataReader for testing
/// </summary>
public class FakeDataReader
{
    /// <summary>
    /// Reads a fake record
    /// </summary>
    /// <returns>FakeDataRecord</returns>
    public static FakeDataRecord Read()
    {
        FakeDataRecord dRec = new();
        RonsRandom rdmObj = new();

        dRec.Value1 = "Val1_" + rdmObj.RandomTextGenerator(8);
        dRec.Value2 = "Val2_" + rdmObj.RandomNumberGenerator();
        dRec.Value3 = "Val3_" + rdmObj.RandomTextGenerator(6);
        dRec.Value4 = "Val4_" + rdmObj.RandomTextGenerator(20);

        return dRec;
    }
}