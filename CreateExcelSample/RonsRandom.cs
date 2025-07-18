using System.Text;

namespace CreateExcelSample;

/// <summary>
/// A class to use to get random strings and numbers for testing
/// </summary>
public class RonsRandom
{
    private Random random;

    public RonsRandom()
    {
        this.random = new Random();
    }

    /// <summary>
    /// Gets random text string
    /// </summary>
    /// <param name="size">Length of random text to return</param>
    /// <returns>string</returns>
    public string RandomTextGenerator(int size)
    {
        StringBuilder builder = new();
        string characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        char character;

        for (int i = 0; i < size; i++)
        {
            character = characters[random.Next(0, characters.Length)];
            builder.Append(character);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Gets a random int between 222222 & 999999
    /// </summary>
    /// <returns>int</returns>
    public int RandomNumberGenerator()
    {
        return random.Next(222222, 999999);
    }
}