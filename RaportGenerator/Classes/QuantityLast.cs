namespace RaportGenerator.Classes
{
    /// <summary>
    /// Класс, представляющий количественные данные за прошедший период.
    /// </summary>
    public class QuantityLast
    {
        // Кол-во
        public int Value { get; set; }

        // Период
        public string? Dates { get; set; }
    }
}