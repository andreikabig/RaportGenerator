namespace RaportGenerator.Classes
{
    /// <summary>
    /// Класс, представляющий количественные данные за текущий период.
    /// </summary>
    public class QuantityCurrent
    {
        // Кол-во
        public int Value { get; set; }

        // Период
        public string? Dates { get; set; }
    }
}