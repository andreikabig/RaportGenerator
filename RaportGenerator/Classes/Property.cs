namespace RaportGenerator.Classes
{
    public class Property<T> : IProperty
    {
        // Название свойства
        public string? Name { get; set; }

        // Значение свойства
        public T? Value { get; set; }
    }
}