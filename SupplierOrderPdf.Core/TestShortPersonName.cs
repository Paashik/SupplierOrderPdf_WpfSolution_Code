using System;

namespace SupplierOrderPdf.Core;

/// <summary>
/// Тестовый класс для проверки работы свойства ShortPersonName в классе AccessUser
/// </summary>
public class TestShortPersonName
{
    public static void RunTests()
    {
        Console.WriteLine("Тестирование свойства ShortPersonName в классе AccessUser");
        Console.WriteLine("============================================================");
        
        // Тест 1: Полные данные персоны (фамилия, имя, отчество)
        var user1 = new AccessUser
        {
            PersonLastName = "Петров",
            PersonFirstName = "Иван",
            PersonSecondName = "Иванович"
        };
        Console.WriteLine($"Тест 1 - Полные данные: '{user1.ShortPersonName}'");
        Console.WriteLine($"Ожидаемый результат: 'Петров И.И.'");
        Console.WriteLine($"Результат корректный: {user1.ShortPersonName == "Петров И.И."}");
        Console.WriteLine();
        
        // Тест 2: Фамилия и имя (без отчества)
        var user2 = new AccessUser
        {
            PersonLastName = "Сидоров",
            PersonFirstName = "Петр"
        };
        Console.WriteLine($"Тест 2 - Фамилия и имя: '{user2.ShortPersonName}'");
        Console.WriteLine($"Ожидаемый результат: 'Сидоров П.'");
        Console.WriteLine($"Результат корректный: {user2.ShortPersonName == "Сидоров П."}");
        Console.WriteLine();
        
        // Тест 3: Только фамилия
        var user3 = new AccessUser
        {
            PersonLastName = "Козлов"
        };
        Console.WriteLine($"Тест 3 - Только фамилия: '{user3.ShortPersonName}'");
        Console.WriteLine($"Ожидаемый результат: 'Козлов'");
        Console.WriteLine($"Результат корректный: {user3.ShortPersonName == "Козлов"}");
        Console.WriteLine();
        
        // Тест 4: Имя и отчество (без фамилии)
        var user4 = new AccessUser
        {
            PersonFirstName = "Александр",
            PersonSecondName = "Сергеевич"
        };
        Console.WriteLine($"Тест 4 - Имя и отчество: '{user4.ShortPersonName}'");
        Console.WriteLine($"Ожидаемый результат: 'А.С.'");
        Console.WriteLine($"Результат корректный: {user4.ShortPersonName == "А.С."}");
        Console.WriteLine();
        
        // Тест 5: Только имя
        var user5 = new AccessUser
        {
            PersonFirstName = "Михаил"
        };
        Console.WriteLine($"Тест 5 - Только имя: '{user5.ShortPersonName}'");
        Console.WriteLine($"Ожидаемый результат: 'М.'");
        Console.WriteLine($"Результат корректный: {user5.ShortPersonName == "М."}");
        Console.WriteLine();
        
        // Тест 6: Нет данных персоны
        var user6 = new AccessUser();
        Console.WriteLine($"Тест 6 - Нет данных персоны: '{user6.ShortPersonName}'");
        Console.WriteLine($"Ожидаемый результат: '' (пустая строка)");
        Console.WriteLine($"Результат корректный: {user6.ShortPersonName == ""}");
        Console.WriteLine();
        
        // Тест 7: Сравнение с PersonName
        var user7 = new AccessUser
        {
            PersonLastName = "Иванов",
            PersonFirstName = "Сергей",
            PersonSecondName = "Петрович"
        };
        Console.WriteLine("Тест 7 - Сравнение PersonName и ShortPersonName");
        Console.WriteLine($"PersonName: '{user7.PersonName}'");
        Console.WriteLine($"ShortPersonName: '{user7.ShortPersonName}'");
        Console.WriteLine($"ShortPersonName не содержит 'Пользователь': {!user7.ShortPersonName.Contains("Пользователь")}");
        Console.WriteLine($"ShortPersonName содержит только фамилию и инициалы: {user7.ShortPersonName == "Иванов С.П."}");
        Console.WriteLine();
        
        Console.WriteLine("Все тесты завершены!");
    }
}