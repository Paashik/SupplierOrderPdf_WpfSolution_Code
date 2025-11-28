using System;

namespace SupplierOrderPdf.Core;

/// <summary>
/// Модель пользователя системы Access.
/// Представляет информацию о пользователе, авторизованном в системе управления заказами.
/// Содержит идентификационные данные, информацию о роли и контактные данные.
/// </summary>
public class AccessUser
{
    /// <summary>
    /// Уникальный идентификатор пользователя в системе Access.
    /// Используется как первичный ключ для ссылок на пользователя.
    /// </summary>
    public int Id { get; set; }
    
    /// <summary>
    /// Логин пользователя для входа в систему.
    /// Уникален в рамках системы и используется для аутентификации.
    /// </summary>
    public string Login { get; set; } = string.Empty;
    
    /// <summary>
    /// Идентификатор роли пользователя в системе.
    /// Определяет права доступа и функциональность, доступную пользователю.
    /// Может быть null для пользователей без назначенной роли.
    /// </summary>
    public int? RoleId { get; set; }

    /// <summary>
    /// Название роли пользователя в системе.
    /// Определяет права доступа и функциональность, доступную пользователю.
    /// </summary>
    public string RoleName { get; set; } = string.Empty;

    /// <summary>
    /// Идентификатор персоны, связанной с пользователем.
    /// Связывает пользователя системы с физическим лицом в базе данных.
    /// Может быть null, если пользователь не привязан к конкретной персоне.
    /// </summary>
    public int? PersonId { get; set; }
    
    /// <summary>
    /// Отображаемое имя персоны, связанной с пользователем.
    /// Формируется в формате "Фамилия Имя Отчество" или только "Имя Отчество" если фамилия отсутствует.
    /// Автоматически вычисляется на основе PersonLastName, PersonFirstName и PersonSecondName без приставки "Пользователь".
    /// </summary>
    public string PersonName 
    { 
        get
        {
            // Формируем полные имена
            var names = new List<string>();
            
            // Добавляем фамилию если есть
            if (!string.IsNullOrWhiteSpace(PersonLastName))
                names.Add(PersonLastName.Trim());
            
            // Добавляем имя если есть
            if (!string.IsNullOrWhiteSpace(PersonFirstName))
                names.Add(PersonFirstName.Trim());
            
            // Добавляем отчество если есть
            if (!string.IsNullOrWhiteSpace(PersonSecondName))
                names.Add(PersonSecondName.Trim());
            
            // Возвращаем объединенные имена или пустую строку если нет данных
            return names.Count > 0 ? string.Join(" ", names) : string.Empty;
        }
    }
    
    /// <summary>
    /// Краткое имя персоны в формате "Фамилия И.О." без приставки "Пользователь".
    /// Используется для компактного отображения в PDF документах, заголовках email и интерфейсах.
    /// Автоматически вычисляется на основе PersonLastName, PersonFirstName и PersonSecondName.
    /// </summary>
    public string ShortPersonName 
    { 
        get
        {
            // Формируем инициалы из имени и отчества
            string initials = "";
            if (!string.IsNullOrWhiteSpace(PersonFirstName))
                initials += PersonFirstName.Trim()[0] + ".";
            if (!string.IsNullOrWhiteSpace(PersonSecondName))
                initials += PersonSecondName.Trim()[0] + ".";
            
            // Формируем краткое имя персоны
            if (!string.IsNullOrWhiteSpace(PersonLastName))
            {
                return string.IsNullOrWhiteSpace(initials)
                    ? PersonLastName.Trim()
                    : $"{PersonLastName.Trim()} {initials}";
            }
            else
            {
                return string.IsNullOrWhiteSpace(initials) 
                    ? string.Empty 
                    : initials;
            }
        }
    }

    /// <summary>
    /// Формат "Логин (Фамилия И.О.)" для использования в LoginWindow.
    /// </summary>
    public string LoginWithShortName
    {
        get { return $"{Login} ({ShortPersonName})"; }
    }

    /// <summary>
    /// Имя персоны, связанной с пользователем.
    /// </summary>
    public string PersonFirstName { get; set; } = string.Empty;
    
    /// <summary>
    /// Фамилия персоны, связанной с пользователем.
    /// </summary>
    public string PersonLastName { get; set; } = string.Empty;
    
    /// <summary>
    /// Отчество персоны, связанной с пользователем.
    /// </summary>
    public string PersonSecondName { get; set; } = string.Empty;
    
   
    
    /// <summary>
    /// Адрес электронной почты пользователя.
    /// Используется для отправки уведомлений и связи с пользователем.
    /// </summary>
    public string Email { get; set; } = string.Empty;
    
    /// <summary>
    /// Номер телефона пользователя.
    /// Может использоваться для SMS уведомлений или связи.
    /// </summary>
    public string Phone { get; set; } = string.Empty;
    
    /// <summary>
    /// Пароль пользователя (хранится в зашифрованном виде в базе данных).
    /// НЕ рекомендуется использовать это свойство в клиентском коде без необходимости.
    /// </summary>
    public string Password { get; set; } = string.Empty;

    
    /// <summary>
    /// Определяет отображаемое имя типа пользователя на основе роли.
    /// </summary>
    /// <returns>Строка с типом пользователя или пустая строка</returns>
    private string GetUserTypeDisplayName()
    {
        // По умолчанию используем "person" как тип пользователя
        // Можно расширить логику на основе RoleId в будущем
        return "person";
    }

    /// <summary>
    /// Возвращает отформатированную строку для отображения контакта пользователя.
    /// Использует ShortPersonName с fallback на PersonName, если ShortPersonName пустой.
    /// Используется для отображения контактной информации в PDF документах.
    /// </summary>
    /// <returns>Строка в формате "Фамилия И.О." или полное ФИО</returns>
    public string ToContactDisplayString()
    {
        return string.IsNullOrWhiteSpace(ShortPersonName) ? PersonName : ShortPersonName;
    }

    /// <summary>
    /// Возвращает отформатированную строку пользователя с типом.
    /// Формат: "Фамилия Имя Отчество" (без префикса "Пользователь").
    /// Используется для отображения пользователя в заголовках и уведомлениях.
    /// </summary>
    /// <returns>Строка в формате "Фамилия Имя Отчество"</returns>
    public string ToLoginWithType()
    {
        return PersonName;
    }

    /// <summary>
    /// Возвращает информацию об отправителе для настроек email.
    /// Формат: включает имя пользователя и email если доступен.
    /// Используется в окне настроек для отображения информации об отправителе.
    /// </summary>
    /// <returns>Строка с информацией об отправителе</returns>
    public string ToEmailFromInfo()
    {
        var contactInfo = ToContactDisplayString();
        
        if (!string.IsNullOrWhiteSpace(Email))
            return $"{contactInfo} <{Email}>";
        
        return contactInfo;
    }

}

/// <summary>
/// Модель роли доступа в системе Access.
/// Представляет группу прав доступа, которая может быть назначена пользователям.
/// Роли определяют функциональность, доступную пользователям в системе.
/// </summary>
public class AccessRole
{
    /// <summary>
    /// Уникальный идентификатор роли в системе.
    /// Используется как первичный ключ для связывания ролей с пользователями.
    /// </summary>
    public int Id { get; set; }
    
    /// <summary>
    /// Название роли в системе.
    /// Должно быть уникальным и описывать набор прав доступа (например, "Администратор", "Пользователь").
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// Возвращает название роли для отображения в пользовательском интерфейсе.
    /// Используется в выпадающих списках и элементах выбора роли.
    /// </summary>
    /// <returns>Название роли</returns>
    public override string ToString() => Name;
}

/// <summary>
/// Модель персоны в системе.
/// Представляет физическое лицо в базе данных, которое может быть связано с пользователем
/// или выступать в качестве контактного лица для заказчиков и поставщиков.
/// </summary>
public class PersonItem
{
    /// <summary>
    /// Уникальный идентификатор персоны в базе данных.
    /// Используется как первичный ключ для ссылок на персону.
    /// </summary>
    public int Id { get; set; }
    
    /// <summary>
    /// Полное имя персоны.
    /// Обычно содержит фамилию, имя и отчество для российских персон.
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// Номер телефона персоны.
    /// Может содержать мобильный или городской номер телефона.
    /// </summary>
    public string Phone { get; set; } = string.Empty;
    
    /// <summary>
    /// Адрес электронной почты персоны.
    /// Используется для отправки уведомлений и связи с персоной.
    /// </summary>
    public string Email { get; set; } = string.Empty;
    
    /// <summary>
    /// Возвращает имя персоны для отображения в пользовательском интерфейсе.
    /// Используется в выпадающих списках и элементах выбора персоны.
    /// </summary>
    /// <returns>Имя персоны</returns>
    public override string ToString() => Name;
}

/// <summary>
/// Модель контактного лица.
/// Представляет контактную информацию, которая может быть привязана к поставщикам,
/// заказчикам или использоваться как независимые контакты.
/// Содержит расширенную информацию по сравнению с PersonItem, включая заметки.
/// </summary>
public class ContactItem
{
    /// <summary>
    /// Уникальный идентификатор контактного лица.
    /// Используется как первичный ключ для ссылок на контактное лицо.
    /// </summary>
    public int Id { get; set; }
    
    /// <summary>
    /// Имя контактного лица.
    /// Может быть полным именем или названием должности/компании.
    /// </summary>
    public string Name { get; set; } = string.Empty;
    
    /// <summary>
    /// Адрес электронной почты контактного лица.
    /// Приоритетный способ связи для профессиональных контактов.
    /// </summary>
    public string Email { get; set; } = string.Empty;
    
    /// <summary>
    /// Номер телефона контактного лица.
    /// Альтернативный способ связи или основной для личных контактов.
    /// </summary>
    public string Phone { get; set; } = string.Empty;
    
    /// <summary>
    /// Дополнительные заметки о контактном лице.
    /// Может содержать должность, компанию, предпочтения по связи или другую полезную информацию.
    /// </summary>
    public string Note { get; set; } = string.Empty;
    
    /// <summary>
    /// Возвращает строковое представление контактного лица для отображения в UI.
    /// 
    /// Формат отображения:
    /// - Если есть и email, и телефон: "Имя (email)" или "Имя (телефон)" (приоритет email)
    /// - Если есть только email: "Имя (email)"
    /// - Если есть только телефон: "Имя (телефон)"
    /// - Если нет ни email, ни телефона: только "Имя"
    /// 
    /// Это обеспечивает информативное отображение контактной информации в списках.
    /// </summary>
    /// <returns>Строка для отображения контактного лица</returns>
    public override string ToString()
        => string.IsNullOrWhiteSpace(Email) && string.IsNullOrWhiteSpace(Phone)
           ? Name
           : $"{Name} ({(string.IsNullOrWhiteSpace(Email) ? Phone : Email)})";
}

/// <summary>
/// Модель заголовка заказа.
/// Содержит общую информацию о заказе, включая данные о поставщике и заказчике.
/// Используется как основа для создания PDF документов заявок на закупку.
/// </summary>
public class OrderHeader
{
    /// <summary>
    /// Уникальный идентификатор заказа в системе.
    /// Используется как первичный ключ для ссылки на заказ.
    /// </summary>
    public int ID { get; set; }
    
    /// <summary>
    /// Дата создания заказа.
    /// Может быть null, если заказ еще не был сохранен или дата не была установлена.
    /// </summary>
    public DateTime? OrderDate { get; set; }
    
    /// <summary>
    /// Номер заказа в системе.
    /// Может быть автоматически сгенерированным или введенным вручную.
    /// Используется для идентификации заказа в документах и переписке.
    /// </summary>
    public string OrderNum { get; set; } = string.Empty;
    
    /// <summary>
    /// Дополнительные заметки к заказу.
    /// Может содержать особые требования, комментарии или инструкции.
    /// </summary>
    public string Note { get; set; } = string.Empty;

    // ==================== Информация о поставщике ====================
    
    /// <summary>
    /// Идентификатор поставщика в базе данных.
    /// Используется для связи с таблицей поставщиков.
    /// </summary>
    public int ProviderId { get; set; }
    
    /// <summary>
    /// Краткое название поставщика.
    /// Используется для отображения в списках и заголовках документов.
    /// </summary>
    public string ProviderName { get; set; } = string.Empty;
    
    /// <summary>
    /// Полное наименование поставщика.
    /// Используется в официальных документах и юридической переписке.
    /// </summary>
    public string ProviderFullName { get; set; } = string.Empty;
    
    /// <summary>
    /// ИНН (Идентификационный номер налогоплательщика) поставщика.
    /// Обязательный реквизит для российских организаций.
    /// </summary>
    public string INN { get; set; } = string.Empty;
    
    /// <summary>
    /// КПП (Код причины постановки на учет) поставщика.
    /// Обязательный реквизит для российских организаций.
    /// </summary>
    public string KPP { get; set; } = string.Empty;
    
    /// <summary>
    /// Юридический или фактический адрес поставщика.
    /// Используется для указания адреса доставки в документах.
    /// </summary>
    public string ProviderAddress { get; set; } = string.Empty;
    
    /// <summary>
    /// Веб-сайт поставщика.
    /// Может использоваться для получения дополнительной информации или связи.
    /// </summary>
    public string Site { get; set; } = string.Empty;
    
    /// <summary>
    /// Основной email адрес поставщика.
    /// Используется для отправки заявок и деловой переписки.
    /// </summary>
    public string Email { get; set; } = string.Empty;
    
    /// <summary>
    /// Основной телефон поставщика.
    /// Используется для связи в случае возникновения вопросов по заказу.
    /// </summary>
    public string Phone { get; set; } = string.Empty;
    
    /// <summary>
    /// Дополнительные заметки о поставщике.
    /// Может содержать информацию об условиях сотрудничества, особенностях работы и т.д.
    /// </summary>
    public string ProviderNote { get; set; } = string.Empty;

    // ==================== Связанные документы и данные ====================
    
    /// <summary>
    /// CSV строка с ID заказчиков, связанных с данным заказом.
    /// Используется для связывания заказа с одним или несколькими заказчиками.
    /// Формат: "1,2,3" - ID заказчиков через запятую.
    /// </summary>
    public string CustomerOrdersCsv { get; set; } = string.Empty;
    
    /// <summary>
    /// Путь к файлу счета, связанного с данным заказом.
    /// Может быть относительным или абсолютным путем.
    /// Используется для прикрепления счета к PDF документу заявки.
    /// </summary>
    public string InvoicePath { get; set; } = string.Empty;
}
