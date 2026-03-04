using GAWELA.Tourenplaner2.Models;

namespace GAWELA.Tourenplaner2.Services;

public static class EmployeeStorageService
{
    public static Employee NormalizeEmployee(Employee? employee)
    {
        return new Employee
        {
            Id = string.IsNullOrWhiteSpace(employee?.Id) ? Guid.NewGuid().ToString() : employee!.Id,
            Name = employee?.Name?.Trim() ?? string.Empty,
            Short = employee?.Short?.Trim() ?? string.Empty,
            Phone = employee?.Phone?.Trim() ?? string.Empty,
            Active = employee?.Active ?? true,
            CreatedAt = string.IsNullOrWhiteSpace(employee?.CreatedAt) ? DateTime.Now.ToString("s") : employee!.CreatedAt,
        };
    }

    public static async Task<List<Employee>> LoadEmployeesAsync(string path)
    {
        try
        {
            var data = await JsonStorageService.LoadJsonFileAsync(path, () => new List<Employee>(), createIfMissing: true);
            return NormalizeList(data);
        }
        catch (InvalidJsonFileException)
        {
            return [];
        }
    }

    public static async Task<List<Employee>> SaveEmployeesAsync(string path, List<Employee> employees)
    {
        var cleaned = NormalizeList(employees);
        await JsonStorageService.AtomicWriteJsonAsync(path, cleaned);
        return cleaned;
    }

    private static List<Employee> NormalizeList(List<Employee> employees)
    {
        var seen = new HashSet<string>();
        var result = new List<Employee>();

        foreach (var entry in employees)
        {
            var normalized = NormalizeEmployee(entry);
            if (string.IsNullOrWhiteSpace(normalized.Name))
            {
                continue;
            }

            if (!seen.Add(normalized.Id))
            {
                normalized.Id = Guid.NewGuid().ToString();
                seen.Add(normalized.Id);
            }

            result.Add(normalized);
        }

        return result.OrderBy(e => !e.Active).ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }
}
