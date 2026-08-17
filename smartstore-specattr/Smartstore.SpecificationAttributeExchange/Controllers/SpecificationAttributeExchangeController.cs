using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Smartstore.Core.Data;
using Smartstore.SpecificationAttributeExchange.Models;
using Smartstore.SpecificationAttributeExchange.Services;
using Smartstore.Web.Controllers;

namespace Smartstore.SpecificationAttributeExchange.Controllers;

public sealed class SpecificationAttributeExchangeController : AdminController
{
    private const long MaxUploadBytes = 20L * 1024L * 1024L;
    private readonly SpecificationAttributeExchangeService _service;

    public SpecificationAttributeExchangeController(SmartDbContext db)
    {
        _service = new SpecificationAttributeExchangeService(db);
    }

    [HttpGet]
    public IActionResult Configure()
        => View(new ImportUploadModel());

    [HttpGet]
    public async Task<IActionResult> Export()
    {
        var bytes = await _service.ExportAsync();
        var filename = $"Spezifikationsattribute_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Preview(ImportUploadModel model)
    {
        if (!ValidateUpload(model.File, out var error))
        {
            ViewBag.Error = error;
            return View("Configure", model);
        }

        await using var stream = model.File!.OpenReadStream();
        try
        {
            ViewBag.Preview = await _service.PreviewAsync(stream);
        }
        catch (Exception ex)
        {
            ViewBag.Error = "Die Datei konnte nicht geprüft werden: " + ex.Message;
        }
        return View("Configure", model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Import(ImportUploadModel model)
    {
        if (!model.ConfirmImport)
        {
            ViewBag.Error = "Bitte bestätigen Sie den Import mit der Checkbox.";
            return View("Configure", model);
        }
        if (!ValidateUpload(model.File, out var error))
        {
            ViewBag.Error = error;
            return View("Configure", model);
        }

        byte[] bytes;
        await using (var upload = model.File!.OpenReadStream())
        await using (var copy = new MemoryStream())
        {
            await upload.CopyToAsync(copy);
            bytes = copy.ToArray();
        }

        try
        {
            await using (var validationStream = new MemoryStream(bytes, writable: false))
            {
                var preview = await _service.PreviewAsync(validationStream);
                if (!preview.CanImport)
                {
                    ViewBag.Preview = preview;
                    ViewBag.Error = "Der Import wurde nicht ausgeführt, weil die Prüfung Fehler enthält.";
                    return View("Configure", model);
                }
            }

            await using var importStream = new MemoryStream(bytes, writable: false);
            ViewBag.Result = await _service.ImportAsync(importStream);
        }
        catch (Exception ex)
        {
            ViewBag.Error = "Der Import wurde abgebrochen und zurückgerollt: " + ex.Message;
        }

        return View("Configure", new ImportUploadModel());
    }

    private static bool ValidateUpload(IFormFile? file, out string error)
    {
        if (file is null || file.Length == 0)
        {
            error = "Bitte eine XLSX-Datei auswählen.";
            return false;
        }
        if (file.Length > MaxUploadBytes)
        {
            error = "Die Datei ist größer als 20 MB.";
            return false;
        }
        if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            error = "Es werden ausschließlich XLSX-Dateien akzeptiert.";
            return false;
        }
        error = string.Empty;
        return true;
    }
}
