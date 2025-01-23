using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelInvocable : BaseInvocable
{
    protected readonly MicrosoftExcelClient Client;

    protected MicrosoftExcelInvocable(InvocationContext invocationContext) : base(invocationContext)
    {
        Client = new();
    }

    protected void ValidateWorksheetParameter(WorksheetRequest worksheetRequest)
    {
        string errorMessage = "Invalid worksheet ID. Worksheet ID example:  \"{00000000-0001-0000-0000-000000000000}\"";
        if (worksheetRequest.Worksheet.FirstOrDefault() != '{' || worksheetRequest.Worksheet.LastOrDefault() != '}' ||
            !Guid.TryParse(worksheetRequest.Worksheet.Substring(1, worksheetRequest.Worksheet.Length - 2), out var _))
            throw new PluginMisconfigurationException(errorMessage);
    }

    protected void ValidateCellAddressParameter(GetCellRequest cellAddress)
    {
        string errorMessage = "Invalid cell address format. Cell address example: \"A1\"";
        var firstAddressChar = cellAddress.CellAddress.FirstOrDefault();
        var lastAddressChar = cellAddress.CellAddress.LastOrDefault();
        if (firstAddressChar == default || !char.IsLetter(firstAddressChar) || !char.IsUpper(firstAddressChar) ||
            lastAddressChar == default || !char.IsDigit(lastAddressChar))
            throw new PluginMisconfigurationException(errorMessage);
    }
}