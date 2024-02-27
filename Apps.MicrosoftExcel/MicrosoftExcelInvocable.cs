using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelInvocable : BaseInvocable
{
    protected readonly MicrosoftExcelClient Client;

    protected MicrosoftExcelInvocable(InvocationContext invocationContext) : base(invocationContext)
    {
        Client = new();
    }
}