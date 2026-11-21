using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchIntentClassifierTests
{
    [Theory]
    [InlineData("ابحث عن مفهوم المقاصد عند الشاطبي")]
    [InlineData("أريد مصادر موثقة عن تطور المصطلح الأصولي")]
    [InlineData("اذكر الدراسات السابقة مع المراجع")]
    [InlineData("وثق هذه الفكرة من كتب التراث")]
    [InlineData("Find studies and references about legal theory")]
    public void RequiresMcpSources_WhenRequestAsksForResearchOrReferences(string prompt)
    {
        Assert.True(ResearchIntentClassifier.RequiresMcpSources(prompt));
    }

    [Theory]
    [InlineData("ما معنى الاجتهاد؟")]
    [InlineData("اشرح الفرق بين العام والخاص")]
    [InlineData("السلام عليكم")]
    public void DoesNotRequireMcpSources_ForGeneralConversation(string prompt)
    {
        Assert.False(ResearchIntentClassifier.RequiresMcpSources(prompt));
    }
}
