using System.Text.Json;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchTaskPlannerTests
{
    [Fact]
    public void ParsesOnlyRequestedScopeItemsAndExplicitLimits()
    {
        const string response = """
            {
              "scope_items": ["تعريف الأمر الوارد بعد الحظر", "أقوال الأصوليين في دلالته"],
              "constraints": ["الاعتماد على المكتبتين المختارتين"],
              "exclusions": ["عدم التوسع في تاريخ علم الأصول"]
            }
            """;

        Assert.True(ResearchTaskPlanner.TryParse(
            response,
            out var scope,
            out var constraints,
            out var exclusions));

        Assert.Equal(2, scope.Count);
        Assert.Equal("تعريف الأمر الوارد بعد الحظر", scope[0]);
        Assert.Single(constraints);
        Assert.Single(exclusions);
    }

    [Theory]
    [InlineData("")]
    [InlineData("هذه ليست JSON")]
    [InlineData("""{"scope_items":[],"constraints":[],"exclusions":[]}""")]
    [InlineData("""{"scope_items":["موضوع"],"constraints":[]}""")]
    public void RejectsIncompleteOrEmptyPlanningResponses(string response)
    {
        Assert.False(ResearchTaskPlanner.TryParse(
            response,
            out _,
            out _,
            out _));
    }

    [Fact]
    public void PlannerSchemaIsStrictAndContainsOnlyTaskBoundaryFields()
    {
        using var document = JsonDocument.Parse(
            ResearchTaskPlanner.BuildResponseFormat().GetRawText());

        var format = document.RootElement;
        Assert.Equal("json_schema", format.GetProperty("type").GetString());
        Assert.True(format.GetProperty("json_schema").GetProperty("strict").GetBoolean());

        var schema = format.GetProperty("json_schema").GetProperty("schema");
        Assert.False(schema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(
            new[] { "scope_items", "constraints", "exclusions" },
            schema.GetProperty("required").EnumerateArray()
                .Select(item => item.GetString()).ToArray());
    }

    [Fact]
    public void BudgetGrowsWithRequestedScopeAndDepthButHasAnEmergencyCeiling()
    {
        var focused = ResearchTaskPlan.Create(
            "ما تعريف الاجتهاد؟",
            ["ما تعريف الاجتهاد؟"],
            [],
            [],
            selectedSourceCount: 1);

        var expanded = ResearchTaskPlan.Create(
            "ابحث وحلل أقوال العلماء وأدلتهم في المسألة مع مقارنة الاتجاهات",
            ["تعريف المسألة", "أقوال العلماء", "الأدلة", "المقارنة"],
            [],
            [],
            selectedSourceCount: 2);

        var deep = ResearchTaskPlan.Create(
            "أريد بحثاً معمقاً شاملاً واستقصائياً ومقارنة بين الاتجاهات",
            Enumerable.Range(1, 8).Select(index => "عنصر مطلوب " + index),
            [],
            [],
            selectedSourceCount: 2);

        Assert.True(expanded.MaxToolCalls > focused.MaxToolCalls);
        Assert.Equal(32, deep.MaxToolCalls);
        Assert.InRange(focused.TimeoutSeconds, 180, 300);
        Assert.Equal(300, deep.TimeoutSeconds);
    }

    [Theory]
    [InlineData("وثق هذه الفكرة من المصادر الأصلية", true)]
    [InlineData("دقق الإحالات والاقتباسات", true)]
    [InlineData("ما تعريف الاجتهاد؟", false)]
    [InlineData("أجر بحثاً معمّقاً في المسألة", true)]
    public void RequiresSourceVerificationOnlyWhenRequestedOrDeep(string goal, bool expected)
    {
        var plan = ResearchTaskPlan.Create(goal, [goal], [], [], selectedSourceCount: 1);
        Assert.Equal(expected, plan.RequiresVerification);
    }

    [Fact]
    public void DeepVerificationTracksEveryRequestedScopeItem()
    {
        var plan = ResearchTaskPlan.Create(
            "أجر بحثاً معمقاً في المسألة",
            ["العنصر الأول", "العنصر الثاني"],
            [],
            [],
            selectedSourceCount: 1);

        var instructions = plan.BuildProgressInstructions(
            new HashSet<string>(["S1", "S2"]),
            [],
            4,
            0,
            verificationRequired: true,
            verifiedScopeItemIds: new HashSet<string>(["S1"]),
            verificationToolsAvailable: true);

        Assert.Contains("التحقق من المواضع الأصلية", instructions);
        Assert.Contains("S2: العنصر الثاني", instructions);
        Assert.DoesNotContain("S1: العنصر الأول؛", instructions);
    }

    [Fact]
    public void CompletedCoverageStillAllowsAnEvidenceSufficiencyReview()
    {
        var plan = ResearchTaskPlan.Create(
            "ابحث عن مفهوم الاجتهاد",
            ["مفهوم الاجتهاد"],
            [],
            [],
            selectedSourceCount: 1);

        var instructions = plan.BuildProgressInstructions(
            new HashSet<string>(["S1"]),
            [],
            2,
            0,
            verificationRequired: false,
            verifiedScopeItemIds: new HashSet<string>(),
            verificationToolsAvailable: false);

        Assert.Contains("اكتملت التغطية الأولية", instructions);
        Assert.Contains("راجع الآن كفاية الأدلة", instructions);
        Assert.Contains("لا تضف عناصر جديدة", instructions);
    }

    [Fact]
    public void ProgressInstructionsExplicitlyForbidAdjacentTopicExpansion()
    {
        var plan = ResearchTaskPlan.Create(
            "ابحث عن مفهوم المقاصد عند الشاطبي",
            ["مفهوم المقاصد عند الشاطبي"],
            [],
            [],
            selectedSourceCount: 1);

        var instructions = plan.BuildProgressInstructions(
            new HashSet<string>(),
            ["aljam3"],
            0,
            0,
            verificationRequired: false,
            verifiedScopeItemIds: new HashSet<string>(),
            verificationToolsAvailable: false);

        Assert.Contains("لا تبحث ولا تشرح ولا تضف محوراً", instructions);
        Assert.Contains("المكتبات المختارة", instructions);
        Assert.Contains("S1", instructions);
    }
}
