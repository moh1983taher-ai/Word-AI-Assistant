# ResearchTools MCP Host

مضيف محلي جديد لـ «أدوات الباحث»، مبني على .NET 10 LTS، ويستخدم مكتبة **Model Context Protocol C# SDK** الرسمية لإدارة عملاء MCP.

## النطاق الحالي

- خادم HTTP غير متزامن على `127.0.0.1:18081`.
- يحافظ على واجهات التوافق مع إضافة Word الحالية: `/ping` و`/health` و`/v1/models` و`/v1/research`.
- واجهة `/v1/mcp/sources` تختبر اكتشاف الأدوات وتعرض حالة الجامع والشاملة وعدد الأدوات المتاحة.
- يتصل بالجامع الرسمي عبر `https://aljam3.com/mcp` وبالمكتبة الشاملة عبر `https://mcp.shamela.ws`.
- يستخدم نقل SDK التلقائي: Streamable HTTP أولاً ثم SSE عند الحاجة، مع التفاوض التلقائي على إصدار البروتوكول.
- يستخدم DuckAI المحلي على المنفذين `8080` و`18080`، ويفضّل نموذج Luna المتاح.
- يحتفظ بفهرس الأدوات خمس دقائق، ويحد تنفيذ البحث إلى طلبين متزامنين وستة استدعاءات أدوات في المهمة الواحدة.
- لا يتيح أدوات الكتابة أو التعديل أو الحذف، ويتحقق محلياً من الحقول المطلوبة وأنواع المدخلات قبل إرسال الاستدعاء.
- فحص الحياة `/ping` مستقل عن مهام البحث الطويلة، ولا يقرأ محتوى الطلبات في سجل التشخيص.

## البناء والاختبار

يتطلب .NET 10 SDK:

```powershell
dotnet restore runtime/McpHost.Tests/ResearchTools.McpHost.Tests.csproj
dotnet test runtime/McpHost.Tests/ResearchTools.McpHost.Tests.csproj --configuration Release
dotnet publish runtime/McpHost/ResearchTools.McpHost.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true
```

يتحقق سير GitHub Actions من البناء واختبارات الوحدات، ثم يشغّل المضيف على Windows runner ويتأكد من اكتشاف الأدوات المتاحة في المصدرين قبل نشر الملف التنفيذي كأثر بناء.

## حالة الدمج

هذه نسخة تطوير مستقلة. لا تستبدل المشغّل أو الجسر الحاليين، ولا تعدّل معرّف إضافة Word أو إعداداتها. لا يجوز الانتقال إلى اعتمادها في التثبيت التلقائي قبل اجتياز CI، ثم اختبار التكامل على Windows مع نسخة DuckAI الفعلية وخادمي MCP.
