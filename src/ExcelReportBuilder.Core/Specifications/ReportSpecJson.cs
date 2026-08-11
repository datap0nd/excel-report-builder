using System;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Transforms;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace ExcelReportBuilder.Core.Specifications
{
    public static class ReportSpecJson
    {
        private static readonly JsonSerializerSettings SerializerSettings = CreateSettings();

        public static string Serialize(ReportSpecV1 specification, Formatting formatting = Formatting.Indented)
        {
            if (specification == null)
            {
                throw new ArgumentNullException(nameof(specification));
            }

            EnsureSupportedVersion(specification.SchemaVersion);

            return JsonConvert.SerializeObject(specification, formatting, SerializerSettings);
        }

        public static ReportSpecV1 Deserialize(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("A report specification JSON document is required.", nameof(json));
            }

            var document = JObject.Parse(json);
            var versionToken = document.GetValue("schemaVersion", StringComparison.OrdinalIgnoreCase);
            if (versionToken == null || versionToken.Type != JTokenType.String)
            {
                throw new UnsupportedReportSpecVersionException(null);
            }

            EnsureSupportedVersion(versionToken.Value<string>());
            ReportSpecJsonShapeValidator.Validate(document);

            var readerSettings = CreateSettings();
            readerSettings.NullValueHandling = NullValueHandling.Include;
            var result = document.ToObject<ReportSpecV1>(JsonSerializer.Create(readerSettings));
            if (result == null)
            {
                throw new JsonSerializationException("The JSON document did not contain a report specification.");
            }

            EnsureSupportedVersion(result.SchemaVersion);

            return result;
        }

        public static JsonSerializerSettings CreateSerializerSettings()
        {
            return CreateSettings();
        }

        private static JsonSerializerSettings CreateSettings()
        {
            var settings = new JsonSerializerSettings
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                DateFormatString = "yyyy-MM-ddTHH:mm:ss.FFFFFFFK",
                DateParseHandling = DateParseHandling.DateTime,
                MissingMemberHandling = MissingMemberHandling.Error,
                NullValueHandling = NullValueHandling.Ignore,
                ObjectCreationHandling = ObjectCreationHandling.Replace
            };

            settings.Converters.Add(new StringEnumConverter(new CamelCaseNamingStrategy(), false));
            settings.Converters.Add(new TransformStepJsonConverter());
            settings.Converters.Add(new MeasureExpressionJsonConverter());
            return settings;
        }

        private static void EnsureSupportedVersion(string? schemaVersion)
        {
            if (!string.Equals(schemaVersion, ReportSpecV1.CurrentSchemaVersion, StringComparison.Ordinal))
            {
                throw new UnsupportedReportSpecVersionException(schemaVersion);
            }
        }
    }

    public sealed class UnsupportedReportSpecVersionException : NotSupportedException
    {
        public UnsupportedReportSpecVersionException(string? schemaVersion)
            : base("Report specification version '" + (schemaVersion ?? "<missing>") + "' is not supported.")
        {
            SchemaVersion = schemaVersion;
        }

        public string? SchemaVersion { get; }
    }

    internal sealed class TransformStepJsonConverter : JsonConverter
    {
        public override bool CanWrite => false;

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(TransformStep);
        }

        public override object? ReadJson(
            JsonReader reader,
            Type objectType,
            object? existingValue,
            JsonSerializer serializer)
        {
            var item = JObject.Load(reader);
            var kind = ReadKind<TransformKind>(item, "kind");
            TransformStep target;
            switch (kind)
            {
                case TransformKind.SelectColumns:
                    target = new SelectColumnsTransform();
                    break;
                case TransformKind.KeepColumns:
                    target = new KeepColumnsTransform();
                    break;
                case TransformKind.RemoveColumns:
                    target = new RemoveColumnsTransform();
                    break;
                case TransformKind.ReorderColumns:
                    target = new ReorderColumnsTransform();
                    break;
                case TransformKind.RenameColumn:
                    target = new RenameColumnTransform();
                    break;
                case TransformKind.ChangeColumnType:
                    target = new ChangeColumnTypeTransform();
                    break;
                case TransformKind.TrimText:
                    target = new TrimTextTransform();
                    break;
                case TransformKind.ReplaceValue:
                    target = new ReplaceValueTransform();
                    break;
                case TransformKind.NormalizeBlanks:
                    target = new NormalizeBlanksTransform();
                    break;
                case TransformKind.NormalizeErrors:
                    target = new NormalizeErrorsTransform();
                    break;
                case TransformKind.FillDown:
                    target = new FillDownTransform();
                    break;
                case TransformKind.MapValues:
                    target = new MapValuesTransform();
                    break;
                case TransformKind.FilterRows:
                    target = new FilterRowsTransform();
                    break;
                case TransformKind.ExcludeTotalRows:
                    target = new ExcludeTotalRowsTransform();
                    break;
                case TransformKind.DerivePeriodParts:
                    target = new DerivePeriodPartsTransform();
                    break;
                case TransformKind.AddArithmeticColumn:
                    target = new AddArithmeticColumnTransform();
                    break;
                case TransformKind.NormalizePeriods:
                    target = new NormalizePeriodsTransform();
                    break;
                default:
                    throw new JsonSerializationException("Unsupported transform kind: " + kind + ".");
            }

            using (var objectReader = item.CreateReader())
            {
                serializer.Populate(objectReader, target);
            }

            return target;
        }

        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer)
        {
            throw new NotSupportedException();
        }

        private static TEnum ReadKind<TEnum>(JObject item, string propertyName)
            where TEnum : struct
        {
            var token = item.GetValue(propertyName, StringComparison.OrdinalIgnoreCase);
            if (token == null || token.Type != JTokenType.String)
            {
                throw new JsonSerializationException("Property '" + propertyName + "' is required.");
            }

            var raw = token.Value<string>();
            TEnum value;
            if (string.IsNullOrWhiteSpace(raw)
                || !Enum.GetNames(typeof(TEnum)).Any(name => string.Equals(name, raw, StringComparison.OrdinalIgnoreCase))
                || !Enum.TryParse(raw, true, out value))
            {
                throw new JsonSerializationException(
                    "Unsupported " + propertyName + " value: " + raw + ".");
            }

            return value;
        }
    }

    internal sealed class MeasureExpressionJsonConverter : JsonConverter
    {
        public override bool CanWrite => false;

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(MeasureExpression);
        }

        public override object? ReadJson(
            JsonReader reader,
            Type objectType,
            object? existingValue,
            JsonSerializer serializer)
        {
            var item = JObject.Load(reader);
            var token = item.GetValue("kind", StringComparison.OrdinalIgnoreCase);
            if (token == null || token.Type != JTokenType.String)
            {
                throw new JsonSerializationException("Property 'kind' is required for a measure expression.");
            }

            var rawKind = token.Value<string>();
            MeasureExpressionKind kind;
            if (string.IsNullOrWhiteSpace(rawKind)
                || !Enum.GetNames(typeof(MeasureExpressionKind)).Any(name => string.Equals(name, rawKind, StringComparison.OrdinalIgnoreCase))
                || !Enum.TryParse(rawKind, true, out kind))
            {
                throw new JsonSerializationException(
                    "Unsupported measure expression kind: " + rawKind + ".");
            }

            MeasureExpression target;
            switch (kind)
            {
                case MeasureExpressionKind.Aggregate:
                    target = new AggregateMeasureExpression();
                    break;
                case MeasureExpressionKind.FilteredAggregate:
                    target = new FilteredAggregateMeasureExpression();
                    break;
                case MeasureExpressionKind.WeightedAggregate:
                    target = new WeightedAggregateMeasureExpression();
                    break;
                case MeasureExpressionKind.Reference:
                    target = new ReferenceMeasureExpression();
                    break;
                case MeasureExpressionKind.Constant:
                    target = new ConstantMeasureExpression();
                    break;
                case MeasureExpressionKind.Binary:
                    target = new BinaryMeasureExpression();
                    break;
                case MeasureExpressionKind.SafeDivide:
                    target = new SafeDivideMeasureExpression();
                    break;
                case MeasureExpressionKind.Ratio:
                    target = new RatioMeasureExpression();
                    break;
                case MeasureExpressionKind.Difference:
                    target = new DifferenceMeasureExpression();
                    break;
                case MeasureExpressionKind.Share:
                    target = new ShareMeasureExpression();
                    break;
                default:
                    throw new JsonSerializationException("Unsupported measure expression kind: " + kind + ".");
            }

            using (var objectReader = item.CreateReader())
            {
                serializer.Populate(objectReader, target);
            }

            return target;
        }

        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer)
        {
            throw new NotSupportedException();
        }
    }
}
