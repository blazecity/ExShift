using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExShift.Util
{
    public class PersistableConverter : JsonConverter<IPersistable>
    {
        public override IPersistable Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            return JsonSerializer.Deserialize(ref reader, typeToConvert, options) as IPersistable;
        }

        public override void Write(Utf8JsonWriter writer, IPersistable value, JsonSerializerOptions options)
        {
            if (value == null)
            {
                JsonSerializer.Serialize(writer, null, options);
            }
            else
            {
                JsonSerializer.Serialize(writer, value, value.GetType(), options);
            }
        }
    }
}
