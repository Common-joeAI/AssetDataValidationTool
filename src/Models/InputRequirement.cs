using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace AssetDataValidationTool.Models
{
    internal sealed class InputRequirement
    {
        [JsonPropertyName("label")]
        public string Label { get; set; } = "Source";

        [JsonPropertyName("description")]
        public string? Description { get; set; }

        [JsonPropertyName("patterns")]
        public List<string>? Patterns { get; set; }
    }
}
