using System;
using System.Collections.Generic;
using System.Linq;
using TypeRight.Roles;
using TypeRight.Scanning;

namespace TypeRight.Llm
{
    /// <summary>
    /// 仅对低置信段落（R11/R12）调用 LLM；失败/超时自动回落到规则结果（PRD §4.3）。
    /// </summary>
    public sealed class AmbiguityResolver
    {
        private readonly LlmConfig _config;

        public AmbiguityResolver(LlmConfig config)
        {
            _config = config ?? new LlmConfig { Enabled = false };
        }

        public sealed class ResolveResult
        {
            public bool UsedLlm;
            public int AdjustedCount;
            public string Error;
        }

        /// <summary>就地修改低置信段落的 Role；任何失败保持规则结果。</summary>
        public ResolveResult Resolve(IList<ParagraphFeatures> features)
        {
            var result = new ResolveResult();
            if (!_config.Enabled) return result;
            if (string.IsNullOrWhiteSpace(_config.BaseUrl) || string.IsNullOrWhiteSpace(_config.Model))
            {
                result.Error = "LLM 配置不完整（BaseURL / 模型名缺失）。";
                return result;
            }

            var lowConfidence = features.Where(f => !f.HighConfidence && f.Role != ParagraphRole.Skip).ToList();
            if (lowConfidence.Count == 0)
            {
                result.UsedLlm = true;
                return result;
            }

            try
            {
                var inputs = lowConfidence.Select(f => new ParagraphGuessInput
                {
                    Index = f.Index,
                    Text = f.Text.Length > 50 ? f.Text.Substring(0, 50) : f.Text,
                    FeatureSummary = string.Format(
                        "len={0};size={1};bold={2};align={3};pat={4}",
                        f.Text.Length,
                        double.IsNaN(f.FontSize) ? "?" : f.FontSize.ToString("0.#"),
                        (f.Bold || f.BoldMixed) ? 1 : 0,
                        f.Alignment,
                        f.NumberingPattern)
                }).ToList();

                var client = new LlmClient(_config);
                var guesses = client.ClassifyBatch(inputs);

                var byIndex = new Dictionary<int, string>();
                foreach (var g in guesses) byIndex[g.Index] = g.Role;

                foreach (var f in lowConfidence)
                {
                    string roleKey;
                    if (!byIndex.TryGetValue(f.Index, out roleKey)) continue;
                    var role = ParagraphRoles.TryParse(roleKey);
                    if (role.HasValue && role.Value != f.Role)
                    {
                        f.Role = role.Value;
                        result.AdjustedCount++;
                    }
                }
                result.UsedLlm = true;
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
                // 失败回落：保持规则结果，不抛出
            }
            return result;
        }
    }
}
