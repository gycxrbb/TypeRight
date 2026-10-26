using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TypeRight.Llm
{
    /// <summary>低置信段落送检模型（PRD §4.3）。</summary>
    public struct ParagraphGuessInput
    {
        public int Index;
        public string Text;          // 前 50 字
        public string FeatureSummary;
    }

    public struct RoleGuess
    {
        public int Index;
        public string Role;
    }

    /// <summary>OpenAI 兼容 chat/completions 客户端。抽象出传输层以便单测。</summary>
    public class LlmClient
    {
        private readonly LlmConfig _config;

        public LlmClient(LlmConfig config)
        {
            if (config == null) throw new ArgumentNullException("config");
            _config = config;
        }

        /// <summary>构造请求体（单测用）。</summary>
        public static string BuildRequestBody(string model, IList<ParagraphGuessInput> items)
        {
            var systemPrompt =
                "你是 Word 文档段落角色标注器。给定段落列表，为每段返回一个角色。" +
                "角色取值只能是: title, heading1, heading2, heading3, body, caption, listItem, tableText, formulaLine, skip。" +
                "依据：编号模式（第X章/1.1/1.1.1/一、）、长度（标题通常≤40字）、加粗字号（大于正文的加粗短段多为标题）、" +
                "图注模式（图X/表X/Figure/Table开头）、整行公式特征（含=且几乎全是数学字符）。" +
                "只输出 JSON：{\"results\":[{\"i\":段序号,\"role\":\"角色\"}]}，不要输出任何其他内容。";

            var itemArray = new JArray();
            foreach (var it in items)
            {
                itemArray.Add(new JObject
                {
                    { "i", it.Index },
                    { "t", it.Text ?? "" },
                    { "f", it.FeatureSummary ?? "" }
                });
            }

            var body = new JObject
            {
                { "model", model },
                { "temperature", 0 },
                { "messages", new JArray
                    {
                        new JObject { { "role", "system" }, { "content", systemPrompt } },
                        new JObject { { "role", "user" }, { "content", itemArray.ToString(Newtonsoft.Json.Formatting.None) } }
                    }
                }
            };
            return body.ToString(Newtonsoft.Json.Formatting.None);
        }

        /// <summary>解析响应内容（单测用）。返回 i→role；无法解析时抛异常。</summary>
        public static List<RoleGuess> ParseResponse(string contentJson)
        {
            var result = new List<RoleGuess>();
            JObject root = JObject.Parse(ExtractJson(contentJson));
            JArray results = root["results"] as JArray;
            if (results == null) throw new InvalidDataException("响应缺少 results 数组。");
            foreach (var token in results)
            {
                JObject obj = token as JObject;
                if (obj == null) continue;
                int i;
                if (!int.TryParse(Convert.ToString(obj["i"]), out i)) continue;
                string role = Convert.ToString(obj["role"]);
                if (string.IsNullOrEmpty(role)) continue;
                result.Add(new RoleGuess { Index = i, Role = role });
            }
            if (result.Count == 0) throw new InvalidDataException("results 为空。");
            return result;
        }

        private static string ExtractJson(string content)
        {
            if (string.IsNullOrEmpty(content)) throw new InvalidDataException("响应为空。");
            string s = content.Trim();
            int start = s.IndexOf('{');
            int end = s.LastIndexOf('}');
            if (start < 0 || end <= start) throw new InvalidDataException("响应中未找到 JSON。");
            return s.Substring(start, end - start + 1);
        }

        /// <summary>批量请求。传输层可注入（单测）。</summary>
        public List<RoleGuess> ClassifyBatch(IList<ParagraphGuessInput> items, Func<string, string, string> transport = null)
        {
            if (items == null || items.Count == 0) return new List<RoleGuess>();
            string url = _config.BaseUrl.TrimEnd('/') + "/chat/completions";
            string body = BuildRequestBody(_config.Model, items);
            string content;
            if (transport != null)
            {
                content = transport(url, body);
            }
            else
            {
                content = Post(url, body);
            }
            return ParseResponse(content);
        }

        private string Post(string url, string body)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Timeout = Math.Max(1, _config.TimeoutSeconds) * 1000;
            request.ReadWriteTimeout = Math.Max(1, _config.TimeoutSeconds) * 1000;
            request.Headers.Add("Authorization", "Bearer " + _config.ApiKey);
            byte[] data = Encoding.UTF8.GetBytes(body);
            request.ContentLength = data.Length;
            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }
            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                string raw = reader.ReadToEnd();
                JObject j = JObject.Parse(raw);
                return Convert.ToString(j["choices"][0]["message"]["content"]);
            }
        }

        /// <summary>连通性测试。</summary>
        public string Ping()
        {
            var probe = new List<ParagraphGuessInput>
            {
                new ParagraphGuessInput { Index = 0, Text = "第一章 绪论", FeatureSummary = "len=6" }
            };
            var guesses = ClassifyBatch(probe);
            return guesses.Count > 0 ? guesses[0].Role : "(empty)";
        }
    }
}
