using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;

namespace TypeRight.Profiles
{
    public sealed class LoadedProfile
    {
        public Profile Profile { get; set; }
        public string FilePath { get; set; }          // null = 内置
        public bool IsBuiltIn { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>方案读写/导入导出：内置方案（嵌入资源）+ %APPDATA%\TypeRight\profiles\ 下的用户方案。</summary>
    public sealed class ProfileStore
    {
        public const string BuiltInResourceId = "TypeRight.Profiles.BuiltIn.academic-lecture.json";
        public static readonly string ProfilesDirectory =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TypeRight", "profiles");

        private static readonly Lazy<Profile> _builtIn = new Lazy<Profile>(LoadBuiltInProfile, true);

        public static Profile BuiltInProfile { get { return _builtIn.Value; } }

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            NullValueHandling = NullValueHandling.Include,
            Formatting = Newtonsoft.Json.Formatting.Indented
        };

        public ProfileStore()
        {
            Directory.CreateDirectory(ProfilesDirectory);
        }

        /// <summary>内置 + 用户目录下全部可用方案；非法文件被拒绝并记录到 loadErrors。</summary>
        public List<LoadedProfile> LoadAll(List<string> loadErrors)
        {
            var result = new List<LoadedProfile>();
            var builtin = new LoadedProfile
            {
                Profile = BuiltInProfile,
                FilePath = null,
                IsBuiltIn = true,
                Warnings = ProfileValidator.WarnNotInstalledFonts(BuiltInProfile)
            };
            result.Add(builtin);

            if (!Directory.Exists(ProfilesDirectory)) return result;
            foreach (string file in Directory.GetFiles(ProfilesDirectory, "*.json").OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
            {
                Profile p;
                List<string> errors;
                if (TryLoad(file, out p, out errors))
                {
                    result.Add(new LoadedProfile
                    {
                        Profile = p,
                        FilePath = file,
                        IsBuiltIn = false,
                        Warnings = ProfileValidator.WarnNotInstalledFonts(p)
                    });
                }
                else
                {
                    loadErrors.Add(Path.GetFileName(file) + ": " + string.Join(" ", errors));
                }
            }
            return result;
        }

        public static bool TryLoad(string file, out Profile profile, out List<string> errors)
        {
            profile = null;
            errors = new List<string>();
            try
            {
                string json = File.ReadAllText(file, Encoding.UTF8);
                profile = Parse(json);
                errors = ProfileValidator.Validate(profile);
                if (errors.Count > 0)
                {
                    profile = null;
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                errors.Add("JSON 解析失败: " + ex.Message);
                profile = null;
                return false;
            }
        }

        /// <summary>校验失败抛 ProfileValidationException，调用方提示用户，不产生半成品。</summary>
        public static Profile Parse(string json)
        {
            var errors = new List<string>();
            Profile p;
            try
            {
                p = JsonConvert.DeserializeObject<Profile>(json);
            }
            catch (Exception ex)
            {
                throw new ProfileValidationException("JSON 解析失败: " + ex.Message, new List<string> { ex.Message });
            }
            errors = ProfileValidator.Validate(p);
            if (errors.Count > 0)
                throw new ProfileValidationException("方案校验未通过。", errors);
            return p;
        }

        /// <summary>保存（导出）到用户方案目录。同名覆盖前先备份为 .bak。</summary>
        public string Save(Profile profile)
        {
            if (profile == null) throw new ArgumentNullException("profile");
            var errors = ProfileValidator.Validate(profile);
            if (errors.Count > 0)
                throw new ProfileValidationException("方案校验未通过，已拒绝保存。", errors);

            Directory.CreateDirectory(ProfilesDirectory);
            string file = Path.Combine(ProfilesDirectory, profile.Id + ".json");
            if (File.Exists(file))
                File.Copy(file, file + ".bak", true);
            string json = JsonConvert.SerializeObject(profile, JsonSettings);
            File.WriteAllText(file, json, new UTF8Encoding(false));
            return file;
        }

        public bool Delete(string profileId)
        {
            string file = Path.Combine(ProfilesDirectory, profileId + ".json");
            if (File.Exists(file))
            {
                File.Delete(file);
                return true;
            }
            return false;
        }

        /// <summary>把内置方案复制到用户目录，便于在其基础上修改。</summary>
        public string DuplicateBuiltIn(string newId, string newName)
        {
            var copy = JsonConvert.DeserializeObject<Profile>(JsonConvert.SerializeObject(BuiltInProfile));
            copy.Id = newId;
            copy.Name = newName;
            return Save(copy);
        }

        private static Profile LoadBuiltInProfile()
        {
            Assembly asm = typeof(ProfileStore).Assembly;
            using (Stream stream = asm.GetManifestResourceStream(BuiltInResourceId))
            {
                if (stream == null)
                    throw new InvalidOperationException("内置方案资源缺失: " + BuiltInResourceId);
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    string json = reader.ReadToEnd();
                    var p = JsonConvert.DeserializeObject<Profile>(json);
                    var errors = ProfileValidator.Validate(p);
                    if (errors.Count > 0)
                        throw new ProfileValidationException("内置方案校验失败。", errors);
                    return p;
                }
            }
        }
    }

    [Serializable]
    public sealed class ProfileValidationException : Exception
    {
        public List<string> Errors { get; private set; }

        public ProfileValidationException(string message, List<string> errors)
            : base(message)
        {
            Errors = errors ?? new List<string>();
        }
    }
}
