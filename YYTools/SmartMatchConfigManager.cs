using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace YYTools
{
    /// <summary>
    /// 智能匹配配置管理器
    /// </summary>
    public class SmartMatchConfigManager
    {
        private static readonly Lazy<SmartMatchConfigManager> _instance = new Lazy<SmartMatchConfigManager>(() => new SmartMatchConfigManager());
        public static SmartMatchConfigManager Instance => _instance.Value;

        private SmartMatchConfiguration _configuration;
        private readonly string _configPath;

        private SmartMatchConfigManager()
        {
            _configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "YYTools",
                "SmartMatchConfig.xml"
            );
            LoadConfiguration();
        }

        #region 配置管理

        /// <summary>
        /// 加载配置
        /// </summary>
        public void LoadConfiguration()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var serializer = new XmlSerializer(typeof(SmartMatchConfiguration));
                    using (var reader = new StreamReader(_configPath))
                    {
                        _configuration = (SmartMatchConfiguration)serializer.Deserialize(reader);
                    }
                    Logger.LogInfo("智能匹配配置已加载");
                }
                else
                {
                    CreateDefaultConfiguration();
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("加载智能匹配配置失败", ex);
                CreateDefaultConfiguration();
            }
        }

        /// <summary>
        /// 保存配置
        /// </summary>
        public void SaveConfiguration()
        {
            try
            {
                var directory = Path.GetDirectoryName(_configPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var serializer = new XmlSerializer(typeof(SmartMatchConfiguration));
                using (var writer = new StreamWriter(_configPath))
                {
                    serializer.Serialize(writer, _configuration);
                }

                Logger.LogInfo("智能匹配配置已保存");
            }
            catch (Exception ex)
            {
                Logger.LogError("保存智能匹配配置失败", ex);
            }
        }

        /// <summary>
        /// 创建默认配置
        /// </summary>
        private void CreateDefaultConfiguration()
        {
            _configuration = new SmartMatchConfiguration
            {
                WorksheetRules = new List<WorksheetMatchRule>
                {
                    // 发货明细工作表规则
                    new WorksheetMatchRule
                    {
                        Name = "发货明细",
                        Keywords = new[] { "发货明细", "发货" },
                        Priority = 10,
                        ExactMatch = true,
                        TargetType = "Shipping"
                    },
                    new WorksheetMatchRule
                    {
                        Name = "发货",
                        Keywords = new[] { "发货" },
                        Priority = 8,
                        ExactMatch = false,
                        TargetType = "Shipping"
                    },
                    
                    // 账单明细工作表规则
                    new WorksheetMatchRule
                    {
                        Name = "账单明细",
                        Keywords = new[] { "账单明细", "账单" },
                        Priority = 10,
                        ExactMatch = true,
                        TargetType = "Bill"
                    },
                    new WorksheetMatchRule
                    {
                        Name = "账单",
                        Keywords = new[] { "账单" },
                        Priority = 8,
                        ExactMatch = false,
                        TargetType = "Bill"
                    }
                },
                
                ColumnRules = new List<ColumnMatchRule>
                {
                    // 运单号列规则
                    new ColumnMatchRule
                    {
                        Name = "运单号",
                        Keywords = new[] { "快递单号", "运单号", "邮件号", "物流单号", "快递号" },
                        Priority = 10,
                        ExactMatch = false,
                        ColumnType = "TrackColumn"
                    },
                    new ColumnMatchRule
                    {
                        Name = "运单号通用",
                        Keywords = new[] { "运单", "快递", "物流", "单号", "tracking" },
                        Priority = 8,
                        ExactMatch = false,
                        ColumnType = "TrackColumn"
                    },
                    
                    // 商品编码列规则
                    new ColumnMatchRule
                    {
                        Name = "商品编码",
                        Keywords = new[] { "商品编码", "商品编号", "产品编码", "sku" },
                        Priority = 10,
                        ExactMatch = false,
                        ColumnType = "ProductColumn"
                    },
                    new ColumnMatchRule
                    {
                        Name = "商品编码通用",
                        Keywords = new[] { "编码" },
                        Priority = 8,
                        ExactMatch = false,
                        ColumnType = "ProductColumn"
                    },
                    new ColumnMatchRule
                    {
                        Name = "商品",
                        Keywords = new[] { "商品" },
                        Priority = 6,
                        ExactMatch = true,
                        ColumnType = "ProductColumn"
                    },
                    
                    // 商品名称列规则
                    new ColumnMatchRule
                    {
                        Name = "商品名称",
                        Keywords = new[] { "商品名称", "品名" },
                        Priority = 10,
                        ExactMatch = false,
                        ColumnType = "NameColumn"
                    },
                    new ColumnMatchRule
                    {
                        Name = "商品名称通用",
                        Keywords = new[] { "名称" },
                        Priority = 7,
                        ExactMatch = false,
                        ColumnType = "NameColumn"
                    },
                    
                    // 数量列规则
                    new ColumnMatchRule
                    {
                        Name = "数量",
                        Keywords = new[] { "数量", "件数", "qty", "quantity" },
                        Priority = 9,
                        ExactMatch = false,
                        ColumnType = "QuantityColumn"
                    },
                    
                    // 价格列规则
                    new ColumnMatchRule
                    {
                        Name = "价格",
                        Keywords = new[] { "价格", "单价", "金额", "price", "amount" },
                        Priority = 9,
                        ExactMatch = false,
                        ColumnType = "PriceColumn"
                    }
                },
                
                GlobalSettings = new GlobalMatchSettings
                {
                    EnableSmartMatching = true,
                    EnableExactMatchPriority = true,
                    EnableKeywordWeighting = true,
                    DefaultPriority = 5,
                    MaxKeywordsPerRule = 10,
                    MinMatchScore = 0.5
                }
            };
        }

        #endregion

        #region 工作表匹配

        /// <summary>
        /// 获取工作表匹配规则
        /// </summary>
        public List<WorksheetMatchRule> GetWorksheetRules(string targetType = null)
        {
            if (string.IsNullOrEmpty(targetType))
            {
                return _configuration.WorksheetRules.OrderByDescending(r => r.Priority).ToList();
            }
            
            return _configuration.WorksheetRules
                .Where(r => string.IsNullOrEmpty(r.TargetType) || r.TargetType.Equals(targetType, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(r => r.Priority)
                .ToList();
        }

        /// <summary>
        /// 智能匹配工作表
        /// </summary>
        public WorksheetMatchResult MatchWorksheet(string sheetName, string targetType = null)
        {
            try
            {
                if (string.IsNullOrEmpty(sheetName))
                {
                    return new WorksheetMatchResult { Success = false, Message = "工作表名称为空" };
                }

                var rules = GetWorksheetRules(targetType);
                var bestMatch = new WorksheetMatchResult { Success = false };

                foreach (var rule in rules)
                {
                    var score = CalculateWorksheetMatchScore(sheetName, rule);
                    if (score > bestMatch.Score)
                    {
                        bestMatch = new WorksheetMatchResult
                        {
                            Success = true,
                            SheetName = sheetName,
                            MatchedRule = rule,
                            Score = score,
                            Message = $"匹配规则: {rule.Name} (优先级: {rule.Priority})"
                        };
                    }
                }

                if (bestMatch.Success)
                {
                    Logger.LogUserAction("工作表智能匹配", $"工作表: {sheetName}, 规则: {bestMatch.MatchedRule.Name}, 分数: {bestMatch.Score:F2}", "成功");
                }
                else
                {
                    Logger.LogUserAction("工作表智能匹配", $"工作表: {sheetName}", "未找到匹配规则");
                }

                return bestMatch;
            }
            catch (Exception ex)
            {
                Logger.LogError($"工作表智能匹配失败: {sheetName}", ex);
                return new WorksheetMatchResult { Success = false, Message = $"匹配失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 计算工作表匹配分数
        /// </summary>
        private double CalculateWorksheetMatchScore(string sheetName, WorksheetMatchRule rule)
        {
            if (string.IsNullOrEmpty(sheetName) || rule?.Keywords == null)
                return 0;

            var sheetNameLower = sheetName.ToLowerInvariant();
            double totalScore = 0;
            int matchedKeywords = 0;

            foreach (var keyword in rule.Keywords)
            {
                var keywordLower = keyword.ToLowerInvariant();
                double keywordScore = 0;

                if (rule.ExactMatch)
                {
                    // 完全匹配模式
                    if (sheetNameLower.Equals(keywordLower))
                    {
                        keywordScore = 100;
                    }
                    else if (sheetNameLower.Contains(keywordLower))
                    {
                        keywordScore = 50;
                    }
                }
                else
                {
                    // 包含匹配模式
                    if (sheetNameLower.Equals(keywordLower))
                    {
                        keywordScore = 100;
                    }
                    else if (sheetNameLower.Contains(keywordLower))
                    {
                        keywordScore = keyword.Length * 2;
                    }
                }

                if (keywordScore > 0)
                {
                    totalScore += keywordScore;
                    matchedKeywords++;
                }
            }

            if (matchedKeywords == 0)
                return 0;

            // 计算平均分数并应用优先级权重
            var averageScore = totalScore / matchedKeywords;
            var priorityMultiplier = rule.Priority / 10.0;
            
            return averageScore * priorityMultiplier;
        }

        #endregion

        #region 列匹配

        /// <summary>
        /// 获取列匹配规则
        /// </summary>
        public List<ColumnMatchRule> GetColumnRules(string columnType = null)
        {
            if (string.IsNullOrEmpty(columnType))
            {
                return _configuration.ColumnRules.OrderByDescending(r => r.Priority).ToList();
            }
            
            return _configuration.ColumnRules
                .Where(r => string.IsNullOrEmpty(r.ColumnType) || r.ColumnType.Equals(columnType, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(r => r.Priority)
                .ToList();
        }

        /// <summary>
        /// 智能匹配列
        /// </summary>
        public ColumnMatchResult MatchColumn(string headerText, string columnType = null)
        {
            try
            {
                if (string.IsNullOrEmpty(headerText))
                {
                    return new ColumnMatchResult { Success = false, Message = "列标题为空" };
                }

                var rules = GetColumnRules(columnType);
                var bestMatch = new ColumnMatchResult { Success = false };

                foreach (var rule in rules)
                {
                    var score = CalculateColumnMatchScore(headerText, rule);
                    if (score > bestMatch.Score)
                    {
                        bestMatch = new ColumnMatchResult
                        {
                            Success = true,
                            HeaderText = headerText,
                            MatchedRule = rule,
                            Score = score,
                            Message = $"匹配规则: {rule.Name} (优先级: {rule.Priority})"
                        };
                    }
                }

                if (bestMatch.Success)
                {
                    Logger.LogUserAction("列智能匹配", $"列标题: {headerText}, 规则: {bestMatch.MatchedRule.Name}, 分数: {bestMatch.Score:F2}", "成功");
                }
                else
                {
                    Logger.LogUserAction("列智能匹配", $"列标题: {headerText}", "未找到匹配规则");
                }

                return bestMatch;
            }
            catch (Exception ex)
            {
                Logger.LogError($"列智能匹配失败: {headerText}", ex);
                return new ColumnMatchResult { Success = false, Message = $"匹配失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 计算列匹配分数
        /// </summary>
        private double CalculateColumnMatchScore(string headerText, ColumnMatchRule rule)
        {
            if (string.IsNullOrEmpty(headerText) || rule?.Keywords == null)
                return 0;

            var headerLower = headerText.ToLowerInvariant();
            double totalScore = 0;
            int matchedKeywords = 0;

            foreach (var keyword in rule.Keywords)
            {
                var keywordLower = keyword.ToLowerInvariant();
                double keywordScore = 0;

                if (rule.ExactMatch)
                {
                    // 完全匹配模式
                    if (headerLower.Equals(keywordLower))
                    {
                        keywordScore = 100;
                    }
                    else if (headerLower.Contains(keywordLower))
                    {
                        keywordScore = 50;
                    }
                }
                else
                {
                    // 包含匹配模式
                    if (headerLower.Equals(keywordLower))
                    {
                        keywordScore = 100;
                    }
                    else if (headerLower.Contains(keywordLower))
                    {
                        keywordScore = keyword.Length * 2;
                    }
                }

                if (keywordScore > 0)
                {
                    totalScore += keywordScore;
                    matchedKeywords++;
                }
            }

            if (matchedKeywords == 0)
                return 0;

            // 计算平均分数并应用优先级权重
            var averageScore = totalScore / matchedKeywords;
            var priorityMultiplier = rule.Priority / 10.0;
            
            return averageScore * priorityMultiplier;
        }

        #endregion

        #region 配置编辑

        /// <summary>
        /// 添加工作表匹配规则
        /// </summary>
        public bool AddWorksheetRule(WorksheetMatchRule rule)
        {
            try
            {
                if (rule == null || string.IsNullOrEmpty(rule.Name))
                    return false;

                // 检查是否已存在同名规则
                var existingRule = _configuration.WorksheetRules.FirstOrDefault(r => r.Name.Equals(rule.Name, StringComparison.OrdinalIgnoreCase));
                if (existingRule != null)
                {
                    _configuration.WorksheetRules.Remove(existingRule);
                }

                _configuration.WorksheetRules.Add(rule);
                SaveConfiguration();
                
                Logger.LogUserAction("添加工作表匹配规则", $"规则名称: {rule.Name}, 关键字: {string.Join(", ", rule.Keywords)}", "成功");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"添加工作表匹配规则失败: {rule?.Name}", ex);
                return false;
            }
        }

        /// <summary>
        /// 添加列匹配规则
        /// </summary>
        public bool AddColumnRule(ColumnMatchRule rule)
        {
            try
            {
                if (rule == null || string.IsNullOrEmpty(rule.Name))
                    return false;

                // 检查是否已存在同名规则
                var existingRule = _configuration.ColumnRules.FirstOrDefault(r => r.Name.Equals(rule.Name, StringComparison.OrdinalIgnoreCase));
                if (existingRule != null)
                {
                    _configuration.ColumnRules.Remove(existingRule);
                }

                _configuration.ColumnRules.Add(rule);
                SaveConfiguration();
                
                Logger.LogUserAction("添加列匹配规则", $"规则名称: {rule.Name}, 关键字: {string.Join(", ", rule.Keywords)}", "成功");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"添加列匹配规则失败: {rule?.Name}", ex);
                return false;
            }
        }

        /// <summary>
        /// 删除工作表匹配规则
        /// </summary>
        public bool RemoveWorksheetRule(string ruleName)
        {
            try
            {
                var rule = _configuration.WorksheetRules.FirstOrDefault(r => r.Name.Equals(ruleName, StringComparison.OrdinalIgnoreCase));
                if (rule != null)
                {
                    _configuration.WorksheetRules.Remove(rule);
                    SaveConfiguration();
                    
                    Logger.LogUserAction("删除工作表匹配规则", $"规则名称: {ruleName}", "成功");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Logger.LogError($"删除工作表匹配规则失败: {ruleName}", ex);
                return false;
            }
        }

        /// <summary>
        /// 删除列匹配规则
        /// </summary>
        public bool RemoveColumnRule(string ruleName)
        {
            try
            {
                var rule = _configuration.ColumnRules.FirstOrDefault(r => r.Name.Equals(ruleName, StringComparison.OrdinalIgnoreCase));
                if (rule != null)
                {
                    _configuration.ColumnRules.Remove(rule);
                    SaveConfiguration();
                    
                    Logger.LogUserAction("删除列匹配规则", $"规则名称: {ruleName}", "成功");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Logger.LogError($"删除列匹配规则失败: {ruleName}", ex);
                return false;
            }
        }

        /// <summary>
        /// 更新全局设置
        /// </summary>
        public bool UpdateGlobalSettings(GlobalMatchSettings settings)
        {
            try
            {
                if (settings == null)
                    return false;

                _configuration.GlobalSettings = settings;
                SaveConfiguration();
                
                Logger.LogUserAction("更新全局匹配设置", "智能匹配配置", "成功");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError("更新全局匹配设置失败", ex);
                return false;
            }
        }

        #endregion

        #region 配置导出导入

        /// <summary>
        /// 导出配置到文件
        /// </summary>
        public bool ExportConfiguration(string filePath)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(SmartMatchConfiguration));
                using (var writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, _configuration);
                }
                
                Logger.LogUserAction("导出智能匹配配置", $"文件路径: {filePath}", "成功");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"导出智能匹配配置失败: {filePath}", ex);
                return false;
            }
        }

        /// <summary>
        /// 从文件导入配置
        /// </summary>
        public bool ImportConfiguration(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return false;

                var serializer = new XmlSerializer(typeof(SmartMatchConfiguration));
                using (var reader = new StreamReader(filePath))
                {
                    var importedConfig = (SmartMatchConfiguration)serializer.Deserialize(reader);
                    _configuration = importedConfig;
                }
                
                SaveConfiguration();
                
                Logger.LogUserAction("导入智能匹配配置", $"文件路径: {filePath}", "成功");
                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError($"导入智能匹配配置失败: {filePath}", ex);
                return false;
            }
        }

        #endregion

        #region 配置验证

        /// <summary>
        /// 验证配置有效性
        /// </summary>
        public List<string> ValidateConfiguration()
        {
            var errors = new List<string>();

            try
            {
                // 验证工作表规则
                if (_configuration.WorksheetRules != null)
                {
                    foreach (var rule in _configuration.WorksheetRules)
                    {
                        if (string.IsNullOrEmpty(rule.Name))
                            errors.Add("工作表规则名称不能为空");
                        
                        if (rule.Keywords == null || rule.Keywords.Length == 0)
                            errors.Add($"工作表规则 '{rule.Name}' 关键字不能为空");
                        
                        if (rule.Priority < 1 || rule.Priority > 100)
                            errors.Add($"工作表规则 '{rule.Name}' 优先级必须在1-100之间");
                    }
                }

                // 验证列规则
                if (_configuration.ColumnRules != null)
                {
                    foreach (var rule in _configuration.ColumnRules)
                    {
                        if (string.IsNullOrEmpty(rule.Name))
                            errors.Add("列规则名称不能为空");
                        
                        if (rule.Keywords == null || rule.Keywords.Length == 0)
                            errors.Add($"列规则 '{rule.Name}' 关键字不能为空");
                        
                        if (rule.Priority < 1 || rule.Priority > 100)
                            errors.Add($"列规则 '{rule.Name}' 优先级必须在1-100之间");
                    }
                }

                // 验证全局设置
                if (_configuration.GlobalSettings != null)
                {
                    if (_configuration.GlobalSettings.DefaultPriority < 1 || _configuration.GlobalSettings.DefaultPriority > 100)
                        errors.Add("默认优先级必须在1-100之间");
                    
                    if (_configuration.GlobalSettings.MinMatchScore < 0 || _configuration.GlobalSettings.MinMatchScore > 1)
                        errors.Add("最小匹配分数必须在0-1之间");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"配置验证失败: {ex.Message}");
            }

            return errors;
        }

        #endregion
    }

    #region 配置类

    /// <summary>
    /// 智能匹配配置
    /// </summary>
    [Serializable]
    public class SmartMatchConfiguration
    {
        public List<WorksheetMatchRule> WorksheetRules { get; set; } = new List<WorksheetMatchRule>();
        public List<ColumnMatchRule> ColumnRules { get; set; } = new List<ColumnMatchRule>();
        public GlobalMatchSettings GlobalSettings { get; set; } = new GlobalMatchSettings();
    }

    /// <summary>
    /// 工作表匹配规则
    /// </summary>
    [Serializable]
    public class WorksheetMatchRule
    {
        public string Name { get; set; }
        public string[] Keywords { get; set; }
        public int Priority { get; set; }
        public bool ExactMatch { get; set; }
        public string TargetType { get; set; } // Shipping, Bill, 或其他
    }

    /// <summary>
    /// 列匹配规则
    /// </summary>
    [Serializable]
    public class ColumnMatchRule
    {
        public string Name { get; set; }
        public string[] Keywords { get; set; }
        public int Priority { get; set; }
        public bool ExactMatch { get; set; }
        public string ColumnType { get; set; } // TrackColumn, ProductColumn, NameColumn, 等
    }

    /// <summary>
    /// 全局匹配设置
    /// </summary>
    [Serializable]
    public class GlobalMatchSettings
    {
        public bool EnableSmartMatching { get; set; } = true;
        public bool EnableExactMatchPriority { get; set; } = true;
        public bool EnableKeywordWeighting { get; set; } = true;
        public int DefaultPriority { get; set; } = 5;
        public int MaxKeywordsPerRule { get; set; } = 10;
        public double MinMatchScore { get; set; } = 0.5;
    }

    #endregion

    #region 匹配结果类

    /// <summary>
    /// 工作表匹配结果
    /// </summary>
    public class WorksheetMatchResult
    {
        public bool Success { get; set; }
        public string SheetName { get; set; }
        public WorksheetMatchRule MatchedRule { get; set; }
        public double Score { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// 列匹配结果
    /// </summary>
    public class ColumnMatchResult
    {
        public bool Success { get; set; }
        public string HeaderText { get; set; }
        public ColumnMatchRule MatchedRule { get; set; }
        public double Score { get; set; }
        public string Message { get; set; }
    }

    #endregion
}