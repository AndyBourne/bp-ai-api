# 胰腺囊肿治疗协变量筛选与倾向性评分分析报告
**生成时间**: 2026-02-15 01:07:00.417964
**数据源**: /Users/wangguotao/Downloads/ISAR/Doctor/数据分析总表.xlsx
**结果目录**: /Users/wangguotao/bp-ai-api/AndyBourne-Blog/posts/Test/R_分析结果

## 一、分析流程
1. 核心4变量强制纳入PS：病灶位置、囊肿最大径、包裹性坏死、改良CTSI评分
2. 单因素筛选：按治疗P<0.2/结局P<0.1分层变量
3. VIF控制共线性：保留VIF<4的变量
4. 1:1 PS匹配：卡尺=0.1×PS标准差
5. 敏感性验证：模型稳定性、阈值鲁棒性、A-IPTW

## 二、核心结果
| 指标                | 数值                  |
|---------------------|-----------------------|
| 总样本量            | 143例       |
| PS模型变量数        | 7个       |
| PS模型AUC           | 0.734           |
| 匹配对数            | 11对      |
| 变量平衡率          | 0%        |

## 三、输出文件清单
### 3.1 统计结果文件
1. R_缺失值统计.xlsx - 变量缺失情况
2. R_单因素分析结果.xlsx - 变量P值与分类
3. R_VIF检验结果.xlsx - VIF值与最终变量
4. R_倾向性评分结果.xlsx - PS值与匹配数据
5. R_敏感性分析结果.xlsx - 稳定性验证

### 3.2 可视化图表
1. R_变量分类饼图.png - 变量分类分布
2. R_匹配后PS值分布.png - PS值对比
3. R_匹配前后样本量对比.png - 样本量变化

3. R_分析报告.md - 本报告

## 四、后续分析代码
# 1. 主要结局：影像学缓解率（卡方检验）
cont_tab <- table(matched_data$imaging_response, matched_data$treatment)
chisq_result <- chisq.test(cont_tab)
cat("缓解率（外科）：", round(sum(matched_data$imaging_response[matched_data$treatment==1])/sum(matched_data$treatment==1)*100,1), "%\n")
cat("缓解率（内镜）：", round(sum(matched_data$imaging_response[matched_data$treatment==0])/sum(matched_data$treatment==0)*100,1), "%\n")
cat("P值：", round(chisq_result$p.value,4), "\n")

# 2. 次要结局：住院时长（Wilcoxon检验）
if ("hospital_stay" %in% colnames(matched_data)) {
  wilcox_result <- wilcox.test(hospital_stay ~ treatment, data = matched_data)
  cat("住院时长中位数（外科）：", median(matched_data$hospital_stay[matched_data$treatment==1], na.rm=TRUE), "天\n")
  cat("住院时长中位数（内镜）：", median(matched_data$hospital_stay[matched_data$treatment==0], na.rm=TRUE), "天\n")
  cat("P值：", round(wilcox_result$p.value,4), "\n")
}

