# =========================================================
# Part 2. Measurement checks (CFA + reliability) [FINAL]
# =========================================================

rm(list = ls())
set.seed(123)
options(stringsAsFactors = FALSE)

library(dplyr)
library(tidyr)
library(lavaan)
library(psych)
library(openxlsx)

# ---------------------------------------------------------
# 0) Load data
# ---------------------------------------------------------
in_rds <- "combined_data_clean.rds"   # Part 1 output
out_keyed_rds <- "combined_data_clean_keyed.rds"  # Part 2에서 BFI reverse 자동반영 후 저장

combined_data <- readRDS(in_rds)

# 필수 변수 체크
need_cols <- c("year",
               "extraversion_1","extraversion_2",
               "agreeableness_1","agreeableness_2",
               "conscientiousness_1","conscientiousness_2",
               "neuroticism_1","neuroticism_2",
               "openness_1","openness_2",
               "enjoyment","calm","vitality","worry","sadness","depression","anger","stress","tiredness","loneliness")
miss_cols <- setdiff(need_cols, names(combined_data))
if (length(miss_cols) > 0) stop("Missing columns in combined_data: ", paste(miss_cols, collapse=", "))

# ---------------------------------------------------------
# A) BFI-10: year-wise reverse auto-diagnosis & apply
# ---------------------------------------------------------

# 진단용: "reverse 후보 문항"을 뒤집을 때(6-x) vs 안 뒤집을 때 상관 비교
# 기대: 같은 trait 2문항은 '정(+) 상관'이 되어야 함
bfi_diag <- combined_data %>%
  select(year,
         extraversion_1, extraversion_2,
         agreeableness_1, agreeableness_2,
         conscientiousness_1, conscientiousness_2,
         neuroticism_1, neuroticism_2,
         openness_1, openness_2) %>%
  group_by(year) %>%
  summarise(
    # Extraversion: reverse candidate = extraversion_1 (introverted item)
    ext_norev = cor(extraversion_1, extraversion_2, method="spearman"),
    ext_rev   = cor(6 - extraversion_1, extraversion_2, method="spearman"),
    ext_need_rev = ext_rev > ext_norev,
    
    # Agreeableness: reverse candidate = agreeableness_2 (critical item)
    agr_norev = cor(agreeableness_1, agreeableness_2, method="spearman"),
    agr_rev   = cor(agreeableness_1, 6 - agreeableness_2, method="spearman"),
    agr_need_rev = agr_rev > agr_norev,
    
    # Conscientiousness: reverse candidate = conscientiousness_1 (careless item)
    con_norev = cor(conscientiousness_1, conscientiousness_2, method="spearman"),
    con_rev   = cor(6 - conscientiousness_1, conscientiousness_2, method="spearman"),
    con_need_rev = con_rev > con_norev,
    
    # Neuroticism: reverse candidate = neuroticism_1 (calm item)
    neu_norev = cor(neuroticism_1, neuroticism_2, method="spearman"),
    neu_rev   = cor(6 - neuroticism_1, neuroticism_2, method="spearman"),
    neu_need_rev = neu_rev > neu_norev,
    
    # Openness: reverse candidate = openness_1 (not creative item)
    ope_norev = cor(openness_1, openness_2, method="spearman"),
    ope_rev   = cor(6 - openness_1, openness_2, method="spearman"),
    ope_need_rev = ope_rev > ope_norev,
    
    .groups = "drop"
  )

print(bfi_diag)

# 연도별 reverse 적용 (원 컬럼을 덮어써서 이후 Part 3에서도 동일 컬럼명 사용 가능)
combined_keyed <- combined_data %>%
  left_join(bfi_diag %>% select(year, ext_need_rev, agr_need_rev, con_need_rev, neu_need_rev, ope_need_rev),
            by = "year") %>%
  mutate(
    extraversion_1       = ifelse(ext_need_rev, 6 - extraversion_1, extraversion_1),
    agreeableness_2      = ifelse(agr_need_rev, 6 - agreeableness_2, agreeableness_2),
    conscientiousness_1  = ifelse(con_need_rev, 6 - conscientiousness_1, conscientiousness_1),
    neuroticism_1        = ifelse(neu_need_rev, 6 - neuroticism_1, neuroticism_1),
    openness_1           = ifelse(ope_need_rev, 6 - openness_1, openness_1)
  ) %>%
  select(-ext_need_rev, -agr_need_rev, -con_need_rev, -neu_need_rev, -ope_need_rev)

# 저장(Part 3에서 이 파일을 우선 사용)
saveRDS(combined_keyed, out_keyed_rds)

# ---------------------------------------------------------
# A1) BFI QC after keying
# ---------------------------------------------------------
bfi_vars <- c("extraversion_1","extraversion_2",
              "agreeableness_1","agreeableness_2",
              "conscientiousness_1","conscientiousness_2",
              "neuroticism_1","neuroticism_2",
              "openness_1","openness_2")

bfi_raw <- combined_keyed %>% select(all_of(bfi_vars))

cat("\n[BFI-10] Range check (should be 1–5)\n")
print(sapply(bfi_raw, function(x) range(x, na.rm=TRUE)))

cat("\n[BFI-10] Missing count\n")
print(sapply(bfi_raw, function(x) sum(is.na(x))))

bfi_cc <- bfi_raw %>% drop_na()

# 2-item pair correlations (overall)
pair_spearman <- tibble(
  Trait = c("Extraversion","Agreeableness","Conscientiousness","Neuroticism","Openness"),
  r_spearman = c(
    cor(bfi_cc$extraversion_1, bfi_cc$extraversion_2, method="spearman"),
    cor(bfi_cc$agreeableness_1, bfi_cc$agreeableness_2, method="spearman"),
    cor(bfi_cc$conscientiousness_1, bfi_cc$conscientiousness_2, method="spearman"),
    cor(bfi_cc$neuroticism_1, bfi_cc$neuroticism_2, method="spearman"),
    cor(bfi_cc$openness_1, bfi_cc$openness_2, method="spearman")
  )
)
cat("\n[BFI-10] Pair correlations (Spearman)\n")
print(pair_spearman)

# year-stratified pair correlations (after auto-keying)
pair_by_year <- combined_keyed %>%
  select(year, all_of(bfi_vars)) %>%
  drop_na() %>%
  group_by(year) %>%
  summarise(
    ext = cor(extraversion_1, extraversion_2, method="spearman"),
    agr = cor(agreeableness_1, agreeableness_2, method="spearman"),
    con = cor(conscientiousness_1, conscientiousness_2, method="spearman"),
    neu = cor(neuroticism_1, neuroticism_2, method="spearman"),
    ope = cor(openness_1, openness_2, method="spearman"),
    .groups="drop"
  )
cat("\n[BFI-10] Year-stratified pair correlations (Spearman, AFTER keying)\n")
print(pair_by_year)

# polychoric correlation matrix + PD check
pc_all <- psych::polychoric(bfi_cc)
rho <- pc_all$rho
eig_min <- min(eigen(rho, symmetric = TRUE)$values)
cat("\n[BFI-10] Polychoric rho min eigenvalue (should be > 0):", eig_min, "\n")

pair_poly <- tibble(
  Trait = c("Extraversion","Agreeableness","Conscientiousness","Neuroticism","Openness"),
  r_polychoric = c(
    rho["extraversion_1","extraversion_2"],
    rho["agreeableness_1","agreeableness_2"],
    rho["conscientiousness_1","conscientiousness_2"],
    rho["neuroticism_1","neuroticism_2"],
    rho["openness_1","openness_2"]
  )
)
cat("\n[BFI-10] Pair correlations (Polychoric)\n")
print(pair_poly)

# Spearman–Brown (2-item reliability)
sb_tbl <- pair_spearman %>%
  rename(r = r_spearman) %>%
  mutate(SB_spearman = (2*r)/(1+r)) %>%
  left_join(pair_poly, by="Trait") %>%
  mutate(SB_polychoric = (2*r_polychoric)/(1+r_polychoric))

cat("\n[BFI-10] Spearman–Brown reliability\n")
print(sb_tbl)

# ---------------------------------------------------------
# A2) BFI-10 CFA (Ordinal WLSMV; equal-loading constraints)
# ---------------------------------------------------------
# NOTE: 2-item factors는 CFA가 구조적으로 불안정/민감합니다.
# (리뷰 대응용 '진단/보고' 목적. 해석은 신중히.)

bfi_ord <- bfi_cc %>% mutate(across(everything(), ~ ordered(.)))

bfi10_model_eq <- '
Extraversion =~ lE*extraversion_1 + lE*extraversion_2
Agreeableness =~ lA*agreeableness_1 + lA*agreeableness_2
Conscientiousness =~ lC*conscientiousness_1 + lC*conscientiousness_2
Neuroticism =~ lN*neuroticism_1 + lN*neuroticism_2
Openness =~ lO*openness_1 + lO*openness_2
'

fit_bfi_wlsmv <- cfa(
  model = bfi10_model_eq,
  data = bfi_ord,
  ordered = names(bfi_ord),
  estimator = "WLSMV",
  parameterization = "theta",
  std.lv = TRUE,
  auto.fix.first = FALSE,
  control = list(iter.max = 5000)
)

conv_bfi <- lavInspect(fit_bfi_wlsmv, "converged")
cat("\n[BFI CFA WLSMV] Converged?:", conv_bfi, "\n")

bfi_fitmeas <- NA
bfi_loadings <- NULL
bfi_warn_covlv <- NULL

if (conv_bfi) {
  bfi_fitmeas <- fitMeasures(fit_bfi_wlsmv,
                             c("chisq","df","cfi","tli","rmsea","rmsea.ci.lower","rmsea.ci.upper","srmr"))
  bfi_fitmeas <- as.data.frame(t(bfi_fitmeas))
  bfi_fitmeas$model <- "BFI10_WLSMV_equalLoadings"
  
  pe <- parameterEstimates(fit_bfi_wlsmv, standardized = TRUE)
  bfi_loadings <- pe %>% filter(op=="=~") %>%
    select(lhs, rhs, est, se, z, pvalue, std.all)
  
  # latent covariance PD warning 확인용
  bfi_warn_covlv <- tryCatch(lavInspect(fit_bfi_wlsmv, "cov.lv"), error = function(e) NULL)
}

# Sensitivity CFA: treat as continuous (MLR)
fit_bfi_mlr <- cfa(
  model = bfi10_model_eq,
  data = bfi_cc,
  estimator = "MLR",
  std.lv = TRUE,
  auto.fix.first = FALSE,
  control = list(iter.max = 5000)
)
conv_bfi_mlr <- lavInspect(fit_bfi_mlr, "converged")
cat("\n[BFI CFA MLR] Converged?:", conv_bfi_mlr, "\n")

bfi_fitmeas_mlr <- NA
if (conv_bfi_mlr) {
  bfi_fitmeas_mlr <- fitMeasures(fit_bfi_mlr,
                                 c("chisq","df","cfi","tli","rmsea","rmsea.ci.lower","rmsea.ci.upper","srmr"))
  bfi_fitmeas_mlr <- as.data.frame(t(bfi_fitmeas_mlr))
  bfi_fitmeas_mlr$model <- "BFI10_MLR_equalLoadings"
}

bfi_fitmeas_mlr
# ---------------------------------------------------------
# A3) Composite reliability (compRelSEM) — robust fallback
# ---------------------------------------------------------
# compRelSEM에서 "isShared" 에러가 나면 (semTools/lavaan 버전 mismatch),
# 1) 업데이트 권장: install.packages(c("lavaan","semTools"))
# 2) 일단은 'Spearman–Brown/ordinal alpha'를 주 보고치로 사용

rel_comp <- NULL
if (requireNamespace("semTools", quietly=TRUE)) {
  # compRelSEM 시도
  rel_comp <- tryCatch(
    semTools::compRelSEM(fit_bfi_wlsmv),
    error = function(e) e
  )
  cat("\n[BFI] compRelSEM result (or error):\n")
  print(rel_comp)
  
  # 실패하면 semTools::reliability(구버전/Deprecated) fallback (가능하면)
  if (inherits(rel_comp, "error")) {
    rel_fallback <- tryCatch(
      semTools::reliability(fit_bfi_wlsmv),
      error = function(e) e
    )
    cat("\n[BFI] Fallback semTools::reliability (or error):\n")
    print(rel_fallback)
  }
} else {
  cat("\n[BFI] semTools not installed. Skip compRelSEM.\n")
}

# ---------------------------------------------------------
# B) Affect CFA + reliability
# ---------------------------------------------------------
affect_vars <- c("enjoyment","calm","vitality",
                 "worry","sadness","depression","anger","stress","tiredness","loneliness")

affect_raw <- combined_keyed %>% select(all_of(affect_vars)) %>% drop_na()

cat("\n[Affect] Range check (should be 0–10)\n")
print(sapply(affect_raw, function(x) range(x, na.rm=TRUE)))

affect_model <- '
Positive =~ enjoyment + calm + vitality
Negative =~ worry + sadness + depression + anger + stress + tiredness + loneliness
'

# Primary: MLR (treat as continuous)
fit_aff_mlr <- cfa(affect_model, data=affect_raw, estimator="MLR")
cat("\n[Affect CFA MLR] Converged?:", lavInspect(fit_aff_mlr, "converged"), "\n")

aff_fit_mlr <- fitMeasures(fit_aff_mlr,
                           c("chisq","df","cfi","tli","rmsea","rmsea.ci.lower","rmsea.ci.upper","srmr"))
aff_fit_mlr <- as.data.frame(t(aff_fit_mlr))
aff_fit_mlr$model <- "Affect_MLR"

pe_aff_mlr <- parameterEstimates(fit_aff_mlr, standardized=TRUE)
aff_load_mlr <- pe_aff_mlr %>% filter(op=="=~") %>%
  select(lhs, rhs, est, se, z, pvalue, std.all)

# Sensitivity: WLSMV (treat as ordinal 0–10)
fit_aff_wlsmv <- cfa(affect_model, data=affect_raw,
                     ordered = names(affect_raw),
                     estimator="WLSMV",
                     parameterization="theta",
                     std.lv=TRUE)
cat("\n[Affect CFA WLSMV] Converged?:", lavInspect(fit_aff_wlsmv, "converged"), "\n")

aff_fit_wlsmv <- fitMeasures(fit_aff_wlsmv,
                             c("chisq","df","cfi","tli","rmsea","rmsea.ci.lower","rmsea.ci.upper","srmr"))
aff_fit_wlsmv <- as.data.frame(t(aff_fit_wlsmv))
aff_fit_wlsmv$model <- "Affect_WLSMV"

pe_aff_wlsmv <- parameterEstimates(fit_aff_wlsmv, standardized=TRUE)
aff_load_wlsmv <- pe_aff_wlsmv %>% filter(op=="=~") %>%
  select(lhs, rhs, est, se, z, pvalue, std.all)

# Reliability (alpha) — 경고 줄이기 위해 max=11 명시 (0~10 총 11 범주)
alpha_pos <- psych::alpha(affect_raw[,c("enjoyment","calm","vitality")], max=11)
alpha_neg <- psych::alpha(affect_raw[,c("worry","sadness","depression","anger","stress","tiredness","loneliness")], max=11)

# ---------------------------------------------------------
# C) Save outputs (one Excel workbook)
# ---------------------------------------------------------
wb <- createWorkbook()

addWorksheet(wb, "BFI_reverse_diagnosis")
writeData(wb, "BFI_reverse_diagnosis", bfi_diag)

addWorksheet(wb, "BFI_pair_spearman")
writeData(wb, "BFI_pair_spearman", pair_spearman)

addWorksheet(wb, "BFI_pair_by_year")
writeData(wb, "BFI_pair_by_year", pair_by_year)

addWorksheet(wb, "BFI_pair_polychoric")
writeData(wb, "BFI_pair_polychoric", pair_poly)

addWorksheet(wb, "BFI_SpearmanBrown")
writeData(wb, "BFI_SpearmanBrown", sb_tbl)

addWorksheet(wb, "BFI_CFA_fit")
if (is.data.frame(bfi_fitmeas)) writeData(wb, "BFI_CFA_fit", bfi_fitmeas)
if (is.data.frame(bfi_fitmeas_mlr)) writeData(wb, "BFI_CFA_fit", bfi_fitmeas_mlr, startRow = 5)

addWorksheet(wb, "BFI_CFA_loadings")
if (!is.null(bfi_loadings)) writeData(wb, "BFI_CFA_loadings", bfi_loadings)

addWorksheet(wb, "Affect_CFA_fit")
writeData(wb, "Affect_CFA_fit", aff_fit_mlr)
writeData(wb, "Affect_CFA_fit", aff_fit_wlsmv, startRow = 5)

addWorksheet(wb, "Affect_CFA_loadings_MLR")
writeData(wb, "Affect_CFA_loadings_MLR", aff_load_mlr)

addWorksheet(wb, "Affect_CFA_loadings_WLSMV")
writeData(wb, "Affect_CFA_loadings_WLSMV", aff_load_wlsmv)

addWorksheet(wb, "Affect_alpha")
alpha_tbl <- bind_rows(
  tibble(Scale="Positive(3)", t(alpha_pos$total)),
  tibble(Scale="Negative(7)", t(alpha_neg$total))
)
writeData(wb, "Affect_alpha", alpha_tbl)

saveWorkbook(wb, "Part2_Measurement_Results.xlsx", overwrite = TRUE)

cat("\nSaved:\n",
    "- keyed data:", out_keyed_rds, "\n",
    "- results workbook: Part2_Measurement_Results.xlsx\n")