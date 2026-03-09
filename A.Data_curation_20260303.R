# =========================================================
# Part 1. Data curation (2020–2023 KHS) [FIXED]
#  - Fix: id variable 'no' not found
#  - Create: unique id, gender, income_binary, generation_group, gen_mz
#  - Optional: BFI reverse-by-year diagnostic (kept simple)
# =========================================================

rm(list = ls())
set.seed(123)
options(stringsAsFactors = FALSE)

library(haven)
library(dplyr)
library(tidyr)
library(openxlsx)

# ---- File paths (as you requested)
file_2020 <- "C:/Users/leeeo/Desktop/Research/Data/한국인 행복조사 (국회미래연구원)/kor_data_20200086.dta"
file_2021 <- "C:/Users/leeeo/Desktop/Research/Data/한국인 행복조사 (국회미래연구원)/kor_data_20210003_V2.1.dta"
file_2022 <- "C:/Users/leeeo/Desktop/Research/Data/한국인 행복조사 (국회미래연구원)/kor_data_20220060.dta"
file_2023 <- "C:/Users/leeeo/Desktop/Research/Data/한국인 행복조사 (국회미래연구원)/kor_data_20230039.DTA"

dt_2020 <- read_dta(file_2020)
dt_2021 <- read_dta(file_2021)
dt_2022 <- read_dta(file_2022)
dt_2023 <- read_dta(file_2023)

# ---- Fix id_raw safely (avoid 'no' not found)
# If 'id' exists use it, else if 'no' exists use it, else create row index
if ("id" %in% names(dt_2020)) dt_2020$id_raw <- dt_2020$id else if ("no" %in% names(dt_2020)) dt_2020$id_raw <- dt_2020$no else dt_2020$id_raw <- seq_len(nrow(dt_2020))
if ("id" %in% names(dt_2021)) dt_2021$id_raw <- dt_2021$id else if ("no" %in% names(dt_2021)) dt_2021$id_raw <- dt_2021$no else dt_2021$id_raw <- seq_len(nrow(dt_2021))
if ("id" %in% names(dt_2022)) dt_2022$id_raw <- dt_2022$id else if ("no" %in% names(dt_2022)) dt_2022$id_raw <- dt_2022$no else dt_2022$id_raw <- seq_len(nrow(dt_2022))
if ("id" %in% names(dt_2023)) dt_2023$id_raw <- dt_2023$id else if ("no" %in% names(dt_2023)) dt_2023$id_raw <- dt_2023$no else dt_2023$id_raw <- seq_len(nrow(dt_2023))

# ---- Select + rename (network vars + demographics)
dt_2020_rev <- dt_2020 %>%
  transmute(
    id_raw = id_raw,
    year = 2020,
    gender = sq2,
    age = age,
    income_level = dq10_2,
    
    happiness = a1,
    
    enjoyment = b1_1, calm = b1_2, worry = b1_3, sadness = b1_4, depression = b1_5,
    anger = b1_6, stress = b1_7, tiredness = b1_8, vitality = b1_9, loneliness = b1_10,
    
    standard_of_living = c7_1, health_satisfaction = c7_2, relationships_satisfaction = c7_3,
    safety = c7_4, community_belonging = c7_5, future_security = c7_6,
    time_for_hobbies = c7_7, local_environment = c7_8,
    
    political_orientation = dq16,
    
    # BFI-10: items 1-5 then 6-10 (per your questionnaire)
    extraversion_1 = d11_1,  agreeableness_1 = d11_2,  conscientiousness_1 = d11_3, neuroticism_1 = d11_4, openness_1 = d11_5,
    extraversion_2 = d11_6,  agreeableness_2 = d11_7,  conscientiousness_2 = d11_8, neuroticism_2 = d11_9, openness_2 = d11_10
  )

dt_2021_rev <- dt_2021 %>%
  transmute(
    id_raw = id_raw,
    year = 2021,
    gender = sq1_3,
    age = age,
    income_level = dq12_2,
    
    happiness = a1_1,
    
    enjoyment = b1_1, calm = b1_2, worry = b1_3, sadness = b1_4, depression = b1_5,
    anger = b1_6, stress = b1_7, tiredness = b1_8, vitality = b1_9, loneliness = b1_10,
    
    standard_of_living = c7_1, health_satisfaction = c7_2, relationships_satisfaction = c7_3,
    safety = c7_4, community_belonging = c7_5, future_security = c7_6,
    time_for_hobbies = c7_7, local_environment = c7_8,
    
    political_orientation = dq19,
    
    extraversion_1 = d8_1,  agreeableness_1 = d8_2,  conscientiousness_1 = d8_3, neuroticism_1 = d8_4, openness_1 = d8_5,
    extraversion_2 = d8_6,  agreeableness_2 = d8_7,  conscientiousness_2 = d8_8, neuroticism_2 = d8_9, openness_2 = d8_10
  )

dt_2022_rev <- dt_2022 %>%
  transmute(
    id_raw = id_raw,
    year = 2022,
    gender = SQ1_3,
    age = SQ1_4,
    income_level = DQ12_2,
    
    happiness = A1,
    
    enjoyment = B1_1, calm = B1_2, worry = B1_3, sadness = B1_4, depression = B1_5,
    anger = B1_6, stress = B1_7, tiredness = B1_8, vitality = B1_9, loneliness = B1_10,
    
    standard_of_living = C7_1, health_satisfaction = C7_2, relationships_satisfaction = C7_3,
    safety = C7_4, community_belonging = C7_5, future_security = C7_6,
    time_for_hobbies = C7_7, local_environment = C7_8,
    
    political_orientation = DQ20_1,
    
    extraversion_1 = D8_1,  agreeableness_1 = D8_2,  conscientiousness_1 = D8_3, neuroticism_1 = D8_4, openness_1 = D8_5,
    extraversion_2 = D8_6,  agreeableness_2 = D8_7,  conscientiousness_2 = D8_8, neuroticism_2 = D8_9, openness_2 = D8_10
  )

dt_2023_rev <- dt_2023 %>%
  transmute(
    id_raw = id_raw,
    year = 2023,
    gender = SQ1_3,
    age = age,
    income_level = DQ12_2,
    
    happiness = A1,
    
    enjoyment = B1_1, calm = B1_2, worry = B1_3, sadness = B1_4, depression = B1_5,
    anger = B1_6, stress = B1_7, tiredness = B1_8, vitality = B1_9, loneliness = B1_10,
    
    standard_of_living = C5_1, health_satisfaction = C5_2, relationships_satisfaction = C5_3,
    safety = C5_4, community_belonging = C5_5, future_security = C5_6,
    time_for_hobbies = C5_7, local_environment = C5_8,
    
    political_orientation = DQ17,
    
    extraversion_1 = D8_1,  agreeableness_1 = D8_2,  conscientiousness_1 = D8_3, neuroticism_1 = D8_4, openness_1 = D8_5,
    extraversion_2 = D8_6,  agreeableness_2 = D8_7,  conscientiousness_2 = D8_8, neuroticism_2 = D8_9, openness_2 = D8_10
  )

combined_data <- bind_rows(dt_2020_rev, dt_2021_rev, dt_2022_rev, dt_2023_rev)

# ---- Make UNIQUE id across years
combined_data <- combined_data %>%
  mutate(id = paste0(year, "_", id_raw))

# ---- Safe numeric conversion
to_numeric_safe <- function(x) {
  if (is.factor(x)) return(as.numeric(as.character(x)))
  return(as.numeric(x))  # haven_labelled -> underlying numeric code
}

vars_0_10 <- c(
  "happiness",
  "enjoyment","calm","worry","sadness","depression","anger","stress","tiredness","vitality","loneliness",
  "standard_of_living","health_satisfaction","relationships_satisfaction","safety","community_belonging",
  "future_security","time_for_hobbies","local_environment"
)
vars_1_10 <- c("political_orientation")
vars_1_5  <- c(
  "extraversion_1","agreeableness_1","conscientiousness_1","neuroticism_1","openness_1",
  "extraversion_2","agreeableness_2","conscientiousness_2","neuroticism_2","openness_2"
)

combined_data <- combined_data %>%
  mutate(across(all_of(c("gender","age","income_level", vars_0_10, vars_1_10, vars_1_5)), to_numeric_safe))

# ---- Range validation
combined_data <- combined_data %>%
  mutate(
    across(all_of(vars_0_10), ~ ifelse(. < 0 | . > 10, NA, .)),
    political_orientation = ifelse(political_orientation < 1 | political_orientation > 10, NA, political_orientation),
    across(all_of(vars_1_5), ~ ifelse(. < 1 | . > 5, NA, .)),
    age = ifelse(age < 0 | age > 120, NA, age)
  )

# ---- BFI reverse coding (AUTO by year, based on raw pair sign)
# Reverse-worded items (per questionnaire):
# 1 introverted -> extraversion_1
# 7 critical -> agreeableness_2
# 3 careless -> conscientiousness_1
# 4 calm -> neuroticism_1
# 5 not creative -> openness_1
bfi_qc_by_year <- combined_data %>%
  select(year, all_of(vars_1_5)) %>%
  group_by(year) %>%
  summarise(
    ext_raw = cor(extraversion_1, extraversion_2, method="spearman", use="complete.obs"),
    agr_raw = cor(agreeableness_1, agreeableness_2, method="spearman", use="complete.obs"),
    con_raw = cor(conscientiousness_1, conscientiousness_2, method="spearman", use="complete.obs"),
    neu_raw = cor(neuroticism_1, neuroticism_2, method="spearman", use="complete.obs"),
    ope_raw = cor(openness_1, openness_2, method="spearman", use="complete.obs"),
    .groups = "drop"
  )

rev_flags <- bfi_qc_by_year %>%
  transmute(
    year,
    rev_extraversion_1      = ext_raw < 0,
    rev_agreeableness_2     = agr_raw < 0,
    rev_conscientiousness_1 = con_raw < 0,
    rev_neuroticism_1       = neu_raw < 0,
    rev_openness_1          = ope_raw < 0
  )

combined_data <- combined_data %>%
  left_join(rev_flags, by = "year") %>%
  mutate(
    extraversion_1      = ifelse(rev_extraversion_1,      6 - extraversion_1,      extraversion_1),
    agreeableness_2     = ifelse(rev_agreeableness_2,     6 - agreeableness_2,     agreeableness_2),
    conscientiousness_1 = ifelse(rev_conscientiousness_1, 6 - conscientiousness_1, conscientiousness_1),
    neuroticism_1       = ifelse(rev_neuroticism_1,       6 - neuroticism_1,       neuroticism_1),
    openness_1          = ifelse(rev_openness_1,          6 - openness_1,          openness_1)
  ) %>%
  select(-starts_with("rev_"))

# Save diagnostic (optional but recommended)
write.xlsx(
  list(
    BFI_raw_pair_cor_by_year = bfi_qc_by_year,
    BFI_reverse_flags_by_year = rev_flags
  ),
  "BFI_reverse_diagnostics.xlsx",
  rowNames = FALSE
)

# ---- Recode gender (final)
combined_data <- combined_data %>%
  mutate(
    gender = case_when(
      gender == 1 ~ "Male",
      gender == 2 ~ "Female",
      TRUE ~ NA_character_
    )
  )

# ---- Income binary (keep your original rule)
combined_data <- combined_data %>%
  mutate(
    income_binary = case_when(
      income_level %in% c(1,2,3,4,5) ~ "Low (≤4M KRW/month)",
      income_level %in% c(6,7,8,9,10,11,12) ~ "High (>4M KRW/month)",
      TRUE ~ NA_character_
    )
  )

# ---- Generation group + gen_mz  (THIS FIXES YOUR ERROR)
combined_data <- combined_data %>%
  mutate(
    generation_group = case_when(
      year == 2020 & age >= 57 & age <= 65 ~ "Baby Boomers",
      year == 2020 & age >= 41 & age <= 56 ~ "Generation X",
      year == 2020 & age >= 24 & age <= 40 ~ "Millennials",
      year == 2020 & age >= 15 & age <= 23 ~ "Generation Z",
      year == 2020 & age >= 66            ~ "Silent",
      
      year == 2021 & age >= 58 & age <= 66 ~ "Baby Boomers",
      year == 2021 & age >= 42 & age <= 57 ~ "Generation X",
      year == 2021 & age >= 25 & age <= 41 ~ "Millennials",
      year == 2021 & age >= 15 & age <= 24 ~ "Generation Z",
      year == 2021 & age >= 67             ~ "Silent",
      
      year == 2022 & age >= 59 & age <= 67 ~ "Baby Boomers",
      year == 2022 & age >= 43 & age <= 58 ~ "Generation X",
      year == 2022 & age >= 26 & age <= 42 ~ "Millennials",
      year == 2022 & age >= 15 & age <= 25 ~ "Generation Z",
      year == 2022 & age >= 68             ~ "Silent",
      
      year == 2023 & age >= 60 & age <= 68 ~ "Baby Boomers",
      year == 2023 & age >= 44 & age <= 59 ~ "Generation X",
      year == 2023 & age >= 27 & age <= 43 ~ "Millennials",
      year == 2023 & age >= 15 & age <= 26 ~ "Generation Z",
      year == 2023 & age >= 69             ~ "Silent",
      TRUE ~ NA_character_
    ),
    gen_mz = case_when(
      generation_group %in% c("Generation Z","Millennials") ~ "MZ",
      generation_group == "Generation X" ~ "GenX",
      generation_group == "Baby Boomers" ~ "Boomers",
      generation_group == "Silent" ~ "Silent",
      TRUE ~ NA_character_
    )
  )

# ---- HARD CHECK (prevents Part 3 error)
stopifnot(all(c("gender","income_binary","gen_mz") %in% names(combined_data)))

# ---- Save cleaned data (overwrite!)
saveRDS(combined_data, "combined_data_clean.rds")
write.csv(combined_data, "combined_data_clean.csv", row.names = FALSE)

# ---- Report missingness + analytic N
node_vars <- c(vars_0_10, "political_orientation", vars_1_5)
n_total <- nrow(combined_data)
n_complete <- combined_data %>%
  select(gender, income_binary, gen_mz, all_of(node_vars)) %>%
  drop_na() %>% nrow()

cat("Total N:", n_total, "\n")
cat("Complete-case analytic N:", n_complete, "\n")