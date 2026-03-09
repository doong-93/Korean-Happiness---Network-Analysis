# =========================================================
# Part 3. Network analysis (FULL INTEGRATED)
#  - Overall + subgroup network plots
#  - Centrality tables/figures
#  - CS stability
#  - Edge-weight 95% bootstrap CI for overall + subgroup networks
#  - NCT tables by sex / income / generation / (optional) year
#  - Optional reviewer-oriented extras:
#      * predictability
#      * bridge centrality
#      * gamma sensitivity
#      * year-stratified sensitivity
# =========================================================

rm(list = ls())
set.seed(123)
options(stringsAsFactors = FALSE)

library(dplyr)
library(tidyr)
library(qgraph)
library(bootnet)
library(NetworkComparisonTest)
library(networktools)
library(ggplot2)
library(openxlsx)
library(RColorBrewer)

# ---------------------------------------------------------
# 0) I/O and runtime options
# ---------------------------------------------------------
out_dir <- "Outputs_Part3_0307"
dir.create(out_dir, showWarnings = FALSE)

# Quick test vs final run
# quick test: N_BOOT_EDGE = 200, N_BOOT_CASE = 200, N_NCT_ITER = 100
# final run:  N_BOOT_EDGE = 1000~2000, N_BOOT_CASE = 1000~2000, N_NCT_ITER = 1000~2000
N_BOOT_EDGE <- 1000
N_BOOT_CASE <- 1000
N_NCT_ITER  <- 1000

# optional separate count for year sensitivity edge CI
N_BOOT_EDGE_YEAR <- 500

RUN_BOOT_EDGE <- TRUE
RUN_BOOT_CASE <- TRUE
RUN_NCT       <- TRUE
RUN_EXTRAS    <- TRUE
RUN_YEAR_SENS <- TRUE

# edge CI tables for year sensitivity networks
RUN_BOOT_EDGE_YEAR <- FALSE

# save full CI forest plots for every subgroup network?
# FALSE recommended because it creates many large files
SAVE_EDGE_CI_FIGURES_ALL <- FALSE

# ---------------------------------------------------------
# 1) Load data
# ---------------------------------------------------------
if (file.exists("combined_data_clean_keyed.rds")) {
  combined_data <- readRDS("combined_data_clean_keyed.rds")
} else {
  combined_data <- readRDS("combined_data_clean.rds")
}

# ---------------------------------------------------------
# 2) Defensive demographic checks / creation
# ---------------------------------------------------------
if (!("gender" %in% names(combined_data))) {
  stop("Missing demographic column: gender. Check Part 1.")
}

if (!("income_binary" %in% names(combined_data))) {
  if ("income_level" %in% names(combined_data)) {
    combined_data <- combined_data %>%
      mutate(
        income_binary = case_when(
          income_level %in% c(1,2,3,4,5) ~ "Low",
          income_level %in% c(6,7,8,9,10,11,12) ~ "High",
          TRUE ~ NA_character_
        )
      )
  } else {
    stop("Missing demographic column: income_binary (and income_level not found). Check Part 1.")
  }
}

if (!("gen_mz" %in% names(combined_data))) {
  if ("generation_group" %in% names(combined_data)) {
    combined_data <- combined_data %>%
      mutate(
        gen_mz = case_when(
          generation_group %in% c("Generation Z","Millennials") ~ "MZ",
          generation_group == "Generation X" ~ "GenX",
          generation_group == "Baby Boomers" ~ "Boomers",
          generation_group == "Silent" ~ "Silent",
          TRUE ~ NA_character_
        )
      )
  } else if (all(c("year","age") %in% names(combined_data))) {
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
  } else {
    stop("Missing demographic column: gen_mz. Check Part 1.")
  }
}

# ---------------------------------------------------------
# 3) Variables for network
# ---------------------------------------------------------
node_vars <- c(
  "happiness",
  "enjoyment","calm","worry","sadness","depression","anger","stress","tiredness","vitality","loneliness",
  "standard_of_living","health_satisfaction","relationships_satisfaction","safety","community_belonging",
  "future_security","time_for_hobbies","local_environment",
  "political_orientation",
  "extraversion_1","agreeableness_1","conscientiousness_1","neuroticism_1","openness_1",
  "extraversion_2","agreeableness_2","conscientiousness_2","neuroticism_2","openness_2"
)

node_names <- c("H1", paste0("A", 1:10), paste0("W", 1:8), "P1", paste0("B", 1:10))
stopifnot(length(node_vars) == length(node_names))

node_group <- c(
  "Happiness",
  rep("Affect", 10),
  rep("Satisfaction", 8),
  "Politics",
  rep("Personality", 10)
)

group_colors <- c(
  Happiness    = "#FFFFB3",
  Affect       = "#8DD3C7",
  Satisfaction = "#80B1D3",
  Politics     = "#FB8072",
  Personality  = "#BEBADA"
)

# ---------------------------------------------------------
# 4) Complete-case analysis dataset
# ---------------------------------------------------------
df_all <- combined_data %>%
  select(all_of(c("year","gender","income_binary","gen_mz", node_vars))) %>%
  drop_na()

cat("Analytic N (complete-case):", nrow(df_all), "\n")

dat_all <- df_all %>% select(all_of(node_vars))
colnames(dat_all) <- node_names

# subgroup datasets
dat_male   <- df_all %>% filter(gender == "Male")   %>% select(all_of(node_vars)); colnames(dat_male)   <- node_names
dat_female <- df_all %>% filter(gender == "Female") %>% select(all_of(node_vars)); colnames(dat_female) <- node_names

dat_lowinc  <- df_all %>% filter(grepl("^Low",  income_binary, ignore.case = TRUE)) %>% select(all_of(node_vars)); colnames(dat_lowinc)  <- node_names
dat_highinc <- df_all %>% filter(grepl("^High", income_binary, ignore.case = TRUE)) %>% select(all_of(node_vars)); colnames(dat_highinc) <- node_names

dat_mz   <- df_all %>% filter(gen_mz == "MZ")      %>% select(all_of(node_vars)); colnames(dat_mz)   <- node_names
dat_genx <- df_all %>% filter(gen_mz == "GenX")    %>% select(all_of(node_vars)); colnames(dat_genx) <- node_names
dat_boom <- df_all %>% filter(gen_mz == "Boomers") %>% select(all_of(node_vars)); colnames(dat_boom) <- node_names

# year-stratified
dat_2020 <- df_all %>% filter(year == 2020) %>% select(all_of(node_vars)); colnames(dat_2020) <- node_names
dat_2021 <- df_all %>% filter(year == 2021) %>% select(all_of(node_vars)); colnames(dat_2021) <- node_names
dat_2022 <- df_all %>% filter(year == 2022) %>% select(all_of(node_vars)); colnames(dat_2022) <- node_names
dat_2023 <- df_all %>% filter(year == 2023) %>% select(all_of(node_vars)); colnames(dat_2023) <- node_names

# ---------------------------------------------------------
# 5) Helper functions
# ---------------------------------------------------------

canonicalize_edge_order <- function(df, nodes, col1 = "Node1", col2 = "Node2") {
  idx1 <- match(df[[col1]], nodes)
  idx2 <- match(df[[col2]], nodes)
  
  node_min <- ifelse(idx1 <= idx2, df[[col1]], df[[col2]])
  node_max <- ifelse(idx1 <= idx2, df[[col2]], df[[col1]])
  
  df[[col1]] <- node_min
  df[[col2]] <- node_max
  df
}

estimate_network_and_adj <- function(dat, gamma = 0.5) {
  net <- bootnet::estimateNetwork(
    dat,
    default = "EBICglasso",
    corMethod = "spearman",
    tuning = gamma,
    threshold = TRUE
  )
  
  adj <- net$graph
  colnames(adj) <- rownames(adj) <- colnames(dat)
  
  list(net = net, adj = adj)
}

estimate_adj <- function(dat, gamma = 0.5) {
  out <- estimate_network_and_adj(dat, gamma = gamma)
  out$adj
}

save_edge_table <- function(adj, file) {
  idx <- which(upper.tri(adj), arr.ind = TRUE)
  
  out <- data.frame(
    Node1 = rownames(adj)[idx[,1]],
    Node2 = colnames(adj)[idx[,2]],
    Strength = adj[idx],
    stringsAsFactors = FALSE
  ) %>%
    mutate(
      Edge = paste(Node1, Node2, sep = " - ")
    ) %>%
    select(Edge, Node1, Node2, Strength) %>%
    arrange(desc(abs(Strength)))
  
  write.xlsx(out, file, rowNames = FALSE)
}

plot_network_png <- function(adj, file, title_text, layout_fixed, node_names, node_group, group_colors) {
  edge_col <- ifelse(
    adj > 0,
    adjustcolor("#008FD5", alpha.f = 0.85),
    adjustcolor("#FF2700", alpha.f = 0.85)
  )
  edge_col[adj == 0] <- NA
  
  png(file, width = 5000, height = 5000, res = 900)
  qgraph(
    adj,
    layout = layout_fixed,
    labels = node_names,
    color = group_colors[node_group],
    groups = split(node_names, node_group),
    edge.color = edge_col,
    vsize = 7,
    esize = 25,
    minimum = 0.01,
    repulsion = 0.85,
    label.cex = 1.4,
    legend = FALSE,
    title = title_text,
    title.cex = 0.9
  )
  dev.off()
}

pick_boot_summary_col <- function(df, patterns, required = TRUE) {
  nm_orig <- names(df)
  nm_synt <- tolower(make.names(nm_orig, unique = TRUE))
  
  idx <- unique(unlist(lapply(patterns, function(p) grep(p, nm_synt, perl = TRUE))))
  idx <- idx[!is.na(idx)]
  
  if (length(idx) == 0) {
    if (required) {
      stop("Could not find required bootstrap summary column. Available columns: ",
           paste(nm_orig, collapse = ", "))
    } else {
      return(NULL)
    }
  }
  
  nm_orig[idx[1]]
}

make_boot_edge_ci_table <- function(boot_obj, node_names, only_nonzero_sample = FALSE) {
  tol <- sqrt(.Machine$double.eps)
  
  smry <- summary(boot_obj, statistics = "edge")
  smry_df <- as.data.frame(smry, stringsAsFactors = FALSE)
  
  col_node1 <- pick_boot_summary_col(smry_df, c("^node1$", "^var1$"))
  col_node2 <- pick_boot_summary_col(smry_df, c("^node2$", "^var2$"))
  col_type  <- pick_boot_summary_col(smry_df, c("^type$"), required = FALSE)
  col_sample <- pick_boot_summary_col(smry_df, c("^sample$", "^original$"))
  col_mean   <- pick_boot_summary_col(smry_df, c("^mean$", "^bootmean$"))
  col_sd     <- pick_boot_summary_col(smry_df, c("^sd$", "^bootsd$"))
  col_low    <- pick_boot_summary_col(
    smry_df,
    c("^q2\\.5$", "^x2\\.5\\.$", "^lower$", "^lower95$", "^ci_low$", "^cilow$", "^lowerci$")
  )
  col_high   <- pick_boot_summary_col(
    smry_df,
    c("^q97\\.5$", "^x97\\.5\\.$", "^upper$", "^upper95$", "^ci_high$", "^cihigh$", "^upperci$")
  )
  col_prop0  <- pick_boot_summary_col(
    smry_df,
    c("^prop0$", "^propzero$", "^p0$"),
    required = FALSE
  )
  
  out <- smry_df %>%
    {
      if (!is.null(col_type)) dplyr::filter(., .data[[col_type]] == "edge") else .
    } %>%
    transmute(
      Node1 = as.character(.data[[col_node1]]),
      Node2 = as.character(.data[[col_node2]]),
      sample = as.numeric(.data[[col_sample]]),
      mean_boot = as.numeric(.data[[col_mean]]),
      sd_boot = as.numeric(.data[[col_sd]]),
      ci_low = as.numeric(.data[[col_low]]),
      ci_high = as.numeric(.data[[col_high]]),
      prop_zero = if (!is.null(col_prop0)) as.numeric(.data[[col_prop0]]) else NA_real_
    ) %>%
    canonicalize_edge_order(nodes = node_names, col1 = "Node1", col2 = "Node2") %>%
    mutate(
      Edge = paste(Node1, Node2, sep = " - "),
      present_in_sample = abs(sample) > tol,
      sign = case_when(
        sample > tol  ~ "positive",
        sample < -tol ~ "negative",
        TRUE ~ "zero"
      )
    ) %>%
    select(
      Edge, Node1, Node2,
      sample, mean_boot, sd_boot,
      ci_low, ci_high,
      prop_zero, present_in_sample, sign
    ) %>%
    arrange(desc(abs(sample)), Node1, Node2)
  
  if (only_nonzero_sample) {
    out <- out %>% filter(present_in_sample)
  }
  
  out
}

save_edge_ci_figure <- function(edge_tbl, out_file, title_text,
                                only_nonzero = TRUE,
                                bg = "white") {
  df <- edge_tbl
  
  if (only_nonzero) {
    df <- df %>% filter(present_in_sample)
  }
  
  if (nrow(df) == 0) return(invisible(NULL))
  
  df <- df %>%
    arrange(sample) %>%
    mutate(Edge = factor(Edge, levels = Edge))
  
  p <- ggplot(df, aes(x = sample, y = Edge, color = sign)) +
    geom_errorbarh(
      aes(xmin = ci_low, xmax = ci_high),
      height = 0.15,
      linewidth = 0.45,
      color = "grey70"
    ) +
    geom_point(size = 1.8) +
    geom_vline(xintercept = 0, linetype = "dashed", linewidth = 0.35, color = "grey60") +
    scale_color_manual(values = c(
      positive = "#008FD5",
      negative = "#FF2700",
      zero = "grey40"
    )) +
    labs(
      title = title_text,
      x = "Edge weight",
      y = NULL
    ) +
    theme_bw(base_size = 11) +
    theme(
      panel.grid.major.y = element_blank(),
      panel.grid.minor = element_blank(),
      legend.position = "none",
      plot.title = element_text(face = "bold")
    )
  
  h <- max(8, min(28, 0.16 * nrow(df)))
  
  ggsave(
    filename = out_file,
    plot = p,
    width = 9,
    height = h,
    dpi = 600,
    bg = bg
  )
}

run_edge_boot_ci_export <- function(net_obj, label, out_dir, nboots,
                                    node_names,
                                    save_ci_plot = FALSE) {
  
  boot_edge <- bootnet::bootnet(
    net_obj,
    nBoots = nboots,
    type = "nonparametric",
    statistics = "edge",
    nCores = 1
  )
  
  edge_ci_all <- make_boot_edge_ci_table(
    boot_obj = boot_edge,
    node_names = node_names,
    only_nonzero_sample = FALSE
  )
  
  edge_ci_nonzero <- edge_ci_all %>%
    filter(present_in_sample)
  
  if (label == "Overall") {
    table_file <- file.path(out_dir, "Supplementary_Table_Edges_with_CI.xlsx")
    fig_png    <- file.path(out_dir, "Supplementary_EdgeWeight_CI.png")
    fig_pdf    <- file.path(out_dir, "Supplementary_EdgeWeight_CI.pdf")
  } else {
    table_file <- file.path(out_dir, paste0("Supplementary_Table_Edges_with_95CI_", label, ".xlsx"))
    fig_png    <- file.path(out_dir, paste0("Supplementary_EdgeWeight_CI_", label, ".png"))
    fig_pdf    <- file.path(out_dir, paste0("Supplementary_EdgeWeight_CI_", label, ".pdf"))
  }
  
  write.xlsx(
    list(
      All_edges = edge_ci_all,
      Nonzero_edges = edge_ci_nonzero
    ),
    file = table_file,
    rowNames = FALSE
  )
  
  if (save_ci_plot) {
    save_edge_ci_figure(
      edge_ci_all,
      fig_png,
      title_text = paste0("Bootstrapped 95% CIs for Edge Weights: ", label),
      only_nonzero = TRUE,
      bg = "white"
    )
    
    save_edge_ci_figure(
      edge_ci_all,
      fig_pdf,
      title_text = paste0("Bootstrapped 95% CIs for Edge Weights: ", label),
      only_nonzero = TRUE,
      bg = "white"
    )
  }
  
  invisible(edge_ci_all)
}

# ---- NCT helpers
extract_nct_edge_pvals <- function(nct_obj, nodes) {
  
  p_obj <- nct_obj$einv.pvals
  
  if (is.data.frame(p_obj)) {
    nm <- names(p_obj)
    nm_clean <- make.names(nm)
    names(p_obj) <- nm_clean
    
    if (all(c("Var1","Var2") %in% names(p_obj))) {
      p_col <- names(p_obj)[grepl("^p(\\.value)?$", names(p_obj), ignore.case = TRUE)]
      if (length(p_col) == 0) p_col <- names(p_obj)[grepl("p.*value", names(p_obj), ignore.case = TRUE)]
      
      e_col <- names(p_obj)[grepl("Test.*statistic.*E", names(p_obj), ignore.case = TRUE)]
      if (length(e_col) == 0) e_col <- names(p_obj)[grepl("statistic.*E", names(p_obj), ignore.case = TRUE)]
      
      out <- data.frame(
        Node1 = as.character(p_obj$Var1),
        Node2 = as.character(p_obj$Var2),
        p_value = if (length(p_col) >= 1) as.numeric(p_obj[[p_col[1]]]) else NA_real_,
        Test_stat_E = if (length(e_col) >= 1) as.numeric(p_obj[[e_col[1]]]) else NA_real_,
        stringsAsFactors = FALSE
      )
      return(out)
    }
  }
  
  if (is.matrix(p_obj)) {
    if (nrow(p_obj) == length(nodes) && ncol(p_obj) == length(nodes)) {
      rownames(p_obj) <- colnames(p_obj) <- nodes
      idx <- which(upper.tri(p_obj), arr.ind = TRUE)
      out <- data.frame(
        Node1 = nodes[idx[,1]],
        Node2 = nodes[idx[,2]],
        p_value = p_obj[idx],
        Test_stat_E = NA_real_,
        stringsAsFactors = FALSE
      )
      return(out)
    }
  }
  
  if (is.numeric(p_obj) && !is.null(names(p_obj))) {
    out <- data.frame(
      raw_name = names(p_obj),
      p_value = as.numeric(p_obj),
      stringsAsFactors = FALSE
    )
    tmp <- strsplit(out$raw_name, "[[:space:]-]+")
    out$Node1 <- sapply(tmp, `[`, 1)
    out$Node2 <- sapply(tmp, `[`, 2)
    out$Test_stat_E <- NA_real_
    out <- out %>% select(Node1, Node2, p_value, Test_stat_E)
    return(out)
  }
  
  stop("Unable to parse nct_obj$einv.pvals. Inspect str(nct_obj$einv.pvals).")
}

make_edge_diff_table_from_nct <- function(adj_g1, adj_g2, nct_obj, label_g1, label_g2) {
  nodes <- rownames(adj_g1)
  
  idx <- which(upper.tri(adj_g1), arr.ind = TRUE)
  edge_weights <- data.frame(
    Node1 = nodes[idx[,1]],
    Node2 = nodes[idx[,2]],
    Strength_g1 = adj_g1[idx],
    Strength_g2 = adj_g2[idx],
    stringsAsFactors = FALSE
  )
  names(edge_weights)[3:4] <- c(paste0(label_g1, "_Strength"), paste0(label_g2, "_Strength"))
  
  p_df <- extract_nct_edge_pvals(nct_obj, nodes)
  
  edge_weights <- canonicalize_edge_order(edge_weights, nodes, "Node1", "Node2")
  p_df         <- canonicalize_edge_order(p_df, nodes, "Node1", "Node2")
  
  out <- edge_weights %>%
    left_join(p_df, by = c("Node1","Node2")) %>%
    mutate(
      Difference = .data[[paste0(label_g1, "_Strength")]] - .data[[paste0(label_g2, "_Strength")]],
      Test_stat_E = ifelse(is.na(Test_stat_E), abs(Difference), Test_stat_E),
      p_fdr_BH = p.adjust(p_value, method = "BH"),
      Edge = paste(Node1, Node2, sep = " - ")
    ) %>%
    select(
      Edge, Node1, Node2,
      all_of(paste0(label_g1, "_Strength")),
      all_of(paste0(label_g2, "_Strength")),
      Difference, p_value, p_fdr_BH, Test_stat_E
    ) %>%
    arrange(p_value)
  
  out
}

est_for_nct <- function(data, gamma = 0.5, ...) {
  estimate_adj(data, gamma = gamma)
}

run_nct_and_export <- function(dat_g1, dat_g2, adj_g1, adj_g2,
                               label_g1, label_g2, file_stub,
                               it = 1000, gamma = 0.5, out_dir = ".") {
  
  stopifnot(identical(colnames(dat_g1), colnames(dat_g2)))
  stopifnot(identical(colnames(dat_g1), rownames(adj_g1)))
  stopifnot(identical(colnames(dat_g2), rownames(adj_g2)))
  
  nct <- NetworkComparisonTest::NCT(
    dat_g1, dat_g2,
    it = it,
    binary.data = FALSE,
    paired = FALSE,
    weighted = TRUE,
    test.edges = TRUE,
    edges = "all",
    test.centrality = FALSE,
    p.adjust.methods = "none",
    estimator = est_for_nct,
    estimatorArgs = list(gamma = gamma),
    progressbar = TRUE,
    verbose = TRUE
  )
  
  edge_tbl <- make_edge_diff_table_from_nct(adj_g1, adj_g2, nct, label_g1, label_g2)
  
  summary_tbl <- data.frame(
    comparison = paste(label_g1, "vs", label_g2),
    glstrinv_p = nct$glstrinv.pval,
    glstrinv_diff = nct$glstrinv.real,
    nwinv_p = nct$nwinv.pval
  )
  
  sig_raw <- edge_tbl %>% dplyr::filter(!is.na(p_value) & p_value < 0.05)
  sig_fdr <- edge_tbl %>% dplyr::filter(!is.na(p_fdr_BH) & p_fdr_BH < 0.05)
  
  openxlsx::write.xlsx(
    list(
      "NCT_summary" = summary_tbl,
      "All_edges" = edge_tbl,
      "Sig_raw_p<0.05" = sig_raw,
      "Sig_FDR_BH<0.05" = sig_fdr
    ),
    file = file.path(out_dir, paste0(file_stub, ".xlsx")),
    rowNames = FALSE
  )
  
  invisible(list(nct = nct, edges = edge_tbl, summary = summary_tbl))
}

extract_bridge_metric <- function(bridge_obj, candidates_norm) {
  nm <- names(bridge_obj)
  nm_norm <- gsub("[^a-z0-9]", "", tolower(nm))
  idx <- which(nm_norm %in% candidates_norm)
  if (length(idx) == 0) return(NULL)
  as.numeric(bridge_obj[[idx[1]]])
}

# ---------------------------------------------------------
# 6) Overall network + layout
# ---------------------------------------------------------
gamma_main <- 0.5

est_all <- estimate_network_and_adj(dat_all, gamma_main)
net_all <- est_all$net
adj_all <- est_all$adj

layout_all <- qgraph(adj_all, layout = "spring", DoNotPlot = TRUE)$layout
idx_h1 <- which(node_names == "H1")
layout_all <- layout_all - matrix(layout_all[idx_h1, ], nrow = nrow(layout_all), ncol = 2, byrow = TRUE)

save_edge_table(adj_all, file.path(out_dir, "Table_Edges_Overall.xlsx"))

plot_network_png(
  adj_all,
  file.path(out_dir, "Figure_Network_Overall.png"),
  "Overall Network",
  layout_all, node_names, node_group, group_colors
)

# ---------------------------------------------------------
# 7) Subgroup networks (Sex / Income / Generation / Year sensitivity)
# ---------------------------------------------------------

# --- Sex
est_male   <- estimate_network_and_adj(dat_male, gamma_main)
est_female <- estimate_network_and_adj(dat_female, gamma_main)

net_male   <- est_male$net
net_female <- est_female$net
adj_male   <- est_male$adj
adj_female <- est_female$adj

save_edge_table(adj_male,   file.path(out_dir, "Table_Edges_Male.xlsx"))
save_edge_table(adj_female, file.path(out_dir, "Table_Edges_Female.xlsx"))

plot_network_png(adj_male,   file.path(out_dir, "Figure_Network_Male.png"),   "Male Network",   layout_all, node_names, node_group, group_colors)
plot_network_png(adj_female, file.path(out_dir, "Figure_Network_Female.png"), "Female Network", layout_all, node_names, node_group, group_colors)

# --- Income
est_lowinc  <- estimate_network_and_adj(dat_lowinc, gamma_main)
est_highinc <- estimate_network_and_adj(dat_highinc, gamma_main)

net_lowinc  <- est_lowinc$net
net_highinc <- est_highinc$net
adj_lowinc  <- est_lowinc$adj
adj_highinc <- est_highinc$adj

save_edge_table(adj_lowinc,  file.path(out_dir, "Table_Edges_LowIncome.xlsx"))
save_edge_table(adj_highinc, file.path(out_dir, "Table_Edges_HighIncome.xlsx"))

plot_network_png(adj_lowinc,  file.path(out_dir, "Figure_Network_LowIncome.png"),  "Low Income Network",  layout_all, node_names, node_group, group_colors)
plot_network_png(adj_highinc, file.path(out_dir, "Figure_Network_HighIncome.png"), "High Income Network", layout_all, node_names, node_group, group_colors)

# --- Generation
est_mz   <- estimate_network_and_adj(dat_mz, gamma_main)
est_genx <- estimate_network_and_adj(dat_genx, gamma_main)
est_boom <- estimate_network_and_adj(dat_boom, gamma_main)

net_mz   <- est_mz$net
net_genx <- est_genx$net
net_boom <- est_boom$net
adj_mz   <- est_mz$adj
adj_genx <- est_genx$adj
adj_boom <- est_boom$adj

save_edge_table(adj_mz,   file.path(out_dir, "Table_Edges_MZ.xlsx"))
save_edge_table(adj_genx, file.path(out_dir, "Table_Edges_GenX.xlsx"))
save_edge_table(adj_boom, file.path(out_dir, "Table_Edges_Boomers.xlsx"))

plot_network_png(adj_mz,   file.path(out_dir, "Figure_Network_MZ.png"),      "MZ Network",      layout_all, node_names, node_group, group_colors)
plot_network_png(adj_genx, file.path(out_dir, "Figure_Network_GenX.png"),    "GenX Network",    layout_all, node_names, node_group, group_colors)
plot_network_png(adj_boom, file.path(out_dir, "Figure_Network_Boomers.png"), "Boomers Network", layout_all, node_names, node_group, group_colors)

# --- Year sensitivity
if (RUN_YEAR_SENS) {
  est_2020 <- estimate_network_and_adj(dat_2020, gamma_main)
  est_2021 <- estimate_network_and_adj(dat_2021, gamma_main)
  est_2022 <- estimate_network_and_adj(dat_2022, gamma_main)
  est_2023 <- estimate_network_and_adj(dat_2023, gamma_main)
  
  net_2020 <- est_2020$net
  net_2021 <- est_2021$net
  net_2022 <- est_2022$net
  net_2023 <- est_2023$net
  adj_2020 <- est_2020$adj
  adj_2021 <- est_2021$adj
  adj_2022 <- est_2022$adj
  adj_2023 <- est_2023$adj
  
  save_edge_table(adj_2020, file.path(out_dir, "Table_Edges_2020.xlsx"))
  save_edge_table(adj_2021, file.path(out_dir, "Table_Edges_2021.xlsx"))
  save_edge_table(adj_2022, file.path(out_dir, "Table_Edges_2022.xlsx"))
  save_edge_table(adj_2023, file.path(out_dir, "Table_Edges_2023.xlsx"))
  
  plot_network_png(adj_2020, file.path(out_dir, "Figure_Network_2020.png"), "2020 Network", layout_all, node_names, node_group, group_colors)
  plot_network_png(adj_2021, file.path(out_dir, "Figure_Network_2021.png"), "2021 Network", layout_all, node_names, node_group, group_colors)
  plot_network_png(adj_2022, file.path(out_dir, "Figure_Network_2022.png"), "2022 Network", layout_all, node_names, node_group, group_colors)
  plot_network_png(adj_2023, file.path(out_dir, "Figure_Network_2023.png"), "2023 Network", layout_all, node_names, node_group, group_colors)
}

# =========================================================
# 8) Centrality Measures (raw + standardized)
#    - node order fixed by node_names
#    - standardized figure uses geom_path()
# =========================================================
centrality_keep <- c("Strength", "Closeness", "Betweenness", "ExpectedInfluence")
node_order <- node_names

# ---------- Raw centrality ----------
cent_long_raw <- qgraph::centralityTable(
  adj_all,
  labels = node_names,
  standardized = FALSE,
  relative = FALSE,
  weighted = TRUE,
  signed = TRUE
) %>%
  filter(measure %in% centrality_keep) %>%
  mutate(
    measure = factor(measure, levels = centrality_keep),
    node = factor(node, levels = node_order),
    node_id = match(as.character(node), node_order)
  ) %>%
  arrange(measure, node_id)

# ---------- Standardized centrality ----------
cent_long_std <- qgraph::centralityTable(
  adj_all,
  labels = node_names,
  standardized = TRUE,
  relative = FALSE,
  weighted = TRUE,
  signed = TRUE
) %>%
  filter(measure %in% centrality_keep) %>%
  mutate(
    measure = factor(measure, levels = centrality_keep),
    node = factor(node, levels = node_order),
    node_id = match(as.character(node), node_order)
  ) %>%
  arrange(measure, node_id)

# ---------- Save wide tables ----------
cent_wide_raw <- cent_long_raw %>%
  transmute(node = as.character(node), measure, value) %>%
  pivot_wider(names_from = measure, values_from = value)

cent_wide_std <- cent_long_std %>%
  transmute(node = as.character(node), measure, value) %>%
  pivot_wider(names_from = measure, values_from = value)

write.xlsx(
  cent_wide_raw,
  file.path(out_dir, "Supplementary_Table_Centrality_Raw.xlsx"),
  rowNames = FALSE
)

write.xlsx(
  cent_wide_std,
  file.path(out_dir, "Supplementary_Table_Centrality_Standardized.xlsx"),
  rowNames = FALSE
)

# ---------- Raw figure ----------
p_cent_raw <- ggplot(cent_long_raw, aes(x = value, y = node)) +
  geom_col(width = 0.72, fill = "grey35") +
  geom_vline(
    xintercept = 0,
    linetype = "dashed",
    linewidth = 0.35,
    color = "grey60"
  ) +
  facet_wrap(~ measure, scales = "free_x", ncol = 2) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid.major.y = element_blank(),
    panel.grid.minor = element_blank(),
    strip.text = element_text(face = "bold")
  ) +
  labs(
    x = NULL,
    y = NULL,
    title = "Centrality Measures of Network Variables (Raw)"
  )

ggsave(
  file.path(out_dir, "Figure_Centrality_Raw.png"),
  plot = p_cent_raw,
  width = 12,
  height = 8,
  dpi = 600,
  bg = "white"
)

# ---------- Standardized figure ----------
p_cent_std <- ggplot(cent_long_std, aes(x = value, y = node_id)) +
  geom_path(
    aes(group = measure),
    linewidth = 0.65,
    color = "black",
    lineend = "round",
    linejoin = "round"
  ) +
  geom_point(size = 2.2, color = "black") +
  geom_vline(
    xintercept = 0,
    linetype = "dashed",
    linewidth = 0.35,
    color = "grey60"
  ) +
  facet_grid(. ~ measure, scales = "free_x") +
  scale_y_continuous(
    breaks = seq_along(node_order),
    labels = node_order,
    expand = expansion(mult = c(0.02, 0.02))
  ) +
  theme_bw(base_size = 12) +
  theme(
    panel.grid.major.y = element_line(color = "grey85", linewidth = 0.3),
    panel.grid.major.x = element_line(color = "grey85", linewidth = 0.3),
    panel.grid.minor = element_blank(),
    strip.text = element_text(face = "bold"),
    axis.title = element_blank()
  ) +
  labs(
    title = "Centrality Measures of Network Variables (Standardized)"
  )

ggsave(
  file.path(out_dir, "Figure_Centrality_Standardized.png"),
  plot = p_cent_std,
  width = 14,
  height = 8,
  dpi = 600,
  bg = "white"
)

ggsave(
  file.path(out_dir, "Figure_Centrality_Standardized.pdf"),
  plot = p_cent_std,
  width = 14,
  height = 8,
  bg = "white"
)

# =========================================================
# 9) CS stability (case-dropping bootstrap)
# =========================================================
if (RUN_BOOT_CASE) {
  
  cs_stats <- c("strength", "closeness", "betweenness", "expectedInfluence")
  
  boot_case <- bootnet::bootnet(
    net_all,
    nBoots = N_BOOT_CASE,
    type = "case",
    statistics = cs_stats,
    nCores = 1
  )
  
  cs_tbl <- data.frame(
    statistic = cs_stats,
    CS_coefficient = sapply(
      cs_stats,
      function(s) {
        tryCatch(
          bootnet::corStability(boot_case, statistics = s),
          error = function(e) NA_real_
        )
      }
    )
  )
  
  write.xlsx(
    cs_tbl,
    file.path(out_dir, "Table_CSstability.xlsx"),
    rowNames = FALSE
  )
  
  p_cs <- plot(
    boot_case,
    statistics = cs_stats,
    plot = "area",
    labels = TRUE,
    legend = TRUE
  ) +
    labs(
      title = "Correlation Stability of Centrality Indices",
      x = "Proportion of cases dropped",
      y = "Correlation with original sample"
    ) +
    theme_bw(base_size = 12) +
    theme(
      strip.text = element_text(face = "bold"),
      plot.title = element_text(face = "bold"),
      legend.position = "bottom"
    )
  
  ggsave(
    file.path(out_dir, "Supplementary_Figure_CSstability.png"),
    plot = p_cs,
    width = 11,
    height = 8,
    dpi = 600,
    bg = "white"
  )
  
  ggsave(
    file.path(out_dir, "Supplementary_Figure_CSstability.pdf"),
    plot = p_cs,
    width = 11,
    height = 8,
    bg = "white"
  )
}

# =========================================================
# 10) Edge-weight bootstrap CI + Top/Bottom edge plots
#    - 95% CI for overall + subgroup networks
#    - overall keeps CI figure + top/bottom figures
# =========================================================
if (RUN_BOOT_EDGE) {
  
  edge_ci_results <- list()
  
  # overall + subgroup targets
  boot_targets <- list(
    Overall    = net_all,
    Male       = net_male,
    Female     = net_female,
    LowIncome  = net_lowinc,
    HighIncome = net_highinc,
    MZ         = net_mz,
    GenX       = net_genx,
    Boomers    = net_boom
  )
  
  if (RUN_YEAR_SENS && RUN_BOOT_EDGE_YEAR) {
    boot_targets$Year2020 <- net_2020
    boot_targets$Year2021 <- net_2021
    boot_targets$Year2022 <- net_2022
    boot_targets$Year2023 <- net_2023
  }
  
  for (nm in names(boot_targets)) {
    message("Running edge bootstrap CI: ", nm)
    
    nboots_use <- if (grepl("^Year", nm)) N_BOOT_EDGE_YEAR else N_BOOT_EDGE
    save_plot_use <- if (nm == "Overall") TRUE else SAVE_EDGE_CI_FIGURES_ALL
    
    edge_ci_results[[nm]] <- run_edge_boot_ci_export(
      net_obj = boot_targets[[nm]],
      label = nm,
      out_dir = out_dir,
      nboots = nboots_use,
      node_names = node_names,
      save_ci_plot = save_plot_use
    )
  }
  
  # -------- Overall top/bottom figures --------
  edge_ci2 <- edge_ci_results[["Overall"]]
  
  top10 <- edge_ci2 %>%
    filter(present_in_sample, sample > 0) %>%
    arrange(desc(sample)) %>%
    slice(1:10)
  
  bot10 <- edge_ci2 %>%
    filter(present_in_sample, sample < 0) %>%
    arrange(sample) %>%
    slice(1:10) %>%
    arrange(desc(sample))
  
  make_edge_plot <- function(df, point_col, out_file) {
    df <- df %>%
      mutate(Edge = factor(Edge, levels = rev(Edge)))
    
    p <- ggplot(df, aes(x = sample, y = Edge)) +
      geom_errorbarh(
        aes(xmin = ci_low, xmax = ci_high),
        height = 0.16,
        linewidth = 0.8,
        color = "grey60"
      ) +
      geom_point(size = 3, color = point_col) +
      theme_minimal(base_size = 13) +
      theme(
        panel.background = element_rect(fill = "transparent", color = NA),
        plot.background = element_rect(fill = "transparent", color = NA),
        legend.background = element_rect(fill = "transparent", color = NA),
        legend.box.background = element_rect(fill = "transparent", color = NA),
        panel.grid.major.y = element_blank(),
        panel.grid.minor = element_blank(),
        panel.grid.major.x = element_line(color = "grey75", linewidth = 0.5),
        axis.text = element_text(color = "black"),
        axis.title = element_blank()
      )
    
    ggsave(
      filename = out_file,
      plot = p,
      width = 9,
      height = 6,
      dpi = 600,
      bg = "transparent"
    )
  }
  
  make_edge_plot(
    top10,
    "#00A6FF",
    file.path(out_dir, "Figure_Top10_PositiveEdges.png")
  )
  
  make_edge_plot(
    bot10,
    "#FF3B30",
    file.path(out_dir, "Figure_Bottom10_NegativeEdges.png")
  )
}

# =========================================================
# 11) NCT subgroup analyses
# =========================================================
if (RUN_NCT) {
  
  # --- Sex
  res_sex <- run_nct_and_export(
    dat_male, dat_female,
    adj_male, adj_female,
    label_g1 = "Male", label_g2 = "Female",
    file_stub = "Supplementary_Table_EdgeDiff_Sex",
    it = N_NCT_ITER, gamma = gamma_main, out_dir = out_dir
  )
  
  # --- Income
  res_inc <- run_nct_and_export(
    dat_lowinc, dat_highinc,
    adj_lowinc, adj_highinc,
    label_g1 = "LowIncome", label_g2 = "HighIncome",
    file_stub = "Supplementary_Table_EdgeDiff_Income",
    it = N_NCT_ITER, gamma = gamma_main, out_dir = out_dir
  )
  
  # --- Generation
  res_mz_genx <- run_nct_and_export(
    dat_mz, dat_genx,
    adj_mz, adj_genx,
    label_g1 = "MZ", label_g2 = "GenX",
    file_stub = "Supplementary_Table_EdgeDiff_MZ_vs_GenX",
    it = N_NCT_ITER, gamma = gamma_main, out_dir = out_dir
  )
  
  res_mz_boom <- run_nct_and_export(
    dat_mz, dat_boom,
    adj_mz, adj_boom,
    label_g1 = "MZ", label_g2 = "Boomers",
    file_stub = "Supplementary_Table_EdgeDiff_MZ_vs_Boomers",
    it = N_NCT_ITER, gamma = gamma_main, out_dir = out_dir
  )
  
  res_genx_boom <- run_nct_and_export(
    dat_genx, dat_boom,
    adj_genx, adj_boom,
    label_g1 = "GenX", label_g2 = "Boomers",
    file_stub = "Supplementary_Table_EdgeDiff_GenX_vs_Boomers",
    it = N_NCT_ITER, gamma = gamma_main, out_dir = out_dir
  )
  
  # Optional: Year sensitivity NCT (2020 vs 2023)
  if (RUN_YEAR_SENS && nrow(dat_2020) > 100 && nrow(dat_2023) > 100) {
    res_2020_2023 <- run_nct_and_export(
      dat_2020, dat_2023,
      adj_2020, adj_2023,
      label_g1 = "Year2020", label_g2 = "Year2023",
      file_stub = "Supplementary_Table_EdgeDiff_2020_vs_2023",
      it = N_NCT_ITER, gamma = gamma_main, out_dir = out_dir
    )
  }
}

# =========================================================
# 12) Reviewer-oriented extras
# =========================================================
if (RUN_EXTRAS) {
  
  # ---- Predictability (neighbor-based R2)
  dat_z <- as.data.frame(scale(dat_all))
  pred_R2 <- rep(NA_real_, ncol(dat_z))
  names(pred_R2) <- colnames(dat_z)
  
  for (i in seq_len(ncol(dat_z))) {
    nbr <- which(adj_all[i, ] != 0)
    nbr <- setdiff(nbr, i)
    
    if (length(nbr) == 0) {
      pred_R2[i] <- 0
    } else {
      df_tmp <- as.data.frame(dat_z[, c(i, nbr), drop = FALSE])
      colnames(df_tmp)[1] <- "y"
      fit <- lm(y ~ ., data = df_tmp)
      pred_R2[i] <- summary(fit)$r.squared
    }
  }
  
  pred_tbl <- data.frame(
    node = names(pred_R2),
    predictability_R2 = as.numeric(pred_R2),
    group = node_group[match(names(pred_R2), node_names)]
  ) %>%
    arrange(desc(predictability_R2))
  
  write.xlsx(pred_tbl, file.path(out_dir, "Table_Predictability_R2.xlsx"), rowNames = FALSE)
  
  p_pred <- ggplot(pred_tbl, aes(x = predictability_R2, y = factor(node, levels = pred_tbl$node))) +
    geom_col(aes(fill = group), width = 0.75) +
    scale_fill_manual(values = group_colors) +
    labs(x = "Predictability (R² from neighbors)", y = NULL, title = "Node Predictability") +
    theme_bw(base_size = 13) +
    theme(legend.position = "bottom")
  
  ggsave(
    file.path(out_dir, "Supplementary_Figure_Predictability_R2.png"),
    p_pred, width = 8, height = 9, dpi = 600
  )
  
  # ---- Bridge centrality
  comm <- as.numeric(factor(node_group, levels = unique(node_group)))
  bridge_out <- networktools::bridge(adj_all, communities = comm)
  
  bridge_tbl <- data.frame(
    node = node_names,
    group = node_group,
    stringsAsFactors = FALSE
  )
  
  bridge_strength <- extract_bridge_metric(bridge_out, c("bridgestrength"))
  bridge_ei1 <- extract_bridge_metric(bridge_out, c("bridgeexpectedinfluence1step", "bridgeexpectedinfluence"))
  bridge_ei2 <- extract_bridge_metric(bridge_out, c("bridgeexpectedinfluence2step"))
  
  if (!is.null(bridge_strength)) bridge_tbl$bridgeStrength <- bridge_strength
  if (!is.null(bridge_ei1)) bridge_tbl$bridgeExpectedInfluence_1step <- bridge_ei1
  if (!is.null(bridge_ei2)) bridge_tbl$bridgeExpectedInfluence_2step <- bridge_ei2
  
  write.xlsx(bridge_tbl, file.path(out_dir, "Table_BridgeCentrality.xlsx"), rowNames = FALSE)
  
  # ---- Gamma sensitivity
  gammas <- c(0, 0.25, 0.5)
  adj_list <- lapply(gammas, function(g) estimate_adj(dat_all, gamma = g))
  names(adj_list) <- paste0("gamma_", gammas)
  
  key_edges <- data.frame(
    Edge1 = c("H1","H1","H1","H1"),
    Edge2 = c("A1","A4","W1","W2"),
    stringsAsFactors = FALSE
  )
  
  sens_tbl <- key_edges
  for (nm in names(adj_list)) {
    w <- mapply(function(e1, e2) adj_list[[nm]][e1, e2], sens_tbl$Edge1, sens_tbl$Edge2)
    sens_tbl[[nm]] <- w
  }
  
  write.xlsx(sens_tbl, file.path(out_dir, "Supplementary_Table_Sensitivity_gamma.xlsx"), rowNames = FALSE)
}

cat("\n[Done] All Part 3 outputs saved to:", out_dir, "\n")