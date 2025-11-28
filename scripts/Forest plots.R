# Load packages
pacman::p_load(
  tidyverse,
  meta,
  metafor,
  dmetar,
  grid
)


sheet_names <- readxl::excel_sheets(
  here::here(
    "data",
    "data_aspirin_discontinuation_pci.xlsx"
  )
)


#### Retrieving Outcome measures ####

ma <- purrr::map(
  sheet_names, \(x) {
    tibble::as_tibble(
      readxl::read_excel(
        here::here(
          "data",
          "data_aspirin_discontinuation_pci.xlsx"
        ),
        sheet = x
      )
    )
  }
)


names(ma) <- sheet_names

sheet_names

# Import data
# mb_data <- readxl::read_xlsx(here::here("data", "data_aspirin_discontinuation_pci.xlsx"), sheet = "all_cause_mortality")

# Create forest plot

#### Meta analysis of all cause mortality outcome ####

outcome_acm <- "All cause mortality"

ma_acm <- metabin(
  event.e,
  n.e,
  event.c,
  n.c,
  data              = ma$all_cause_mortality,
  method            = "MH",
  method.tau        = "REML",
  method.I2         = "tau2",
  sm                = "RR",
  method.random.ci  = "HK",
  studlab           = glue::glue("{author} {year}"),

  # Add these for better documentation and output:
  title             = outcome_acm, # Documents the outcome
  label.e           = "Short DAPT", # Clearer than default "Experimental"
  label.c           = "Standard DAPT", # Clearer than default "Control"

  # Consider adding:
  common            = FALSE, # Only show random effects (more appropriate)
  prediction        = TRUE # Show prediction interval (already in forest plot)
)

summary(ma_acm)


# Create forest plot PDF
pdf(
  here::here(
    "outputs",
    "forest_plots",
    "data_aspirin_discontinuation_pci_acm.pdf"
  ),
  width = 12, height = 10
)

forest(
  ma_acm,
  title = outcome_acm,
  layout = "RevMan5",
  sortvar = TE,

  # Labels
  label.e = "Short DAPT",
  label.c = "Standard DAPT",
  label.left = "Favours short DAPT",
  label.right = "Favours standard DAPT",
  col.label.left = "grey30",
  col.label.right = "grey30",
  bottom.lr = TRUE,
  fs.lr = 10,

  # Outcome label
  #  smlab                 = outcome,

  # Only show random-effects model
  random = TRUE,
  common = FALSE,
  test.overall.random = TRUE,
  text.random = "Pooled (random effects, 95% CI)",

  # Columns
  leftcols = c("studlab", "event.e", "n.e", "event.c", "n.c", "w.random", "effect", "ci"),
  leftlabs = c("Study", "Events", "N", "Events", "N", "Weight (%)", "RR", "95% CI"),
  rightcols = FALSE,

  # Plot style
  col.square = "blue",
  col.square.lines = "black",
  col.predict = "red",
  col.predict.lines = "black",
  colgap = "3mm",
  colgap.left = "3mm",
  colgap.forest.left = "5mm",
  just.studlab = "left",

  # Prediction interval
  prediction = TRUE,
  text.predict = "Prediction interval",

  # Font sizes and face
  fontsize = 11,
  fs.study = 11,
  fs.heading = 10,
  fs.hetstat = 10,
  fs.predict = 11,
  fs.axis = 10,
  fs.xlab = 12,
  ff.study = "plain",
  ff.random = "bold",
  ff.hetstat = "italic",
  ff.axis = "plain",
  ff.lr = "bold",
  ff.heading = "bold",

  # Optional pooled stats
  pooled.events = TRUE,
  pooled.totals = TRUE,

  # Spacing
  addrows.below.overall = 2
)

grid.text(outcome_acm,
  x = 0.05, # Left side (0 = far left, 1 = far right)
  y = 0.75, # Top (0 = bottom, 1 = top)
  just = "left", # Left-justify the text
  gp = gpar(fontsize = 14, fontface = "bold")
)

dev.off()


#### Meta analysis of major or clinically relevant bleeding outcome ####

outcome_mb <- "Major or clinically relevant bleeding"

ma_mb <- metabin(
  event.e,
  n.e,
  event.c,
  n.c,
  data              = ma$major_or_relevant_bleeding,
  method            = "MH",
  method.tau        = "REML",
  method.I2         = "tau2",
  sm                = "RR",
  method.random.ci  = "HK",
  studlab           = glue::glue("{author} {year}"),
  title = outcome_mb
)

summary(ma_mb)

# Forest plot PDF

pdf(
  here::here(
    "outputs",
    "forest_plots",
    "data_aspirin_discontinuation_pci_mb.pdf"
  ),
  width = 12, height = 10
)

forest(
  ma_mb,
  layout = "RevMan5",
  sortvar = TE,

  # Labels
  label.e = "Short DAPT",
  label.c = "Standard DAPT",
  label.left = "Favours short DAPT",
  label.right = "Favours standard DAPT",
  col.label.left = "grey30",
  col.label.right = "grey30",
  bottom.lr = TRUE,
  fs.lr = 10,

  # Outcome label
  #  smlab                 = outcome,

  # Only show random-effects model
  random = TRUE,
  common = FALSE,
  test.overall.random = TRUE,
  text.random = "Pooled (random effects, 95% CI)",

  # Columns
  leftcols = c("studlab", "event.e", "n.e", "event.c", "n.c", "w.random", "effect", "ci"),
  leftlabs = c("Study", "Events", "N", "Events", "N", "Weight (%)", "RR", "95% CI"),
  rightcols = FALSE,

  # Plot style
  col.square = "blue",
  col.square.lines = "black",
  col.predict = "red",
  col.predict.lines = "black",
  colgap = "3mm",
  colgap.left = "3mm",
  colgap.forest.left = "5mm",
  just.studlab = "left",

  # Prediction interval
  prediction = TRUE,
  text.predict = "Prediction interval",

  # Font sizes and face
  fontsize = 11,
  fs.study = 11,
  fs.heading = 10,
  fs.hetstat = 10,
  fs.predict = 11,
  fs.axis = 10,
  fs.xlab = 12,
  ff.study = "plain",
  ff.random = "bold",
  ff.hetstat = "italic",
  ff.axis = "plain",
  ff.lr = "bold",
  ff.heading = "bold",

  # Optional pooled stats
  pooled.events = TRUE,
  pooled.totals = TRUE,

  # Spacing
  addrows.below.overall = 2
)

grid.text(outcome_mb,
  x = 0.05, # Left side (0 = far left, 1 = far right)
  y = 0.75, # Top (0 = bottom, 1 = top)
  just = "left", # Left-justify the text
  gp = gpar(fontsize = 14, fontface = "bold")
)

dev.off()


# Meta analysis of MI

outcome_mi <- "Myocardial infarction"

ma_mi <- metabin(
  event.e, 
  n.e, 
  event.c, 
  n.c,
  data              = ma$mi,
  method            = "MH",
  method.tau        = "REML",
  method.I2         = "tau2",
  sm                = "RR",
  method.random.ci  = "HK",
  studlab           = glue::glue("{author} {year}"),
  title             = outcome_mi 
)

summary(ma_mi)

pdf(
  here::here(
    "outputs",
    "forest_plots",
    "data_aspirin_discontinuation_pci_mi.pdf"
  ),
  width = 12,
  height = 10
)

meta::forest(
  ma_mi,
  layout = "RevMan5",
#  sortvar = TE,

  # Labels
  label.e = "Short DAPT",
  label.c = "Standard DAPT",
  label.left = "Favours short DAPT",
  label.right = "Favours standard DAPT",
  col.label.left = "grey30",
  col.label.right = "grey30",
  bottom.lr = TRUE,
  fs.lr = 10,

  # Outcome label
  #  smlab                 = outcome,

  # Only show random-effects model
  random = TRUE,
  common = FALSE,
  test.overall.random = TRUE,
  text.random = "Pooled (random effects, 95% CI)",

  # Columns
  leftcols = c("studlab", "event.e", "n.e", "event.c", "n.c", "w.random", "effect", "ci"),
  leftlabs = c("Study", "Events", "N", "Events", "N", "Weight (%)", "RR", "95% CI"),
  rightcols = FALSE,

  # Plot style
  col.square = "blue",
  col.square.lines = "black",
  col.predict = "red",
  col.predict.lines = "black",
  colgap = "3mm",
  colgap.left = "3mm",
  colgap.forest.left = "5mm",
  just.studlab = "left",

  # Prediction interval
  prediction = TRUE,
  text.predict = "Prediction interval",

  # Font sizes and face
  fontsize = 11,
  fs.study = 11,
  fs.heading = 10,
  fs.hetstat = 10,
  fs.predict = 11,
  fs.axis = 10,
  fs.xlab = 12,
  ff.study = "plain",
  ff.random = "bold",
  ff.hetstat = "italic",
  ff.axis = "plain",
  ff.lr = "bold",
  ff.heading = "bold",

  # Optional pooled stats
  pooled.events = TRUE,
  pooled.totals = TRUE,

  # Spacing
  addrows.below.overall = 2
)

grid.text(outcome_mi,
          x = 0.05, # Left side (0 = far left, 1 = far right)
          y = 0.75, # Top (0 = bottom, 1 = top)
          just = "left", # Left-justify the text
          gp = gpar(fontsize = 14, fontface = "bold")
)

dev.off()




# Heterogeneity stats

summary(ma_acm)


# Funnel plot

funnel(ma_acm)

eggers.test(ma_acm)


#### Leave one out sensitivity analysis

loo_acm <- ma_acm |> metainf(pooled = "random")

# Calculate key statistics for annotation
rr_values <- exp(loo_acm$TE)
rr_range_text <- sprintf("RR range: %.2f to %.2f", 
                         min(rr_values), max(rr_values))

i2_values <- loo_acm$I2
i2_range_text <- sprintf("I² range: %.1f%% to %.1f%%", 
                         min(i2_values), max(i2_values))

# Find most influential study
baseline_rr <- exp(ma_acm$TE.random)
rr_changes <- abs(rr_values - baseline_rr)
most_influential <- loo_acm$studlab[which.max(rr_changes)]

pdf(
  here::here(
    "outputs",
    "sensitivity_analysis",
    "data_aspirin_discontinuation_pci_acm_loo.pdf"
  ),
  width = 11, height = 9
)

forest(
  loo_acm,
  smlab = "",
  
  # Colors
  col.bg = "#005f73",
  col.border = "#014f5c",
  col.diamond = "#2a9d8f", 
  col.diamond.lines = "#238276",
  col.predict = "#e76f51",
  col.predict.lines = "#d65c3d",
  
  # Column setup
  leftcols = c("studlab", "effect", "ci", "I2"),
  leftlabs = c("Study omitted", "RR", "95% CI", "I²"),
  rightcols = FALSE,
  
  # Adjust column spacing
  colgap.studlab = "5mm",      
  colgap.forest.left = "8mm",   
  colgap = "3mm",               
  
  # Wider xlim to accommodate all intervals
  xlim = c(0.4, 1.8),
  at = c(0.5, 0.75, 1, 1.25, 1.5, 1.75),
  
  # Font styling
  fontsize = 10,
  fs.heading = 11,
  fs.study = 10,
  fs.xlab = 11,
  ff.heading = "bold",
  ff.xlab = "bold",
  ff.lr = "bold",
  
  # Labels
  label.left = "\nFavors short DAPT",
  label.right = "\nFavors standard DAPT",
  fs.lr = 9,
  
  # Spacing
  spacing = 0.8,
  
  # Prediction interval
  text.predict = "Prediction interval",
  
  # Overall aesthetics
  just.studlab = "left"
)

# Add title
grid.text(
  glue::glue("{outcome_acm}: Leave-one-out sensitivity analysis"),
  x = 0.12, y = 0.95, 
  just = "left", 
  gp = gpar(fontsize = 14, fontface = "bold")
)

# Add key statistics with box - MORE PADDING
grid.rect(
  x = 0.12, y = 0.06,
  width = 0.40, height = 0.12,  # Increased height and width
  just = c("left", "bottom"),
  gp = gpar(fill = "grey95", col = "grey60", lwd = 1)
)

# Text without bullet points (avoids encoding issues)
grid.text(
  glue::glue(
    "Sensitivity Analysis Summary:\n",
    "  {rr_range_text}\n",
    "  {i2_range_text}\n",
    "  Most influential: {most_influential}\n",
    "  Method: MH-REML with HKSJ adjustment"
  ),
  x = 0.14, y = 0.075,
  just = c("left", "bottom"),
  gp = gpar(fontsize = 8.5, lineheight = 1.4)
)

dev.off()


#### Baujat plot for all cause mortality outcome ####

pdf(
  here::here(
    "outputs",
    "heterogeneity",
    "data_aspirin_discontinuation_pci_acm_baujat.pdf"
  ),
  width = 11, height = 9
)

meta::baujat(
  ma_acm,
  bg="red",
  cex.studlab = 0.8,
  yscale = 10, 
  xmin = 4, 
  ymin = 1
  )

dev.off()
