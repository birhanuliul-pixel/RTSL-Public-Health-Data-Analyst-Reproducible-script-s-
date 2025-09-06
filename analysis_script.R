# ============================
# 1) Setup and Configuration
# ============================
# Ensures all necessary packages are installed and loaded.
packages_to_install <- c(
  "readxl", "readr", "dplyr", "tidyr", "janitor", "stringr", "ggplot2",
  "glue", "officer", "flextable", "here", "fs", "sf", "purrr", "lubridate", "ISOweek", "tibble"
)
need <- setdiff(packages_to_install, rownames(installed.packages()))
if (length(need) > 0) install.packages(need, dependencies = TRUE)

# Load all required libraries.
purrr::walk(packages_to_install, library, character.only = TRUE)

# Define directories.
OUT_DIR <- here("outputs")
FIG_DIR <- here("figures")
dir.create(OUT_DIR, showWarnings = FALSE)
dir.create(FIG_DIR, showWarnings = FALSE)
DATA_DIR <- here("data")

# NOTE: The script is designed to use hardcoded data from the documents you provided.
# If you wish to use external files, you must ensure they exist in the correct directories.
# The map plot specifically requires the shapefile.
required_files <- c(
  here("data", "shapefiles", "eth_admbnda_adm1_csa_bofedb_2021.shp")
)
for (f in required_files) {
  if (!file_exists(f)) {
    message(glue("Warning: Missing required file for map plot: {f}. The map will not be generated."))
  }
}

# ============================
# 2) Data Ingestion
# ============================
# Hardcoded data based on the provided Example_Epidemiological_Bulletin45.pdf
weekly_data <- tribble(
  ~disease, ~region, ~suspected,
  "Afp", "Harari", 3,
  "Afp", "Amhara", 2,
  "Afp", "SNNP", 1,
  "Anthrax", "Harari", 7,
  "Anthrax", "SNNP", 1,
  "Cholera", "Dire Dawa", 4,
  "Measles", "Somali", 14,
  "Measles", "Amhara", 7,
  "Measles", "Afar", 3,
  "Measles", "Oromia", 2,
  "Measles", "SNNP", 1,
  "Measles", "Sidama", 1,
  "Non-bloody Diarrhoea", "Afar", 1950,
  "Non-bloody Diarrhoea", "Amhara", 5419,
  "Non-bloody Diarrhoea", "Benishangul Gumz", 1465,
  "Non-bloody Diarrhoea", "Dire Dawa", 3856,
  "Non-bloody Diarrhoea", "Harari", 5142,
  "Non-bloody Diarrhoea", "Oromia", 2495,
  "Non-bloody Diarrhoea", "Sidama", 3828,
  "Non-bloody Diarrhoea", "SNNP", 4163,
  "Non-bloody Diarrhoea", "Somali", 2289,
  "Non-bloody Diarrhoea", "Tigray", 4568,
  "Malaria", "Afar", 22041,
  "Malaria", "Amhara", 25239,
  "Malaria", "Benishangul Gumz", 7282,
  "Malaria", "Dire Dawa", 23810,
  "Malaria", "Harari", 2889,
  "Malaria", "Oromia", 4643,
  "Malaria", "Sidama", 23453,
  "Malaria", "SNNP", 10414,
  "Malaria", "Somali", 14352,
  "Malaria", "Tigray", 10517,
  "Typhoid Fever", "Dire Dawa", 4,
  "Typhoid Fever", "Oromia", 18,
  "Typhoid Fever", "Somali", 9,
  "Typhoid Fever", "Tigray", 2,
  "Monkeypox", "All", 0,
  "Plague", "All", 0
)

# Hardcoded summary report data from the table in Example_Epidemiological_Bulletin45.pdf
summary_report <- tribble(
  ~disease, ~suspected_week, ~tested_week, ~confirmed_week, ~suspected_cum, ~tested_cum, ~confirmed_cum,
  "AFP", 6, 4, 1, 318, 259, 1,
  "Anthrax", 8, 2, 1, 283, 36, 5,
  "Cholera", 4, 0, 0, 16684, 4410, 2034,
  "Non-bloody Diarrhoea", 35175, 2435, 2171, 813029, 46953, 50502,
  "Malaria", 144640, 138717, 40329, 8011914, 7570205, 3693242,
  "Measles", 28, 21, 3, 4360, 1343, 655,
  "Monkeypox", 0, 0, 0, 8, 4, 0,
  "Plague", 0, 0, 0, 13, 1, 0,
  "Typhoid Fever", 33, 20, 4, 822, 499, 45
) %>% clean_names()

# Hardcoded Measles Lab Data from Example_Epidemiological_Bulletin45.pdf
measles_lab_data <- tribble(
  ~region, ~tested, ~positive,
  "Afar", 70, 20,
  "Amhara", 50, 20,
  "Gambela", 18, 0,
  "Harari", 62, 20,
  "Oromia", 160, 40,
  "SNNP", 20, 0,
  "South West", 62, 28,
  "Tigray", 112, 60
)

# Hardcoded Maternal Deaths Data
# Cumulative data from Example_Epidemiological_Bulletin45.pdf
maternal_deaths_cum <- tribble(
  ~region, ~deaths,
  "Tigray", 48,
  "Amhara", 59,
  "Addis Ababa", 39,
  "Harari", 29,
  "Gambela", 23,
  "South West Ethiopia", 22,
  "SNNP", 29,
  "Oromia", 30,
  "Dire Dawa", 38,
  "Afar", 0,
  "Benishangul Gumz", 0,
  "Somali", 0,
  "Sidama", 28
)

# Weekly data from Example_Epidemiological_Bulletin45.pdf
maternal_deaths_week <- tribble(
  ~cause, ~n,
  "Obstetric haemorrhage", 5,
  "Non-obstetric complications", 4,
  "Unanticipated complications", 3,
  "Hypertensive disorder", 2,
  "Pregnancy-related infection", 1
)

# Set report period
REPORT_YEAR <- 2024
REPORT_WEEK <- 34

# ============================
# 3) Data Processing & Analysis
# ============================
weekly_current <- weekly_data %>%
  rename(org_unit_name = region, disease = disease, value = suspected) %>%
  mutate(status = "suspected")
summary_week <- summary_report %>%
  select(disease, suspected_week, tested_week, confirmed_week) %>%
  rename(suspected = suspected_week, tested = tested_week, confirmed = confirmed_week)
summary_cum <- summary_report %>%
  select(disease, suspected_cum, tested_cum, confirmed_cum) %>%
  rename(suspected = suspected_cum, tested = tested_cum, confirmed = confirmed_cum)

# Summarize measles lab data.
measles_summary <- measles_lab_data %>%
  rename(province_of_residence = region) %>%
  mutate(positivity = round(100 * positive / pmax(tested, 1), 1))

# Summarize maternal causes data for the current week.
maternal_causes <- maternal_deaths_week

# ============================
# 4) Plots & Visualizations
# ============================
# Save plots to the figures directory.
plot_path_bar <- file.path(FIG_DIR, "maternal_causes_week.png")
if (nrow(maternal_causes) > 0) {
  # Calculate percentages for the maternal deaths plot
  maternal_causes_pct <- maternal_causes %>%
    mutate(pct = n / sum(n) * 100) %>%
    mutate(pct_label = paste0(round(pct, 0), "%"))
  
  # Define a buffer to prevent labels from being cut off
  plot_limit <- max(maternal_causes_pct$n) * 1.2
  
  ggsave(plot_path_bar,
         ggplot(maternal_causes_pct, aes(x = reorder(cause, n), y = n)) +
           geom_col(fill = "#0072B2") +
           # Add text labels for percentages
           geom_text(aes(label = pct_label), hjust = -0.2, size = 3, color = "black") +
           coord_flip() +
           labs(
             title = glue("Causes of Maternal Death — Week {REPORT_WEEK}, {REPORT_YEAR}"),
             x = "",
             y = "Number of deaths"
           ) +
           # Expand the x-axis to prevent text from being cut off
           scale_y_continuous(limits = c(0, plot_limit)) +
           theme_minimal() +
           theme(plot.title = element_text(face = "bold", size = 14)),
         width = 4.25, height = 3.5, dpi = 300)
}

# The measles stacked bar chart
plot_path_measles_stack <- file.path(FIG_DIR, "measles_stack_bar.png")
if (nrow(measles_lab_data) > 0) {
  measles_long <- measles_lab_data %>%
    pivot_longer(cols = c(positive, tested), names_to = "status", values_to = "count") %>%
    filter(status == "positive" | status == "tested") %>%
    group_by(region) %>%
    summarise(
      positive = sum(count[status == "positive"], na.rm = TRUE),
      negative = sum(count[status == "tested"], na.rm = TRUE) - positive,
      .groups = "drop"
    ) %>%
    pivot_longer(cols = c(positive, negative), names_to = "status", values_to = "count")
  
  ggsave(plot_path_measles_stack,
         ggplot(measles_long, aes(x = region, y = count, fill = status)) +
           geom_bar(stat = "identity") +
           labs(
             title = "Measles Laboratory Test Results by Region",
             x = "Region",
             y = "Number of Cases",
             fill = "Test Result"
           ) +
           scale_fill_manual(values = c("negative" = "turquoise", "positive" = "salmon")) +
           theme_minimal() +
           theme(
             plot.title = element_text(face = "bold", size = 14),
             axis.text.x = element_text(angle = 45, hjust = 1)
           ),
         width = 4.75, height = 3.5, dpi = 300)
}

# Code for the Ethiopian map plot, requires the shapefile
plot_path_map <- file.path(FIG_DIR, "maternal_deaths_map.png")
map_file_path <- here("data", "shapefiles", "eth_admbnda_adm1_csa_bofedb_2021.shp")
if (file.exists(map_file_path)) {
  map_regions <- st_read(map_file_path, quiet = TRUE)
  maternal_cum_region <- maternal_deaths_cum %>%
    mutate(region = case_when(
      region == "SNNP" ~ "Southern Nations, Nationalities and Peoples",
      region == "Benshangul Gumz" ~ "Benishangul-Gumuz",
      region == "South West" ~ "South West Ethiopia",
      TRUE ~ region
    ))
  
  map_data <- map_regions %>%
    mutate(REGION_NAME = str_to_title(str_trim(ADM1_EN))) %>%
    left_join(maternal_cum_region, by = c("REGION_NAME" = "region")) %>%
    mutate(
      deaths = replace_na(deaths, 0),
      deaths_cat = cut(
        deaths,
        breaks = c(-1, 0, 10, 50, Inf),
        labels = c("0", "1–10", "11–50", ">50")
      )
    )
  
  centroids <- st_centroid(map_data, of_largest_polygon = TRUE)
  
  ggsave(
    plot_path_map,
    ggplot(map_data) +
      geom_sf(aes(fill = deaths_cat), color = "gray30", linewidth = 0.2) +
      geom_sf_text(
        data = centroids,
        aes(label = glue("{REGION_NAME}\n({deaths})")),
        size = 2.5, na.rm = TRUE
      ) +
      scale_fill_manual(
        values = c("0" = "#E0E0FF", "1–10" = "#B3B3FF", "11–50" = "#8080FF", ">50" = "#4B0082"),
        na.value = "white",
        name = "Cumulative Maternal Deaths"
      ) +
      labs(title = glue("Cumulative Maternal Deaths — Week {REPORT_WEEK}, {REPORT_YEAR}")) +
      theme_minimal() +
      theme(
        plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
        axis.text = element_blank(),
        axis.title = element_blank(),
        panel.grid = element_blank(),
        legend.position = c(0, 1),
        legend.justification = c("left", "top")
      ),
    width = 5.5, height = 3.5, dpi = 300
  )
}

# ============================
# 5) Tables & Bullet Points
# ============================
# Generates formatted bullet points.
make_bullet <- function(df, disease) {
  tmp <- df %>%
    filter(disease == !!disease) %>%
    group_by(region) %>%
    summarise(n = sum(suspected, na.rm = TRUE), .groups = "drop") %>%
    filter(n > 0) %>%
    arrange(desc(n))
  
  total <- sum(tmp$n, na.rm = TRUE)
  if (total == 0) return(glue("• {disease}: no suspected cases reported."))
  regional_list <- paste0(tmp$region, '(', tmp$n, ')', collapse = ', ')
  return(glue("• {disease}: {total} suspected cases — {regional_list}"))
}

# Generate bullet points for notifiable and other diseases.
notifiable_diseases <- c("Afp", "Anthrax", "Cholera", "Measles")
other_diseases <- c("Non-bloody Diarrhoea", "Malaria", "Typhoid Fever", "Monkeypox", "Plague")

notifiable_bullets <- map_chr(notifiable_diseases, ~make_bullet(weekly_data, .x))
other_bullets <- map_chr(other_diseases, ~make_bullet(weekly_data, .x))

# Define colors for easy management
cobalt_blue <- "#0047AB"
sky_blue <- "skyblue"

# Create the summary table with grouped headers and blue style.
ft_summary <- flextable(
  left_join(summary_week, summary_cum, by = c("disease" = "disease"))
) %>%
  set_header_labels(
    disease = "Disease/Event/Condition",
    suspected_week = "Suspected", tested_week = "Tested", confirmed_week = "Confirmed",
    suspected_cum = "Suspected", tested_cum = "Tested", confirmed_cum = "Confirmed"
  ) %>%
  add_header_row(
    values = c("", glue("Week {REPORT_WEEK}"), glue("Week 1 to {REPORT_WEEK}, Cumulative Total")),
    colwidths = c(1, 3, 3)
  ) %>%
  bold(part = "header") %>%
  align(align = "center", part = "header") %>%
  valign(valign = "center", part = "header") %>%
  bold(j = 1, part = "body") %>%
  align(j = 2:7, align = "center", part = "body") %>%
  width(j = c(1, 2:7), width = c(2.0, rep(0.8, 6))) %>%
  bg(part = "header", bg = "#00B0F0") %>%
  color(part = "header", color = "white") %>%
  bg(i = ~ !is.na(disease), j = 1, bg = "#E6F5FF", part = "body") %>%
  color(i = ~ !is.na(disease), j = 1, color = "black", part = "body") %>%
  border_outer(border = fp_border(color = "black", width = 1.5), part = "all") %>%
  border_inner_h(border = fp_border(color = "black", width = 0.8), part = "all") %>%
  border_inner_v(border = fp_border(color = "black", width = 0.8), part = "all") %>%
  height_all(height = 0.35)

# Calculate dynamic values for the report text.
total_maternal_deaths_week <- 15
total_maternal_deaths_cum <- 410
total_measles_suspected_cum <- 4360
total_measles_tested <- 553
total_measles_positive <- 192
measles_positivity_rate <- round(total_measles_positive / pmax(total_measles_tested, 1) * 100, 1)

high_death_regions <- c("Amhara", "Tigray", "Addis Ababa")
high_death_regions_text <- paste(high_death_regions, collapse = ", ")

get_week_dates <- function(week, year) {
  start_date <- ISOweek::ISOweek2date(glue("{year}-W{sprintf('%02d', week)}-1"))
  end_date <- start_date + days(6)
  return(glue("{day(start_date)}th - {day(end_date)}th, {month(start_date, label = TRUE, abbr = FALSE)}, {year(start_date)}"))
}

REPORT_DATES <- get_week_dates(REPORT_WEEK, REPORT_YEAR)

# ============================
# 6) Build DOCX Bulletin
# ============================
# Constructs the final Word document using a series of chained `body_add_*` calls.
fname <- file.path(OUT_DIR, glue("Epidemiological_Bulletin_{REPORT_YEAR}W{REPORT_WEEK}.docx"))
if (file.exists(fname)) file.remove(fname)

# Simplified Title Box Creation
create_title_box <- function(title, bg_color, font_color) {
  flextable(data.frame(col = title)) %>%
    width(width = 6.9) %>%
    fontsize(size = 12, part = "body") %>%
    bold(part = "body") %>%
    color(color = font_color, part = "body") %>%
    bg(bg = bg_color, part = "body") %>%
    align(align = "center", part = "body") %>%
    border_remove()
}

# Generate the automated header title string.
header_title <- glue("Week {REPORT_WEEK} Epidemiological Bulletin {REPORT_DATES}")

doc <- read_docx() %>%
  body_add_flextable(create_title_box(header_title, cobalt_blue, "white")) %>%
  body_add_par("") %>%
  body_add_flextable(create_title_box("Summary", sky_blue, "black")) %>%
  body_add_par("Executive Summary", style = "Normal") %>%
  body_add_fpar(fpar(
    ftext("•     Fill in the summary manually", prop = fp_text(color = "red"))
  )) %>%
  body_add_par("Immediately Notifiable Diseases and Events", style = "heading 2") %>%
  body_add_par(paste(notifiable_bullets, collapse = "\n"), style = "Normal") %>%
  body_add_par("Other Diseases and Events", style = "heading 2") %>%
  body_add_par(paste(other_bullets, collapse = "\n"), style = "Normal") %>%
  body_add_flextable(create_title_box("Summary Report Priority Diseases, Conditions and Events", sky_blue, "black")) %>%
  body_add_flextable(ft_summary) %>%
  body_add_flextable(create_title_box("Summary of VPD Surveillance Indicators", sky_blue, "black")) %>%
  body_add_par("") %>%
  body_add_flextable(create_title_box("Measles Laboratory Test Results by Region", cobalt_blue, "white"))

if (file.exists(plot_path_measles_stack)) {
  doc <- doc %>%
    body_add_par(glue("• The country has recorded a total of {total_measles_suspected_cum} suspected measles cases in {REPORT_YEAR}.\n• From the {total_measles_tested} measles specimen that have been tested, {total_measles_positive} have been confirmed positive (PR {measles_positivity_rate} %).")) %>%
    body_add_img(src = plot_path_measles_stack, width = 4.75, height = 3.5)
}

doc <- doc %>%
  body_add_flextable(create_title_box("Maternal Deaths", sky_blue, "black")) %>%
  body_add_par("")

if (file.exists(plot_path_bar) && file.exists(plot_path_map)) {
  doc <- doc %>%
    body_add_fpar(fpar(
      external_img(src = plot_path_bar, width = 4.25, height = 3.5),
      external_img(src = plot_path_map, width = 5.5, height = 3.5)
    ))
} else if (file.exists(plot_path_bar)) {
  doc <- doc %>% body_add_img(src = plot_path_bar, width = 4.25, height = 3.5)
} else if (file.exists(plot_path_map)) {
  doc <- doc %>% body_add_img(src = plot_path_map, width = 5.5, height = 3.5)
}

doc <- doc %>%
  body_add_par(glue("• The bar chart on the left summarizes the causes of deaths of {total_maternal_deaths_week} maternal deaths recorded in week {REPORT_WEEK}."), style = "Normal") %>%
  body_add_par(glue("• Hypertensive disorder, Non-obstetric complications, and Obstetric haemorrhage continue to be the leading causes of maternal deaths this year."), style = "Normal") %>%
  body_add_par(glue("• Cumulatively, in {REPORT_YEAR}, {total_maternal_deaths_cum} maternal deaths have been recorded across the country, as depicted on the map."), style = "Normal") %>%
  body_add_par(glue("• Regions with darker shades ({high_death_regions_text}) indicate those with a higher number of reported maternal deaths."), style = "Normal") %>%
  body_add_par("— End of Bulletin —", style = "Normal")

print(doc, target = fname)
message(glue("✅ Bulletin created: {fname}"))