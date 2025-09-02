# ============================
# 1) Setup and Configuration
# ============================
# Ensures all necessary packages are installed and loaded.
packages_to_install <- c(
  "readxl", "readr", "dplyr", "tidyr", "janitor", "stringr", "ggplot2",
  "glue", "officer", "flextable", "here", "fs", "sf", "purrr", "lubridate", "ISOweek"
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

message("Listing all files in 'data':")
print(list.files(DATA_DIR, recursive = TRUE))

# Check for required input files to ensure the script runs successfully.
required_files <- c(
  here("data", "weekly_data_by_region.csv"),
  here("data", "measles_lab.xlsx"),
  here("data", "Maternal deaths.xlsx"),
  here("data", "shapefiles", "eth_admbnda_adm1_csa_bofedb_2021.shp")
)
for (f in required_files) {
  if (!file_exists(f)) {
    stop(glue("Missing required file: {f}. Please ensure all data files are in the 'data' directory."))
  }
}

# ============================
# 2) Data Ingestion
# ============================
# Read raw data files. The clean_names() function tidies up column headers.
weekly_raw <- read_csv(here("data", "weekly_data_by_region.csv"), show_col_types = FALSE) %>% clean_names()
measles_raw <- read_excel(here("data", "measles_lab.xlsx")) %>% clean_names()
maternal_linelist_raw <- read_excel(here("data", "Maternal deaths.xlsx"), sheet = "MD linelist") %>% clean_names()
maternal_summary_raw <- read_excel(here("data", "Maternal deaths.xlsx"), sheet = "MD summary", skip = 1) %>% clean_names()
map_regions <- st_read(here("data", "shapefiles", "eth_admbnda_adm1_csa_bofedb_2021.shp"), quiet = TRUE)

# ============================
# 3) Data Processing & Analysis
# ============================
# Function to process weekly data. This is a robust way to reshape the data.
process_weekly_data <- function(df) {
  df %>%
    pivot_longer(
      cols = -c(org_unit_name, period),
      names_to = "measure",
      values_to = "value"
    ) %>%
    mutate(
      disease = str_to_title(str_remove(measure, "_confirmed|_sent_to_lab|_suspected|_deaths")),
      status = str_extract(measure, "confirmed|sent_to_lab|suspected|deaths"),
      year = as.integer(str_extract(period, "^[0-9]{4}")),
      week = as.integer(str_extract(period, "(?<=W)\\d+"))
    ) %>%
    filter(!is.na(week) & week > 0)
}

# Process weekly data to get the latest week and filter.
weekly_long <- process_weekly_data(weekly_raw)
latest <- weekly_long %>% arrange(desc(year), desc(week)) %>% slice(1)
REPORT_YEAR <- latest$year
REPORT_WEEK <- latest$week

weekly_current <- weekly_long %>% filter(year == REPORT_YEAR, week == REPORT_WEEK)
weekly_cum <- weekly_long %>% filter(year == REPORT_YEAR, week <= REPORT_WEEK)

# Summarize data by disease.
summary_week <- weekly_current %>%
  group_by(disease) %>%
  summarise(
    suspected = sum(value[status == "suspected"], na.rm = TRUE),
    tested = sum(value[status == "sent_to_lab"], na.rm = TRUE),
    confirmed = sum(value[status == "confirmed"], na.rm = TRUE),
    .groups = "drop"
  )
summary_cum <- weekly_cum %>%
  group_by(disease) %>%
  summarise(
    suspected = sum(value[status == "suspected"], na.rm = TRUE),
    tested = sum(value[status == "sent_to_lab"], na.rm = TRUE),
    confirmed = sum(value[status == "confirmed"], na.rm = TRUE),
    .groups = "drop"
  )

# Summarize measles lab data.
measles_summary <- measles_raw %>%
  group_by(province_of_residence) %>%
  summarise(
    tested = n(),
    positive = sum(as.numeric(ig_m_results) == 1, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  rename(region = province_of_residence) %>%
  mutate(positivity = round(100 * positive / pmax(tested, 1), 1))

# Match region names and death counts.
maternal_cum_region <- maternal_summary_raw %>%
  transmute(
    region = maternal_death_cumulative,
    deaths = as.numeric(x5)
  ) %>%
  filter(!is.na(deaths), region != "Total")

# Summarize maternal causes data for the current week.
maternal_causes <- maternal_linelist_raw %>%
  mutate(date_of_notification = as.Date(date_of_notification, format = "%Y-%m-%d")) %>%
  filter(!is.na(date_of_notification) & !is.na(short_name)) %>%
  filter(year(date_of_notification) == REPORT_YEAR, isoweek(date_of_notification) == REPORT_WEEK) %>%
  count(short_name) %>%
  rename(cause = short_name)


# ============================
# 4) Plots & Visualizations
# ============================
# Save plots to the figures directory.
plot_path_bar <- file.path(FIG_DIR, "maternal_causes_week.png")
if (nrow(maternal_causes) > 0) {
  ggsave(plot_path_bar,
         ggplot(maternal_causes, aes(x = reorder(cause, n), y = n)) +
           geom_col(fill = "#0072B2") +
           coord_flip() +
           labs(
             title = glue("Causes of Maternal Death — Week {REPORT_WEEK}, {REPORT_YEAR}"),
             x = "",
             y = "Number of deaths"
           ) +
           theme_minimal() +
           theme(plot.title = element_text(face = "bold", size = 14)),
         width = 4.25, height = 3.5, dpi = 300)
}

plot_path_map <- file.path(FIG_DIR, "maternal_deaths_map.png")
if (nrow(maternal_cum_region) > 0) {
  maternal_cum_region <- maternal_cum_region %>%
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

plot_path_measles_stack <- file.path(FIG_DIR, "measles_stack_bar.png")
if (nrow(measles_summary) > 0) {
  measles_long <- measles_summary %>%
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


# ============================
# 5) Tables & Bullet Points
# ============================
# Generates formatted bullet points.
make_bullet <- function(df, disease, metric = "suspected") {
  tmp <- df %>%
    filter(disease == !!disease, status == metric, value > 0) %>%
    group_by(org_unit_name) %>%
    summarise(n = sum(value, na.rm = TRUE), .groups = "drop") %>%
    arrange(desc(n))
  
  total <- sum(tmp$n, na.rm = TRUE)
  if (total == 0) return(glue("• {disease}: no {metric} cases reported."))
  regional_list <- paste0(tmp$org_unit_name, '(', tmp$n, ')', collapse = ', ')
  return(glue("• {disease}: {total} {metric} cases — {regional_list}"))
}

# Generate bullet points for notifiable and other diseases.
notifiable_diseases <- c("Afp", "Anthrax", "Cholera", "Measles")
other_diseases <- c("Non-bloody Diarrhoea", "Malaria", "Typhoid Fever", "Monkeypox", "Plague")

notifiable_bullets <- map_chr(notifiable_diseases, ~make_bullet(weekly_current, .x))
other_bullets <- map_chr(other_diseases, ~make_bullet(weekly_current, .x))

# Define colors for easy management
cobalt_blue <- "#0047AB"
sky_blue <- "skyblue"

# Create the summary table with grouped headers and blue style.
# The use of chained pipes here keeps the code clean and readable.
ft_summary <- flextable(
  left_join(summary_week, summary_cum, by = "disease", suffix = c("_week", "_cum"))
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
total_maternal_deaths_week <- sum(maternal_causes$n, na.rm = TRUE)
total_maternal_deaths_cum <- sum(maternal_cum_region$deaths, na.rm = TRUE)
total_measles_suspected_cum <- sum(weekly_cum$value[weekly_cum$disease == 'Measles' & weekly_cum$status == 'suspected'], na.rm = TRUE)
total_measles_tested <- sum(measles_summary$tested, na.rm = TRUE)
total_measles_positive <- sum(measles_summary$positive, na.rm = TRUE)
measles_positivity_rate <- round(total_measles_positive / pmax(total_measles_tested, 1) * 100, 1)

high_death_regions <- maternal_cum_region %>%
  arrange(desc(deaths)) %>%
  slice(1:3) %>%
  pull(region)
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
    ftext("•      Fill in the summary manually", prop = fp_text(color = "red"))
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



