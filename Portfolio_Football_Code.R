# Historical football portfolio
# What this script does
# It reads the Excel workbook, prepares the team-season data, builds the
# rankings used in the report, and saves the tables and charts in one folder.
# The comments explain the choices that are not obvious from the code.
#
# The four measures stay separate throughout:
#   - points per match among non-champions
#   - overall points per match
#   - goals scored per match
#   - goals conceded per match
#
# I use the historical ratios already supplied in the workbook. In particular,
# I do not recalculate old two-point seasons as if three points had been awarded
# for a win. No combined score or "best club overall" ranking is created.

# SETUP AND HELPERS
# Packages

required_packages <- c(
  "readxl",
  "dplyr",
  "tidyr",
  "stringr",
  "purrr",
  "ggplot2",
  "forcats",
  "scales",
  "readr"
)

missing_packages <- required_packages[
  !required_packages %in% rownames(installed.packages())
]

if (length(missing_packages) > 0) {
  install.packages(missing_packages, dependencies = TRUE)
}

invisible(lapply(required_packages, library, character.only = TRUE))


# A few settings to change before running the script

# TRUE = Excel file selection dialog.
# FALSE = use the path specified below.
USE_FILE_DIALOG <- TRUE

excel_file <- "PL.xlsx"

# Hungary is included by default: it is one of the 8
# requested leagues.
INCLUDE_HUNGARY <- TRUE

# For "Top 5 vs others" comparisons, we use:
BIG5 <- c(
  "Premier League",
  "Ligue 1",
  "Serie A",
  "Bundesliga",
  "La Liga"
)

OTHER_LEAGUES <- c(
  "Eredivisie",
  "Primeira Liga",
  "Hungary"
)

ALL_INDIVIDUAL_LEAGUES <- c(BIG5, OTHER_LEAGUES)

# Periods comparable across all leagues.
# The PL 1955-92 sheet is kept for Premier League-only analyses,
# but is never included in cross-league comparisons.
INTER_PERIODS <- c("Since 3 points", "1955+", "All Time")

# Size of homogeneous comparisons.
TOP_N_STANDARD <- 10

# All Time:
#   - each league on its own: top 15
#   - combined Big Five ranking: top 20
TOP_N_ALLTIME_INDIVIDUAL <- 15
TOP_N_ALLTIME_BIG5 <- 20

# Results are exported to a folder next to the Excel file.
if (USE_FILE_DIALOG) {
  excel_file <- file.choose()
}

if (!file.exists(excel_file)) {
  stop("Fichier Excel introuvable : ", excel_file)
}

output_dir <- file.path(
  dirname(normalizePath(excel_file)),
  "PORTFOLIO_FOOTBALL_ENGLISH"
)

dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

plots_dir <- file.path(output_dir, "graphs")
tables_dir <- file.path(output_dir, "tables")
dir.create(plots_dir, recursive = TRUE, showWarnings = FALSE)
dir.create(tables_dir, recursive = TRUE, showWarnings = FALSE)


# Small helper functions

clean_text <- function(x) {
  x <- as.character(x)
  x <- str_replace_all(x, "\u00A0", " ")
  x <- str_replace_all(x, "[\r\n\t]+", " ")
  x <- str_squish(x)
  x
}

clean_header <- function(x) {
  x <- clean_text(x)
  x <- str_to_lower(x)

  x <- chartr(
    "éèêëàâäùûüôöîïçÉÈÊËÀÂÄÙÛÜÔÖÎÏÇ",
    "eeeeaaauuuooiicEEEEAAAUUUOOIIC",
    x
  )

  x <- str_replace_all(x, "\\s+", " ")
  str_trim(x)
}

num <- function(x) {

  if (is.numeric(x)) {
    return(as.numeric(x))
  }

  x <- as.character(x)
  x <- str_trim(x)

  # Remove spaces but keep signs.
  x <- str_replace_all(x, "\\s+", "")

  # French comma -> decimal point.
  x <- str_replace_all(x, ",", ".")

  # Keep digits, decimal points and minus signs; discard other characters.
  x <- str_replace_all(x, "[^0-9.\\-]", "")

  suppressWarnings(as.numeric(x))
}

first_non_na <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) == 0) return(NA)
  x[1]
}

safe_mean <- function(x) {
  x <- x[is.finite(x)]
  if (length(x) == 0) return(NA_real_)
  mean(x)
}

safe_median <- function(x) {
  x <- x[is.finite(x)]
  if (length(x) == 0) return(NA_real_)
  median(x)
}

safe_cor <- function(x, y) {
  ok <- is.finite(x) & is.finite(y)
  if (sum(ok) < 3) return(NA_real_)
  cor(x[ok], y[ok])
}

safe_file_name <- function(x) {
  x %>%
    str_replace_all("[^A-Za-z0-9_-]+", "_") %>%
    str_replace_all("_+", "_") %>%
    str_replace_all("^_|_$", "")
}

period_order <- c(
  "All Time",
  "1955+",
  "1955-1992",
  "Since 3 points"
)

metric_order <- c(
  "points_non_champion",
  "points_general",
  "goals_for",
  "goals_against"
)

metric_labels <- c(
  points_non_champion = "Points / match — without being champion",
  points_general = "Points / match — overall",
  goals_for = "Goals scored / match",
  goals_against = "Goals conceded / match"
)

metric_short <- c(
  points_non_champion = "Non-champion pts",
  points_general = "Overall pts",
  goals_for = "Goals scored",
  goals_against = "Goals conceded"
)

league_colors <- c(
  "Premier League" = "#1F77B4",
  "La Liga" = "#D62728",
  "Serie A" = "#2CA02C",
  "Bundesliga" = "#4A4A4A",
  "Ligue 1" = "#9467BD",
  "Eredivisie" = "#FF7F0E",
  "Primeira Liga" = "#17BECF",
  "Hungary" = "#8C564B"
)

# Club palette: recognizable colors for major clubs,
# then a deterministic fallback color for all other clubs.
# The same club therefore keeps the same color across all charts.
club_color <- function(team) {
  x <- stringr::str_to_lower(stringr::str_squish(team))

  known <- c(
    "liverpool fc" = "#C8102E",
    "real madrid" = "#C9A227",
    "fc barcelone" = "#A50044",
    "fc barcelona" = "#A50044",
    "manchester united" = "#DA291C",
    "arsenal" = "#DB0007",
    "tottenham hotspur" = "#132257",
    "chelsea fc" = "#034694",
    "manchester city" = "#6CABDD",
    "leeds united" = "#FFCD00",
    "nottingham forest" = "#E53233",
    "everton" = "#003399",
    "aston villa" = "#670E36",
    "newcastle united" = "#241F20",
    "west ham united" = "#7A263A",
    "wolverhampton wanderers" = "#FDB913",
    "sunderland" = "#D71920",
    "preston north end" = "#1D3557",
    "juventus turin" = "#171717",
    "inter milan" = "#0068A8",
    "ac milan" = "#FB090B",
    "ac milan (calciopoli)" = "#FB090B",
    "s.s.c. napoli" = "#12A0D7",
    "a.s. roma" = "#8E1F2F",
    "fiorentina" = "#5B2A86",
    "torino" = "#8C1D40",
    "bologna" = "#A41E34",
    "paris sg" = "#004170",
    "rc lens" = "#D20A0A",
    "as monaco" = "#E21E2B",
    "olympique de marseille" = "#2FAEE0",
    "ogc nice" = "#E30613",
    "olymlique lyonnais" = "#D71920",
    "olyimque lyonnais" = "#D71920",
    "olymique lyonnais" = "#D71920",
    "fc nantes" = "#F2D000",
    "as saint etienne" = "#008F5A",
    "stade de reims" = "#E30613",
    "fc sochaux montbéliard" = "#F7C600",
    "rc strasbourg" = "#0066B3",
    "real sociedad" = "#00529F",
    "athletic bilbao" = "#EE2523",
    "fc bayern munich" = "#DC052D",
    "borussia dortmund" = "#FDE100",
    "bayer leverkusen" = "#E32221",
    "vfb stuttgart" = "#E32219",
    "werder bremen" = "#1D9053",
    "schalke 04" = "#0F4C9C",
    "hamburger sv" = "#005CA9",
    "borussia mönchengladbach" = "#111111",
    "ajax amsterdam" = "#D2122E",
    "feyenoord rotterdam" = "#E30613",
    "psv eindhoven" = "#ED1C24",
    "fc twente" = "#E30613",
    "az alkmaar" = "#E30613",
    "sl benfica" = "#ED1C24",
    "sporting cp" = "#008E4F",
    "fc porto" = "#0050A4",
    "s.c. braga" = "#E30613",
    "ferencváros" = "#008751",
    "ferencvaros" = "#008751",
    "újpest" = "#6A1B9A",
    "ujpest" = "#6A1B9A",
    "mtk budapest" = "#0072CE",
    "mtk hungária" = "#0072CE",
    "mtk hungaria" = "#0072CE",
    "budapest honvéd" = "#D71920",
    "budapest honved" = "#D71920",
    "vasas sc" = "#E30613",
    "debrecen" = "#0057A8",
    "eto fc győr" = "#0066B3",
    "eto fc gyor" = "#0066B3",
    "valencia" = "#EE7D11",
    "valence" = "#EE7D11"
  )

  if (x %in% names(known)) return(unname(known[[x]]))

  # A few common variants are handled by pattern.
  if (str_detect(x, "real madrid")) return("#C9A227")
  if (str_detect(x, "barcelone|barcelona")) return("#A50044")
  if (str_detect(x, "bayern")) return("#DC052D")
  if (str_detect(x, "dortmund")) return("#FDE100")
  if (str_detect(x, "juventus")) return("#171717")
  if (str_detect(x, "milan")) return(ifelse(str_detect(x, "inter"), "#0068A8", "#FB090B"))
  if (str_detect(x, "napoli")) return("#12A0D7")
  if (str_detect(x, "paris sg|psg")) return("#004170")
  if (str_detect(x, "monaco")) return("#E21E2B")
  if (str_detect(x, "porto")) return("#0050A4")
  if (str_detect(x, "benfica")) return("#ED1C24")
  if (str_detect(x, "sporting")) return("#008E4F")
  if (str_detect(x, "ajax")) return("#D2122E")
  if (str_detect(x, "psv")) return("#ED1C24")
  if (str_detect(x, "ferenc")) return("#008751")

  # Stable fallback color: it depends only on the club name.
  code <- sum(utf8ToInt(x) * seq_along(utf8ToInt(x)))
  h <- (code %% 360)
  hcl(h = h, c = 65, l = 58)
}

club_colors <- function(teams) {
  teams <- unique(as.character(teams))
  cols <- vapply(teams, club_color, character(1))
  stats::setNames(cols, teams)
}


# READING AND PREPARING THE DATA
# Recognise each league and historical period

get_league <- function(sheet_name) {

  s <- clean_header(sheet_name)

  case_when(

    str_detect(s, "top ?5 ?eu") ~ "Top 5 Europe",

    str_detect(s, "premier league|^pl") ~ "Premier League",

    str_detect(s, "serie a|ital") ~ "Serie A",

    str_detect(s, "ligue ?1|france") ~ "Ligue 1",

    str_detect(s, "primeira|liga nos|portugal|bwin") ~
      "Primeira Liga",

    # The Spanish sheets in the workbook are named "Liga 1995",
    # "Liga 1955", "Liga AT" : do not confuse them with Liga NOS.
    str_detect(s, "la liga|liga esp|espagne|spain|^liga( |$)") ~ "La Liga",

    str_detect(s, "bundes|germany|allem") ~ "Bundesliga",

    str_detect(s, "eredivisie|netherlands|holland|pays.bas") ~
      "Eredivisie",

    str_detect(s, "nemzeti|hungary|hongrie|bajnoksag") ~
      "Hungary",

    TRUE ~ str_trim(sheet_name)
  )
}


get_period <- function(sheet_name) {

  s <- clean_header(sheet_name)

  case_when(

    str_detect(s, "at$|all.?time") ~ "All Time",

    # Specific historical Premier League sheet.
    # It is retained for PL analyses, but excluded from cross-league comparisons
    # inter-leagues because the other leagues do not have this window.
    str_detect(s, "1955.*92|55.*92") ~ "1955-1992",

    # Each league has its own transition year (92/94/95),
    # but all are grouped under the same analytical period.
    str_detect(s, "(^|[^0-9])(1992|1994|1995|92|94|95)([^0-9]|$)") ~ "Since 3 points",

    str_detect(s, "1955") ~ "1955+",

    # Non-analytical sheets (ex. Global) must not become
    # a fake "Other" period. They will simply be excluded.
    TRUE ~ NA_character_
  )
}


# Read the four tables on each worksheet

find_team_headers <- function(raw_chr) {

  mat <- as.matrix(raw_chr)

  which(
    mat == "equipe",
    arr.ind = TRUE
  )
}


find_col <- function(headers, pattern) {

  hits <- which(
    str_detect(headers, regex(pattern, ignore_case = TRUE))
  )

  if (length(hits) == 0) {
    return(NA_integer_)
  }

  hits[1]
}


extract_table <- function(
    raw,
    raw_chr,
    header_row,
    team_col,
    table_right,
    next_header_row,
    metric,
    sheet_name
) {

  # The "Position" header is located just to the left of "Equipe" in
  # the workbook tables. This column must therefore be included in
  # the reading window, otherwise top_rank becomes NA for the tables
  # This matters especially for the four goals tables.
  table_left <- if (team_col > 1 &&
                    identical(raw_chr[header_row, team_col - 1], "position")) {
    team_col - 1
  } else {
    team_col
  }

  headers <- clean_header(
    raw_chr[
      header_row,
      table_left:table_right,
      drop = TRUE
    ]
  )

  team_local <- find_col(headers, "^equipe$")
  season_local <- find_col(headers, "^saison$")
  matches_local <- find_col(headers, "nombre de matchs")
  position_local <- find_col(headers, "^position$")
  rank_local <- find_col(headers, "^classement$")

  points_local <- find_col(headers, "points? pris")
  ppm_local <- find_col(headers, "ratio points")

  goals_for_local <- find_col(headers, "buts marqu")
  goals_against_local <- find_col(headers, "buts encai")
  gpm_local <- find_col(headers, "ratio buts")

  to_global <- function(local_col) {

    if (is.na(local_col)) {
      return(NA_integer_)
    }

    table_left + local_col - 1
  }

  team_col_g <- to_global(team_local)
  season_col_g <- to_global(season_local)
  matches_col_g <- to_global(matches_local)
  position_col_g <- to_global(position_local)
  rank_col_g <- to_global(rank_local)
  points_col_g <- to_global(points_local)
  ppm_col_g <- to_global(ppm_local)
  goals_for_col_g <- to_global(goals_for_local)
  goals_against_col_g <- to_global(goals_against_local)
  gpm_col_g <- to_global(gpm_local)

  first_data_row <- header_row + 1

  last_data_row <- if (!is.na(next_header_row)) {
    next_header_row - 1
  } else {
    nrow(raw)
  }

  if (last_data_row < first_data_row) {
    return(NULL)
  }

  rows <- first_data_row:last_data_row

  get_column <- function(col_number) {

    if (is.na(col_number)) {
      return(rep(NA, length(rows)))
    }

    raw[
      rows,
      col_number,
      drop = TRUE
    ]
  }

  out <- tibble(
    team = as.character(get_column(team_col_g)),
    season = as.character(get_column(season_col_g)),
    matches = num(get_column(matches_col_g)),
    top_rank = as.character(get_column(position_col_g)),
    rank = as.character(get_column(rank_col_g)),
    points = num(get_column(points_col_g)),
    historical_ppm = num(get_column(ppm_col_g)),
    goals_for = num(get_column(goals_for_col_g)),
    goals_against = num(get_column(goals_against_col_g)),
    historical_gpm = num(get_column(gpm_col_g)),
    metric = metric,
    sheet = sheet_name,
    league = get_league(sheet_name),
    period = get_period(sheet_name)
  )

  out <- out %>%
    mutate(
      team = clean_text(team),
      season = clean_text(season),
      top_rank = clean_text(top_rank),
      rank = clean_text(rank)
    ) %>%
    filter(
      !is.na(team),
      team != "",
      !str_detect(str_to_lower(team), "^equipe$"),
      !str_detect(str_to_lower(team), "^position$"),
      !str_detect(str_to_lower(team), "^saison$")
    )

  # The points ratio from the file must have PRIORITY.
  # This is important for 2-point seasons: the sheets
  # historical sheets explicitly indicate a ratio "normalized to 3 pts".
  out <- out %>%
    mutate(
      ppm = if_else(
        !is.na(historical_ppm),
        historical_ppm,
        if_else(
          matches > 0,
          points / matches,
          NA_real_
        )
      ),

      gpm_for = if_else(
        !is.na(historical_gpm) & metric == "goals_for",
        historical_gpm,
        if_else(
          matches > 0 & metric == "goals_for",
          goals_for / matches,
          NA_real_
        )
      ),

      gpm_against = if_else(
        !is.na(historical_gpm) & metric == "goals_against",
        historical_gpm,
        if_else(
          matches > 0 & metric == "goals_against",
          goals_against / matches,
          NA_real_
        )
      )
    )

  out
}


parse_sheet <- function(sheet_name, file) {

  raw <- read_excel(
    file,
    sheet = sheet_name,
    col_names = FALSE,
    .name_repair = "minimal"
  )

  raw <- as.data.frame(
    raw,
    stringsAsFactors = FALSE
  )

  raw_chr <- as.data.frame(
    lapply(
      raw,
      clean_header
    ),
    stringsAsFactors = FALSE
  )

  team_positions <- find_team_headers(raw_chr)

  if (nrow(team_positions) < 4) {

    stop(
      paste0(
        "The sheet '",
        sheet_name,
        "' contains only ",
        nrow(team_positions),
        " 'Equipe' headers. 4 are expected."
      )
    )
  }

  team_positions <- team_positions[
    order(
      team_positions[, "row"],
      team_positions[, "col"]
    ),
    ,
    drop = FALSE
  ]

  # Each worksheet contains four tables arranged in a 2-by-2 grid.
  team_positions <- team_positions[
    seq_len(4),
    ,
    drop = FALSE
  ]

  rows_unique <- sort(unique(team_positions[, "row"]))
  cols_unique <- sort(unique(team_positions[, "col"]))

  top_row <- min(rows_unique)
  bottom_row <- max(rows_unique)
  left_col <- min(cols_unique)
  right_col <- max(cols_unique)

  metric_from_position <- function(r, c) {

    if (r == top_row && c == left_col) {
      return("points_non_champion")
    }

    if (r == bottom_row && c == left_col) {
      return("points_general")
    }

    if (r == top_row && c == right_col) {
      return("goals_for")
    }

    if (r == bottom_row && c == right_col) {
      return("goals_against")
    }

    NA_character_
  }

  result_list <- vector(
    "list",
    nrow(team_positions)
  )

  for (i in seq_len(nrow(team_positions))) {

    r <- team_positions[i, "row"]
    c <- team_positions[i, "col"]

    metric <- metric_from_position(r, c)

    if (is.na(metric)) {
      next
    }

    # Stop at the column just before the neighbouring table on the right.
    right_candidates <- team_positions[
      team_positions[, "row"] == r &
        team_positions[, "col"] > c,
      "col"
    ]

    table_right <- if (length(right_candidates) > 0) {
      min(right_candidates) - 1
    } else {
      ncol(raw)
    }

    # Le prochain header horizontalement n'existe pas ici :
    # The next table header is the next "Equipe" row below this one.
    lower_headers <- team_positions[
      team_positions[, "row"] > r,
      "row"
    ]

    next_header_row <- if (length(lower_headers) > 0) {
      min(lower_headers)
    } else {
      NA_integer_
    }

    result_list[[i]] <- extract_table(
      raw = raw,
      raw_chr = raw_chr,
      header_row = r,
      team_col = c,
      table_right = table_right,
      next_header_row = next_header_row,
      metric = metric,
      sheet_name = sheet_name
    )
  }

  result_list <- compact(result_list)

  if (length(result_list) == 0) {
    stop("No usable table in ", sheet_name)
  }

  bind_rows(result_list)
}


# Load the workbook

cat("\n============================================================\n")
cat("WORKBOOK READING\n")
cat("============================================================\n")
cat("File:", excel_file, "\n")
cat("Output:", output_dir, "\n\n")

sheets <- excel_sheets(excel_file)

data_sheets <- sheets[
  !str_to_lower(str_trim(sheets)) %in% c("global")
]

print(data_sheets)

parsed <- map(
  data_sheets,
  ~ parse_sheet(.x, excel_file)
)

data <- bind_rows(parsed)


# Prepare the variables used in the analysis

data <- data %>%
  mutate(
    league = clean_text(league),
    team = clean_text(team),
    season = clean_text(season),

    top_rank_numeric = num(
      str_extract(top_rank, "\\d+")
    ),

    rank_numeric = num(
      str_extract(rank, "\\d+")
    ),

    champion = str_detect(
      str_to_lower(rank),
      "champion"
    ),

    season_start = num(
      str_extract(season, "^[0-9]{4}")
    ),

    decade = case_when(
      !is.na(season_start) ~
        paste0(floor(season_start / 10) * 10, "s"),
      TRUE ~ NA_character_
    ),

    performance_id = paste(
      league,
      metric,
      team,
      season,
      sep = " | "
    )
  )


# Quick checks before calculating rankings

cat("\n============================================================\n")
cat("CHECKS\n")
cat("============================================================\n")

cat("\nTotal number of rows:", nrow(data), "\n")

cat("\nLeagues :\n")
print(
  data %>%
    count(league, sort = TRUE)
)

cat("\nPeriods :\n")
print(
  data %>%
    count(period, sort = TRUE)
)

cat("\nMetrics :\n")
print(
  data %>%
    count(metric, sort = TRUE)
)

# Consistency checks before analysis.
expected_metrics <- metric_order
missing_metrics <- setdiff(expected_metrics, unique(data$metric))
if (length(missing_metrics) > 0) {
  stop(
    "Metrics missing after parsing: ",
    paste(missing_metrics, collapse = ", ")
  )
}

unexpected_periods <- setdiff(unique(data$period), period_order)
if (length(unexpected_periods) > 0) {
  warning(
    "Non-standard periods detected: ",
    paste(unexpected_periods, collapse = ", ")
  )
}

# Export raw parsing check.
write_csv(
  data,
  file.path(tables_dir, "00_parsed_data.csv")
)


# Keep only the leagues included in the project

if (INCLUDE_HUNGARY) {

  individual_data <- data %>%
    filter(league %in% ALL_INDIVIDUAL_LEAGUES)

} else {

  individual_data <- data %>%
    filter(league %in% setdiff(ALL_INDIVIDUAL_LEAGUES, "Hungary"))
}

# Check: no unknown period should enter the analyses
# for the individual-league analyses.
bad_individual_periods <- individual_data %>%
  filter(is.na(period)) %>%
  distinct(league, sheet)

if (nrow(bad_individual_periods) > 0) {
  stop(
    "Unrecognized period in individual league sheets: ",
    paste(paste0(bad_individual_periods$league, " [", bad_individual_periods$sheet, "]"), collapse = "; ")
  )
}

# For leagues founded after 1955,
# "1955+" and "All Time" refer to exactly the same population.
# We therefore explicitly materialize the 1955+ period from All Time
# so that they appear in ALL heatmaps / comparisons.
PERIOD_1955_EQ_ALLTIME <- c("Bundesliga", "Eredivisie")

alltime_as_1955 <- individual_data %>%
  filter(league %in% PERIOD_1955_EQ_ALLTIME, period == "All Time") %>%
  mutate(period = "1955+")

individual_data <- bind_rows(
  individual_data,
  alltime_as_1955
) %>%
  distinct(league, period, metric, team, season, .keep_all = TRUE)

# Cross-league comparisons never use the PL 1955-1992 window.
# It remains available only for PL-specific analyses.

big5_data <- individual_data %>%
  filter(league %in% BIG5)

other_data <- individual_data %>%
  filter(league %in% OTHER_LEAGUES)


# RANKINGS, RECORDS AND RECURRENCE
# Two ways of counting appearances
# RAW :
#   a season present in 1955+ AND 1992+ counts as two appearances.
# UNIQUE :
#   Bayern 2012-13 present in several periods = one season.
# This is intentionally separated.

raw_appearances <- individual_data

unique_performances <- individual_data %>%
  distinct(
    league,
    metric,
    team,
    season,
    .keep_all = TRUE
  )


# Build comparable Top 10 rankings

# To compare periods consistently, we create Top 10s
# from the rank provided in the file, rather than mixing
# the four metrics.

standard_top10 <- individual_data %>%
  filter(
    top_rank_numeric <= TOP_N_STANDARD
  )

standard_top10_unique <- standard_top10 %>%
  distinct(
    league,
    period,
    metric,
    team,
    season,
    .keep_all = TRUE
  )

# Coverage check: each league must have usable Top 10s
# in each cross-league period.
# This check immediately detects a parsing error
# instead of producing silent gaps in the charts.
expected_top10_coverage <- tidyr::expand_grid(
  league = ALL_INDIVIDUAL_LEAGUES,
  period = INTER_PERIODS,
  metric = metric_order
)

top10_coverage_check <- standard_top10 %>%
  filter(period %in% INTER_PERIODS) %>%
  distinct(league, period, metric, team, season) %>%
  count(league, period, metric, name = "n_top10") %>%
  right_join(expected_top10_coverage, by = c("league", "period", "metric")) %>%
  mutate(n_top10 = dplyr::coalesce(n_top10, 0L)) %>%
  arrange(league, factor(period, levels = INTER_PERIODS), metric)

write_csv(
  top10_coverage_check,
  file.path(tables_dir, "00b_top10_coverage_check.csv")
)

bad_top10_coverage <- top10_coverage_check %>%
  filter(n_top10 == 0)

if (nrow(bad_top10_coverage) > 0) {
  stop(
    "Top 10 missing after parsing for: ",
    paste(
      paste(bad_top10_coverage$league, bad_top10_coverage$period, bad_top10_coverage$metric, sep = " | "),
      collapse = "; "
    )
  )
}


# Use larger rankings for the all-time comparisons

alltime_individual <- individual_data %>%
  filter(
    period == "All Time",
    top_rank_numeric <= TOP_N_ALLTIME_INDIVIDUAL
  )

# Top 20 specific to Top 5 Europe.
# It is handled separately because it is NOT a league.
top5_europe <- data %>%
  filter(
    league == "Top 5 Europe"
  )

top5_europe_standard <- top5_europe %>%
  filter(
    top_rank_numeric <= TOP_N_STANDARD
  )

top5_europe_alltime <- top5_europe %>%
  filter(
    period == "All Time",
    top_rank_numeric <= TOP_N_ALLTIME_BIG5
  )


# Helpers for selecting records

value_column <- function(metric) {

  case_when(
    metric %in% c(
      "points_non_champion",
      "points_general"
    ) ~ "ppm",

    metric == "goals_for" ~ "gpm_for",

    metric == "goals_against" ~ "gpm_against",

    TRUE ~ NA_character_
  )
}


add_value <- function(df) {

  df %>%
    mutate(
      value = case_when(
        metric %in% c(
          "points_non_champion",
          "points_general"
        ) ~ ppm,

        metric == "goals_for" ~ gpm_for,

        metric == "goals_against" ~ gpm_against,

        TRUE ~ NA_real_
      )
    )
}


# Find the best performance in each period

records_by_period <- individual_data %>%
  add_value() %>%
  filter(
    !is.na(value)
  ) %>%
  group_by(
    league,
    period,
    metric
  ) %>%
  summarise(
    record_value = if (first(metric) == "goals_against") {
      min(value, na.rm = TRUE)
    } else {
      max(value, na.rm = TRUE)
    },
    .groups = "drop"
  ) %>%
  left_join(
    individual_data %>%
      add_value() %>%
      select(
        league,
        period,
        metric,
        team,
        season,
        value,
        rank,
        champion
      ),
    by = c(
      "league",
      "period",
      "metric",
      "record_value" = "value"
    )
  ) %>%
  arrange(
    league,
    metric,
    factor(period, levels = period_order),
    desc(metric != "goals_against"),
    record_value
  )


write_csv(
  records_by_period,
  file.path(tables_dir, "01_records_by_period.csv")
)


# Compare the single record with the Top 10

top10_profile <- standard_top10 %>%
  add_value() %>%
  group_by(
    league,
    period,
    metric
  ) %>%
  summarise(
    n = n(),
    mean_value = safe_mean(value),
    median_value = safe_median(value),
    max_value = if (first(metric) == "goals_against") {
      min(value, na.rm = TRUE)
    } else {
      max(value, na.rm = TRUE)
    },
    min_value = if (first(metric) == "goals_against") {
      max(value, na.rm = TRUE)
    } else {
      min(value, na.rm = TRUE)
    },
    .groups = "drop"
  )

write_csv(
  top10_profile,
  file.path(tables_dir, "02_top10_profile_by_period.csv")
)


# Club recurrence: count every appearance

club_recurrence_raw <- individual_data %>%
  group_by(
    league,
    metric,
    team
  ) %>%
  summarise(
    appearances = n(),
    unique_seasons = n_distinct(season),
    periods_present = n_distinct(period),
    period_list = paste(
      sort(unique(period)),
      collapse = " / "
    ),
    .groups = "drop"
  ) %>%
  arrange(
    league,
    metric,
    desc(appearances),
    desc(unique_seasons)
  )

write_csv(
  club_recurrence_raw,
  file.path(tables_dir, "03_raw_club_recurrence.csv")
)


# Club recurrence: count unique seasons

club_recurrence_unique <- individual_data %>%
  distinct(
    league,
    metric,
    team,
    season,
    .keep_all = TRUE
  ) %>%
  group_by(
    league,
    metric,
    team
  ) %>%
  summarise(
    unique_seasons = n(),
    periods_present = n_distinct(period),
    .groups = "drop"
  ) %>%
  arrange(
    league,
    metric,
    desc(unique_seasons)
  )

write_csv(
  club_recurrence_unique,
  file.path(tables_dir, "04_unique_season_club_recurrence.csv")
)


# HISTORICAL PATTERNS AND COMPARISONS
# Check for repeated seasons across periods

period_overlap_check <- individual_data %>%
  group_by(
    league,
    metric,
    team,
    season
  ) %>%
  summarise(
    n_periods = n_distinct(period),
    periods = paste(
      sort(unique(period)),
      collapse = " / "
    ),
    .groups = "drop"
  ) %>%
  filter(
    n_periods > 1
  ) %>%
  arrange(
    desc(n_periods),
    league,
    metric,
    team,
    season
  )

write_csv(
  period_overlap_check,
  file.path(tables_dir, "05_performances_in_more_than_one_period.csv")
)


# Measure how much recurrence comes from overlapping periods

dup_rate <- individual_data %>%
  group_by(
    league,
    metric
  ) %>%
  summarise(
    raw_appearances = n(),
    unique_performances = n_distinct(performance_id),
    repeated_period_appearances =
      raw_appearances - unique_performances,
    repeated_share =
      repeated_period_appearances / raw_appearances,
    .groups = "drop"
  )

write_csv(
  dup_rate,
  file.path(tables_dir, "06_period_duplication_rate.csv")
)


# Compare recurrence across metrics
# This does NOT create a score.
# It only answers a descriptive question:
# In plain terms: which clubs show up in more than one of the four rankings?
# A season is counted only once within each metric.

club_metric_presence <- individual_data %>%
  distinct(
    league,
    team,
    metric,
    season
  ) %>%
  group_by(
    league,
    team
  ) %>%
  summarise(
    metrics_present = n_distinct(metric),
    metrics = paste(
      sort(unique(metric)),
      collapse = " + "
    ),
    unique_seasons = n_distinct(season),
    .groups = "drop"
  ) %>%
  arrange(
    league,
    desc(metrics_present),
    desc(unique_seasons)
  )

write_csv(
  club_metric_presence,
  file.path(tables_dir, "07_club_presence_across_4_metrics.csv")
)


# Decade patterns in the Top 10
# Here we compare performances within ONE ranking only.
# The 4 metrics remain separate.

decade_top10 <- standard_top10 %>%
  distinct(
    league,
    period,
    metric,
    team,
    season,
    .keep_all = TRUE
  ) %>%
  filter(
    !is.na(decade)
  ) %>%
  count(
    league,
    period,
    metric,
    decade,
    name = "performances"
  )

write_csv(
  decade_top10,
  file.path(tables_dir, "08_top10_decades.csv")
)


# Decade patterns in the all-time Top 15 and Top 20

decade_alltime <- bind_rows(
  alltime_individual,
  top5_europe_alltime
) %>%
  distinct(
    league,
    metric,
    team,
    season,
    .keep_all = TRUE
  ) %>%
  filter(
    !is.na(decade)
  ) %>%
  count(
    league,
    metric,
    decade,
    name = "performances"
  )

write_csv(
  decade_alltime,
  file.path(tables_dir, "09_alltime_decades.csv")
)


# How widely elite places are shared between clubs
# A raw share of ranking places would be misleading here:
# each league mechanically provides 10 places in a Top 10.
# We therefore measure Top 10 diversity here:
#   100 % = 10 seasons / clubs distincts
#   40% = 10 appearances concentrated among 4 distinct clubs
# This informs elite concentration, not average level.
# PL 1955-92 is excluded because this window does not exist for the other leagues.

league_top10_diversity <- standard_top10 %>%
  filter(period %in% INTER_PERIODS) %>%
  distinct(league, period, metric, team, season) %>%
  group_by(league, period, metric) %>%
  summarise(
    top10_appearances = n(),
    distinct_clubs = n_distinct(team),
    diversity = distinct_clubs / top10_appearances,
    .groups = "drop"
  )

write_csv(
  league_top10_diversity,
  file.path(tables_dir, "10_top10_diversity_leagues_periods.csv")
)

for (m in metric_order) {
  d <- league_top10_diversity %>%
    filter(metric == m) %>%
    mutate(
      period = factor(period, levels = INTER_PERIODS),
      league = factor(league, levels = rev(ALL_INDIVIDUAL_LEAGUES))
    )

  if (nrow(d) == 0) next

  p <- ggplot(d, aes(x = period, y = league, fill = diversity)) +
    geom_tile(colour = "white", linewidth = 0.8) +
    geom_text(aes(label = sprintf("%.0f%%", diversity * 100)), size = 3.4) +
    scale_fill_gradient(
      limits = c(0, 1),
      labels = scales::percent_format(accuracy = 1),
      name = "Top 10\ndiversity"
    ) +
    labs(
      title = paste("Club diversity in the Top 10s —", metric_labels[[m]]),
      subtitle = "% = distinct clubs among the 10 appearances in the period; the lower the rate, the more often the same clubs return.",
      x = "Period",
      y = NULL,
      caption = 'Cross-league comparison: "Since 3 points", "1955+" and "All Time". The PL 1955–1992 window is excluded.'
    ) +
    theme_portfolio() +
    theme(
      axis.text.y = element_text(size = 10),
      plot.caption = element_text(hjust = 0, size = 9, colour = "grey35")
    )

  ggsave(
    file.path(plots_dir, paste0("02_top10_diversity_", safe_file_name(m), ".png")),
    p,
    width = 11,
    height = 7,
    dpi = 300
  )
}

# Compare the Big Five with the other leagues

group_top10 <- standard_top10 %>%
  mutate(
    league_group = case_when(
      league %in% BIG5 ~ "Big 5",
      league %in% OTHER_LEAGUES ~ "Autres leagues",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(
    !is.na(league_group)
  ) %>%
  count(
    period,
    metric,
    league_group,
    name = "appearances"
  ) %>%
  group_by(
    period,
    metric
  ) %>%
  mutate(
    share = appearances / sum(appearances)
  ) %>%
  ungroup()

write_csv(
  group_top10,
  file.path(tables_dir, "11_big5_vs_other_leagues_top10.csv")
)


# Club share in the combined European ranking
# The "Top 5 EU" sheets are already a direct aggregation
# across the Big Five leagues.
# They are never mixed with the individual league sheets.

top5_club_recurrence <- top5_europe %>%
  group_by(
    metric,
    team
  ) %>%
  summarise(
    appearances = n(),
    unique_seasons = n_distinct(season),
    periods = n_distinct(period),
    .groups = "drop"
  ) %>%
  arrange(
    metric,
    desc(appearances),
    desc(unique_seasons)
  )

write_csv(
  top5_club_recurrence,
  file.path(tables_dir, "12_top5_europe_recurrence.csv")
)


# Identify the leading club in each period

leaders_top10 <- standard_top10 %>%
  add_value() %>%
  filter(
    top_rank_numeric == min(top_rank_numeric, na.rm = TRUE)
  ) %>%
  group_by(
    league,
    period,
    metric
  ) %>%
  summarise(
    first_rank = min(top_rank_numeric, na.rm = TRUE),
    record_value = if (first(metric) == "goals_against") {
      min(value, na.rm = TRUE)
    } else {
      max(value, na.rm = TRUE)
    },
    leaders = paste(
      unique(paste(team, season, sep = " — ")),
      collapse = " ; "
    ),
    .groups = "drop"
  )

write_csv(
  leaders_top10,
  file.path(tables_dir, "13_leaders_by_period.csv")
)


# Count clubs that return to the Top 10

club_top10_recurrence <- standard_top10 %>%
  distinct(
    league,
    period,
    metric,
    team,
    season,
    .keep_all = TRUE
  ) %>%
  count(
    league,
    metric,
    team,
    name = "top10_unique_seasons"
  ) %>%
  arrange(
    league,
    metric,
    desc(top10_unique_seasons)
  )

write_csv(
  club_top10_recurrence,
  file.path(tables_dir, "14_club_recurrence_top10.csv")
)


# Find clubs that appear in more than one period

club_period_persistence <- individual_data %>%
  distinct(
    league,
    metric,
    team,
    period,
    .keep_all = TRUE
  ) %>%
  count(
    league,
    metric,
    team,
    name = "n_periods"
  ) %>%
  arrange(
    league,
    metric,
    desc(n_periods),
    team
  )

write_csv(
  club_period_persistence,
  file.path(tables_dir, "15_club_persistence_between_periods.csv")
)


# Measure overlap between period rankings
# Jaccard = intersection / union.
# Calculated separately for each metric and league.

make_period_overlap <- function(df) {

  periods <- sort(unique(df$period))

  if (length(periods) < 2) {
    return(tibble())
  }

  pairs <- combn(periods, 2, simplify = FALSE)

  map_dfr(pairs, function(p) {

    a <- df %>%
      filter(period == p[1]) %>%
      distinct(team)

    b <- df %>%
      filter(period == p[2]) %>%
      distinct(team)

    inter <- intersect(a$team, b$team)
    uni <- union(a$team, b$team)

    tibble(
      period_1 = p[1],
      period_2 = p[2],
      intersection = length(inter),
      union = length(uni),
      jaccard = if_else(
        length(uni) > 0,
        length(inter) / length(uni),
        NA_real_
      ),
      common_clubs = paste(
        inter,
        collapse = " ; "
      )
    )
  })
}

period_overlap_clubs <- individual_data %>%
  group_by(
    league,
    metric
  ) %>%
  group_modify(
    ~ make_period_overlap(.x)
  ) %>%
  ungroup()

write_csv(
  period_overlap_clubs,
  file.path(tables_dir, "16_club_overlap_between_periods.csv")
)


# Look at points and goals together
# This is the only direct relationship analysis between two metrics,
# and it is presented as a descriptive relationship, not as
# un score de performance.
# Goals conceded remain completely separate.

points_vs_goals <- individual_data %>%
  filter(
    metric == "points_general"
  ) %>%
  select(
    league,
    period,
    team,
    season,
    ppm = ppm
  ) %>%
  inner_join(
    individual_data %>%
      filter(
        metric == "goals_for"
      ) %>%
      select(
        league,
        period,
        team,
        season,
        gpm_for = gpm_for
      ),
    by = c(
      "league",
      "period",
      "team",
      "season"
    )
  ) %>%
  filter(
    is.finite(ppm),
    is.finite(gpm_for)
  )

cor_points_goals <- points_vs_goals %>%
  group_by(
    league,
    period
  ) %>%
  summarise(
    n = n(),
    correlation = safe_cor(ppm, gpm_for),
    .groups = "drop"
  )

write_csv(
  points_vs_goals,
  file.path(tables_dir, "17_points_vs_goals_scored_data.csv")
)

write_csv(
  cor_points_goals,
  file.path(tables_dir, "18_points_goals_correlation.csv")
)


# Compare champions with non-champions using points only

champion_nonchampion <- individual_data %>%
  filter(
    metric == "points_general"
  ) %>%
  mutate(
    status = if_else(
      champion,
      "Champion",
      "Non-champion"
    )
  )

champion_nonchampion_summary <- champion_nonchampion %>%
  group_by(
    league,
    period,
    status
  ) %>%
  summarise(
    n = n(),
    mean_ppm = safe_mean(ppm),
    median_ppm = safe_median(ppm),
    .groups = "drop"
  )

write_csv(
  champion_nonchampion_summary,
  file.path(tables_dir, "19_champions_vs_nonchampions_points.csv")
)


# ANALYSIS TABLES AND VISUAL SUMMARIES
# Build the analysis tables and chart-ready summaries
# Analytical flow:
#   A. Records: who holds the extremes, by league and era?
#   B. Top 10 depth: is the record isolated or accompanied?
#   C. Comparisons: Big 5 against each other, then Big 5 vs other leagues.
#   D. Recurrence: which clubs return within ONE ranking?
#   E. Time: which decades concentrate performances?
#   F. Profiles: points, attack and defense remain three separate stories.
# I keep the metrics separate rather than creating an overall score.
# No chart combines the four rankings into a single metric.

# ------------------------------------------------------------
# Chart theme and scales
# ------------------------------------------------------------

theme_portfolio <- function() {
  theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(face = "bold", size = 17, colour = "#171717", margin = margin(b = 5)),
      plot.subtitle = element_text(size = 11, colour = "#555555", margin = margin(b = 10)),
      plot.caption = element_text(size = 9.5, colour = "#666666", hjust = 0, margin = margin(t = 10)),
      axis.title = element_text(face = "bold", colour = "#222222"),
      axis.text = element_text(colour = "#333333"),
      legend.position = "bottom",
      legend.title = element_text(face = "bold"),
      legend.text = element_text(size = 10),
      panel.grid.major.y = element_blank(),
      panel.grid.major.x = element_line(colour = "#E7E7E7", linewidth = 0.45),
      panel.grid.minor = element_blank(),
      panel.background = element_rect(fill = "#FFFFFF", colour = NA),
      plot.background = element_rect(fill = "#FFFFFF", colour = NA),
      strip.background = element_rect(fill = "#F2F2F2", colour = NA),
      strip.text = element_text(face = "bold", colour = "#333333", margin = margin(5, 8, 5, 8)),
      panel.spacing = unit(1.1, "lines"),
      plot.margin = margin(14, 30, 14, 14)
    )
}

metric_direction_text <- function(metric_name) {
  if (metric_name == "goals_against") {
    "For defense, a lower value corresponds to better performance."
  } else {
    "A higher value corresponds to better performance."
  }
}

metric_best <- function(x, metric_name) {
  if (metric_name == "goals_against") {
    min(x, na.rm = TRUE)
  } else {
    max(x, na.rm = TRUE)
  }
}

metric_worst <- function(x, metric_name) {
  if (metric_name == "goals_against") {
    max(x, na.rm = TRUE)
  } else {
    min(x, na.rm = TRUE)
  }
}

# ------------------------------------------------------------
# Reference table of records by league and period
# ------------------------------------------------------------

record_matrix <- individual_data %>%
  filter(period %in% INTER_PERIODS) %>%
  add_value() %>%
  filter(is.finite(value)) %>%
  group_by(league, period, metric) %>%
  summarise(
    record_value = metric_best(value, first(metric)),
    .groups = "drop"
  )

record_holders <- individual_data %>%
  filter(period %in% INTER_PERIODS) %>%
  add_value() %>%
  filter(is.finite(value)) %>%
  group_by(league, period, metric) %>%
  mutate(
    record_value = metric_best(value, first(metric))
  ) %>%
  filter(value == record_value) %>%
  summarise(
    record_holders = paste(unique(paste(team, season, sep = " — ")), collapse = " ; "),
    n_holders = n_distinct(paste(team, season)),
    .groups = "drop"
  )

records_comparison <- record_matrix %>%
  left_join(record_holders, by = c("league", "period", "metric")) %>%
  arrange(metric, factor(period, levels = period_order), league)

write_csv(records_comparison, file.path(tables_dir, "23_records_comparison_leagues_periods.csv"))

# ------------------------------------------------------------
# Top 10 profile: average, median and spread
# ------------------------------------------------------------

# Here we compare elite depth, not the average league level.
top10_stats <- standard_top10 %>%
  add_value() %>%
  filter(is.finite(value)) %>%
  group_by(league, period, metric) %>%
  summarise(
    n = n(),
    mean = mean(value),
    median = median(value),
    q1 = quantile(value, 0.25, names = FALSE),
    q3 = quantile(value, 0.75, names = FALSE),
    sd = if (n() > 1) sd(value) else NA_real_,
    best = metric_best(value, first(metric)),
    worst = metric_worst(value, first(metric)),
    spread_best_to_10th = if (first(metric) == "goals_against") {
      max(value) - min(value)
    } else {
      max(value) - min(value)
    },
    .groups = "drop"
  )

write_csv(top10_stats, file.path(tables_dir, "24_top10_depth_leagues_periods.csv"))

# The gap between the single best result and the Top 10 average.
top10_record_gap <- top10_stats %>%
  mutate(
    gap_record_mean = if_else(
      metric == "goals_against",
      mean - best,
      best - mean
    ),
    gap_record_mean_relative = if_else(
      mean != 0,
      gap_record_mean / abs(mean),
      NA_real_
    )
  )

write_csv(top10_record_gap, file.path(tables_dir, "25_record_gap_top10_mean.csv"))

# ------------------------------------------------------------
# Compare the Big Five with one another and with other leagues
# ------------------------------------------------------------

league_groups <- tibble(
  league = ALL_INDIVIDUAL_LEAGUES,
  league_group = case_when(
    league %in% BIG5 ~ "Big 5",
    league %in% OTHER_LEAGUES ~ "Other leagues",
    TRUE ~ NA_character_
  )
)

top10_group_comparison <- standard_top10 %>%
  filter(period %in% INTER_PERIODS) %>%
  add_value() %>%
  filter(is.finite(value)) %>%
  left_join(league_groups, by = "league") %>%
  group_by(league_group, period, metric) %>%
  summarise(
    appearances = n(),
    mean_value = mean(value),
    median_value = median(value),
    .groups = "drop"
  )

write_csv(top10_group_comparison, file.path(tables_dir, "26_big5_vs_other_top10_depth.csv"))

# Specific comparison of the five major leagues: each league remains visible.
big5_top10_stats <- top10_stats %>%
  filter(league %in% BIG5, period %in% INTER_PERIODS)

write_csv(big5_top10_stats, file.path(tables_dir, "27_big5_internal_top10_comparison.csv"))

# ------------------------------------------------------------
# Recurrence within each ranking
# ------------------------------------------------------------

# Here, each team-season counts once; only the three shared periods are used.
# comparable, so PL 1955-1992 is never included here.
club_top10_recurrence_v6 <- standard_top10 %>%
  filter(period %in% INTER_PERIODS) %>%
  distinct(league, period, metric, team, season) %>%
  group_by(league, metric, team) %>%
  summarise(
    top10_seasons = n_distinct(season),
    top10_periods = n_distinct(period),
    periods = paste(sort(unique(period)), collapse = " / "),
    .groups = "drop"
  ) %>%
  arrange(league, metric, desc(top10_seasons), desc(top10_periods), team)

write_csv(club_top10_recurrence_v6, file.path(tables_dir, "28_club_recurrence_by_ranking.csv"))

# True persistence across different periods, but always for
# ONE ranking at a time.
club_top10_period_persistence <- standard_top10 %>%
  filter(period %in% INTER_PERIODS) %>%
  distinct(league, period, metric, team, season) %>%
  group_by(league, metric, team) %>%
  summarise(
    n_periods = n_distinct(period),
    periods = paste(sort(unique(period)), collapse = " / "),
    seasons = n_distinct(season),
    .groups = "drop"
  ) %>%
  arrange(league, metric, desc(n_periods), desc(seasons), team)

write_csv(club_top10_period_persistence, file.path(tables_dir, "29_club_persistence_top10_between_periods.csv"))

# ------------------------------------------------------------
# Overlap between periods, using the Top 10 only
# ------------------------------------------------------------

make_top10_period_overlap <- function(df) {
  periods <- sort(unique(df$period))
  if (length(periods) < 2) return(tibble())

  pairs <- combn(periods, 2, simplify = FALSE)

  map_dfr(pairs, function(p) {
    a <- df %>% filter(period == p[1]) %>% distinct(team)
    b <- df %>% filter(period == p[2]) %>% distinct(team)
    inter <- intersect(a$team, b$team)
    uni <- union(a$team, b$team)

    tibble(
      period_1 = p[1],
      period_2 = p[2],
      intersection = length(inter),
      union = length(uni),
      jaccard = if (length(uni) > 0) length(inter) / length(uni) else NA_real_,
      common_clubs = paste(inter, collapse = " ; ")
    )
  })
}

top10_period_overlap <- standard_top10 %>%
  filter(period %in% INTER_PERIODS) %>%
  distinct(league, period, metric, team, season) %>%
  group_by(league, metric) %>%
  group_modify(~ make_top10_period_overlap(.x)) %>%
  ungroup()

write_csv(top10_period_overlap, file.path(tables_dir, "30_top10_overlap_between_periods.csv"))

# ------------------------------------------------------------
# Decade concentration in the rankings
# ------------------------------------------------------------

decade_top10_v6 <- standard_top10 %>%
  distinct(league, period, metric, team, season, decade) %>%
  filter(!is.na(decade)) %>%
  count(league, period, metric, decade, name = "performances") %>%
  group_by(league, period, metric) %>%
  mutate(share = performances / sum(performances)) %>%
  ungroup()

write_csv(decade_top10_v6, file.path(tables_dir, "31_top10_decades_detail.csv"))

# All Time: Top 15 individuels / Top 20 Top 5 Europe.
decade_alltime_v6 <- bind_rows(
  alltime_individual,
  top5_europe_alltime
) %>%
  distinct(league, metric, team, season, decade) %>%
  filter(!is.na(decade)) %>%
  count(league, metric, decade, name = "performances") %>%
  group_by(league, metric) %>%
  mutate(share = performances / sum(performances)) %>%
  ungroup()

write_csv(decade_alltime_v6, file.path(tables_dir, "32_alltime_decades_detail.csv"))

# ------------------------------------------------------------
# Combined Top Five Europe ranking
# ------------------------------------------------------------

top5_europe_stats <- top5_europe %>%
  add_value() %>%
  filter(is.finite(value)) %>%
  group_by(period, metric) %>%
  summarise(
    n = n(),
    mean = mean(value),
    median = median(value),
    best = metric_best(value, first(metric)),
    worst = metric_worst(value, first(metric)),
    .groups = "drop"
  )

write_csv(top5_europe_stats, file.path(tables_dir, "33_top5_europe_profile_by_period.csv"))

# Recurrence dans le Top 5 Europe, par ranking.
top5_club_recurrence_v6 <- top5_europe %>%
  distinct(period, metric, team, season) %>%
  group_by(metric, team) %>%
  summarise(
    top10_seasons = n_distinct(season),
    periods = n_distinct(period),
    .groups = "drop"
  ) %>%
  arrange(metric, desc(top10_seasons), desc(periods), team)

write_csv(top5_club_recurrence_v6, file.path(tables_dir, "34_top5_europe_club_recurrence.csv"))

# ------------------------------------------------------------
# Points and goals: describe the relationship without combining them
# ------------------------------------------------------------

points_vs_goals_v6 <- individual_data %>%
  filter(metric == "points_general") %>%
  select(league, period, team, season, ppm) %>%
  inner_join(
    individual_data %>%
      filter(metric == "goals_for") %>%
      select(league, period, team, season, gpm_for),
    by = c("league", "period", "team", "season")
  ) %>%
  filter(is.finite(ppm), is.finite(gpm_for))

cor_points_goals_v6 <- points_vs_goals_v6 %>%
  group_by(league, period) %>%
  summarise(
    n = n(),
    correlation = safe_cor(ppm, gpm_for),
    .groups = "drop"
  )

write_csv(points_vs_goals_v6, file.path(tables_dir, "35_points_vs_goals_data.csv"))
write_csv(cor_points_goals_v6, file.path(tables_dir, "36_points_goals_correlations.csv"))

# ------------------------------------------------------------
# Champions and non-champions: points comparison
# ------------------------------------------------------------

champion_nonchampion_v6 <- individual_data %>%
  filter(metric == "points_general") %>%
  mutate(status = if_else(champion, "Champion", "Non-champion"))

champion_nonchampion_summary_v6 <- champion_nonchampion_v6 %>%
  group_by(league, period, status) %>%
  summarise(
    n = n(),
    mean_ppm = safe_mean(ppm),
    median_ppm = safe_median(ppm),
    .groups = "drop"
  )

write_csv(champion_nonchampion_summary_v6, file.path(tables_dir, "37_champions_vs_nonchampions_points.csv"))

# CHARTS
# Charts: records by league and period

plot_record_heatmap <- function(metric_name) {
  df <- records_comparison %>%
    filter(metric == metric_name, period %in% INTER_PERIODS) %>%
    mutate(
      period = factor(period, levels = INTER_PERIODS),
      league = factor(league, levels = rev(ALL_INDIVIDUAL_LEAGUES)),
      heat_score = if (metric_name == "goals_against") -record_value else record_value
    )
  if (nrow(df) == 0) return(NULL)

  fill_scale <- if (metric_name == "goals_against") {
    scale_fill_gradient(
      low = "white",
      high = "#1F77B4",
      breaks = scales::breaks_pretty(n = 4),
      labels = function(x) sprintf("%.2f", abs(x)),
      guide = guide_colorbar(barheight = grid::unit(70, "pt"))
    )
  } else {
    scale_fill_gradient(
      low = "white",
      high = "#1F77B4",
      breaks = scales::breaks_pretty(n = 5),
      guide = guide_colorbar(barheight = grid::unit(70, "pt"))
    )
  }

  ggplot(df, aes(x = period, y = league, fill = heat_score)) +
    geom_tile(colour = "white", linewidth = 0.7) +
    geom_text(aes(label = sprintf("%.2f", record_value)), size = 3.2) +
    labs(
      title = paste("Records by league and period —", metric_labels[[metric_name]]),
      subtitle = metric_direction_text(metric_name),
      x = "Period",
      y = NULL,
      fill = "Performance\n(better = darker)"
    ) +
    fill_scale +
    theme_portfolio()
}

for (m in metric_order) {
  p <- plot_record_heatmap(m)
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("01_records_leagues_periods_", safe_file_name(m), ".png")), p,
           width = 12, height = 7, dpi = 300)
  }
}

# Charts: depth of the Top 10

plot_top10_mean <- function(metric_name, data_frame = top10_stats, suffix = "") {
  df <- data_frame %>% filter(metric == metric_name, period %in% INTER_PERIODS)
  if (nrow(df) == 0) return(NULL)

  ggplot(df, aes(x = factor(league, levels = rev(ALL_INDIVIDUAL_LEAGUES)), y = mean, fill = league)) +
    geom_col(width = 0.68) +
    geom_point(aes(y = best), size = 2.8, colour = "black") +
    facet_wrap(~ period, ncol = 1) +
    scale_fill_manual(values = league_colors, drop = FALSE) +
    labs(
      title = paste("Top 10 depth —", metric_labels[[metric_name]]),
      subtitle = "Bar = Top 10 mean; black dot = period record",
      x = NULL,
      y = "Ratio",
      fill = "League"
    ) +
    coord_flip() +
    theme_portfolio() +
    theme(legend.position = "none")
}

for (m in metric_order) {
  p <- plot_top10_mean(m)
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("02_profondeur_top10_", safe_file_name(m), ".png")), p,
           width = 12, height = 12, dpi = 300, limitsize = FALSE)
  }
}

plot_top10_distribution <- function(metric_name, df_source = standard_top10, suffix = "") {
  df <- df_source %>% filter(metric == metric_name, period %in% INTER_PERIODS) %>% add_value() %>% filter(is.finite(value))
  if (nrow(df) == 0) return(NULL)

  ggplot(df, aes(x = fct_reorder(league, value, .fun = median), y = value, fill = league)) +
    geom_boxplot(alpha = 0.75, outlier.shape = NA) +
    facet_wrap(~ period, ncol = 1) +
    scale_fill_manual(values = league_colors, drop = FALSE) +
    coord_flip() +
    labs(
      title = paste("Top 10 distribution —", metric_labels[[metric_name]]),
      subtitle = "Each box shows the 10 performances in the ranking; the four metrics remain separate",
      x = NULL,
      y = "Ratio",
      fill = "League"
    ) +
    theme_portfolio() +
    theme(legend.position = "none")
}

for (m in metric_order) {
  p <- plot_top10_distribution(m)
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("03_distribution_top10_", safe_file_name(m), ".png")), p,
           width = 12, height = 12, dpi = 300, limitsize = FALSE)
  }
}

# Charts: comparisons within the Big Five

for (m in metric_order) {
  p <- plot_top10_distribution(m, standard_top10 %>% filter(league %in% BIG5), "_big5")
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("04_big5_top10_distribution_", safe_file_name(m), ".png")), p,
           width = 12, height = 10, dpi = 300, limitsize = FALSE)
  }
}

# Top 10 mean heatmap: quick Big 5 vs other leagues comparison.
plot_top10_mean_heatmap <- function(df_source, filename, title_text) {
  df <- df_source %>%
    select(league, period, metric, mean) %>%
    filter(period %in% INTER_PERIODS, is.finite(mean))
  if (nrow(df) == 0) return(invisible(NULL))

  for (m in metric_order) {
    d <- df %>% filter(metric == m)
    if (nrow(d) == 0) next

    d <- d %>%
      mutate(
        league = factor(league, levels = rev(ALL_INDIVIDUAL_LEAGUES)),
        heat_score = if (first(metric) == "goals_against") -mean else mean
      )

    p <- ggplot(d, aes(x = factor(period, levels = INTER_PERIODS), y = league, fill = heat_score)) +
      geom_tile(colour = "white") +
      geom_text(aes(label = sprintf("%.2f", mean)), size = 3.2) +
      labs(title = paste(title_text, "—", metric_labels[[m]]),
           subtitle = "Mean of the 10 best performances for each league and period",
           x = "Period", y = NULL, fill = "Performance\n(better = darker)") +
      scale_fill_gradient(low = "white", high = "#1F77B4") +
      theme_portfolio()

    ggsave(file.path(plots_dir, paste0(filename, "_", safe_file_name(m), ".png")), p,
           width = 12, height = 7, dpi = 300)
  }
}

plot_top10_mean_heatmap(top10_stats, "05_all_leagues_top10_depth", "Comparison of Top 10 depth across all leagues")
plot_top10_mean_heatmap(big5_top10_stats, "06_big5_internal_comparison", "Big 5 internal comparison")

# Charts: Big Five versus other leagues

for (m in metric_order) {
  d <- top10_group_comparison %>% filter(metric == m)
  if (nrow(d) == 0) next

  p <- ggplot(d, aes(x = factor(period, levels = period_order), y = mean_value, fill = league_group)) +
    geom_col(position = position_dodge(width = 0.72), width = 0.65) +
    labs(title = paste("Big 5 vs other leagues —", metric_labels[[m]]),
         subtitle = "Mean of performances appearing in the Top 10s",
         x = "Period", y = "Mean ratio", fill = NULL) +
    theme_portfolio()

  ggsave(file.path(plots_dir, paste0("07_big5_vs_others_", safe_file_name(m), ".png")), p,
         width = 11, height = 7, dpi = 300)
}

# Charts: club recurrence

period_scope_label <- paste(INTER_PERIODS, collapse = " • ")
period_scope_note <- paste0(
  "Periods included: ", period_scope_label,
  '. For Bundesliga and Eredivisie, "1955+" = "All Time" because these leagues started after 1955.\nThe PL 1955–1992 window is excluded.\n')

plot_recurrence <- function(metric_name) {
  d <- club_top10_recurrence_v6 %>%
    filter(metric == metric_name) %>%
    group_by(league) %>%
    slice_max(top10_seasons, n = 5, with_ties = TRUE) %>%
    ungroup()
  if (nrow(d) == 0) return(NULL)

  team_cols <- club_colors(d$team)
  ggplot(d, aes(x = top10_seasons, y = fct_reorder(team, top10_seasons), fill = team)) +
    geom_col(width = 0.68) +
    geom_text(aes(label = top10_seasons), hjust = -0.15, size = 3) +
    facet_wrap(~ league, scales = "free_y") +
    scale_fill_manual(values = team_cols, drop = FALSE) +
    scale_x_continuous(expand = expansion(mult = c(0, 0.18))) +
    coord_cartesian(clip = "off") +
    labs(title = paste("Recurrence within the same Top 10 —", metric_labels[[metric_name]]),
         subtitle = "Number of unique seasons",
         caption = period_scope_note,
         x = "Seasons in the Top 10", y = NULL, fill = "Club") +
    theme_portfolio() + theme(legend.position = "none")
}

for (m in metric_order) {
  p <- plot_recurrence(m)
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("08_club_recurrence_", safe_file_name(m), ".png")), p,
           width = 14, height = 11, dpi = 300, limitsize = FALSE)
  }
}

# Persistence between periods, but always metric by metric.
plot_persistence <- function(metric_name) {
  d <- club_top10_period_persistence %>%
    filter(metric == metric_name, n_periods > 1) %>%
    group_by(league) %>%
    slice_max(n_periods, n = 5, with_ties = TRUE) %>%
    ungroup()
  if (nrow(d) == 0) return(NULL)

  team_cols <- club_colors(d$team)
  ggplot(d, aes(x = n_periods, y = fct_reorder(team, n_periods), fill = team)) +
    geom_col(width = 0.68) +
    geom_text(aes(label = n_periods), hjust = -0.15, size = 3) +
    facet_wrap(~ league, scales = "free_y") +
    scale_fill_manual(values = team_cols, drop = FALSE) +
    scale_x_continuous(breaks = seq_len(max(d$n_periods)), expand = expansion(mult = c(0, 0.18))) +
    coord_cartesian(clip = "off") +
    labs(title = paste("Clubs appearing across multiple periods —", metric_labels[[metric_name]]),
         subtitle = "Presence across multiple periods of the same ranking only",
         caption = period_scope_note,
         x = "Number of periods", y = NULL, fill = "Club") +
    theme_portfolio() + theme(legend.position = "none")
}

for (m in metric_order) {
  p <- plot_persistence(m)
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("09_club_persistence_", safe_file_name(m), ".png")), p,
           width = 14, height = 10, dpi = 300, limitsize = FALSE)
  }
}

# Charts: decade patterns

plot_decade_concentration <- function(metric_name, alltime = FALSE) {
  d <- if (alltime) decade_alltime_v6 %>% filter(metric == metric_name) else decade_top10_v6 %>% filter(metric == metric_name)
  if (nrow(d) == 0) return(NULL)

  facet_variable <- if (alltime) "league" else "period"

  d <- d %>%
    mutate(facet_value = .data[[facet_variable]])

  ggplot(d, aes(x = decade, y = performances, fill = league)) +
    geom_col(position = "stack") +
    facet_wrap(~ facet_value, scales = "free_x") +
    scale_fill_manual(values = league_colors) +
    labs(title = paste("Concentration by decade —", metric_labels[[metric_name]]),
         subtitle = if (alltime) "All Time Top: performance concentration by decade" else "Number of Top 10 performances by decade",
         x = "Decade", y = "Number of performances", fill = "League") +
    theme_portfolio()
}

for (m in metric_order) {
  p <- plot_decade_concentration(m, FALSE)
  if (!is.null(p)) ggsave(file.path(plots_dir, paste0("10_top10_decades_", safe_file_name(m), ".png")), p,
                          width = 13, height = 8, dpi = 300)

  p2 <- plot_decade_concentration(m, TRUE)
  if (!is.null(p2)) ggsave(file.path(plots_dir, paste0("11_alltime_decades_", safe_file_name(m), ".png")), p2,
                           width = 14, height = 9, dpi = 300)
}


# ------------------------------------------------------------
# Club concentration by decade
# ------------------------------------------------------------

club_decade_top10 <- standard_top10 %>%
  distinct(league, period, metric, team, season, decade) %>%
  filter(!is.na(decade)) %>%
  count(league, metric, team, decade, name = "top10_seasons")

write_csv(club_decade_top10, file.path(tables_dir, "41_club_recurrence_by_decade.csv"))

plot_club_decade_heatmap <- function(metric_name) {
  base <- club_top10_recurrence_v6 %>%
    filter(metric == metric_name) %>%
    group_by(team) %>%
    summarise(total = sum(top10_seasons), .groups = "drop") %>%
    slice_max(total, n = 15, with_ties = TRUE)

  d <- club_decade_top10 %>%
    filter(metric == metric_name, team %in% base$team) %>%
    group_by(team, decade) %>%
    summarise(top10_seasons = sum(top10_seasons), .groups = "drop")

  if (nrow(d) == 0) return(NULL)

  d <- d %>%
    complete(team = base$team, decade, fill = list(top10_seasons = 0))

  ggplot(d, aes(x = decade, y = fct_reorder(team, top10_seasons, .fun = sum), fill = top10_seasons)) +
    geom_tile(colour = "white") +
    geom_text(aes(label = if_else(top10_seasons > 0, as.character(top10_seasons), "")), size = 3) +
    scale_fill_gradient(low = "grey92", high = "#1F77B4") +
    labs(title = paste("Club recurrence by decade —", metric_labels[[metric_name]]),
         subtitle = "Top 15 clubs by total number of seasons appearing in this ranking's Top 10",
         x = "Decade", y = NULL, fill = "Top 10 seasons") +
    theme_portfolio()
}

for (m in metric_order) {
  p <- plot_club_decade_heatmap(m)
  if (!is.null(p)) {
    ggsave(file.path(plots_dir, paste0("11b_club_recurrence_decades_", safe_file_name(m), ".png")), p,
           width = 14, height = 9, dpi = 300, limitsize = FALSE)
  }
}

# Charts: all-time Top Five Europe

plot_top5_alltime <- function(metric_name) {
  d <- top5_europe_alltime %>% filter(metric == metric_name) %>% add_value() %>% filter(is.finite(value))
  if (nrow(d) == 0) return(NULL)

  d <- d %>%
    mutate(label = paste(team, season, sep = " — "),
           plot_value = if_else(metric == "goals_against", -value, value))

  team_cols <- club_colors(d$team)
  ggplot(d, aes(x = value, y = fct_reorder(label, plot_value), fill = team)) +
    geom_col(width = 0.65) +
    geom_text(aes(label = sprintf("%.2f", value)), hjust = -0.12, size = 2.8) +
    scale_fill_manual(values = team_cols, drop = FALSE) +
    scale_x_continuous(expand = expansion(mult = c(0, 0.15))) +
    coord_cartesian(clip = "off") +
    labs(title = paste("Top 20 All Time — Top 5 Europe —", metric_labels[[metric_name]]),
         subtitle = "Aggregated ranking provided by the Top 5 Europe sheet; color = club, with a stable color for each team",
         x = "Ratio", y = NULL, fill = "Club") +
    theme_portfolio() + theme(legend.position = "none")
}

for (m in metric_order) {
  p <- plot_top5_alltime(m)
  if (!is.null(p)) ggsave(file.path(plots_dir, paste0("12_top20_alltime_top5_europe_", safe_file_name(m), ".png")), p,
                          width = 14, height = 12, dpi = 300, limitsize = FALSE)
}


# Top 20 across the eight leagues and three periods
# The Top 8 Europe pool is built from the eight individual
# domestic leagues in the project. The PL 1955-1992 window is
# intentionally excluded from cross-league comparisons.
# The same four indicators are ranked separately.
# Outputs are produced for the three comparable periods:
# All Time, 1955+ and Since 3 points.
# The purpose is to create a Top 8 Europe perimeter that can be
# compared directly with the existing Top 5 Europe outputs,
# without creating a composite score.

TOP8_PERIODS <- c(
  "All Time",
  "1955+",
  "Depuis 3 points"
)

TOP8_EUROPE_LEAGUES <- ALL_INDIVIDUAL_LEAGUES

# Keep the raw performances from the eight domestic leagues.
# The PL 1955-1992 window is excluded because there is no
# equivalent comparison window for the other championships.
top8_europe <- individual_data %>%
  filter(
    league %in% TOP8_EUROPE_LEAGUES,
    period %in% TOP8_PERIODS
  ) %>%
  distinct(
    league,
    period,
    metric,
    team,
    season,
    .keep_all = TRUE
  )

plot_top8_period <- function(metric_name, period_name) {
  d <- top8_europe %>%
    filter(
      metric == metric_name,
      period == period_name
    ) %>%
    add_value() %>%
    filter(is.finite(value)) %>%
    arrange(
      if (metric_name == "goals_against") value else desc(value)
    ) %>%
    slice_head(n = 20)

  if (nrow(d) == 0) return(NULL)

  d <- d %>%
    mutate(
      label = paste(team, season, sep = " — "),
      plot_value = if_else(metric == "goals_against", -value, value)
    )

  team_cols <- club_colors(d$team)

  ggplot(
    d,
    aes(
      x = value,
      y = fct_reorder(label, plot_value),
      fill = team
    )
  ) +
    geom_col(width = 0.65) +
    geom_text(
      aes(label = sprintf("%.2f", value)),
      hjust = -0.12,
      size = 2.8
    ) +
    scale_fill_manual(values = team_cols, drop = FALSE) +
    scale_x_continuous(expand = expansion(mult = c(0, 0.15))) +
    coord_cartesian(clip = "off") +
    labs(
      title = paste(
        "Top 20 — Top 8 Europe —",
        period_name,
        "—",
        metric_labels[[metric_name]]
      ),
      subtitle = "Top 20 performances across the eight domestic leagues; color = club",
      x = "Ratio",
      y = NULL,
      fill = "Club"
    ) +
    theme_portfolio() +
    theme(legend.position = "none")
}

for (p_name in TOP8_PERIODS) {
  for (m in metric_order) {
    p <- plot_top8_period(m, p_name)
    if (!is.null(p)) {
      ggsave(
        file.path(
          plots_dir,
          paste0(
            "12_top20_top8_europe_",
            safe_file_name(p_name), "_",
            safe_file_name(m), ".png"
          )
        ),
        p,
        width = 14,
        height = 12,
        dpi = 300,
        limitsize = FALSE
      )
    }
  }
}

# Export the rankings for direct comparison with the Top 5 Europe.
top8_europe_top20 <- map_dfr(
  TOP8_PERIODS,
  function(p_name) {
    map_dfr(
      metric_order,
      function(m) {
        top8_europe %>%
          filter(metric == m, period == p_name) %>%
          add_value() %>%
          filter(is.finite(value)) %>%
          arrange(if (m == "goals_against") value else desc(value)) %>%
          slice_head(n = 20) %>%
          mutate(period = p_name)
      }
    )
  }
)

write_csv(
  top8_europe_top20,
  file.path(tables_dir, "42_top20_top8_europe_3_periods.csv")
)

# Charts: football profiles

if (nrow(points_vs_goals_v6) > 0) {
  team_cols <- club_colors(points_vs_goals_v6$team)
  p <- ggplot(points_vs_goals_v6, aes(x = gpm_for, y = ppm, colour = team)) +
    geom_point(alpha = 0.70, size = 2.5) +
    facet_wrap(~ period, scales = "free") +
    scale_colour_manual(values = team_cols, drop = FALSE) +
    labs(title = "Results and attacking output",
         subtitle = "Each point = a team/season appearing in both rankings; descriptive relationship",
         x = "Goals scored / match", y = "Points / match", colour = "Club") +
    theme_portfolio() + theme(legend.position = "none")
  ggsave(file.path(plots_dir, "13_points_vs_goals_scored.png"), p, width = 13, height = 8, dpi = 300)
}

# Defense: points / goals conceded relationship, separate from attack.
points_vs_goals_against <- individual_data %>%
  filter(metric == "points_general") %>%
  select(league, period, team, season, ppm) %>%
  inner_join(
    individual_data %>% filter(metric == "goals_against") %>% select(league, period, team, season, gpm_against),
    by = c("league", "period", "team", "season")
  ) %>%
  filter(is.finite(ppm), is.finite(gpm_against))

cor_points_goals_against <- points_vs_goals_against %>%
  group_by(league, period) %>%
  summarise(n = n(), correlation = safe_cor(ppm, gpm_against), .groups = "drop")

write_csv(points_vs_goals_against, file.path(tables_dir, "38_points_vs_goals_conceded_data.csv"))
write_csv(cor_points_goals_against, file.path(tables_dir, "39_points_goals_conceded_correlations.csv"))

if (nrow(points_vs_goals_against) > 0) {
  team_cols <- club_colors(points_vs_goals_against$team)
  p <- ggplot(points_vs_goals_against, aes(x = gpm_against, y = ppm, colour = team)) +
    geom_point(alpha = 0.70, size = 2.5) +
    facet_wrap(~ period, scales = "free") +
    scale_colour_manual(values = team_cols, drop = FALSE) +
    labs(title = "Results and defensive solidity",
         subtitle = "Goals conceded remain a separate dimension from goals scored",
         x = "Goals conceded / match", y = "Points / match", colour = "Club") +
    theme_portfolio() + theme(legend.position = "none")
  ggsave(file.path(plots_dir, "14_points_vs_goals_conceded.png"), p, width = 13, height = 8, dpi = 300)
}

# Correlations: only calculable groups, to avoid NA values
# This avoids problems with fct_reorder and empty plots.
cor_plot_data <- cor_points_goals_v6 %>% filter(is.finite(correlation))
if (nrow(cor_plot_data) > 0) {
  p <- ggplot(cor_plot_data, aes(x = correlation,
                                 y = fct_reorder(paste(league, period, sep = " — "), correlation),
                                 colour = league)) +
    geom_point(size = 3.5) +
    geom_vline(xintercept = 0, linetype = 2, colour = "grey50") +
    scale_colour_manual(values = league_colors) +
    scale_x_continuous(limits = c(-1, 1)) +
    labs(title = "Association between points / match and goals scored / match",
         subtitle = "Only league-period pairs for which the correlation can be calculated",
         x = "Pearson correlation", y = NULL, colour = "League") +
    theme_portfolio()
  ggsave(file.path(plots_dir, "15_points_goals_correlations.png"), p, width = 12, height = 9, dpi = 300)
}

# OUTPUTS AND FINAL CHECKS
# Save the main summary tables

analysis_summary <- records_comparison %>%
  select(league, period, metric, record_value, record_holders) %>%
  left_join(
    top10_stats %>% select(league, period, metric, mean, median, q1, q3, best, worst),
    by = c("league", "period", "metric")
  ) %>%
  left_join(
    club_top10_recurrence_v6 %>%
      group_by(league, metric) %>%
      summarise(max_club_seasons = max(top10_seasons), .groups = "drop"),
    by = c("league", "metric")
  ) %>%
  arrange(metric, factor(period, levels = period_order), league)

write_csv(analysis_summary, file.path(tables_dir, "40_summary_records_top10_recurrence.csv"))

# Final checks and notes for running the script

readme_v6 <- c(
  "HISTORICAL FOOTBALL PORTFOLIO — FOOTBALL DATA VERSION",
  "",
  "ANALYTICAL FLOW",
  "1. Records: comparison across leagues and eras.",
  "2. Top 10: mean, median and dispersion to measure elite depth.",
  "3. Big 5: internal comparison, then Big 5 / other leagues comparison.",
  "4. Recurrence: recurring clubs within ONE ranking at a time.",
  "5. Time: Top 10 and All Time concentration by decade.",
  "6. Profiles: results, attack and defense analyzed separately.",
  "",
  "RULES",
  "- The four rankings are never transformed into a composite score.",
  "- A season can belong to several historical windows; recurrence tables use unique seasons.",
  "- Periods remain comparable only with the precautions imposed by their historical coverage.",
  "- Top 10s measure the presence of elite performances, not the average level of a league.",
  "- Goals scored and goals conceded remain two different dimensions.",
  "- Correlations are descriptive and do not demonstrate causality.",
  ""
)

writeLines(readme_v6, file.path(output_dir, "README.txt"), useBytes = TRUE)

cat("\n============================================================\n")
cat("FOOTBALL PORTFOLIO — DATA VERSION COMPLETE\n")
cat("============================================================\n")
cat("Rows analyzed:", nrow(individual_data), "\n")
cat("Leagues :", paste(sort(unique(individual_data$league)), collapse = " | "), "\n")
cat("Periods :", paste(sort(unique(individual_data$period)), collapse = " | "), "\n")
cat("Charts:", plots_dir, "\n")
cat("Tables:", tables_dir, "\n")
cat("============================================================\n")
