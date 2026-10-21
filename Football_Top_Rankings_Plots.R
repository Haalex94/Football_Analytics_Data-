# Create one chart for every Top 10, Top 15 and Top 20 ranking in the workbook.
# This script only reads the existing ranking tables and exports PNG charts.

required_packages <- c("readxl", "ggplot2")
missing_packages <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_packages) > 0) {
  install.packages(missing_packages, repos = "https://cloud.r-project.org")
}

library(readxl)
library(ggplot2)

# Select the Excel workbook when the script starts.
excel_path <- file.choose()
if (!file.exists(excel_path)) stop("The selected Excel file could not be found.")

output_dir <- file.path(dirname(excel_path), "FOOTBALL_TOP_RANKING_CHARTS")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

# Keep team names consistent between charts without changing the source labels.
clean_text <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- NA_character_
  x <- gsub("\\s+", " ", x)
  trimws(x)
}

safe_filename <- function(x) {
  x <- iconv(x, from = "", to = "ASCII//TRANSLIT", sub = "")
  x <- tolower(x)
  x <- gsub("[^a-z0-9]+", "_", x)
  x <- gsub("^_+|_+$", "", x)
  ifelse(nzchar(x), x, "ranking")
}

# Convert Excel cells to numeric values while tolerating decimal commas.
parse_numeric <- function(x) {
  x <- clean_text(x)
  x <- gsub(" ", "", x, fixed = TRUE)
  x <- gsub(",", ".", x, fixed = TRUE)
  suppressWarnings(as.numeric(x))
}

# Find ranking titles, then read the table directly underneath each title.
read_ranking_tables <- function(path) {
  sheet_names <- readxl::excel_sheets(path)
  all_tables <- list()
  table_id <- 0L
  title_pattern <- "^\\s*Top\\s*(10|15|20)\\s+teams\\b"

  for (sheet_name in sheet_names) {
    raw <- readxl::read_excel(
      path,
      sheet = sheet_name,
      col_names = FALSE,
      .name_repair = "minimal"
    )
    cells <- as.data.frame(raw, stringsAsFactors = FALSE)
    if (nrow(cells) < 3 || ncol(cells) < 3) next

    # Look for Top 10 / Top 15 / Top 20 headings anywhere on the sheet.
    title_locations <- list()
    for (r in seq_len(nrow(cells))) {
      for (cc in seq_len(ncol(cells))) {
        value <- cells[[cc]][r]
        if (!is.na(value) && is.character(value) && grepl(title_pattern, value, ignore.case = TRUE)) {
          title_locations[[length(title_locations) + 1L]] <- c(row = r, col = cc)
        }
      }
    }
    if (length(title_locations) == 0) next

    for (loc in title_locations) {
      title_row <- as.integer(loc[["row"]])
      title_col <- as.integer(loc[["col"]])
      title <- clean_text(cells[[title_col]][title_row])
      header_row <- title_row + 1L
      if (header_row > nrow(cells)) next

      header_values <- vapply(seq_len(ncol(cells)), function(j) {
        x <- cells[[j]][header_row]
        if (length(x) == 0 || is.na(x)) "" else clean_text(x)
      }, character(1))

      # The table starts at its Position column and ends before the next table.
      position_cols <- which(tolower(header_values) == "position")
      position_cols <- position_cols[position_cols >= title_col]
      if (length(position_cols) == 0 || position_cols[1] != title_col) next
      next_position <- position_cols[position_cols > title_col]
      end_col <- if (length(next_position) > 0) next_position[1] - 1L else ncol(cells)
      if (end_col <= title_col) next

      local_headers <- header_values[title_col:end_col]
      find_header <- function(pattern) {
        hits <- which(grepl(pattern, local_headers, ignore.case = TRUE))
        if (length(hits) == 0) NA_integer_ else hits[1] + title_col - 1L
      }

      team_col <- find_header("^Team$")
      season_col <- find_header("^Season$")
      ratio_col <- find_header("ratio")
      position_col <- find_header("^Position$")
      if (anyNA(c(team_col, season_col, ratio_col, position_col))) next

      # Read downward until the Team field is empty or another section begins.
      data_rows <- list()
      if (header_row + 1L > nrow(cells)) next
      for (rr in seq.int(header_row + 1L, nrow(cells))) {
        team_value <- cells[[team_col]][rr]
        if (is.null(team_value) || length(team_value) == 0 || is.na(team_value) || !nzchar(clean_text(team_value))) break
        possible_title <- FALSE
        row_values <- unlist(cells[rr, , drop = FALSE], use.names = FALSE)
        possible_title <- any(vapply(row_values, function(z) {
          !is.na(z) && is.character(z) && grepl(title_pattern, z, ignore.case = TRUE)
        }, logical(1)))
        if (possible_title) break

        team <- clean_text(team_value)
        season <- clean_text(cells[[season_col]][rr])
        value <- parse_numeric(cells[[ratio_col]][rr])
        position <- clean_text(cells[[position_col]][rr])
        if (!is.finite(value)) next
        if (is.na(season) || !nzchar(season)) season <- ""
        if (is.na(position) || !nzchar(position)) position <- as.character(length(data_rows) + 1L)

        data_rows[[length(data_rows) + 1L]] <- data.frame(
          position = position,
          team = team,
          season = season,
          value = value,
          stringsAsFactors = FALSE
        )
      }

      if (length(data_rows) == 0) next
      ranking <- do.call(rbind, data_rows)
      ranking$sheet <- sheet_name
      ranking$title <- title
      ranking$top_n <- as.integer(sub("^\\s*Top\\s*(10|15|20).*", "\\1", title, ignore.case = TRUE))
      ranking$source_order <- seq_len(nrow(ranking))

      # Use a unique sequence number so similar sheet names cannot overwrite charts.
      table_id <- table_id + 1L
      ranking$table_id <- table_id
      all_tables[[length(all_tables) + 1L]] <- ranking
    }
  }

  all_tables
}

rankings <- read_ranking_tables(excel_path)
if (length(rankings) == 0) {
  stop("No Top 10, Top 15 or Top 20 ranking tables were detected in the selected workbook.")
}

# Club colours: explicit, stable colours for the clubs in the workbook.
# Colours are inspired by each club's traditional identity. White is avoided
# as a bar fill; white-dominant clubs use a recognisable secondary colour.
# Similar club colours (especially reds and blues) use different shades.
all_teams <- sort(unique(unlist(lapply(rankings, function(x) x$team))))
all_teams <- all_teams[!is.na(all_teams) & nzchar(all_teams)]

normalise_team <- function(x) {
  key <- tolower(iconv(x, from = "", to = "ASCII//TRANSLIT", sub = ""))
  key <- gsub("[^a-z0-9]+", " ", key)
  trimws(gsub("\\s+", " ", key))
}

# Keys are normalised club names. The aliases below account for spelling,
# accents, historical names, and common workbook variants.
club_colours <- c(
  "ac milan" = "#FB090B", "ac milan calciopoli" = "#B5121B",
  "as monaco" = "#E31B23", "as roma" = "#8E1F2F",
  "as saint etienne" = "#087A45", "az alkmaar" = "#D71920",
  "ajax amsterdam" = "#D2122E", "ajax" = "#D2122E",
  "arsenal" = "#EF0107", "aston villa" = "#670E36",
  "atalanta" = "#1B4E9B", "athletic club" = "#C8102E",
  "atletico madrid" = "#CB3524", "atletico de madrid" = "#CB3524",
  "bayer leverkusen" = "#E32219", "belenenses" = "#0072BC",
  "blackburn rovers" = "#009FE3", "bologna" = "#183B72",
  "bolton wanderers" = "#931A1D", "borussia dortmund" = "#FDE100",
  "borussia monchengladbach" = "#168447",
  "budapest honved" = "#C8102E", "budapest honved fc" = "#A6192E",
  "budapesti postas se" = "#D71920", "budapesti tc" = "#222222",
  "co roubaix tourcoing" = "#C8102E", "cagliari" = "#B01F36",
  "chelsea fc" = "#034694", "chelsea" = "#034694",
  "debrecen" = "#D71920", "deportivo" = "#007FC8",
  "dunaferr" = "#D71920", "eto fc gyor" = "#168447",
  "everton" = "#003399", "fc barcelona" = "#A50044",
  "barcelona" = "#A50044", "fc bayern munich" = "#DC052D",
  "bayern munich" = "#DC052D", "fc porto" = "#0050A4",
  "porto" = "#0050A4", "fc nantes" = "#F7D117",
  "fc sochaux montbeliard" = "#0055A4", "fc sete" = "#0072BC",
  "fc twente" = "#E30613", "fehervar" = "#D71920",
  "videoton fehervar" = "#D71920", "ferencvaros" = "#007A3D",
  "feyenoord" = "#E30613", "fiorentina" = "#482E92",
  "girondins de bordeaux" = "#001E62", "gyor" = "#168447",
  "hamburger sv" = "#005CA9", "inter milan" = "#0068A8",
  "juventus" = "#222222", "juventus turin" = "#222222",
  "juventus turin calciopoli" = "#555555",
  "losc lille" = "#E01E13", "lille" = "#E01E13",
  "leeds united" = "#1D428A", "leeds united " = "#1D428A",
  "liverpool fc" = "#C8102E", "liverpool" = "#C8102E",
  "mtk budapest" = "#1769AA", "magyar ac" = "#222222",
  "manchester city" = "#6CABDD", "manchester united" = "#DA291C",
  "nottingham forest" = "#DD0000", "ogc nice" = "#C8102E",
  "olymique lyonnais" = "#DA0812", "olympique lyonnais" = "#DA0812",
  "olympique lillois" = "#E01E13", "olympique de marseille" = "#2FAEE0",
  "marseille" = "#2FAEE0", "psv eindhoven" = "#ED1C24",
  "paksi fc" = "#168447", "paris sg" = "#004170", "psg" = "#004170",
  "perugia" = "#C8102E", "preston north end" = "#FFFFFF",
  "queens park rangers" = "#005CAB", "rc lens" = "#D71920",
  "rc paris" = "#0055A4", "rc strasbourg" = "#009FE3",
  "racing paris" = "#0055A4", "real betis" = "#087A45",
  "real madrid" = "#174A8B", "real sociedad" = "#0067B1",
  "roda jc" = "#F2D21B", "s c braga" = "#C8102E", "sc braga" = "#C8102E",
  "sl benfica" = "#E83030", "benfica" = "#E83030",
  "ss lazio" = "#87D8F7", "lazio" = "#87D8F7",
  "ssc napoli" = "#00AEEF", "napoli" = "#00AEEF",
  "schalke 04" = "#004D9D", "sevilla fc" = "#D71920", "sheffield wednesday" = "#0066B3",
  "sporting cp" = "#008C45", "stade de reims" = "#C8102E",
  "sunderland" = "#C8102E", "torino" = "#8A1538",
  "tottenham hotspur" = "#132257", "unidos de lisboa" = "#087A45",
  "vfb stuttgart" = "#D71920", "valencia cf" = "#EE3524",
  "vasas sc" = "#D71920", "vac fc" = "#D71920",
  "werder bremen" = "#008B6A", "west bromwich albion" = "#1B4E9B",
  "wolverhampton wanderers" = "#FDB913",
  "wolverhamton wanderers" = "#FDB913", "ujpest" = "#6F2C91",
  "ujpest fc" = "#6F2C91"
)

# Real Madrid and other white-dominant kits use a blue/charcoal fill so bars
# remain visible on the white plot background.
club_colours["real madrid"] <- "#174A8B"
club_colours["preston north end"] <- "#4A4A4A"

get_club_colour <- function(team) {
  key <- normalise_team(team)
  if (key %in% names(club_colours)) return(unname(club_colours[[key]]))

  # Stable fallback for any unexpected historical club name; never white.
  codepoints <- utf8ToInt(key)
  hash <- sum(codepoints * seq_along(codepoints))
  hue <- (hash * 137.508) %% 360
  grDevices::hcl(h = hue, c = 65, l = 48)
}

team_colours <- stats::setNames(
  vapply(all_teams, get_club_colour, character(1)),
  all_teams
)

make_ranking_plot <- function(data) {
  data$label <- ifelse(
    nzchar(data$season),
    paste0(data$position, ". ", data$team, " — ", data$season),
    paste0(data$position, ". ", data$team)
  )
  data$label <- make.unique(data$label)
  # Keep the ranking order from Excel, with rank 1 displayed at the top.
  data$label <- factor(data$label, levels = rev(data$label))
  data$team <- factor(data$team, levels = names(team_colours))

  max_value <- max(data$value, na.rm = TRUE)
  value_digits <- if (grepl("points/match|goals scored/match|goals conceded/match", data$title[1], ignore.case = TRUE)) 2 else 1

  ggplot(data, aes(x = value, y = label, fill = team)) +
    geom_col(width = 0.72, show.legend = FALSE) +
    geom_text(
      aes(label = formatC(value, format = "f", digits = value_digits)),
      hjust = -0.12,
      size = 3.1,
      colour = "#222222"
    ) +
    scale_fill_manual(values = team_colours, na.value = "grey60", drop = FALSE) +
    scale_x_continuous(expand = expansion(mult = c(0, 0.18))) +
    coord_cartesian(clip = "off") +
    labs(
      title = data$title[1],
      subtitle = paste0("Source: Excel sheet ‘", data$sheet[1], "’ | Club colours remain consistent across charts"),
      x = "Ratio",
      y = NULL
    ) +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      plot.subtitle = element_text(size = 9, colour = "grey35"),
      axis.text.y = element_text(size = if (nrow(data) >= 20) 8.5 else 9.5),
      panel.grid.major.y = element_blank(),
      panel.grid.minor = element_blank(),
      plot.margin = margin(10, 42, 10, 10)
    )
}

# Export one PNG for every ranking table found in the workbook.
for (i in seq_along(rankings)) {
  data <- rankings[[i]]
  plot <- make_ranking_plot(data)
  file_name <- sprintf(
    "%03d_%s_%s.png",
    data$table_id[1],
    safe_filename(data$sheet[1]),
    safe_filename(data$title[1])
  )
  output_path <- file.path(output_dir, file_name)
  chart_height <- max(5.5, 0.34 * nrow(data) + 2.2)

  ggsave(
    filename = output_path,
    plot = plot,
    width = 13.5,
    height = chart_height,
    units = "in",
    dpi = 300,
    limitsize = FALSE,
    bg = "white"
  )
}

message(length(rankings), " ranking charts saved to: ", output_dir)
