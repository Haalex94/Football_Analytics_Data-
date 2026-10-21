# Football_Analytics_Data-

Historical Performance in European Football

A data-driven analysis of elite performance, recurrence and evolution across eight European leagues

This project explores historical team performance across eight European domestic football leagues. Using team-season data, it examines how exceptional performances differ across competitions, historical periods and performance indicators.

The objective is to identify historical records, compare the depth of elite performance, study the recurrence of clubs among top-performing teams, and explore how attacking output, defensive solidity and league results relate to one another.

The project was developed independently using R, with the findings presented in an English-language report.

## Leagues covered

The analysis includes eight domestic leagues:

- Premier League (English top flight)
- La Liga
- Serie A
- Bundesliga
- Ligue 1
- Primeira Liga
- Eredivisie
- Hungarian top division

## Research questions

The project investigates several questions:

- Which teams produced the most exceptional historical performances in each league?
- How do the strongest team-seasons compare across leagues and historical periods?
- How frequently do the same clubs appear among the top-performing teams?
- How has the concentration of elite performances changed across decades?
- How are points, attacking output and defensive performance related?
- Do the patterns observed in the Big Five leagues differ from those in the other competitions included in the study?

## Performance indicators

Four indicators are analysed separately throughout the project:

| Indicator | What it measures |
|---|---|
| Points per match | Overall results relative to matches played |
| Points per match among non-champions | The level of performance achieved by teams that did not win the league |
| Goals scored per match | Attacking output |
| Goals conceded per match | Defensive solidity; lower values indicate stronger defensive records |

These indicators represent different dimensions of performance. They are not combined into a single score, and the project does not attempt to identify one universally best club across all metrics.

## Historical periods and methodology

The analysis compares three main periods:

- **All Time:** the full historical period available in the dataset.
- **1955+:** seasons from 1955 onwards.
- **Since 3 Points:** seasons played under the three-points-for-a-win system.

A separate historical period is also used for Premier League-specific analyses.

Because leagues were founded at different times and historical scoring systems have changed, the periods do not always cover identical samples. In particular, the Bundesliga and Eredivisie were established after 1955, so their All-Time and 1955+ samples overlap completely.

Historical points are handled using the ratios supplied in the workbook rather than by reconstructing every older season as if the modern three-points system had applied. Per-match indicators are used to facilitate comparisons between seasons of different lengths.

The methodology includes historical rankings, comparisons between leagues, club recurrence, decade-level distributions and analyses of relationships between performance indicators. Cross-league comparisons use the defined common periods, excluding the separate Premier League 1955–1992 window where no equivalent period is available for all other leagues.

## Main areas of analysis

The report is organised around the following themes:

1. Historical records by league, period and indicator.
2. Depth of elite performance beyond individual record seasons.
3. Comparisons between the Big Five leagues and the other competitions.
4. Rankings of leading European team-seasons.
5. Concentration of elite performances across decades.
6. Recurrence and persistence of clubs across historical periods.
7. Relationships between points, goals scored and goals conceded.
8. Historical context and limitations of cross-era comparisons.

## Tools and technologies

- **R** for data preparation, analysis and visualisation.
- **readxl** for importing Excel data.
- **dplyr**, **tidyr**, **stringr** and **purrr** for data manipulation.
- **ggplot2**, **forcats** and **scales** for visualisation and chart formatting.
- **readr** for data export.

The script generates tables and charts in a dedicated output folder.

## Repository contents

- `Portfolio_Foot.pdf` — the final report, including methodology, visualisations, results and discussion.
- `Portfolio_Football_Code_Annotated.R` — the annotated R script used to prepare the data, conduct the analyses and generate the outputs.
- `PortfolioFoot_Data.xlsx` — the Excel workbook used as the data source, if redistribution is permitted.

## Running the analysis

1. Install R and preferably RStudio.
2. Download the R script and the Excel workbook.
3. Open the script in RStudio.
4. Run the script and select the Excel workbook when prompted.
5. The script installs any missing required R packages and saves the generated tables and charts in a folder named `PORTFOLIO_FOOTBALL_ENGLISH` next to the workbook.

The workbook must retain the expected sheet structure for the script to read it correctly.

## Limitations

Historical comparisons require caution. League formats, season lengths, points systems and playing conditions have changed over time, and the eight leagues do not all cover the same historical span.

The indicators used here describe observed team-season outcomes; they do not establish why a team performed exceptionally well. Historical context is discussed alongside the results, but the analysis does not establish causal relationships between historical events and performance.

The findings should therefore be interpreted as a comparative exploration of historical league data rather than a definitive ranking of clubs across all eras.

## Project purpose

This independent project was an opportunity to apply data analysis to a subject I follow closely: football. It combines statistical exploration, data visualisation and historical research to investigate how elite team performance has evolved across European leagues.

The project also reflects my interest in developing further skills in data analysis and applying quantitative methods to sports-related questions.
