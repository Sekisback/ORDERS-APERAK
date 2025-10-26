# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#
# ORDERS-APERAK Automatisches Update                                        ----
#
# Author : Sascha Kornberger
# Datum  : 26.10.2025
# Version: 1.0.0
#
# History:
# 1.0.0  Funktion: Initiale Freigabe
#
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --

## OPTIONS ----
# Deaktiviert die die Meldungen der Quellpakete während der Installation.
options(install.packages.check.source = "no")

# Verhindert die wissenschaftliche Notation für Zahlen.
options(scipen = 999)

# BENOETIGTE PAKETE ----
## Liste der Pakete ----
pakete <- c(
  "readr", "readxl", "dplyr", "stringr", "lubridate"
)

## Installiere fehlende Pakete ohne Rückfragen ----
installiere_fehlende <- pakete[!pakete %in% installed.packages()[, "Package"]]
if (length(installiere_fehlende) > 0) {
  install.packages(
    installiere_fehlende,
    repos = "https://cran.r-project.org",
    quiet = TRUE
  )
}

## Lade alle Pakete
invisible(lapply(pakete, function(pkg) {
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}))


# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#                                 UNC-Pfade                                 ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --

## Pfade basierend auf der Umgebung ----
if (Sys.info()[["sysname"]] == "Windows") {
  main_path <- "//swnor.eeg-powerbox.de/eeg/Technik$/DLZMBS/05 irgendwas/ORDERS und APERAKs"
} else {
  main_path <- "/media/archive/RStudio/EEG/swnor.eeg-powerbox.de/eeg/Technik$/DLZMSB/ORDERS und APERAKs"
}

# Unterordner für Reports 
reports_path <- file.path(main_path, "_ECOUNT_REPORTS")


# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                           FUNKTIONEN                                 ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#

# formatierte Konsolenausgabe 
log_pretty <- function(label, value, width = 53) {
  if (is.na(value)) value <- ""
  msg <- paste0(str_pad(label, width = width, side = "right"), ": ", value)
  cat(msg, "\n")
}

# Ausgabe-Trennlinie 
print_line <- function(char = "-", length = 66) {
  cat(strrep(char, length), "\n", sep = "")
}

# CSV-Dateien finden 
get_csv_files <- function(main_path) {
  reports_path <- file.path(main_path, "_ECOUNT_REPORTS")
  csv_files <- list.files(reports_path, pattern = "\\.csv$", full.names = TRUE)
  
  if (length(csv_files) == 0) {
    print_line()
    log_pretty("Status", "Keine CSV-Dateien im Ordner _ECOUNT_REPORTS gefunden.")
    print_line()
    stop(invisible(NULL))
  }
  
  log_pretty("Gefundene CSV-Dateien", length(csv_files))
  return(csv_files)
}


# Kundennamen extrahieren
extract_customer_name <- function(filename) {
  customer <- str_extract(basename(filename), "(?<=_MSB_)[^.]+")
  if (is.na(customer)) {
    log_pretty("Warnung", paste("Kein Kundenname in", basename(filename), "gefunden"))
  }
  return(customer)
}


# Netzbetreiber aus Dateinamen extrahieren (Tennet, 50Hertz)
extract_network_operator <- function(filename) {
  op <- str_extract(basename(filename), "(?<=Orders_)[^_]+")
  return(op)
}


# CSV einlesen
read_orders_csv <- function(csv_path) {
  log_pretty("Lese CSV-Datei", basename(csv_path))
  
  df <- suppressMessages(
    read_delim(
      csv_path, 
      delim = ";", 
      escape_double = FALSE, 
      trim_ws = TRUE, 
      locale = locale(encoding = "UTF-8")
    )
  ) |> 
    mutate(
      DOC_DATE = case_when(
        nchar(DOC_DATE) == 8  ~ as.Date(DOC_DATE, format = "%d.%m.%y"),
        nchar(DOC_DATE) == 10 ~ as.Date(DOC_DATE, format = "%d.%m.%Y"),
        TRUE ~ as.Date(NA)
      )
    ) |>
    select(DOC_DATE, EMPFAENGER_ID, SENDER_ID, LS_ZPT) |>
    group_by(LS_ZPT) |>
    slice_max(DOC_DATE, n = 1, with_ties = FALSE) |>
    ungroup()
  
  return(df)
}


# XLSX einlesen
read_orders_xlsx <- function(main_path, customer, operator = NULL) {
  customer_path <- file.path(main_path, customer)
  xlsx_files <- list.files(customer_path, pattern = "\\.xlsx$", full.names = TRUE)
  
  if (length(xlsx_files) == 0) {
    log_pretty("Warnung", paste("Keine XLSX-Datei im Kundenordner", customer, "gefunden"))
    return(NULL)
  }
  
  xlsx_file <- xlsx_files[1]
  wb_sheets <- openxlsx::getSheetNames(xlsx_file)
  
  # Auswahl Tabellenblatt
  sheet_to_use <- if (customer == "VBE" && grepl("50Hertz", operator, ignore.case = TRUE)) {
    if (length(wb_sheets) >= 2) wb_sheets[2] else wb_sheets[1]
  } else {
    wb_sheets[1]
  }
  
  log_pretty("Lese XLSX-Datei", paste(basename(xlsx_file), "→ Blatt:", sheet_to_use))
  
  df <- read_excel(xlsx_file, sheet = sheet_to_use, skip = 1)
  df[[1]] <- convert_excel_date(df[[1]])
  attr(df, "sheet_name") <- sheet_to_use  # merken für später
  
  return(df)
}



# Excel-Datum konvertieren
convert_excel_date <- function(x) {
  x <- as.vector(x)
  if (is.factor(x)) x <- as.character(x)
  
  suppressWarnings({
    if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30"))
  })
  
  suppressWarnings({
    x_clean <- gsub(",", ".", x)
    x_num <- as.numeric(x_clean)
    if (any(!is.na(x_num))) return(as.Date(x_num, origin = "1899-12-30"))
  })
  
  suppressWarnings({
    out <- ymd(x)
    if (all(is.na(out))) out <- dmy(x)
  })
  
  return(out)
}


# Kunde verarbeiten
process_customer <- function(csv_path, main_path) {
  customer <- extract_customer_name(csv_path)
  operator <- extract_network_operator(csv_path)
  if (is.na(customer)) return(NULL)
  
  print_line()
  log_pretty("Verarbeite Orders für Kunde", customer)
  log_pretty("Netzbetreiber erkannt", operator)
  
  df_orders_neu <- read_orders_csv(csv_path)
  df_orders_alt <- read_orders_xlsx(main_path, customer, operator)
  
  if (is.null(df_orders_alt)) return(NULL)
  
  return(list(
    kunde = customer,
    operator = operator,
    df_orders_neu = df_orders_neu,
    df_orders_alt = df_orders_alt
  ))
}



# Vergleich Neu vs Alt
compare_orders <- function(df_orders_neu, df_orders_alt) {
  log_pretty("Vergleiche Daten", "Neue vs. alte Orders")
  
  if (!("LS_ZPT" %in% names(df_orders_neu))) {
    stop("Spalte 'LS_ZPT' fehlt in df_orders_neu.")
  }
  if (!("MaLo" %in% names(df_orders_alt))) {
    stop("Spalte 'MaLo' fehlt in df_orders_alt.")
  }
  
  df_diff <- dplyr::anti_join(df_orders_neu, df_orders_alt, by = c("LS_ZPT" = "MaLo"))
  log_pretty("Neue Zählpunkte gefunden", nrow(df_diff))
  
  return(df_diff)
}


# Hauptverarbeitung
run_all_customers <- function(main_path) {
  csv_files <- get_csv_files(main_path)
  ergebnisse <- list()
  
  for (csv in csv_files) {
    res <- process_customer(csv, main_path)
    if (is.null(res)) next
    
    log_pretty("Vergleiche", "df_orders_neu mit df_orders_alt")
    
    df_diff <- compare_orders(res$df_orders_neu, res$df_orders_alt)
    ergebnisse[[res$kunde]] <- list(
      df_orders_neu = res$df_orders_neu,
      df_orders_alt = res$df_orders_alt,
      df_diff = df_diff
    )
    
    log_pretty("Neue Zählpunkte für Kunde", paste0(res$kunde, " → ", nrow(df_diff)))
    update_orders_xlsx(main_path, res$kunde, df_diff, res$operator)
  }
  
  print_line()
  log_pretty("Verarbeitung abgeschlossen. Kunden eingelesen", 
             paste(names(ergebnisse), collapse = ", "))
  print_line()
  
  return(ergebnisse)
}


# XLSX aktualisieren
update_orders_xlsx <- function(main_path, customer, df_diff, operator = NULL) {
  library(openxlsx)
  
  if (nrow(df_diff) == 0) {
    log_pretty("Status", paste("Keine neuen Zählpunkte für", customer, "– kein Update nötig"))
    return(invisible(NULL))
  }
  
  # Hilfsfunktion: flexible Spaltensuche
  find_col <- function(df, pattern) {
    cols <- trimws(names(df))
    match <- cols[grepl(pattern, cols, ignore.case = TRUE)]
    if (length(match) == 0) return(NA_character_)
    match[1]
  }
  
  # Datei & Blatt bestimmen
  customer_path <- file.path(main_path, customer)
  xlsx_files <- list.files(customer_path, pattern = "\\.xlsx$", full.names = TRUE)
  if (length(xlsx_files) == 0) {
    log_pretty("Warnung", paste("Keine XLSX-Datei im Kundenordner", customer))
    return(NULL)
  }
  
  xlsx_file <- xlsx_files[1]
  wb <- loadWorkbook(xlsx_file)
  wb_sheets <- sheets(wb)
  
  sheet_name <- if (identical(customer, "VBE") && grepl("50Hertz", operator %||% "", ignore.case = TRUE)) {
    if (length(wb_sheets) >= 2) wb_sheets[2] else wb_sheets[1]
  } else {
    wb_sheets[1]
  }
  
  log_pretty("Aktualisiere XLSX-Datei", paste(basename(xlsx_file), "→ Blatt:", sheet_name))
  
  # Daten lesen (exakt das Blatt)
  df_xlsx <- read.xlsx(xlsx_file, sheet = sheet_name, skipEmptyRows = FALSE, startRow = 2)
  
  # Spalten finden
  col_letzter_eingang <- find_col(df_xlsx, "^letzter")
  col_malo            <- find_col(df_xlsx, "^MaLo")
  col_fehler          <- find_col(df_xlsx, "^Fehlerbearbeitung")
  col_bemerkung       <- find_col(df_xlsx, "^Bemerkung")
  
  needed_cols <- c(col_letzter_eingang, col_malo, col_fehler, col_bemerkung)
  if (any(is.na(needed_cols) | needed_cols == "")) {
    stop("Eine oder mehrere benötigte Spalten wurden nicht gefunden. Vorhandene Spalten: ",
         paste(names(df_xlsx), collapse = ", "))
  }
  
  # Neue Zeilen vorbereiten
  df_append <- as.data.frame(matrix(NA, nrow = nrow(df_diff), ncol = ncol(df_xlsx)))
  names(df_append) <- names(df_xlsx)
  df_append[[col_letzter_eingang]] <- df_diff$DOC_DATE
  df_append[[col_malo]] <- df_diff$LS_ZPT
  df_append[[col_fehler]] <- sprintf(
    "%s - Bitte Stammdatenänderung mit %s durchführen oder Stammdaten korrigieren",
    customer,
    operator
  )
  
  # (Optional) Filter auf leere Bemerkungen
  df_append <- df_append[is.na(df_append[[col_bemerkung]]) | df_append[[col_bemerkung]] == "", , drop = FALSE]
  
  # Nächste freie Zeile robust ermitteln (genau dieses Blatt!)
  existing_rows <- nrow(read.xlsx(xlsx_file, sheet = sheet_name, startRow = 2, skipEmptyRows = FALSE))
  start_row <- existing_rows + 3  # +1 Header, +1 nächste freie Zeile (Sicherheitsabstand)
  
  # Schreiben & speichern
  writeData(wb, sheet = sheet_name, x = df_append, startRow = start_row, colNames = FALSE)
  saveWorkbook(wb, xlsx_file, overwrite = TRUE)
  
  log_pretty("XLSX Update abgeschlossen", paste(customer, "-", nrow(df_append), "neue Zeilen angehängt"))
  print_line()
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                               MAIN                                   ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#

alles <- run_all_customers(main_path)

