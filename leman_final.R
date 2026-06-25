# =============================================================
# SETUP INICIAL — Execute UMA VEZ ao clonar o projeto.
# Cria a estrutura de pastas e move arquivos para os lugares.
# =============================================================

setwd("~/Desktop/leman_cem")

dir.create("R",              showWarnings = FALSE)
dir.create("data/raw",       showWarnings = FALSE, recursive = TRUE)
dir.create("data/processed", showWarnings = FALSE, recursive = TRUE)
dir.create("outputs",        showWarnings = FALSE)

# Move CSVs INEP para data/raw
arquivos_raw <- c(
  list.files(pattern = "^microdados_ed_basica_.*\\.csv$", ignore.case = TRUE),
  list.files(pattern = "^Tabela_.*_2025\\.csv$",          ignore.case = TRUE)
)
if (length(arquivos_raw) > 0)
  file.rename(arquivos_raw, file.path("data/raw", arquivos_raw))

# Move .rds para data/processed
arquivos_rds <- list.files(pattern = "\\.rds$")
if (length(arquivos_rds) > 0)
  file.rename(arquivos_rds, file.path("data/processed", arquivos_rds))

# Move XLSX para outputs
arquivos_out <- c(
  list.files(pattern = "^codebook_.*\\.xlsx$"),
  list.files(pattern = "^tabelas_analiticas_.*\\.xlsx$")
)
if (length(arquivos_out) > 0)
  file.rename(arquivos_out, file.path("outputs", arquivos_out))

message("Estrutura inicial criada. Próximo passo: source('run.R')")



# =============================================================
# LEMANN-CEM — Censo Escolar 2015–2025
# Orquestrador único. Executa os 8 blocos em sequência.
# Anderson Henrique — versão final consolidada (jun/2026).
# =============================================================

setwd("~/Desktop/leman_cem")

t0 <- Sys.time()
message("============================================================")
message("LEMANN-CEM — Pipeline Censo Escolar 2015–2025")
message("Início: ", format(t0, "%Y-%m-%d %H:%M:%S"))
message("============================================================\n")

source("R/00_setup.R",        echo = FALSE)
source("R/01_helpers.R",      echo = FALSE)
source("R/02_leitura.R",      echo = FALSE)
source("R/03_tipagem.R",      echo = FALSE)
source("R/04_empilhamento.R", echo = FALSE)
source("R/05_tabelas.R",      echo = FALSE)
source("R/06_validacao.R",    echo = FALSE)   # GATE
source("R/07_export.R",       echo = FALSE)

t1 <- Sys.time()
message("\n============================================================")
message("Pipeline concluído em ",
        round(difftime(t1, t0, units = "mins"), 1), " min")
message("============================================================")



# =============================================================
# 00 — SETUP
# Pacotes, paths e constantes do projeto.
# =============================================================

suppressPackageStartupMessages({
  pacotes <- c("educabR", "dplyr", "readr", "janitor", "purrr",
               "stringr", "tidyr", "openxlsx", "httr", "jsonlite")
  invisible(lapply(pacotes, function(p) {
    if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
    library(p, character.only = TRUE)
  }))
})

# ── Paths ────────────────────────────────────────────────────
PATHS <- list(
  raw       = "data/raw",
  processed = "data/processed",
  outputs   = "outputs"
)

# ── Constantes do domínio ────────────────────────────────────
CAPITAIS_CO <- c(
  1100205, 1302603, 1200401, 1600303, 1721000, 1400100,
  1501402, 2111300, 2211001, 2304400, 2408102, 2507507,
  2611606, 2704302, 2800308, 2927408, 3106200, 3205309,
  3304557, 3550308, 4106902, 4205407, 4314902, 5002704,
  5103403, 5208707, 5300108
)

CO_MUN_BRASILIA <- 5300108
CO_MUN_SP       <- 3550308

ANOS <- 2015:2025

# ── 66 variáveis canônicas (estrutura do banco) ──────────────
VARIAVEIS_INTERESSE <- c(
  "NU_ANO_CENSO", "SG_UF", "NO_MUNICIPIO", "CO_MUNICIPIO",
  "NO_MICRORREGIAO", "CO_MICRORREGIAO", "NO_DISTRITO", "CO_DISTRITO",
  "CO_ENTIDADE", "DS_ENDERECO", "NU_ENDERECO", "DS_COMPLEMENTO", "CO_CEP",
  "TP_DEPENDENCIA", "TP_CATEGORIA_ESCOLA_PRIVADA", "TP_LOCALIZACAO",
  "TP_LOCALIZACAO_DIFERENCIADA", "TP_SITUACAO_FUNCIONAMENTO",
  "IN_PODER_PUBLICO_PARCERIA", "TP_PODER_PUBLICO_PARCERIA",
  "IN_MANT_ESCOLA_PRIVADA_EMP", "IN_MANT_ESCOLA_PRIV_ONG_OSCIP",
  "IN_MANT_ESCOLA_PRIVADA_SIND", "IN_ALIMENTACAO",
  "IN_INF", "IN_INF_CRE", "IN_INF_PRE",
  "QT_MAT_BAS", "QT_MAT_INF_CRE", "QT_MAT_INF_PRE",
  "QT_MAT_FUND_AI", "QT_MAT_FUND_AF", "QT_MAT_MED", "QT_MAT_PROF",
  "QT_MAT_EJA", "QT_MAT_EJA_FUND_AI", "QT_MAT_EJA_FUND_AF", "QT_MAT_EJA_MED",
  "QT_MAT_ESP", "QT_MAT_ESP_CC", "QT_MAT_ESP_CE",
  "QT_MAT_BAS_FEM", "QT_MAT_BAS_MASC", "QT_MAT_BAS_ND",
  "QT_MAT_BAS_BRANCA", "QT_MAT_BAS_PRETA", "QT_MAT_BAS_PARDA",
  "QT_MAT_BAS_AMARELA", "QT_MAT_BAS_INDIGENA",
  "QT_MAT_INF_INT", "QT_MAT_INF_CRE_INT", "QT_MAT_INF_PRE_INT",
  "QT_MAT_FUND_INT", "QT_MAT_FUND_AI_INT", "QT_MAT_FUND_AF_INT",
  "QT_MAT_MED_INT", "QT_MAT_ZR_URB", "QT_MAT_ZR_RUR",
  "QT_TRANSP_PUBLICO",
  "QT_DOC_INF_CRE", "QT_DOC_INF_PRE",
  "QT_TUR_INF_CRE", "QT_TUR_INF_PRE", "QT_TUR_INF_INT",
  "QT_TUR_INF_CRE_INT", "QT_TUR_INF_PRE_INT"
)

# ── Variáveis de infraestrutura ─────────────────────────────
NOVAS_VARS_INFRA <- c(
  "IN_AREA_VERDE", "IN_BANHEIRO_EI",
  "IN_MATERIAL_PED_JOGOS", "IN_MATERIAL_PED_ARTISTICAS"
)

# ── Variáveis numéricas para coerção em massa ────────────────
VARS_NUMERICAS <- c(
  "CO_MUNICIPIO", "CO_MICRORREGIAO", "CO_DISTRITO",
  "CO_ENTIDADE", "CO_CEP", "NU_ENDERECO", "NU_ANO_CENSO",
  "QT_MAT_BAS", "QT_MAT_INF_CRE", "QT_MAT_INF_PRE",
  "QT_MAT_FUND_AI", "QT_MAT_FUND_AF", "QT_MAT_MED", "QT_MAT_PROF",
  "QT_MAT_EJA", "QT_MAT_EJA_FUND_AI", "QT_MAT_EJA_FUND_AF", "QT_MAT_EJA_MED",
  "QT_MAT_ESP", "QT_MAT_ESP_CC", "QT_MAT_ESP_CE",
  "QT_MAT_BAS_FEM", "QT_MAT_BAS_MASC", "QT_MAT_BAS_ND",
  "QT_MAT_BAS_BRANCA", "QT_MAT_BAS_PRETA", "QT_MAT_BAS_PARDA",
  "QT_MAT_BAS_AMARELA", "QT_MAT_BAS_INDIGENA",
  "QT_MAT_INF_INT", "QT_MAT_INF_CRE_INT", "QT_MAT_INF_PRE_INT",
  "QT_MAT_FUND_INT", "QT_MAT_FUND_AI_INT", "QT_MAT_FUND_AF_INT",
  "QT_MAT_MED_INT", "QT_MAT_ZR_URB", "QT_MAT_ZR_RUR",
  "QT_TRANSP_PUBLICO",
  "QT_DOC_INF_CRE", "QT_DOC_INF_PRE",
  "QT_TUR_INF_CRE", "QT_TUR_INF_PRE", "QT_TUR_INF_INT",
  "QT_TUR_INF_CRE_INT", "QT_TUR_INF_PRE_INT"
)

# ── Variáveis-atributo (safe_max em vez de sum) ─────────────
VARS_ATRIBUTO_ESCOLA <- c("QT_TRANSP_PUBLICO")

# ── Mapas de rótulos ─────────────────────────────────────────
LOC_MAP <- c("1" = "urb", "2" = "rur")

LOC_DIF_MAP <- c(
  "Escola fora de área diferenciada"              = "fora",
  "Área de assentamento"                          = "asset",
  "Terra indígena"                                = "indig",
  "Área com comunidade remanescente de quilombos" = "quilom",
  "Área com povos e comunidades tradicionais"     = "trad"
)

LOC_DIF_NIVEIS <- unname(LOC_DIF_MAP)

# ── Chaves de agrupamento ────────────────────────────────────
G_MUN  <- c("NU_ANO_CENSO", "SG_UF", "NO_MUNICIPIO", "CO_MUNICIPIO")
G_DIST <- c("NU_ANO_CENSO", "CO_DISTRITO")

message("[00] Setup concluído.")



# =============================================================
# 01 — HELPERS
# Define todas as funções utilitárias. Sem side-effects.
# =============================================================

# ── safe_max() ───────────────────────────────────────────────
safe_max <- function(x) {
  v <- suppressWarnings(max(x, na.rm = TRUE))
  if (is.infinite(v)) NA_real_ else v
}

# ── padronizar_df() ──────────────────────────────────────────
padronizar_df <- function(df, ano) {
  names(df) <- toupper(names(df))
  
  if ("IN_CONVENIADA_PP" %in% names(df) &&
      !"IN_PODER_PUBLICO_PARCERIA" %in% names(df)) {
    df <- rename(df, IN_PODER_PUBLICO_PARCERIA = IN_CONVENIADA_PP)
    message("  [renomeado] IN_CONVENIADA_PP → IN_PODER_PUBLICO_PARCERIA")
  }
  if ("TP_CONVENIO_PODER_PUBLICO" %in% names(df) &&
      !"TP_PODER_PUBLICO_PARCERIA" %in% names(df)) {
    df <- rename(df, TP_PODER_PUBLICO_PARCERIA = TP_CONVENIO_PODER_PUBLICO)
    message("  [renomeado] TP_CONVENIO_PODER_PUBLICO → TP_PODER_PUBLICO_PARCERIA")
  }
  
  df <- mutate(df, NU_ANO_CENSO = as.integer(ano))
  
  disponiveis <- intersect(VARIAVEIS_INTERESSE, names(df))
  ausentes    <- setdiff(VARIAVEIS_INTERESSE, names(df))
  
  if (length(ausentes) > 0)
    message("  [NA estrutural] Ausentes em ", ano, ": ",
            paste(ausentes, collapse = ", "))
  
  df <- select(df, all_of(disponiveis))
  for (col in ausentes) df[[col]] <- NA
  df <- select(df, all_of(VARIAVEIS_INTERESSE))
  
  message("  [ok] ", format(nrow(df), big.mark = "."),
          " escolas | ", ncol(df), " colunas")
  df
}

# ── carregar_ano() ───────────────────────────────────────────
carregar_ano <- function(ano) {
  message("\n>>> Carregando ano: ", ano)
  df <- tryCatch(
    educabR::get_censo_escolar(ano, quiet = FALSE),
    error = function(e) { warning("Erro ", ano, ": ", e$message); NULL }
  )
  if (is.null(df)) return(NULL)
  padronizar_df(df, ano)
}

# ── ler_csv_inep() ───────────────────────────────────────────
ler_csv_inep <- function(arquivo) {
  message("  Lendo: ", arquivo)
  df <- read_delim(
    arquivo, delim = ";",
    locale = locale(encoding = "latin1", decimal_mark = ",", grouping_mark = "."),
    col_types = cols(.default = col_character()),
    progress = TRUE, show_col_types = FALSE
  )
  names(df) <- toupper(names(df))
  message("  [ok] ", format(nrow(df), big.mark = ","),
          " linhas | ", ncol(df), " colunas")
  df
}

# ── agregar_por_escola() ─────────────────────────────────────
agregar_por_escola <- function(df, vars_ok, nome) {
  vars_num <- setdiff(vars_ok, "CO_ENTIDADE")
  
  resultado <- df |>
    select(all_of(intersect(vars_ok, names(df)))) |>
    mutate(across(-CO_ENTIDADE,
                  ~ as.numeric(str_replace_all(., ",", ".")))) |>
    group_by(CO_ENTIDADE)
  
  cols_disponiveis <- intersect(vars_num, names(df))
  cols_atributo    <- intersect(VARS_ATRIBUTO_ESCOLA, cols_disponiveis)
  cols_soma        <- setdiff(cols_disponiveis, cols_atributo)
  
  resultado <- resultado |>
    summarise(
      across(all_of(cols_soma),     \(x) sum(x, na.rm = TRUE)),
      across(all_of(cols_atributo), \(x) safe_max(x)),
      .groups = "drop"
    )
  
  message("  [", nome, "] ",
          format(nrow(resultado), big.mark = ","), " escolas únicas",
          if (length(cols_atributo) > 0)
            paste0(" (safe_max: ", paste(cols_atributo, collapse = ", "), ")")
          else "")
  resultado
}

# ── extrair_infra() ──────────────────────────────────────────
extrair_infra <- function(df, ano) {
  names(df) <- toupper(names(df))
  df <- mutate(df, NU_ANO_CENSO = as.character(ano))
  
  presentes <- intersect(NOVAS_VARS_INFRA, names(df))
  ausentes  <- setdiff(NOVAS_VARS_INFRA, names(df))
  
  if (length(ausentes) > 0)
    message("  [NA estrutural] Ausentes em ", ano, ": ",
            paste(ausentes, collapse = ", "))
  
  df <- select(df, CO_ENTIDADE, NU_ANO_CENSO, all_of(presentes))
  for (col in ausentes) df[[col]] <- NA_character_
  mutate(df, across(everything(), as.character))
}

# ── buscar_scaffold_sp() — IBGE 96 distritos ─────────────────
buscar_scaffold_sp <- function() {
  message(">>> Buscando scaffold IBGE de distritos de SP...")
  url <- paste0("https://servicodados.ibge.gov.br/api/v1/localidades/",
                "municipios/", CO_MUN_SP, "/distritos")
  
  resp <- tryCatch(
    httr::GET(url, httr::timeout(30)),
    error = function(e) NULL
  )
  
  if (is.null(resp) || httr::status_code(resp) != 200) {
    warning("API IBGE indisponível. Usando fallback dos próprios dados.")
    return(NULL)
  }
  
  dados <- jsonlite::fromJSON(httr::content(resp, "text", encoding = "UTF-8"),
                              flatten = TRUE)
  
  scaffold <- tibble::tibble(
    CO_DISTRITO = as.numeric(dados$id),
    NO_DISTRITO_IBGE = stringr::str_to_title(dados$nome)
  )
  
  stopifnot(
    "API IBGE retornou número errado de distritos para SP" =
      nrow(scaffold) == 96L
  )
  
  message("  [ok] ", nrow(scaffold), " distritos obtidos da API IBGE")
  scaffold
}

# ── Filtros de rede ──────────────────────────────────────────
f_dir <- function(df) {
  filter(df,
         TP_DEPENDENCIA            == "3",
         TP_SITUACAO_FUNCIONAMENTO == "1")
}

f_ind_mun <- function(df) {
  filter(df,
         TP_DEPENDENCIA            == "4",
         TP_SITUACAO_FUNCIONAMENTO == "1",
         TP_PODER_PUBLICO_PARCERIA == 1L)
}

f_ind_mun_est <- function(df) {
  filter(df,
         TP_DEPENDENCIA            == "4",
         TP_SITUACAO_FUNCIONAMENTO == "1",
         TP_PODER_PUBLICO_PARCERIA == 3L)
}

f_dir_capitais <- function(df) {
  filter(df,
         TP_SITUACAO_FUNCIONAMENTO == "1",
         TP_DEPENDENCIA == "3" |
           (TP_DEPENDENCIA == "2" & CO_MUNICIPIO == CO_MUN_BRASILIA))
}

# ── completar_colunas() ──────────────────────────────────────
completar_colunas <- function(tabela, referencia) {
  cols_ausentes <- setdiff(names(referencia), names(tabela))
  cols_extras   <- setdiff(names(tabela),     names(referencia))
  if (length(cols_ausentes) > 0) {
    message("  Completando ", length(cols_ausentes), " colunas: ",
            paste(cols_ausentes, collapse = ", "))
    for (col in cols_ausentes) tabela[[col]] <- 0
  }
  if (length(cols_extras) > 0)
    message("  Preservando ", length(cols_extras),
            " extras: ", paste(cols_extras, collapse = ", "))
  select(tabela, all_of(c(names(referencia), cols_extras)))
}

# ── construir_tabela() ───────────────────────────────────────
construir_tabela <- function(base, g_vars, filtro_dir = f_dir) {
  
  message(">>> Construindo: ", paste(g_vars, collapse = " × "))
  
  dir     <- filtro_dir(base)
  ind     <- f_ind_mun(base)
  ind_est <- f_ind_mun_est(base)
  tot     <- bind_rows(dir, ind)
  
  base_chaves <- bind_rows(
    dir     |> distinct(across(all_of(g_vars))),
    ind     |> distinct(across(all_of(g_vars))),
    ind_est |> distinct(across(all_of(g_vars)))
  ) |>
    distinct() |>
    filter(if_all(all_of(g_vars), ~ !is.na(.)))
  
  message("  [base] ", format(nrow(base_chaves), big.mark = ","),
          " grupos únicos (dir ∪ ind ∪ ind_est)")
  
  agg <- function(df) group_by(df, across(all_of(g_vars)))
  
  agg_loc <- function(df, qt_col, pref) {
    df |>
      filter(!is.na(TP_LOCALIZACAO)) |>
      mutate(loc = LOC_MAP[TP_LOCALIZACAO]) |>
      group_by(across(all_of(c(g_vars, "loc")))) |>
      summarise(mat = sum(.data[[qt_col]], na.rm = TRUE), .groups = "drop") |>
      mutate(loc = factor(loc, levels = unname(LOC_MAP))) |>
      pivot_wider(names_from = loc, values_from = mat,
                  names_prefix = paste0(pref, "_"),
                  values_fill = 0, names_expand = TRUE)
  }
  
  agg_locdif <- function(df, qt_col, pref) {
    df |>
      filter(!is.na(TP_LOCALIZACAO_DIFERENCIADA)) |>
      mutate(ld = LOC_DIF_MAP[TP_LOCALIZACAO_DIFERENCIADA]) |>
      filter(!is.na(ld)) |>
      group_by(across(all_of(c(g_vars, "ld")))) |>
      summarise(mat = sum(.data[[qt_col]], na.rm = TRUE), .groups = "drop") |>
      mutate(ld = factor(ld, levels = LOC_DIF_NIVEIS)) |>
      pivot_wider(names_from = ld, values_from = mat,
                  names_prefix = paste0(pref, "_"),
                  values_fill = 0, names_expand = TRUE)
  }
  
  message("  [A] Escolas...")
  bloco_a <- base_chaves |>
    left_join(agg(dir)     |> summarise(n_esc_dir         = n_distinct(CO_ENTIDADE), .groups = "drop"), by = g_vars) |>
    left_join(agg(ind)     |> summarise(n_esc_ind_mun     = n_distinct(CO_ENTIDADE), .groups = "drop"), by = g_vars) |>
    left_join(agg(ind_est) |> summarise(n_esc_ind_mun_est = n_distinct(CO_ENTIDADE), .groups = "drop"), by = g_vars)
  
  message("  [B] Matrículas EB...")
  bloco_b <- base_chaves |>
    left_join(agg(dir) |> summarise(mat_bas_dir = sum(QT_MAT_BAS, na.rm = TRUE), .groups = "drop"), by = g_vars) |>
    left_join(agg(ind) |> summarise(mat_bas_ind = sum(QT_MAT_BAS, na.rm = TRUE), .groups = "drop"), by = g_vars) |>
    mutate(mat_bas_tot = replace_na(mat_bas_dir, 0) + replace_na(mat_bas_ind, 0))
  
  message("  [C] Matrículas EI...")
  c_dir <- agg(dir) |>
    summarise(
      mat_ei_dir     = sum(QT_MAT_INF_CRE, na.rm=TRUE) + sum(QT_MAT_INF_PRE, na.rm=TRUE),
      mat_ei_ti_dir  = sum(QT_MAT_INF_INT,     na.rm=TRUE),
      mat_cre_dir    = sum(QT_MAT_INF_CRE,     na.rm=TRUE),
      mat_cre_ti_dir = sum(QT_MAT_INF_CRE_INT, na.rm=TRUE),
      mat_pre_dir    = sum(QT_MAT_INF_PRE,     na.rm=TRUE),
      mat_pre_ti_dir = sum(QT_MAT_INF_PRE_INT, na.rm=TRUE),
      .groups = "drop")
  
  c_ind <- agg(ind) |>
    summarise(
      mat_ei_ind     = sum(QT_MAT_INF_CRE, na.rm=TRUE) + sum(QT_MAT_INF_PRE, na.rm=TRUE),
      mat_ei_ti_ind  = sum(QT_MAT_INF_INT,     na.rm=TRUE),
      mat_cre_ind    = sum(QT_MAT_INF_CRE,     na.rm=TRUE),
      mat_cre_ti_ind = sum(QT_MAT_INF_CRE_INT, na.rm=TRUE),
      mat_pre_ind    = sum(QT_MAT_INF_PRE,     na.rm=TRUE),
      mat_pre_ti_ind = sum(QT_MAT_INF_PRE_INT, na.rm=TRUE),
      .groups = "drop")
  
  bloco_c <- base_chaves |>
    left_join(c_dir, by = g_vars) |>
    left_join(c_ind, by = g_vars) |>
    mutate(
      mat_ei_tot    = replace_na(mat_ei_dir, 0)    + replace_na(mat_ei_ind,    0),
      mat_ei_ti_tot = replace_na(mat_ei_ti_dir, 0) + replace_na(mat_ei_ti_ind, 0),
      mat_cre_tot   = replace_na(mat_cre_dir, 0)   + replace_na(mat_cre_ind,   0),
      mat_pre_tot   = replace_na(mat_pre_dir, 0)   + replace_na(mat_pre_ind,   0))
  
  message("  [D] EI por localização...")
  bloco_d <- base_chaves |>
    left_join(agg_loc(dir, "QT_MAT_INF_CRE", "mat_cre_dir_loc"), by = g_vars) |>
    left_join(agg_loc(ind, "QT_MAT_INF_CRE", "mat_cre_ind_loc"), by = g_vars) |>
    left_join(agg_loc(dir, "QT_MAT_INF_PRE", "mat_pre_dir_loc"), by = g_vars) |>
    left_join(agg_loc(ind, "QT_MAT_INF_PRE", "mat_pre_ind_loc"), by = g_vars)
  
  message("  [E] EI por loc. diferenciada...")
  bloco_e <- base_chaves |>
    left_join(agg_locdif(tot, "QT_MAT_INF_CRE", "mat_cre_tot_ld"), by = g_vars) |>
    left_join(agg_locdif(tot, "QT_MAT_INF_PRE", "mat_pre_tot_ld"), by = g_vars)
  
  message("  [F] Docentes/turma...")
  doc_tur <- function(df, suf) {
    agg(df) |>
      summarise(
        sd_cre = sum(QT_DOC_INF_CRE, na.rm = TRUE),
        st_cre = sum(QT_TUR_INF_CRE, na.rm = TRUE),
        sd_pre = sum(QT_DOC_INF_PRE, na.rm = TRUE),
        st_pre = sum(QT_TUR_INF_PRE, na.rm = TRUE),
        .groups = "drop") |>
      mutate(
        !!paste0("doc_tur_cre_", suf) := if_else(st_cre > 0, sd_cre/st_cre, NA_real_),
        !!paste0("doc_tur_pre_", suf) := if_else(st_pre > 0, sd_pre/st_pre, NA_real_)) |>
      select(-sd_cre, -st_cre, -sd_pre, -st_pre)
  }
  bloco_f <- base_chaves |>
    left_join(doc_tur(dir, "dir"), by = g_vars) |>
    left_join(doc_tur(ind, "ind"), by = g_vars)
  
  message("  [G] Matrículas EI por escola...")
  mat_esc <- function(df, suf) {
    agg(df) |>
      summarise(
        s_ei  = sum(QT_MAT_INF_CRE, na.rm=TRUE) + sum(QT_MAT_INF_PRE, na.rm=TRUE),
        n_inf = sum(as.integer(!is.na(IN_INF) & IN_INF == "1")),
        .groups = "drop") |>
      mutate(!!paste0("mat_ei_esc_", suf) :=
               if_else(n_inf > 0, s_ei/n_inf, NA_real_)) |>
      select(-s_ei, -n_inf)
  }
  bloco_g <- base_chaves |>
    left_join(mat_esc(dir, "dir"), by = g_vars) |>
    left_join(mat_esc(ind, "ind"), by = g_vars)
  
  message("  [H] Infraestrutura EI...")
  infra <- function(df, suf) {
    df |>
      mutate(
        mat_ei = rowSums(cbind(QT_MAT_INF_CRE, QT_MAT_INF_PRE), na.rm = TRUE),
        i_av   = !is.na(IN_AREA_VERDE)              & IN_AREA_VERDE              == "1",
        i_ban  = !is.na(IN_BANHEIRO_EI)             & IN_BANHEIRO_EI             == "1",
        i_jog  = !is.na(IN_MATERIAL_PED_JOGOS)      & IN_MATERIAL_PED_JOGOS      == "1",
        i_art  = !is.na(IN_MATERIAL_PED_ARTISTICAS) & IN_MATERIAL_PED_ARTISTICAS == "1",
        i_4    = i_av & i_ban & i_jog & i_art,
        i_s4   = !i_av & !i_ban & !i_jog & !i_art) |>
      group_by(across(all_of(g_vars))) |>
      summarise(
        !!paste0("mat_ei_av_",  suf) := sum(mat_ei * i_av,  na.rm = TRUE),
        !!paste0("mat_ei_ban_", suf) := sum(mat_ei * i_ban, na.rm = TRUE),
        !!paste0("mat_ei_jog_", suf) := sum(mat_ei * i_jog, na.rm = TRUE),
        !!paste0("mat_ei_art_", suf) := sum(mat_ei * i_art, na.rm = TRUE),
        !!paste0("mat_ei_4_",   suf) := sum(mat_ei * i_4,   na.rm = TRUE),
        !!paste0("mat_ei_s4_",  suf) := sum(mat_ei * i_s4,  na.rm = TRUE),
        .groups = "drop")
  }
  bloco_h <- base_chaves |>
    left_join(infra(dir, "dir"), by = g_vars) |>
    left_join(infra(ind, "ind"), by = g_vars)
  
  message("  [I] Fund/EM/EJA/EP/EE...")
  bloco_i <- base_chaves |>
    left_join(
      agg(dir) |> summarise(
        mat_fund_ai_dir = sum(QT_MAT_FUND_AI, na.rm = TRUE),
        mat_fund_af_dir = sum(QT_MAT_FUND_AF, na.rm = TRUE),
        mat_med_dir     = sum(QT_MAT_MED,     na.rm = TRUE),
        mat_eja_dir     = sum(QT_MAT_EJA,     na.rm = TRUE),
        .groups = "drop"), by = g_vars) |>
    left_join(agg_locdif(tot, "QT_MAT_PROF", "mat_prof_tot_ld"), by = g_vars) |>
    left_join(agg_locdif(tot, "QT_MAT_ESP",  "mat_esp_tot_ld"),  by = g_vars)
  
  message("  [J] Cor/raça...")
  raca_cols <- c("QT_MAT_BAS_ND","QT_MAT_BAS_BRANCA","QT_MAT_BAS_PRETA",
                 "QT_MAT_BAS_PARDA","QT_MAT_BAS_AMARELA","QT_MAT_BAS_INDIGENA")
  raca_suf  <- c("nd","branca","preta","parda","amarela","indigena")
  raca <- function(df, seg) {
    agg(df) |>
      summarise(across(all_of(raca_cols), \(x) sum(x, na.rm = TRUE)), .groups = "drop") |>
      rename_with(~ paste0("mat_bas_", raca_suf, "_", seg), all_of(raca_cols))
  }
  bloco_j <- base_chaves |>
    left_join(raca(dir, "dir"), by = g_vars) |>
    left_join(raca(ind, "ind"), by = g_vars)
  
  message("  [K] EB por loc. diferenciada...")
  bloco_k <- base_chaves |>
    left_join(agg_locdif(dir, "QT_MAT_BAS", "mat_bas_dir_ld"), by = g_vars) |>
    left_join(agg_locdif(ind, "QT_MAT_BAS", "mat_bas_ind_ld"), by = g_vars)
  
  message("  [JOIN] Montando...")
  tabela <- list(bloco_a, bloco_b, bloco_c, bloco_d, bloco_e,
                 bloco_f, bloco_g, bloco_h, bloco_i, bloco_j, bloco_k) |>
    reduce(left_join, by = g_vars)
  
  message("  [ok] ", format(nrow(tabela), big.mark = ","),
          " linhas | ", ncol(tabela), " colunas")
  tabela
}

message("[01] Helpers carregados.")




# =============================================================
# 02 — LEITURA CRUA
# Lê microdados INEP 2015–2025:
#   • 2015–2021, 2023–2024: via educabR
#   • 2022: CSV monolítico direto
#   • 2025: join de 4 CSVs temáticos por CO_ENTIDADE
# Saída: censo_2015_2024, censo_2025
# =============================================================

# ── educabR: 2015–2021 e 2023–2024 ──────────────────────────
anos_educabr <- c(2015:2021, 2023:2024)
lista_anos <- map(anos_educabr, carregar_ano) |>
  setNames(as.character(anos_educabr)) |>
  Filter(Negate(is.null), x = _)

message("\n>>> Anos carregados via educabR: ",
        paste(names(lista_anos), collapse = ", "))

# ── CSV manual: 2022 ────────────────────────────────────────
message("\n>>> Carregando 2022 via CSV...")
censo_2022_bruto <- read_delim(
  file.path(PATHS$raw, "microdados_ed_basica_2022.csv"),
  delim  = ";",
  locale = locale(encoding = "latin1", decimal_mark = ",", grouping_mark = "."),
  col_types = cols(.default = col_character()),
  progress = TRUE, show_col_types = FALSE
)
message("  Bruto 2022: ", format(nrow(censo_2022_bruto), big.mark = "."), " linhas")

censo_2022 <- padronizar_df(censo_2022_bruto, 2022)

# ── Empilhamento 2015–2024 ──────────────────────────────────
lista_chr      <- map(lista_anos, ~ mutate(.x, across(everything(), as.character)))
censo_2022_chr <- mutate(censo_2022, across(everything(), as.character))

censo_2015_2024 <- bind_rows(unname(lista_chr), censo_2022_chr) |>
  arrange(NU_ANO_CENSO)

message(">>> BANCO 2015–2024: ",
        format(nrow(censo_2015_2024), big.mark = ","), " linhas | ",
        ncol(censo_2015_2024), " colunas")

# ── 2025: Join dos 4 CSVs temáticos ─────────────────────────
message("\n>>> Lendo arquivos CSV de 2025...")
tab_escola    <- ler_csv_inep(file.path(PATHS$raw, "Tabela_Escola_2025.csv"))
tab_matricula <- ler_csv_inep(file.path(PATHS$raw, "Tabela_Matricula_2025.csv"))
tab_docente   <- ler_csv_inep(file.path(PATHS$raw, "Tabela_Docente_2025.csv"))
tab_turma     <- ler_csv_inep(file.path(PATHS$raw, "Tabela_Turma_2025.csv"))

# Variáveis esperadas por tabela 2025
# NOTA: IN_INF / IN_INF_CRE / IN_INF_PRE ausentes em 2025
vars_escola_2025 <- c(
  "NU_ANO_CENSO","SG_UF","NO_MUNICIPIO","CO_MUNICIPIO",
  "NO_MICRORREGIAO","CO_MICRORREGIAO","NO_DISTRITO","CO_DISTRITO",
  "CO_ENTIDADE","DS_ENDERECO","NU_ENDERECO","DS_COMPLEMENTO","CO_CEP",
  "TP_DEPENDENCIA","TP_CATEGORIA_ESCOLA_PRIVADA","TP_LOCALIZACAO",
  "TP_LOCALIZACAO_DIFERENCIADA","TP_SITUACAO_FUNCIONAMENTO",
  "IN_PODER_PUBLICO_PARCERIA","TP_PODER_PUBLICO_PARCERIA",
  "IN_MANT_ESCOLA_PRIVADA_EMP","IN_MANT_ESCOLA_PRIV_ONG_OSCIP",
  "IN_MANT_ESCOLA_PRIVADA_SIND","IN_ALIMENTACAO","QT_TRANSP_PUBLICO"
)

vars_matricula_2025 <- c(
  "CO_ENTIDADE",
  "QT_MAT_BAS","QT_MAT_INF_CRE","QT_MAT_INF_PRE",
  "QT_MAT_FUND_AI","QT_MAT_FUND_AF","QT_MAT_MED","QT_MAT_PROF",
  "QT_MAT_EJA","QT_MAT_EJA_FUND_AI","QT_MAT_EJA_FUND_AF","QT_MAT_EJA_MED",
  "QT_MAT_ESP","QT_MAT_ESP_CC","QT_MAT_ESP_CE",
  "QT_MAT_BAS_FEM","QT_MAT_BAS_MASC","QT_MAT_BAS_ND",
  "QT_MAT_BAS_BRANCA","QT_MAT_BAS_PRETA","QT_MAT_BAS_PARDA",
  "QT_MAT_BAS_AMARELA","QT_MAT_BAS_INDIGENA",
  "QT_MAT_INF_INT","QT_MAT_INF_CRE_INT","QT_MAT_INF_PRE_INT",
  "QT_MAT_FUND_INT","QT_MAT_FUND_AI_INT","QT_MAT_FUND_AF_INT",
  "QT_MAT_MED_INT","QT_MAT_ZR_URB","QT_MAT_ZR_RUR",
  "QT_TRANSP_PUBLICO"
)

vars_docente_2025 <- c("CO_ENTIDADE","QT_DOC_INF_CRE","QT_DOC_INF_PRE")

vars_turma_2025 <- c(
  "CO_ENTIDADE",
  "QT_TUR_INF_CRE","QT_TUR_INF_PRE","QT_TUR_INF_INT",
  "QT_TUR_INF_CRE_INT","QT_TUR_INF_PRE_INT"
)

# Agregação por escola (com safe_max para QT_TRANSP_PUBLICO)
mat_ag <- agregar_por_escola(tab_matricula, vars_matricula_2025, "Matricula")
doc_ag <- agregar_por_escola(tab_docente,   vars_docente_2025,   "Docente")
tur_ag <- agregar_por_escola(tab_turma,     vars_turma_2025,     "Turma")

# Join final: Escola ⊕ Matrícula ⊕ Docente ⊕ Turma
# CRÍTICO: left_join (não inner) — escolas sem matrícula contam
base_escola_2025 <- tab_escola |>
  select(all_of(intersect(vars_escola_2025, names(tab_escola))))

censo_2025_raw <- base_escola_2025 |>
  left_join(mat_ag, by = "CO_ENTIDADE") |>
  left_join(doc_ag, by = "CO_ENTIDADE") |>
  left_join(tur_ag, by = "CO_ENTIDADE") |>
  mutate(NU_ANO_CENSO = "2025")

message("  Após joins: ", format(nrow(censo_2025_raw), big.mark = ","),
        " linhas | ", ncol(censo_2025_raw), " colunas")

stopifnot(
  "Join de 2025 não pode duplicar linhas" =
    nrow(censo_2025_raw) == nrow(base_escola_2025)
)

# Preenche colunas ausentes em 2025 com NA estrutural
ausentes_2025 <- setdiff(VARIAVEIS_INTERESSE, names(censo_2025_raw))
if (length(ausentes_2025) > 0) {
  message("  [NA estrutural 2025]: ", paste(ausentes_2025, collapse = ", "))
  for (col in ausentes_2025) censo_2025_raw[[col]] <- NA
}

censo_2025 <- censo_2025_raw |>
  mutate(across(everything(), as.character)) |>
  select(all_of(VARIAVEIS_INTERESSE))

message("  [ok] 2025 final: ", format(nrow(censo_2025), big.mark = ","),
        " escolas | ", ncol(censo_2025), " colunas")

message("\n[02] Leitura crua concluída.")






# =============================================================
# 03 — TIPAGEM E HARMONIZAÇÃO SEMÂNTICA
# Saída: censo_completo (todos os anos consolidados e tipados)
# =============================================================

message("\n>>> Consolidando banco completo 2015–2025...")
censo_completo <- bind_rows(censo_2015_2024, censo_2025) |>
  arrange(NU_ANO_CENSO)

message(">>> BANCO COMPLETO: ",
        format(nrow(censo_completo), big.mark = ","), " linhas | ",
        ncol(censo_completo), " colunas")

# ── 3.1 Coerção numérica ────────────────────────────────────
censo_completo <- censo_completo |>
  mutate(across(all_of(VARS_NUMERICAS),
                \(x) as.numeric(str_replace_all(x, ",", "."))))

# ── 3.2 TP_PODER_PUBLICO_PARCERIA: NA→0 a partir de 2022 ────
# Mudança INEP: a partir de 2022 só conveniadas têm valor preenchido;
# escolas sem convênio passam a ter NA em vez de 0.
# Substituição preserva comparabilidade longitudinal.
message("\n>>> Corrigindo NA semântico em TP_PODER_PUBLICO_PARCERIA...")
censo_completo <- censo_completo |>
  mutate(
    TP_PODER_PUBLICO_PARCERIA = case_when(
      NU_ANO_CENSO >= 2022 & is.na(TP_PODER_PUBLICO_PARCERIA) ~ 0L,
      TRUE ~ as.integer(TP_PODER_PUBLICO_PARCERIA)
    ),
    IN_CONVENIO = if_else(TP_PODER_PUBLICO_PARCERIA > 0, 1L, 0L)
  )

# ── 3.3 TP_LOCALIZACAO_DIFERENCIADA: replicação retroativa ──
# Classificação estrutural — propaga do ano mais recente disponível.
message("\n>>> Replicando TP_LOCALIZACAO_DIFERENCIADA por CO_ENTIDADE...")
ref_localizacao <- censo_completo |>
  filter(NU_ANO_CENSO == 2025, !is.na(TP_LOCALIZACAO_DIFERENCIADA)) |>
  group_by(CO_ENTIDADE) |>
  slice_max(NU_ANO_CENSO, n = 1, with_ties = FALSE) |>
  ungroup() |>
  select(CO_ENTIDADE, TP_LOCALIZACAO_DIFERENCIADA_REF = TP_LOCALIZACAO_DIFERENCIADA)

censo_completo <- censo_completo |>
  left_join(ref_localizacao, by = "CO_ENTIDADE") |>
  mutate(
    TP_LOCALIZACAO_DIFERENCIADA = case_when(
      !is.na(TP_LOCALIZACAO_DIFERENCIADA_REF) ~ TP_LOCALIZACAO_DIFERENCIADA_REF,
      !is.na(TP_LOCALIZACAO_DIFERENCIADA)     ~ TP_LOCALIZACAO_DIFERENCIADA,
      TRUE ~ NA_character_
    )
  ) |>
  select(-TP_LOCALIZACAO_DIFERENCIADA_REF)

# ── 3.4 Labels semânticos ───────────────────────────────────
censo_completo <- censo_completo |>
  mutate(
    TP_LOCALIZACAO_DIFERENCIADA = case_when(
      TP_LOCALIZACAO_DIFERENCIADA == "0" ~ "Escola fora de área diferenciada",
      TP_LOCALIZACAO_DIFERENCIADA == "1" ~ "Área de assentamento",
      TP_LOCALIZACAO_DIFERENCIADA == "2" ~ "Terra indígena",
      TP_LOCALIZACAO_DIFERENCIADA == "3" ~ "Área com comunidade remanescente de quilombos",
      TP_LOCALIZACAO_DIFERENCIADA == "8" ~ "Área com povos e comunidades tradicionais",
      TRUE ~ NA_character_
    ),
    TP_CATEGORIA_ESCOLA_PRIVADA = case_when(
      TP_CATEGORIA_ESCOLA_PRIVADA == "1" ~ "Particular",
      TP_CATEGORIA_ESCOLA_PRIVADA == "2" ~ "Comunitária",
      TP_CATEGORIA_ESCOLA_PRIVADA == "3" ~ "Confessional",
      TP_CATEGORIA_ESCOLA_PRIVADA == "4" ~ "Filantrópica",
      TRUE ~ NA_character_
    )
  )

# ── 3.5 Patch: variáveis de infraestrutura ──────────────────
message("\n>>> Injetando variáveis de infraestrutura EI...")

lista_infra <- map(anos_educabr, function(ano) {
  message("  ano: ", ano)
  df <- tryCatch(educabR::get_censo_escolar(ano, quiet = TRUE),
                 error = function(e) NULL)
  if (is.null(df)) return(NULL)
  extrair_infra(df, ano)
}) |> Filter(Negate(is.null), x = _)

infra_2022 <- read_delim(
  file.path(PATHS$raw, "microdados_ed_basica_2022.csv"),
  delim  = ";", locale = locale(encoding = "latin1"),
  col_select = c("CO_ENTIDADE", all_of(NOVAS_VARS_INFRA)),
  col_types  = cols(.default = col_character()),
  show_col_types = FALSE
) |> extrair_infra(2022)

infra_2025 <- read_delim(
  file.path(PATHS$raw, "Tabela_Escola_2025.csv"),
  delim = ";", locale = locale(encoding = "latin1"),
  col_select = c("CO_ENTIDADE", all_of(NOVAS_VARS_INFRA)),
  col_types  = cols(.default = col_character()),
  show_col_types = FALSE
) |> extrair_infra(2025)

infra_completo <- bind_rows(unname(lista_infra), infra_2022, infra_2025) |>
  arrange(NU_ANO_CENSO)

dupl <- infra_completo |> count(CO_ENTIDADE, NU_ANO_CENSO) |> filter(n > 1)
stopifnot("Duplicatas em infra_completo!" = nrow(dupl) == 0)

censo_completo <- censo_completo |>
  mutate(CO_ENTIDADE  = as.character(CO_ENTIDADE),
         NU_ANO_CENSO = as.character(NU_ANO_CENSO))

n_antes <- nrow(censo_completo)
censo_completo <- censo_completo |>
  left_join(infra_completo, by = c("CO_ENTIDADE", "NU_ANO_CENSO"))
stopifnot("Join de infra duplicou linhas!" = nrow(censo_completo) == n_antes)

message("\n[03] Tipagem e harmonização concluídas.")







# =============================================================
# 04 — EMPILHAMENTO DOS TRÊS CENSOS PANEL
# Saída: censo_completo, censo_capitais, censo_saopaulo (+ .rds)
# =============================================================

# ── Scaffold IBGE para garantir 96 distritos em SP ──────────
scaffold_sp <- buscar_scaffold_sp()

# ── Subsets de capitais e São Paulo ─────────────────────────
censo_capitais <- censo_completo |>
  filter(CO_MUNICIPIO %in% as.character(CAPITAIS_CO))

# Para SP: força aparição de todos os 96 distritos em todos os anos
# via crossing × scaffold (mesmo distritos só-indiretos como Marsilac)
censo_saopaulo_base <- censo_completo |>
  filter(CO_MUNICIPIO == as.character(CO_MUN_SP)) |>
  mutate(CO_DISTRITO = as.numeric(CO_DISTRITO))

if (!is.null(scaffold_sp)) {
  ano_distrito_scaffold <- crossing(
    NU_ANO_CENSO = as.character(ANOS),
    scaffold_sp |> select(CO_DISTRITO)
  )
  censo_saopaulo <- censo_saopaulo_base |>
    right_join(ano_distrito_scaffold,
               by = c("NU_ANO_CENSO", "CO_DISTRITO"))
} else {
  warning("Scaffold IBGE indisponível — usando distritos dos próprios dados.")
  censo_saopaulo <- censo_saopaulo_base
}

# ── Sanity checks ───────────────────────────────────────────
stopifnot(
  "censo_completo deve ter todos os anos 2015-2025" =
    all(as.character(ANOS) %in% unique(censo_completo$NU_ANO_CENSO)),
  "censo_capitais deve ter 27 municípios" =
    n_distinct(censo_capitais$CO_MUNICIPIO) == 27L
)

if (!is.null(scaffold_sp)) {
  stopifnot(
    "censo_saopaulo deve cobrir 96 distritos" =
      n_distinct(censo_saopaulo$CO_DISTRITO) == 96L
  )
}

message("\n[04] Censos empilhados:")
message("    censo_completo: ", format(nrow(censo_completo),  big.mark = ","))
message("    censo_capitais: ", format(nrow(censo_capitais),  big.mark = ","),
        " (", n_distinct(censo_capitais$CO_MUNICIPIO), " municípios)")
message("    censo_saopaulo: ", format(nrow(censo_saopaulo),  big.mark = ","),
        " (", n_distinct(censo_saopaulo$CO_DISTRITO), " distritos)")

# ── Persistência ────────────────────────────────────────────
saveRDS(censo_completo, file.path(PATHS$processed, "censo_completo_2015_2025.rds"))
saveRDS(censo_capitais, file.path(PATHS$processed, "censo_capitais_2015_2025.rds"))
saveRDS(censo_saopaulo, file.path(PATHS$processed, "censo_saopaulo_2015_2025.rds"))

fs <- file.size(file.path(PATHS$processed,
                          c("censo_completo_2015_2025.rds",
                            "censo_capitais_2015_2025.rds",
                            "censo_saopaulo_2015_2025.rds")))
names(fs) <- c("completo", "capitais", "saopaulo")
message("Tamanhos (MB): ",
        paste(names(fs), round(fs/1024^2, 1), sep=":", collapse = " | "))








# =============================================================
# 05 — TABELAS ANALÍTICAS
# Regenera as 3 tabs SEMPRE a partir dos censos atuais.
# Princípio: estado é tão crítico quanto código.
# =============================================================

message("\n>>> TABELA: Todos os municípios")
tab_municipios <- construir_tabela(censo_completo, G_MUN)

message("\n>>> TABELA: Capitais")
tab_capitais <- construir_tabela(censo_capitais, G_MUN,
                                 filtro_dir = f_dir_capitais)

message("\n>>> TABELA: São Paulo — por distrito")
tab_saopaulo <- construir_tabela(censo_saopaulo, G_DIST)

# ── Acoplar labels via scaffold IBGE + constantes de SP ─────
# Labels só entram DEPOIS da agregação. Isto é o que evita o
# bug de 2015–2023 onde NO_DISTRITO é NA estrutural.
labels_sp <- scaffold_sp |>
  rename(NO_DISTRITO = NO_DISTRITO_IBGE) |>
  mutate(
    SG_UF            = "SP",
    NO_MUNICIPIO     = "São Paulo",
    CO_MUNICIPIO     = as.character(CO_MUN_SP),
    NO_MICRORREGIAO  = "Metropolitana de São Paulo",
    CO_MICRORREGIAO  = "35061"
  )

tab_saopaulo <- tab_saopaulo |>
  left_join(labels_sp, by = "CO_DISTRITO") |>
  select(NU_ANO_CENSO, SG_UF, NO_MUNICIPIO, CO_MUNICIPIO,
         NO_MICRORREGIAO, CO_MICRORREGIAO, NO_DISTRITO, CO_DISTRITO,
         everything())

# ── Suplemento SP: escolas por tipo (CRE / PRE) ─────────────
# IN_INF_CRE/IN_INF_PRE são NA em 2025 — registro fica vazio nesse ano.
message("\n>>> Suplemento SP: escola-tipo por distrito...")
sp_tipo <- local({
  dir_sp <- f_dir(censo_saopaulo)
  ind_sp <- f_ind_mun(censo_saopaulo)
  list(
    dir_sp |> filter(!is.na(IN_INF_CRE), IN_INF_CRE == "1") |>
      group_by(CO_DISTRITO, NU_ANO_CENSO) |>
      summarise(n_cre_dir = n_distinct(CO_ENTIDADE), .groups = "drop"),
    ind_sp |> filter(!is.na(IN_INF_CRE), IN_INF_CRE == "1") |>
      group_by(CO_DISTRITO, NU_ANO_CENSO) |>
      summarise(n_cre_ind = n_distinct(CO_ENTIDADE), .groups = "drop"),
    dir_sp |> filter(!is.na(IN_INF_PRE), IN_INF_PRE == "1") |>
      group_by(CO_DISTRITO, NU_ANO_CENSO) |>
      summarise(n_pre_dir = n_distinct(CO_ENTIDADE), .groups = "drop"),
    ind_sp |> filter(!is.na(IN_INF_PRE), IN_INF_PRE == "1") |>
      group_by(CO_DISTRITO, NU_ANO_CENSO) |>
      summarise(n_pre_ind = n_distinct(CO_ENTIDADE), .groups = "drop")
  ) |> reduce(left_join, by = c("CO_DISTRITO", "NU_ANO_CENSO"))
})

tab_saopaulo <- left_join(tab_saopaulo, sp_tipo,
                          by = c("CO_DISTRITO", "NU_ANO_CENSO"))

# ── Equalização de schema entre as 3 tabs ───────────────────
tab_capitais <- completar_colunas(tab_capitais, tab_municipios)
tab_saopaulo <- completar_colunas(tab_saopaulo, tab_municipios)

# ── Zerar NAs em colunas de contagem/indiretas ──────────────
cols_zerar <- tab_municipios |>
  select(matches("_ind")) |>
  select(-matches("doc_tur_|mat_ei_esc_")) |>
  names()

for (tbl_nm in c("tab_municipios", "tab_capitais", "tab_saopaulo")) {
  tbl  <- get(tbl_nm)
  cols <- intersect(cols_zerar, names(tbl))
  assign(tbl_nm, mutate(tbl, across(all_of(cols), ~ replace_na(., 0))))
}

# ── Persistência ────────────────────────────────────────────
saveRDS(tab_municipios, file.path(PATHS$processed, "tab_municipios_2015_2025.rds"))
saveRDS(tab_capitais,   file.path(PATHS$processed, "tab_capitais_2015_2025.rds"))
saveRDS(tab_saopaulo,   file.path(PATHS$processed, "tab_saopaulo_2015_2025.rds"))

message("\n[05] Tabelas analíticas:")
message("    tab_municipios: ", paste(dim(tab_municipios), collapse = " × "))
message("    tab_capitais:   ", paste(dim(tab_capitais),   collapse = " × "))
message("    tab_saopaulo:   ", paste(dim(tab_saopaulo),   collapse = " × "))




# =============================================================
# 06 — VALIDAÇÃO TERMINAL (GATE) — VERSÃO TOTAL
# Reconstrói cada tab a partir do CSV bruto INEP e compara
# COLUNA A COLUNA. Cobertura: 100% das ~98 variáveis.
# =============================================================

message("\n[06] Iniciando validação total (coluna a coluna)...\n")

# ─────────────────────────────────────────────────────────────
# 6.1 — Helpers internos do validador total
# ─────────────────────────────────────────────────────────────

# Lê CSV bruto monolítico (2015–2024)
ler_bruto_monolitico <- function(ano) {
  csv_path <- file.path(PATHS$raw, sprintf("microdados_ed_basica_%d.csv", ano))
  if (!file.exists(csv_path)) {
    csv_alt <- file.path(PATHS$raw, sprintf("microdados_ed_basica_%d.CSV", ano))
    if (file.exists(csv_alt)) csv_path <- csv_alt
    else stop("CSV não encontrado para ", ano)
  }
  message(">>> Lendo bruto ", ano, "...")
  read_delim(csv_path, delim = ";",
             locale = locale(encoding = "latin1", decimal_mark = ",",
                             grouping_mark = "."),
             col_types = cols(.default = col_character()),
             show_col_types = FALSE, progress = FALSE) |>
    rename_with(toupper)
}

# Reconstrói o bruto 2025 a partir dos 4 CSVs temáticos
ler_bruto_2025 <- function() {
  message(">>> Reconstituindo bruto 2025...")
  loc <- locale(encoding = "latin1", decimal_mark = ",", grouping_mark = ".")
  
  esc <- read_delim(file.path(PATHS$raw, "Tabela_Escola_2025.csv"),
                    delim = ";", locale = loc,
                    col_types = cols(.default = col_character()),
                    show_col_types = FALSE, progress = FALSE) |>
    rename_with(toupper)
  
  mat <- read_delim(file.path(PATHS$raw, "Tabela_Matricula_2025.csv"),
                    delim = ";", locale = loc,
                    col_types = cols(.default = col_character()),
                    show_col_types = FALSE, progress = FALSE) |>
    rename_with(toupper) |>
    select(any_of(c("CO_ENTIDADE", setdiff(vars_matricula_2025, "CO_ENTIDADE"))))
  
  doc <- read_delim(file.path(PATHS$raw, "Tabela_Docente_2025.csv"),
                    delim = ";", locale = loc,
                    col_types = cols(.default = col_character()),
                    show_col_types = FALSE, progress = FALSE) |>
    rename_with(toupper) |>
    select(any_of(c("CO_ENTIDADE", setdiff(vars_docente_2025, "CO_ENTIDADE"))))
  
  tur <- read_delim(file.path(PATHS$raw, "Tabela_Turma_2025.csv"),
                    delim = ";", locale = loc,
                    col_types = cols(.default = col_character()),
                    show_col_types = FALSE, progress = FALSE) |>
    rename_with(toupper) |>
    select(any_of(c("CO_ENTIDADE", setdiff(vars_turma_2025, "CO_ENTIDADE"))))
  
  esc |>
    left_join(mat, by = "CO_ENTIDADE") |>
    left_join(doc, by = "CO_ENTIDADE") |>
    left_join(tur, by = "CO_ENTIDADE") |>
    mutate(NU_ANO_CENSO = "2025")
}

# Aplica os tratamentos semânticos do pipeline ao bruto
preparar_bruto_pipeline <- function(bruto, ano, ref_localizacao = NULL) {
  
  if ("TP_CONVENIO_PODER_PUBLICO" %in% names(bruto) &&
      !"TP_PODER_PUBLICO_PARCERIA" %in% names(bruto)) {
    bruto <- rename(bruto, TP_PODER_PUBLICO_PARCERIA = TP_CONVENIO_PODER_PUBLICO)
  }
  if ("IN_CONVENIADA_PP" %in% names(bruto) &&
      !"IN_PODER_PUBLICO_PARCERIA" %in% names(bruto)) {
    bruto <- rename(bruto, IN_PODER_PUBLICO_PARCERIA = IN_CONVENIADA_PP)
  }
  
  vars_completas <- union(VARIAVEIS_INTERESSE, NOVAS_VARS_INFRA)
  ausentes <- setdiff(vars_completas, names(bruto))
  for (col in ausentes) bruto[[col]] <- NA_character_
  bruto <- select(bruto, all_of(vars_completas))
  
  bruto <- mutate(bruto, across(all_of(intersect(VARS_NUMERICAS, names(bruto))),
                                \(x) as.numeric(str_replace_all(x, ",", "."))))
  
  bruto <- mutate(bruto,
                  TP_PODER_PUBLICO_PARCERIA = case_when(
                    ano >= 2022 & is.na(TP_PODER_PUBLICO_PARCERIA) ~ 0L,
                    TRUE ~ as.integer(TP_PODER_PUBLICO_PARCERIA)
                  )
  )
  
  # Garante tipo character em CO_ENTIDADE para o join
  bruto <- mutate(bruto, CO_ENTIDADE = as.character(CO_ENTIDADE))
  
  # Replicação retroativa de TP_LOCALIZACAO_DIFERENCIADA
  if (!is.null(ref_localizacao)) {
    bruto <- bruto |>
      left_join(ref_localizacao, by = "CO_ENTIDADE") |>
      mutate(
        TP_LOCALIZACAO_DIFERENCIADA = case_when(
          !is.na(TP_LOCALIZACAO_DIFERENCIADA_REF) ~ TP_LOCALIZACAO_DIFERENCIADA_REF,
          !is.na(TP_LOCALIZACAO_DIFERENCIADA)     ~ TP_LOCALIZACAO_DIFERENCIADA,
          TRUE ~ NA_character_
        )
      ) |>
      select(-TP_LOCALIZACAO_DIFERENCIADA_REF)
  }
  
  bruto <- mutate(bruto,
                  TP_LOCALIZACAO_DIFERENCIADA = case_when(
                    TP_LOCALIZACAO_DIFERENCIADA == "0" ~ "Escola fora de área diferenciada",
                    TP_LOCALIZACAO_DIFERENCIADA == "1" ~ "Área de assentamento",
                    TP_LOCALIZACAO_DIFERENCIADA == "2" ~ "Terra indígena",
                    TP_LOCALIZACAO_DIFERENCIADA == "3" ~ "Área com comunidade remanescente de quilombos",
                    TP_LOCALIZACAO_DIFERENCIADA == "8" ~ "Área com povos e comunidades tradicionais",
                    TRUE ~ NA_character_
                  )
  )
  
  bruto |> mutate(NU_ANO_CENSO = as.character(ano))
}

# Compara coluna a coluna entre tab do pipeline e tab de referência
comparar_total <- function(tab_pipeline, tab_referencia, chave, ano, nivel) {
  
  tab_p <- tab_pipeline   |> filter(NU_ANO_CENSO == as.character(ano))
  tab_r <- tab_referencia |> filter(NU_ANO_CENSO == as.character(ano))
  
  cols_comuns <- intersect(names(tab_p), names(tab_r))
  cols_chave  <- chave
  cols_label  <- c("NU_ANO_CENSO","SG_UF","NO_MUNICIPIO","CO_MUNICIPIO",
                   "NO_DISTRITO","CO_DISTRITO","NO_MICRORREGIAO","CO_MICRORREGIAO")
  cols_check  <- setdiff(cols_comuns, c(cols_chave, cols_label))
  
  comp <- tab_p |>
    select(all_of(c(cols_chave, cols_check))) |>
    rename_with(~ paste0("p_", .), all_of(cols_check)) |>
    full_join(
      tab_r |>
        select(all_of(c(cols_chave, cols_check))) |>
        rename_with(~ paste0("r_", .), all_of(cols_check)),
      by = cols_chave
    )
  
  divergencias_por_col <- map_int(cols_check, function(col) {
    p <- comp[[paste0("p_", col)]]
    r <- comp[[paste0("r_", col)]]
    
    p_norm <- if_else(is.na(p), 0, as.numeric(p))
    r_norm <- if_else(is.na(r), 0, as.numeric(r))
    
    if (grepl("^(doc_tur|mat_ei_esc)", col)) {
      sum(abs(p_norm - r_norm) > 1e-6, na.rm = TRUE)
    } else {
      sum(p_norm != r_norm, na.rm = TRUE)
    }
  }) |> setNames(cols_check)
  
  list(
    n_unidades         = nrow(comp),
    n_colunas          = length(cols_check),
    cols_divergentes   = divergencias_por_col[divergencias_por_col > 0],
    total_divergencias = sum(divergencias_por_col),
    tudo_ok            = sum(divergencias_por_col) == 0
  )
}

# ─────────────────────────────────────────────────────────────
# 6.2 — Construir ref_localizacao ANTES do loop
# ─────────────────────────────────────────────────────────────

# Mapa de labels → códigos (reverte o que 03_tipagem aplicou)
labels_para_codigos <- c(
  "Escola fora de área diferenciada"              = "0",
  "Área de assentamento"                          = "1",
  "Terra indígena"                                = "2",
  "Área com comunidade remanescente de quilombos" = "3",
  "Área com povos e comunidades tradicionais"     = "8"
)

ref_localizacao <- censo_completo |>
  mutate(CO_ENTIDADE = as.character(CO_ENTIDADE)) |>
  filter(NU_ANO_CENSO == "2025", !is.na(TP_LOCALIZACAO_DIFERENCIADA)) |>
  group_by(CO_ENTIDADE) |>
  slice_max(NU_ANO_CENSO, n = 1, with_ties = FALSE) |>
  ungroup() |>
  select(CO_ENTIDADE,
         TP_LOCALIZACAO_DIFERENCIADA_REF = TP_LOCALIZACAO_DIFERENCIADA) |>
  mutate(TP_LOCALIZACAO_DIFERENCIADA_REF =
           unname(labels_para_codigos[TP_LOCALIZACAO_DIFERENCIADA_REF]))

stopifnot(
  "ref_localizacao precisa ter CO_ENTIDADE como character" =
    is.character(ref_localizacao$CO_ENTIDADE),
  "ref_localizacao precisa ter linhas" =
    nrow(ref_localizacao) > 0
)

message("  [ref_localizacao] ",
        format(nrow(ref_localizacao), big.mark = ","),
        " escolas com label 2025 para replicação retroativa")

# ─────────────────────────────────────────────────────────────
# 6.3 — Loop principal: 11 anos × 3 níveis = 33 testes
# ─────────────────────────────────────────────────────────────

resultados_total <- list()
ANOS_VALIDAR    <- 2015:2025

for (ano in ANOS_VALIDAR) {
  message("\n--- Ano ", ano, " ---")
  
  bruto <- if (ano == 2025) ler_bruto_2025() else ler_bruto_monolitico(ano)
  bruto_tratado <- preparar_bruto_pipeline(bruto, ano, ref_localizacao)
  
  ref_mun <- construir_tabela(bruto_tratado, G_MUN)
  
  ref_cap <- bruto_tratado |>
    filter(CO_MUNICIPIO %in% CAPITAIS_CO) |>
    construir_tabela(G_MUN, filtro_dir = f_dir_capitais)
  
  ref_sp <- bruto_tratado |>
    filter(CO_MUNICIPIO == CO_MUN_SP) |>
    mutate(CO_DISTRITO = as.numeric(CO_DISTRITO)) |>
    construir_tabela(G_DIST)
  
  ref_cap <- completar_colunas(ref_cap, ref_mun)
  ref_sp  <- completar_colunas(ref_sp,  ref_mun)
  
  cols_zerar <- ref_mun |>
    select(matches("_ind")) |>
    select(-matches("doc_tur_|mat_ei_esc_")) |>
    names()
  
  ref_mun <- mutate(ref_mun, across(any_of(cols_zerar), ~ replace_na(., 0)))
  ref_cap <- mutate(ref_cap, across(any_of(cols_zerar), ~ replace_na(., 0)))
  ref_sp  <- mutate(ref_sp,  across(any_of(cols_zerar), ~ replace_na(., 0)))
  
  resultados_total[[sprintf("%d_municipio", ano)]] <-
    comparar_total(tab_municipios, ref_mun, "CO_MUNICIPIO", ano, "municipio")
  resultados_total[[sprintf("%d_capital", ano)]] <-
    comparar_total(tab_capitais,   ref_cap, "CO_MUNICIPIO", ano, "capital")
  resultados_total[[sprintf("%d_distrito_sp", ano)]] <-
    comparar_total(tab_saopaulo,   ref_sp,  "CO_DISTRITO",  ano, "distrito_sp")
  
  for (nivel in c("municipio", "capital", "distrito_sp")) {
    r <- resultados_total[[sprintf("%d_%s", ano, nivel)]]
    status <- if (r$tudo_ok) "OK" else paste0("FALHA (", r$total_divergencias, " divs)")
    message(sprintf("    %s [%s]: %d unidades x %d colunas -> %s",
                    nivel, ano, r$n_unidades, r$n_colunas, status))
    if (!r$tudo_ok && length(r$cols_divergentes) > 0) {
      message("      Colunas divergentes: ",
              paste(names(r$cols_divergentes), collapse = ", "))
    }
  }
  
  rm(bruto, bruto_tratado, ref_mun, ref_cap, ref_sp); gc(verbose = FALSE)
}

# ─────────────────────────────────────────────────────────────
# 6.4 — Placar consolidado
# ─────────────────────────────────────────────────────────────

placar <- tibble(
  combinacao       = names(resultados_total),
  unidades         = map_int(resultados_total, ~ .x$n_unidades),
  colunas_testadas = map_int(resultados_total, ~ .x$n_colunas),
  total_celulas    = unidades * colunas_testadas,
  divergencias     = map_int(resultados_total, ~ .x$total_divergencias),
  status           = if_else(divergencias == 0, "APROVADO", "FALHA")
)

message("\n=== PLACAR VALIDAÇÃO TOTAL ===")
print(placar, n = Inf)

total_celulas <- sum(placar$total_celulas)
total_divs    <- sum(placar$divergencias)

message(sprintf("\n>>> Total: %s células testadas | %d divergências",
                format(total_celulas, big.mark = ","), total_divs))

# ─────────────────────────────────────────────────────────────
# 6.5 — Gate
# ─────────────────────────────────────────────────────────────

if (total_divs > 0) {
  message("\n!! FALHA DE VALIDAÇÃO !!")
  falhas <- placar |> filter(status == "FALHA")
  print(falhas)
  stop("Pipeline interrompido: ", total_divs, " divergência(s) em ",
       nrow(falhas), " combinação(ões).")
}

# Compatibilidade: variável esperada por 07_export.R
resultados <- map(resultados_total, function(r) {
  tibble(tudo_ok = rep(r$tudo_ok, max(r$n_unidades, 1)))
})

message("\n[06] Validação TOTAL aprovada: 0 divergências em ",
        format(total_celulas, big.mark = ","), " células testadas (",
        length(resultados_total), " combinações ano × nível).")




# =============================================================
# 07 — EXPORTAÇÃO XLSX + CODEBOOK
# Gera:
#   • outputs/tabelas_analiticas_2015_2025.xlsx — 3 sheets
#   • outputs/codebook_censo_lemann_cem.xlsx   — notas + dicionário
# =============================================================

message("\n[07] Iniciando exportação...")

# ─────────────────────────────────────────────────────────────
# 7.1 — Planilha consolidada das 3 tabelas analíticas
# ─────────────────────────────────────────────────────────────

wb_tabs <- createWorkbook()

estilo_cabecalho <- createStyle(
  fontName = "Calibri", fontSize = 11, fontColour = "white",
  fgFill = "midnightblue", halign = "center", valign = "center",
  textDecoration = "bold", border = "TopBottom", borderColour = "white"
)

estilo_zebra <- createStyle(fgFill = "aliceblue")

adicionar_sheet <- function(wb, nome, df) {
  addWorksheet(wb, nome, gridLines = FALSE)
  writeData(wb, nome, df, headerStyle = estilo_cabecalho)
  freezePane(wb, nome, firstActiveRow = 2, firstActiveCol = 2)
  setColWidths(wb, nome, cols = 1:ncol(df), widths = "auto")
  if (nrow(df) > 1) {
    addStyle(wb, nome, estilo_zebra,
             rows = seq(3, nrow(df) + 1, by = 2),
             cols = 1:ncol(df), gridExpand = TRUE)
  }
}

adicionar_sheet(wb_tabs, "municipios", tab_municipios)
adicionar_sheet(wb_tabs, "capitais",   tab_capitais)
adicionar_sheet(wb_tabs, "saopaulo",   tab_saopaulo)

path_xlsx_tabs <- file.path(PATHS$outputs, "tabelas_analiticas_2015_2025.xlsx")
saveWorkbook(wb_tabs, path_xlsx_tabs, overwrite = TRUE)
message("    [ok] ", path_xlsx_tabs)

# ─────────────────────────────────────────────────────────────
# 7.2 — Codebook: notas metodológicas + dicionário
# ─────────────────────────────────────────────────────────────

notas_metodologicas <- tibble(
  principio  = sprintf("P%d", 1:8),
  titulo     = c(
    "Filtro de funcionamento ativo",
    "Base universal de chaves",
    "Variáveis-atributo vs. variáveis-soma em 2025",
    "Scaffold IBGE 96 distritos",
    "Schema homogêneo entre tabelas",
    "CSV bruto INEP como única ground truth",
    "Estado é tão crítico quanto código",
    "Validação coluna a coluna contra CSV bruto INEP"
  ),
  descricao  = c(
    "Todos os filtros (f_dir, f_ind_mun, f_ind_mun_est, f_dir_capitais) exigem TP_SITUACAO_FUNCIONAMENTO == '1'. Sem este filtro, escolas inativas inflam as contagens da rede direta e indireta.",
    "construir_tabela() parte de base_chaves = dir uniao ind uniao ind_est. Garante presenca de grupos so-indiretos como Alto de Pinheiros, Barra Funda e Marsilac (SP), que seriam descartados em left_joins encadeados a partir de dir.",
    "Em 2025, QT_TRANSP_PUBLICO migrou para Tabela_Matricula_2025.csv com padrao esparso (1 linha preenchida por escola). agregar_por_escola() distingue VARS_ATRIBUTO_ESCOLA (safe_max) de variaveis somaveis (sum). Previne dupla contagem em fragmentacoes futuras do INEP.",
    "Distritos de SP buscados via API IBGE (servicodados.ibge.gov.br) com stopifnot(nrow == 96). Garante que distritos so-indiretos nao sumam por timeout ou paginacao parcial. Labels acoplados depois da agregacao para nao travar anos com NO_DISTRITO ausente (2015-2023).",
    "agg_locdif() forca todos os niveis de TP_LOCALIZACAO_DIFERENCIADA via factor + names_expand=TRUE. completar_colunas() equaliza schema entre as 3 tabs com 0s nas colunas-extras (ex.: nao ha quilombola em distrito urbano de SP).",
    "Validacao contra produtos downstream (tab vs censo_completo) e circular. Apenas comparacao contra CSV bruto INEP e valida. Implementada em validar_contra_bruto() (2015-2024) e validar_contra_bruto_2025() (Escola + Matricula via left_join).",
    "Pipelines com execucao por blocos podem manter versoes dessincronizadas em memoria. Bloco 5 SEMPRE regenera as tabs a partir dos censos finais. Bloco 6 valida e interrompe se houver drift de estado.",
    "O pipeline foi validado em 5.884.588 celulas (33 combinacoes ano x nivel x 94 colunas analiticas). Para cada ano de 2015 a 2025, as tres tabelas analiticas foram reconstruidas a partir do CSV bruto INEP aplicando exatamente os mesmos tratamentos semanticos do pipeline (coercao numerica, NA->0 em TP_PODER_PUBLICO_PARCERIA pos-2022, replicacao retroativa de TP_LOCALIZACAO_DIFERENCIADA via CO_ENTIDADE) e comparadas coluna a coluna com a tabela analitica em memoria. Zero divergencias encontradas. O resultado e prova empirica completa (nao inferencia) de que cada agregacao documentada no codebook reflete fielmente o microdado INEP."
  )
)

# Resultado da validação atual
if (exists("resultados") && length(resultados) > 0) {
  resumo_validacao <- tibble(
    combinacao  = names(resultados),
    unidades    = purrr::map_int(resultados, nrow),
    ok          = purrr::map_int(resultados, ~ sum(.x$tudo_ok)),
    divergentes = purrr::map_int(resultados, ~ sum(!.x$tudo_ok)),
    status      = if_else(divergentes == 0, "APROVADO", "FALHA")
  )
} else {
  resumo_validacao <- tibble(
    combinacao  = character(),
    unidades    = integer(),
    ok          = integer(),
    divergentes = integer(),
    status      = character()
  )
}

# Universo de cobertura
cobertura <- tibble(
  metrica = c(
    "Anos cobertos",
    "Linhas em censo_completo",
    "Municipios em tab_municipios",
    "Linhas em tab_municipios",
    "Capitais em tab_capitais",
    "Linhas em tab_capitais",
    "Distritos em tab_saopaulo",
    "Linhas em tab_saopaulo",
    "Colunas em tab_municipios",
    "Colunas em tab_capitais",
    "Colunas em tab_saopaulo"
  ),
  valor = c(
    paste(range(as.integer(censo_completo$NU_ANO_CENSO)), collapse = "-"),
    format(nrow(censo_completo), big.mark = ","),
    format(n_distinct(tab_municipios$CO_MUNICIPIO), big.mark = ","),
    format(nrow(tab_municipios), big.mark = ","),
    format(n_distinct(tab_capitais$CO_MUNICIPIO), big.mark = ","),
    format(nrow(tab_capitais), big.mark = ","),
    format(n_distinct(tab_saopaulo$CO_DISTRITO), big.mark = ","),
    format(nrow(tab_saopaulo), big.mark = ","),
    as.character(ncol(tab_municipios)),
    as.character(ncol(tab_capitais)),
    as.character(ncol(tab_saopaulo))
  )
)

# Dicionário automático por tabela
gerar_dicionario <- function(df, nome_tab) {
  tibble(
    tabela   = nome_tab,
    variavel = names(df),
    tipo     = purrr::map_chr(df, ~ class(.x)[1]),
    n_obs    = purrr::map_int(df, ~ sum(!is.na(.x))),
    n_na     = purrr::map_int(df, ~ sum(is.na(.x))),
    pct_preenchido = round(purrr::map_dbl(df, ~ mean(!is.na(.x))) * 100, 1)
  )
}

dic_municipios <- gerar_dicionario(tab_municipios, "tab_municipios")
dic_capitais   <- gerar_dicionario(tab_capitais,   "tab_capitais")
dic_saopaulo   <- gerar_dicionario(tab_saopaulo,   "tab_saopaulo")

# Glossário curto das principais variáveis
glossario_principal <- tribble(
  ~variavel,                ~descricao,
  "NU_ANO_CENSO",           "Ano do Censo Escolar INEP (2015-2025).",
  "SG_UF",                  "Sigla da Unidade da Federacao.",
  "CO_MUNICIPIO",           "Codigo IBGE de 7 digitos do municipio.",
  "NO_MUNICIPIO",           "Nome do municipio.",
  "CO_DISTRITO",            "Codigo IBGE de 9 digitos do distrito (so tab_saopaulo).",
  "NO_DISTRITO",            "Nome do distrito (acoplado via scaffold IBGE).",
  "n_esc_dir",              "Numero de escolas da rede municipal direta (TP_DEPENDENCIA=3).",
  "n_esc_ind_mun",          "Numero de escolas privadas conveniadas com a prefeitura (TP_PODER_PUBLICO_PARCERIA=1).",
  "n_esc_ind_mun_est",      "Numero de escolas privadas em convenio tripartite municipal+estadual (TP_PPP=3).",
  "mat_bas_dir",            "Total de matriculas na educacao basica na rede direta.",
  "mat_bas_ind",            "Total de matriculas na educacao basica na rede indireta municipal.",
  "mat_bas_tot",            "Soma de mat_bas_dir + mat_bas_ind.",
  "mat_ei_dir",             "Matriculas em educacao infantil (creche + pre) na rede direta.",
  "mat_ei_ind",             "Matriculas em educacao infantil na rede indireta.",
  "mat_cre_dir",            "Matriculas em creche (0-3 anos) na rede direta.",
  "mat_pre_dir",            "Matriculas em pre-escola (4-5 anos) na rede direta.",
  "mat_cre_dir_loc_urb",    "Matriculas em creche urbana na rede direta.",
  "mat_cre_dir_loc_rur",    "Matriculas em creche rural na rede direta.",
  "doc_tur_cre_dir",        "Razao docente/turma em creche na rede direta (sd_cre/st_cre).",
  "doc_tur_pre_dir",        "Razao docente/turma em pre-escola na rede direta.",
  "mat_ei_esc_dir",         "Razao matriculas EI por escola com IN_INF=1 (NA em 2025 - mudanca instrumento INEP).",
  "mat_ei_av_dir",          "Matriculas EI ponderadas por escolas com area verde (IN_AREA_VERDE=1).",
  "mat_ei_ban_dir",         "Matriculas EI em escolas com banheiro EI (IN_BANHEIRO_EI=1).",
  "mat_ei_jog_dir",         "Matriculas EI em escolas com material pedagogico de jogos.",
  "mat_ei_art_dir",         "Matriculas EI em escolas com material pedagogico artistico.",
  "mat_ei_4_dir",           "Matriculas EI em escolas com todos os 4 itens de infraestrutura.",
  "mat_ei_s4_dir",          "Matriculas EI em escolas sem nenhum dos 4 itens de infraestrutura.",
  "mat_fund_ai_dir",        "Matriculas em ensino fundamental anos iniciais (rede direta).",
  "mat_fund_af_dir",        "Matriculas em ensino fundamental anos finais (rede direta).",
  "mat_med_dir",            "Matriculas em ensino medio (rede direta).",
  "mat_eja_dir",            "Matriculas em EJA (rede direta).",
  "mat_bas_branca_dir",     "Matriculas EB por cor/raca branca na rede direta (e variantes preta, parda, amarela, indigena, nd).",
  "mat_bas_dir_ld_*",       "Matriculas EB na rede direta por localizacao diferenciada (fora, asset, indig, quilom, trad).",
  "mat_bas_ind_ld_*",       "Idem para rede indireta.",
  "n_cre_dir / n_cre_ind",  "Numero de escolas-tipo creche (so tab_saopaulo). NA em 2025 - IN_INF_CRE ausente."
)

# ─────────────────────────────────────────────────────────────
# 7.3 — Montagem do workbook do codebook
# ─────────────────────────────────────────────────────────────

wb_cb <- createWorkbook()

# Sheet 1: notas metodológicas
addWorksheet(wb_cb, "notas_metodologicas", gridLines = FALSE)
writeData(wb_cb, "notas_metodologicas", notas_metodologicas,
          headerStyle = estilo_cabecalho)
setColWidths(wb_cb, "notas_metodologicas",
             cols = 1:3, widths = c(10, 40, 110))
addStyle(wb_cb, "notas_metodologicas",
         createStyle(wrapText = TRUE, valign = "top"),
         rows = 2:(nrow(notas_metodologicas) + 1),
         cols = 1:3, gridExpand = TRUE)
freezePane(wb_cb, "notas_metodologicas", firstActiveRow = 2)

# Sheet 2: resumo da validação
addWorksheet(wb_cb, "validacao", gridLines = FALSE)
writeData(wb_cb, "validacao", resumo_validacao,
          headerStyle = estilo_cabecalho)
setColWidths(wb_cb, "validacao", cols = 1:5, widths = "auto")
freezePane(wb_cb, "validacao", firstActiveRow = 2)

# Sheet 3: cobertura
addWorksheet(wb_cb, "cobertura", gridLines = FALSE)
writeData(wb_cb, "cobertura", cobertura,
          headerStyle = estilo_cabecalho)
setColWidths(wb_cb, "cobertura", cols = 1:2, widths = c(35, 20))

# Sheet 4: glossário curado
addWorksheet(wb_cb, "glossario", gridLines = FALSE)
writeData(wb_cb, "glossario", glossario_principal,
          headerStyle = estilo_cabecalho)
setColWidths(wb_cb, "glossario", cols = 1:2, widths = c(30, 100))
addStyle(wb_cb, "glossario",
         createStyle(wrapText = TRUE, valign = "top"),
         rows = 2:(nrow(glossario_principal) + 1),
         cols = 1:2, gridExpand = TRUE)
freezePane(wb_cb, "glossario", firstActiveRow = 2)

# Sheets 5-7: dicionários automáticos por tabela
for (nome_dic in c("dic_municipios", "dic_capitais", "dic_saopaulo")) {
  df_dic <- get(nome_dic)
  sheet_nm <- sub("^dic_", "dicionario_", nome_dic)
  addWorksheet(wb_cb, sheet_nm, gridLines = FALSE)
  writeData(wb_cb, sheet_nm, df_dic, headerStyle = estilo_cabecalho)
  setColWidths(wb_cb, sheet_nm, cols = 1:6, widths = "auto")
  freezePane(wb_cb, sheet_nm, firstActiveRow = 2)
}

path_xlsx_cb <- file.path(PATHS$outputs, "codebook_censo_lemann_cem.xlsx")
saveWorkbook(wb_cb, path_xlsx_cb, overwrite = TRUE)
message("    [ok] ", path_xlsx_cb)

# ─────────────────────────────────────────────────────────────
# 7.4 — Sanity log final
# ─────────────────────────────────────────────────────────────

fs_out <- file.size(c(path_xlsx_tabs, path_xlsx_cb))
names(fs_out) <- c("tabelas_analiticas", "codebook")

message("\n[07] Exportacao concluida:")
message("    tabelas_analiticas: ", round(fs_out[1] / 1024^2, 2), " MB")
message("    codebook:           ", round(fs_out[2] / 1024^2, 2), " MB")
message("    diretorio:          ", normalizePath(PATHS$outputs))







# Passo 1 — apenas uma vez ao clonar o projeto
source("setup_inicial.R")

# Passo 2 — rodar o pipeline inteiro
source("run.R")