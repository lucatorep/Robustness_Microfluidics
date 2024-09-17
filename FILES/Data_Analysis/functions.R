#' loadinstall
#'
#' @param pkg vector containing all the packages needed
#'
#' @return upload all the libraries needed and install the packages if not already there
loadinstall <- function(pkg) {
  new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
  if (length(new.pkg)) {
    install.packages(new.pkg, dependencies = TRUE)
  }
  sapply(pkg, require, character.only = TRUE)
}


#' add_info
#'
#' @param df a dataframe
#'
#' @return same dataframe, but with column with experimental informations
add_info <- function(df) {
  df %>% dplyr::mutate(
    date = date,
    strain = strain,
    sensor = sensor,
    oscillation_type = oscillation
  )
}


m4Rn <- function(df, colm, groupm) {
  df %>%
    dplyr::ungroup() %>%
    dplyr::group_by(dplyr::pick({{ groupm }})) %>%
    dplyr::mutate(
      across({{ colm }},
        list(m = ~ mean(., na.rm = TRUE)),
        .names = "{.fn}_{.col}"
      ),
      .keep = "none"
    ) %>%
    dplyr::distinct()
}


#' robust
#' @param dfR Input data frame or matrix with data for which robustness needs to be computed.
#' @param colR Column(s) to compute the robustness for. Suggested (not mandatory) to have one column with all the phenotypes and another one with the numerical values. Put the phenotype column into groupR and the value column in colR.
#' @param groupR Column(s) that should be used by the "dplyr::group_by()" function before computing R for colR.
#' @param normR Logical. If TRUE (default), R is normalized by "m", the normalisation factor in the robustness formula. This value will be computed automatically in the function.
#' @param nadropR Logical. If TRUE (default), removes the rows whose R is an NA.
#' @param dfm Data frame to be used to compute the normalisation factor "m", if different from dfR. Default is NULL.
#' @param groupm Column(s) that should be used by the "dplyr::group_by()" function before computing m for colR. NOTE: not the same as groupR!
#'
#' @return A dataframe with the Robustness (R) or normalised Robustnes (Rn), mean and standard deviation of each phenotype gave as input (colR)
robust <- function(dfR, colR, groupR, groupm,
                   dfm = NULL, nadropR = TRUE, normR = TRUE) {
  vector <- ncol(dfR %>% dplyr::ungroup() %>% dplyr::select({{ colR }}))

  tmp <- dfR %>%
    dplyr::ungroup() %>%
    dplyr::group_by(dplyr::pick({{ groupR }})) %>%
    dplyr::mutate(
      dplyr::across(
        {{ colR }},
        base::list(
          mean = ~ mean(., na.rm = TRUE),
          sd = ~ sd(., na.rm = TRUE)
        ),
        .names = "{.fn}_{.col}"
      ),
      dplyr::across(
        dplyr::starts_with("sd_"),
        ~ -.^2 / base::get(base::sub("sd_", "mean_", dplyr::cur_column())),
        .names = "R_{.col}"
      ),
      .keep = "none"
    ) %>%
    dplyr::rename_with(~ base::sub("R_sd_", "R_", .x, fixed = TRUE)) %>%
    dplyr::distinct()

  if (nadropR) {
    tmp <- tmp %>%
      tidyr::drop_na(dplyr::starts_with("R_"))
  }

  if (normR) {
    base::ifelse(
      base::length(dfm) != 0,
      tmp2 <- dfm,
      tmp2 <- dfR
    )
    tmp2 <- tmp2 %>%
      dplyr::ungroup() %>%
      dplyr::group_by(dplyr::pick({{ groupm }})) %>%
      dplyr::mutate(
        dplyr::across({{ colR }},
          base::list(m = ~ mean(., na.rm = TRUE)),
          .names = "{.fn}_{.col}"
        ),
        .keep = "none"
      ) %>%
      dplyr::distinct() %>%
      dplyr::ungroup()

    tmp <- tmp %>%
      dplyr::ungroup() %>%
      base::merge(., tmp2, all.x = T) %>%
      dplyr::mutate(dplyr::across(dplyr::starts_with("R_"),
        ~ . / base::get(base::sub("R_", "m_", dplyr::cur_column())),
        .names = "Rn_{.col}"
      )) %>%
      dplyr::rename_with(~ base::sub("Rn_R_", "Rn_", .x, fixed = TRUE)) %>%
      dplyr::select(-starts_with("R_")) %>%
      dplyr::distinct()
  }

  if (vector == 1) {
    tmp <- tmp %>%
      dplyr::rename(
        R = dplyr::starts_with("R_"),
        Rn = dplyr::starts_with("Rn_")
      )
  }
}
