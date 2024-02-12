# function to make dummy variable for every level of a factor variable
# make_dummies <- function(v) {
#   s <- sort(unique(v))
#   d <- outer(v, s, function(v, s) 1L * (v == s))
#   # colnames(d) <- s
#   return(d)
# }

wtd_t_test <- function(x,
                       mu2 = 0,
                       weights_x = NULL,
                       prop = FALSE,
                       se_calculation = c("taylor_se", "survey_se"),
                       survey_design){
  n <- sum(!is.na(x))
  
  if(is.null(weights_x)) weights_x <- rep(1, length(x))
  
  mu <- weighted.mean(x = x, w = weights_x, na.rm = TRUE)
  
  # SE^2
  if(se_calculation == "taylor_se"){
    use_x <- !is.na(x)
    x <- x[use_x]
    weights_x <- weights_x[use_x]
    
    se_2 <- (n/((n - 1) * sum(weights_x)^2)) * sum(weights_x^2 * (x - mu)^2)
  }
  
  if(se_calculation == "survey_se"){
    se_2 <- SE(svymean(x = x, design = survey_design, na.rm = TRUE))^2
  }

  t <- (mu2 - mu)/(sqrt(se_2))
  
  if(prop == TRUE){
    p <- pnorm(q = abs(t), lower.tail = FALSE) * 2
  } else {
    df <- n - 1
    p <- pt(q = abs(t), df = df, lower.tail = FALSE) * 2
  }
  
  c(povp = mu, t = t, p = p)
}

# weighted frequency tables for weighting variables
display_tables_weighting_vars <- function(orig_data,
                                          sheet_list_table,
                                          prevec = NULL){
  one_dimensional_raking_variable <- NULL
  two_dimensional_raking_variables <- NULL
  prevec_indicator <- FALSE
  
  df <- sheet_list_table
  
  var_names1 <- c("Frekvenca", "Vzorcni %", "Populacijski (ciljni) %")
  var_names2 <- gsub(pattern = " ", replacement = ".", x = var_names1, fixed = TRUE)

  df <- df[,unique(c(1:2,which(colnames(df) %in% c(var_names1, var_names2)))), drop = FALSE]
  df <- na.omit(df)
  
  n_col <- ncol(df)
  df <- df[-nrow(df),]
  
  prx <- "Neutežen (vzorčni)"
  if (sum(prevec == 1) != length(prevec)) {
    prx <- "Stare uteži (vzorčni)"
    prevec_indicator <- TRUE
  }
    
  column_names <- c(paste(prx, "N"),
                    paste(prx, "%"),
                    "Populacijski (ciljni) %")
  
  if(n_col == 5){
    two_dimensional_raking_variables <- c(names(df)[1], names(df)[2])
    two_dim_names <- paste0(two_dimensional_raking_variables, collapse = " x ")
    df <- cbind(paste0(df[[1]],"_",df[[2]]), df[,3:n_col])
    names(df) <- c(two_dim_names, column_names)
    
    if(prevec_indicator){
      levels <- df[[1]]
      orig_data[[two_dimensional_raking_variables[1]]] <- droplevels(clean_data(orig_data[[two_dimensional_raking_variables[1]]]))
      orig_data[[two_dimensional_raking_variables[2]]] <- droplevels(clean_data(orig_data[[two_dimensional_raking_variables[2]]]))
      
      orig_data[[two_dim_names]] <- as.factor(paste0(orig_data[[two_dimensional_raking_variables[1]]],"_",
                                                     orig_data[[two_dimensional_raking_variables[2]]]))
      
      df[[2]] <- weighted_table(orig_data[[two_dim_names]], weights = prevec)[levels]
      df[[3]] <- (df[[2]]/sum(df[[2]]))*100
    }
    
  } else {
    one_dimensional_raking_variable <- names(df)[1]
    names(df) <- c(one_dimensional_raking_variable, column_names)
    
    if(prevec_indicator){
      levels <- df[[1]]
      orig_data[[one_dimensional_raking_variable]] <- droplevels(clean_data(orig_data[[one_dimensional_raking_variable]]))
      df[[2]] <- weighted_table(orig_data[[one_dimensional_raking_variable]], weights = prevec)[levels]
      df[[3]] <- (df[[2]]/sum(df[[2]]))*100
    }
  }
  
  temp_df <- data.frame(v1 = "Skupaj", t(colSums(df[-1])))
  names(temp_df) <- names(df)
  df <- rbind(df, temp_df)

  return(df)
}

# weighted statistics for numeric variables
weighted_numeric_statistics <- function(numeric_variables, orig_data, weights, p_adjust_method = NULL, ...){
  selected_data <- labelled::user_na_to_na(orig_data[ ,numeric_variables, drop = FALSE])
  
  # select only numeric variables first
  true_numeric <- vapply(selected_data, FUN = function(x) is.numeric(x) | is.integer(x),
                         FUN.VALUE = logical(1))
  
  non_numeric_vars <- numeric_variables[!true_numeric]
  numeric_variables <- numeric_variables[true_numeric]
  selected_data <- selected_data[ ,numeric_variables, drop = FALSE]
  
  # exclude variables with variance 0, for which test statistic could not be calculated
  zero_variance <- vapply(selected_data, FUN = function(x){
    variance <- var(x, na.rm = TRUE)
    variance <- ifelse(is.na(variance), 0, variance)
    0 == variance
    }, FUN.VALUE = logical(1))
  
  non_calculated_vars <- numeric_variables[zero_variance]
  numeric_variables <- numeric_variables[!zero_variance]
  selected_data <- selected_data[ ,numeric_variables, drop = FALSE]
  
  variable_labels <- vapply(numeric_variables, FUN = function(x){
    label <- attr(x = orig_data[[x]], which = "label", exact = TRUE)
    if(is.null(label)) "" else label 
  }, FUN.VALUE = character(1))
  
  temp_df <- data.frame("Spremenljivka" = numeric_variables)
  
  if(!all(variable_labels == "")){
    temp_df[["Labela"]] <- variable_labels
  }
  
  if(length(numeric_variables > 0)) {
    temp_df[["N"]] <- colSums(!is.na(selected_data))
    temp_df[["Min"]] <- vapply(selected_data, min, na.rm = TRUE, FUN.VALUE = numeric(1)) 
    temp_df[["Maks"]] <- vapply(selected_data, max, na.rm = TRUE, FUN.VALUE = numeric(1))
    temp_df[["Neuteženo povprečje"]] <- colMeans(selected_data, na.rm = TRUE)
    
    statistic <- lapply(seq_along(numeric_variables), function(i){
      wtd_t_test(x = selected_data[[i]],
                 mu2 = temp_df[["Neuteženo povprečje"]][[i]],
                 weights_x = weights, ...)
    })
    
    temp_df[["Uteženo povprečje"]] <- sapply(statistic, FUN = function(x) x[["povp"]])
    temp_df[["Absolutna razlika"]] <- temp_df[["Neuteženo povprečje"]] - temp_df[["Uteženo povprečje"]]
    temp_df[["Relativna razlika (%)"]] <- (temp_df[["Absolutna razlika"]]/temp_df[["Uteženo povprečje"]])*100
    temp_df[["t"]] <- sapply(statistic, FUN = function(x) x[["t"]])

    if(is.null(p_adjust_method)){
      temp_df[["p"]] <- sapply(statistic, FUN = function(x) x[["p"]])
      temp_df[["Signifikanca"]] <- weights::starmaker(temp_df[["p"]])
    } else {
      temp_df[[paste0("pop. p (", p_adjust_method, ")")]] <- p.adjust(p = sapply(statistic, FUN = function(x) x[["p"]]), method = p_adjust_method)
      temp_df[["Signifikanca"]] <- weights::starmaker(temp_df[[paste0("pop. p (", p_adjust_method, ")")]])
    }
  }
  
  return(list(calculated_table = temp_df,
              non_numeric_vars = non_numeric_vars,
              non_calculated_vars = non_calculated_vars))
}

# weighted statistics for categorical variables
create_w_table <- function(orig_data, variable, weights, p_adjust_method = NULL, ...){
  temp_var <- droplevels(haven::as_factor(labelled::user_na_to_na(orig_data[[variable]])))
  
  invalid_var <- NULL
  constant_var <- NULL
  temp_df <- NULL
  warning_numerus <- NULL
  
  if(all(is.na(temp_var))){
    invalid_var <- variable
    
  } else if(length(unique(na.omit(temp_var))) == 1){
    constant_var <- variable
    
  } else {
    temp_df <- as.data.frame.table(table(temp_var, useNA = "no"), stringsAsFactors = FALSE)
    names(temp_df)[1:2] <- c(variable, "N pred uteževanjem")
    temp_df[["N po uteževanju"]] <- weighted_table(temp_var, weights = weights)
    temp_df[["Delež (%) pred uteževanjem"]] <- (temp_df[[2]]/sum(temp_df[[2]]))*100
    
    # make dummy (0 1) variable for each level of a factor variable
    dummies <- weights::dummify(x = temp_var, show.na = FALSE, keep.na = TRUE)
    
    # then perform weighted z-test on every dummy variable (mean = proportion)
    statistic <- lapply(seq_len(ncol(dummies)), function(i){
      wtd_t_test(x = dummies[ ,i],
                 mu2 = temp_df[["Delež (%) pred uteževanjem"]][[i]]/100,
                 weights_x = weights,
                 prop = TRUE, ...)
    })
    
    temp_df[["Delež (%) po uteževanju"]] <- (sapply(statistic, FUN = function(x) x[["povp"]]))*100
    temp_df[["Razlika v deležih - absolutna"]] <- temp_df[[4]] - temp_df[[5]]
    temp_df[["Razlika v deležih - relativna (%)"]] <- (temp_df[[6]]/temp_df[[5]])*100
    temp_df[["z"]] <- sapply(statistic, FUN = function(x) x[["t"]])
    
    if(is.null(p_adjust_method)){
      temp_df[["p"]] <- sapply(statistic, FUN = function(x) x[["p"]])
      temp_df[["Signifikanca"]] <- weights::starmaker(temp_df[["p"]])
    } else {
      temp_df[[paste0("pop. p (", p_adjust_method, ")")]] <- p.adjust(p = sapply(statistic, FUN = function(x) x[["p"]]), method = p_adjust_method)
      temp_df[["Signifikanca"]] <- weights::starmaker(temp_df[[paste0("pop. p (", p_adjust_method, ")")]])
    }
    
    # add warning when n is too small for reliable use of asymptotic test for proportions
    warning_numerus <- temp_df[[3]] <= 5 | (sum(temp_df[[3]]) - temp_df[[3]]) <= 5
    temp_df[["Signifikanca"]][warning_numerus] <- paste0(temp_df[["Signifikanca"]][warning_numerus], " !")
    
    # add Total row
    df <- data.frame(v1 = "Skupaj", t(c(colSums(temp_df[-c(1,6:10)]), NA, NA, NA, NA, NA)))
    names(df) <- names(temp_df)
    temp_df <- rbind(temp_df, df)
  }
  
  return(list(calculated_table = temp_df,
              invalid_var = invalid_var,
              constant_var = constant_var,
              warning_indicator = any(warning_numerus, na.rm = TRUE)))
}

# function that counts the number of significant variables beased on relative change values
count_rel_diff <- function(vec, p_vec) {
  vec <- abs(vec)
  frek <- c(sum(vec > 20), sum(vec > 10 & vec <= 20), sum(vec >= 5 & vec <= 10), sum(vec < 5))
  p_frek <- c(sum(vec > 20 & p_vec < 0.05), sum(vec > 10 & vec <= 20 & p_vec < 0.05), sum(vec >= 5 & vec <= 10 & p_vec < 0.05), sum(vec < 5 & p_vec < 0.05))
  
  list(sums = frek,
       cumsums = cumsum(frek),
       p_sums = p_frek,
       p_cumsums = cumsum(p_frek))
}

# functions for tables download
download_analyses_numeric_table <- function(numeric_table, file){
  wb <- createWorkbook()
  
  if(nrow(numeric_table) > 0){
    ## summary worksheet
    addWorksheet(wb = wb,
                 sheetName = "Povzetek",
                 gridLines = FALSE)
    
    # extract relative changes and p values
    freq_rel_change <- count_rel_diff(vec = numeric_table[[9]], p_vec = numeric_table[[11]])
    
    tbl <- data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                      "f" = freq_rel_change$sums,
                      "%" = freq_rel_change$sums/sum(freq_rel_change$sums),
                      "Kumul f" = freq_rel_change$cumsums,
                      "Kumul %" = freq_rel_change$cumsums/sum(freq_rel_change$sums),
                      "f*" = freq_rel_change$p_sums,
                      "% od vseh spremenljivk" = freq_rel_change$p_sums/sum(freq_rel_change$sums),
                      "% od relativne razlike" = freq_rel_change$p_sums/freq_rel_change$sums,
                      "Kumul f*" = freq_rel_change$p_cumsums,
                      "Kumul %*" = freq_rel_change$p_cumsums/sum(freq_rel_change$sums),
                      check.names = FALSE)
    
    tbl[is.na(tbl)] <- 0
    
    writeData(wb = wb,
              sheet = "Povzetek",
              x = tbl,
              borders = "all", startRow = 4, 
              headerStyle = createStyle(textDecoration = "bold",
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center", wrapText = TRUE))
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Povzetek sprememb povprečji po uteževanju za številske spremenljivke", startCol = 1, startRow = 1)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 1:10, rows = 1)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Spremenljivke", startCol = 2, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne spremenljivke (p < 0.05)",
              startCol = 6, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh številskih spremenljivk:", sum(freq_rel_change$sums)),
                          paste("Št. statistično značilnih številskih spremenljivk (p < 0.05):", sum(freq_rel_change$p_sums)))),
              startCol = 1, startRow = 10, rowNames = FALSE, colNames = FALSE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1:2, cols = 1:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 border = c("top", "bottom", "left", "right")),
             rows = 3, cols = 1:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(numFmt = "0%"),
             rows = 5:8, cols = c(3, 5, 7, 8, 10),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(halign = "center"),
             rows = 5:8, cols = 1:10,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:10,
                 widths = c(14, 5, 6, 8, 8, 5, 21, 21, 8, 8.5))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 1:4, heights = c(20, 20, 25, 44))
    
    ## table worksheet
    addWorksheet(wb = wb,
                 sheetName = "Opisne statistike",
                 gridLines = FALSE)
    
    writeData(wb = wb,
              sheet = "Opisne statistike",
              x = numeric_table,
              borders = "all",
              startRow = 1,
              headerStyle = createStyle(textDecoration = "bold",
                                        wrapText = TRUE,
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center"))
    
    n_col <- ncol(numeric_table)
    n_row <- nrow(numeric_table)
    
    addStyle(wb = wb,
             sheet = "Opisne statistike",
             style = createStyle(halign = "center"),
             rows = 2:(n_row+2), cols = (n_col-9):n_col,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb,
             sheet = "Opisne statistike",
             style = createStyle(numFmt = "0.00"),
             rows = 2:(n_row+2), cols = (n_col-6):(n_col-1),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(numFmt = "0"),
             rows = 2:(n_row+2), cols = n_col-3,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb, sheet = "Opisne statistike", cols = 1:n_col, widths = "auto")
    setColWidths(wb = wb, sheet = "Opisne statistike", cols = 2, widths = 20)
    setRowHeights(wb = wb, sheet = "Opisne statistike", rows = 1, heights = 35)
    
    writeData(wb = wb, sheet = "Opisne statistike",
              x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001",
              xy = c(n_col + 1, 1))
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(fontSize = 9, valign = "center"),
             rows = 1, cols = n_col + 1)
  }
  
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}

download_analyses_factor_tables <- function(factor_tables, orig_data, file, warning_indicator){
  wb <- createWorkbook()
  
  if(length(factor_tables) > 0){
    ## summary worksheet
    addWorksheet(wb = wb,
                 sheetName = "Povzetek",
                 gridLines = FALSE)
    
    categories <- abs(na.omit(do.call(rbind, lapply(factor_tables, "[", c(7, 9)))))
    warning_numerus <- grepl(pattern = "!", x = na.omit(do.call(rbind, lapply(factor_tables, "[", 10)))[[1]])
    
    # temporarily replace p-values with small n with 1 so they are not counted in summary
    if(any(warning_numerus)) categories[warning_numerus,][[2]] <- 1
    
    freq_rel_change <- count_rel_diff(vec = categories[[1]], p_vec = categories[[2]])
    
    tbl <- data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                      "f" = freq_rel_change$sums,
                      "%" = freq_rel_change$sums/sum(freq_rel_change$sums),
                      "Kumul f" = freq_rel_change$cumsums,
                      "Kumul %" = freq_rel_change$cumsums/sum(freq_rel_change$sums),
                      "f*" = freq_rel_change$p_sums,
                      "% od vseh kategorij" = freq_rel_change$p_sums/sum(freq_rel_change$sums),
                      "% od relativne razlike" = freq_rel_change$p_sums/freq_rel_change$sums,
                      "Kumul f*" = freq_rel_change$p_cumsums,
                      "Kumul %*" = freq_rel_change$p_cumsums/sum(freq_rel_change$sums),
                      check.names = FALSE)
    
    tbl[is.na(tbl)] <- 0
    
    writeData(wb = wb,
              sheet = "Povzetek",
              x = tbl,
              borders = "all", startRow = 4, 
              headerStyle = createStyle(textDecoration = "bold",
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center", wrapText = TRUE))
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Povzetek sprememb deležev kategorij po uteževanju za opisne spremenljivke", startCol = 1, startRow = 1)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 1:10, rows = 1)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Kategorije", startCol = 2, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne kategorije (p < 0.05)",
              startCol = 6, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh kategorij:", sum(freq_rel_change$sums)),
                          paste("Št. statistično značilnih kategorij (p < 0.05):", sum(freq_rel_change$p_sums)))),
              startCol = 1, startRow = 10, rowNames = FALSE, colNames = FALSE)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh opisnih spremenljivk:", length(factor_tables)),
                          paste("Št. statistično značilnih opisnih spremenljivk (p < 0.05):",
                                sum(vapply(seq_along(factor_tables), function(i){
                                  warning <- grepl(pattern = "!", x = factor_tables[[i]][[10]])
                                  any(factor_tables[[i]][[9]][!warning] < 0.05, na.rm = TRUE)
                                }, FUN.VALUE = logical(1)))))),
              startCol = 1, startRow = 13, rowNames = FALSE, colNames = FALSE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1:2, cols = 1:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 border = c("top", "bottom", "left", "right")),
             rows = 3, cols = 1:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(numFmt = "0%"),
             rows = 5:8, cols = c(3, 5, 7, 8, 10),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(halign = "center"),
             rows = 5:8, cols = 1:10,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:10,
                 widths = c(14, 5, 6, 8, 8, 5, 21, 21, 8, 8.5))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 1:4, heights = c(20, 20, 25, 44))
    
    ## frequency tables worksheet
    addWorksheet(wb = wb,
                 sheetName = "Frekvencne tabele",
                 gridLines = FALSE)
    
    # starting row = number of rows of previous table
    # + 2 (for the header and to add a empty row)
    # + 1 for the first table
    start_rows <- c(0, cumsum(2 + sapply(factor_tables, nrow)[-length(factor_tables)])) + 3
    
    for(i in seq_along(factor_tables)){
      spr <- colnames(factor_tables[[i]])[1]
      labela <- attr(x = orig_data[[names(factor_tables)[i]]], which = "label", exact = TRUE)
      colnames(factor_tables[[i]])[1] <- paste0(spr, ifelse(is.null(labela), "", paste0(" - ", labela)))
      
      writeData(wb = wb,
                sheet = "Frekvencne tabele",
                x = factor_tables[[i]],
                borders = "all",
                startRow = start_rows[i],
                headerStyle = createStyle(textDecoration = "bold",
                                          wrapText = TRUE,
                                          border = c("top", "bottom", "left", "right"),
                                          halign = "center", valign = "center"))
    }
    
    writeData(wb = wb,
              sheet = "Frekvencne tabele",
              x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001", xy = c(1, 1))
    
    if(warning_indicator){
      writeData(wb = wb,
                sheet = "Frekvencne tabele",
                x = "! Število enot v celici je premajhno za zanesljivo oceno p vrednosti (np ≤ 5 ali n(1-p) ≤ 5).", xy = c(1,2))
      
    }
    
    addStyle(wb = wb, sheet = "Frekvencne tabele",
             style = createStyle(fontSize = 9),
             rows = 1:2, cols = 1)
    
    addStyle(wb = wb,
             sheet = "Frekvencne tabele",
             style = createStyle(numFmt = "0.00"),
             rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])),
             cols = 8:9,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb,
             sheet = "Frekvencne tabele",
             style = createStyle(numFmt = "0"),
             rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])),
             cols = 2:7,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb,
             sheet = "Frekvencne tabele",
             style = createStyle(halign = "center"),
             rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])),
             cols = 2:10,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb, sheet = "Frekvencne tabele", cols = 1:10, widths = c(60, rep(15, 6), 10, 10, 12))
  }
  
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}
