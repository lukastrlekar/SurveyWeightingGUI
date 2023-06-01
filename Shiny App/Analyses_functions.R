# weighted frequency tables for weighting variables
# TREBA ŠE IZBOLJŠATI PRIKAZ IN Z DATATABLE
display_tables_weighting_vars <- function(orig_data, sheet_list_table, weights){
  one_dimensional_raking_variable <- NULL
  two_dimensional_raking_variables <- NULL
  
  df <- sheet_list_table
  # df <- df[,-((ncol(df)-2):ncol(df)), drop = FALSE]
  # n_col <- ncol(df)
  #df <- na.omit(df) # tole popraviti
  # df <- df[-nrow(df),]
  
  var_names1 <- c("Frekvenca", "Vzorcni %", "Populacijski (ciljni) %")
  var_names2 <- gsub(pattern = " ", replacement = ".", x = var_names1, fixed = TRUE)

  df <- df[,unique(c(1:2,which(colnames(df) %in% c(var_names1, var_names2)))), drop = FALSE]
  df <- na.omit(df)
  
  n_col <- ncol(df)
  df <- df[-nrow(df),]
  
  if(n_col == 5 && nrow(df) > 0){
    two_dimensional_raking_variables <- c(names(df)[1], names(df)[2])
    two_dim_names <- paste0(two_dimensional_raking_variables, collapse = " x ")
    df <- cbind(paste0(df[[1]],"_",df[[2]]), df[,3:n_col])
    names(df) <- c(two_dim_names, "Vzorčni N", "Vzorčni %", "Populacijski (ciljni) %")
    
  } else{
    one_dimensional_raking_variable <- names(df)[1]
    names(df) <- c(one_dimensional_raking_variable, "Vzorčni N", "Vzorčni %", "Populacijski (ciljni) %")
  }
  
  
  if(!is.null(weights) && is.numeric(weights)){
    levels <- df[[1]]
    
    if(!is.null(one_dimensional_raking_variable)){
      orig_data[[one_dimensional_raking_variable]] <- droplevels(clean_data(orig_data[[one_dimensional_raking_variable]]))
      df[["Utežen N"]] <- weighted_table(orig_data[[one_dimensional_raking_variable]], weights = weights)[levels]
      
    } else{
      orig_data[[two_dimensional_raking_variables[1]]] <- droplevels(clean_data(orig_data[[two_dimensional_raking_variables[1]]]))
      orig_data[[two_dimensional_raking_variables[2]]] <- droplevels(clean_data(orig_data[[two_dimensional_raking_variables[2]]]))
      
      orig_data[[two_dim_names]] <- as.factor(paste0(orig_data[[two_dimensional_raking_variables[1]]],"_",
                                                     orig_data[[two_dimensional_raking_variables[2]]]))
      df[["Utežen N"]] <- weighted_table(orig_data[[two_dim_names]], weights = weights)[levels]
    }
    
    df[["Utežen %"]] <- (df[["Utežen N"]]/sum(df[["Utežen N"]]))*100
  }
  
  temp_df <- data.frame(v1 = "Skupaj", t(colSums(df[-1])))
  names(temp_df) <- names(df)
  df <- rbind(df, temp_df)
  
  return(df)
}


# weighted statistics for numeric variables
weighted_numeric_statistics <- function(numeric_variables, orig_data, weights){
  # select only numeric variables first
  true_numeric <- sapply(orig_data[,numeric_variables, drop = FALSE], FUN = function(x) is.numeric(x)|is.integer(x))
  non_numeric_vars <- numeric_variables[numeric_variables %in% names(true_numeric[true_numeric == FALSE])]
  
  numeric_variables <- numeric_variables[numeric_variables %in% names(true_numeric[true_numeric == TRUE])]
  
  selected_data <- labelled::user_na_to_na(orig_data[ ,numeric_variables, drop = FALSE])
  
  variable_labels <- sapply(numeric_variables, FUN = function(x){
    label <- attr(x = orig_data[[x]], which = "label")
    
    if(is.null(label)) "" else label 
  })
  
  temp_df <- data.frame("Spremenljivka" = numeric_variables)
  
  if(any(!is.na(variable_labels))){
    temp_df[["Ime spremenljivke"]] <- variable_labels
  }
  
  temp_df[["N pred uteževanjem"]] <- colSums(!is.na(selected_data))
  temp_df[["N po uteževanju"]] <- apply(selected_data, 2, FUN = function(x) sum(weights[!is.na(x)]))
  temp_df[["Min"]] <- apply(selected_data, 2, min, na.rm = TRUE)
  temp_df[["Max"]] <- apply(selected_data, 2, max, na.rm = TRUE)
  temp_df[["Neuteženo povprečje"]] <- colMeans(selected_data, na.rm = TRUE)
  temp_df[["Uteženo povprečje"]] <- apply(selected_data, 2, weighted.mean, w = weights, na.rm = TRUE)
  temp_df[["Absolutna razlika"]] <- temp_df$`Uteženo povprečje` - temp_df$`Neuteženo povprečje`
  temp_df[["Relativna razlika (v %)"]] <- (temp_df$`Absolutna razlika`/temp_df$`Neuteženo povprečje`)*100
  
  statistic <- lapply(numeric_variables, function(x){
    test <- weights::wtd.t.test(x = selected_data[[x]],
                                y = mean(selected_data[[x]], na.rm = TRUE),
                                weight = weights)
    
    c("t" = unname(test$coefficients["t.value"]),
      "p" = unname(test$coefficients["p.value"]))
  })
  
  temp_df[["T-vrednost"]] <- abs(sapply(statistic, FUN = function(x) x[["t"]]))
  temp_df[["P-vrednost"]] <- sapply(statistic, FUN = function(x) x[["p"]])
  temp_df[["Signifikanca"]] <- weights::starmaker(temp_df[["P-vrednost"]])
  
  non_calculated_vars <- temp_df[which(is.na(temp_df[["T-vrednost"]])),"Spremenljivka"] # variables for which test statistic could not be calculated
  
  temp_df <- temp_df[which(!is.na(temp_df[["T-vrednost"]])),]
  
  return(list(calculated_table = temp_df,
              non_numeric_vars = non_numeric_vars,
              non_calculated_vars = non_calculated_vars))
}


# function to make dummy variable for every level of a factor variable
# make_dummies <- function(v) {
#   s <- sort(unique(v))
#   d <- outer(v, s, function(v, s) 1L * (v == s))
#   # colnames(d) <- s
#   return(d)
# }

# weighted statistics for categorical variables
create_w_table <- function(orig_data, variable, weights){
  temp_var <- droplevels(haven::as_factor(labelled::user_na_to_na(orig_data[[variable]])))
  temp_df <- as.data.frame(table(temp_var, useNA = "no"))
  
  if(nrow(temp_df) == 0){
    temp_df <- data.frame(" " = paste0("Spremenljivka ", variable, " nima veljavnih (nemanjkajočih) vrednosti."),
                          check.names = FALSE, row.names = "")
    names(temp_df) <- variable
    return(temp_df)
    
  } else if(nrow(temp_df) == 1){
    names(temp_df)[1:2] <- c(variable, "Frekvenca")
    return(temp_df)
    
  } else {
    temp_df[["N po uteževanju"]] <- weighted_table(temp_var, weights = weights)
    temp_df[["% pred uteževanjem"]] <- (temp_df$Freq/sum(temp_df$Freq))*100
    temp_df[["% po uteževanju"]] <- (temp_df[[3]]/sum(temp_df[[3]]))*100
    temp_df[["Absolutna razlika deležev (v odstotnih točkah)"]] <- temp_df[[5]] - temp_df[[4]]
    temp_df[["Relativna razlika deležev (v %)"]] <-(temp_df[[6]]/temp_df[[4]])*100
    
    # n <- sum(temp_df[[3]])
    # 
    # statistic <- lapply(1:nrow(temp_df), FUN = function(i){
    #   test <- prop.test(x = temp_df[[3]][i], n = n, p = temp_df[[4]][i]/100, alternative = "two.sided", correct = TRUE)
    #   
    #   # square test statistic because R reports chi-square value for proportion test, squaring this gives z value
    #   list("z" = unname(sqrt(test$statistic)),
    #        "p" = unname(test$p.value))
    # })
    
    # make dummy (0 1) variable for each level of a factor variable
    dummies <- weights::dummify(x = temp_var, show.na = FALSE, keep.na = TRUE)
    
    # then perform weighted t-test on every dummy variable (mean = proportion)
    statistic <- lapply(seq_len(ncol(dummies)), function(i){
      test <- weights::wtd.t.test(x = dummies[ ,i],
                                  y = mean(dummies[ ,i], na.rm = TRUE),
                                  weight = weights)
      
      c("t" = unname(test$coefficients["t.value"]),
        "p" = unname(test$coefficients["p.value"]))
    })
    
    temp_df[["T-vrednost"]] <- abs(sapply(statistic, FUN = function(x) x[["t"]]))
    temp_df[["P-vrednost"]] <- sapply(statistic, FUN = function(x) x[["p"]])
    temp_df[["Signifikanca"]] <- weights::starmaker(temp_df[["P-vrednost"]])
      
    names(temp_df)[1:2] <- c(variable, "N pred uteževanjem")
    
    # add Total row
    df <- data.frame(v1 = "Skupaj", t(c(colSums(temp_df[-c(1,6:10)]), NA, NA, NA, NA, NA)))
    names(df) <- names(temp_df)
    temp_df <- rbind(temp_df, df)
    
    # ? prikazovanje manjkajočih vrednosti
    # + popravljanje p vrednosti
    # DODATI ŠE EN ! POD SIGNIFIKACNA ČE JE PREMALO ENOT V CELICI IN DODATI V LEGENDO
    
    return(temp_df)
  }
}

# functions for tables download
download_analyses_numeric_table <- function(numeric_table, file){
  wb <- createWorkbook()
  
  addWorksheet(wb = wb,
               sheetName = "Opisne_statistike",
               gridLines = FALSE)
  
  writeDataTable(wb = wb,
                 sheet = 1,
                 x = numeric_table,
                 xy = c(1,1),
                 withFilter = FALSE,
                 tableStyle = "TableStyleLight9",
                 bandedRows = TRUE,
                 bandedCols = TRUE,
                 headerStyle = createStyle(halign = "center"),
                 keepNA = FALSE)
  
  n_col <- ncol(numeric_table)
  n_row <- nrow(numeric_table)
  
  addStyle(wb = wb,
           sheet = 1, 
           style = createStyle(numFmt = "NUMBER"), 
           rows = 1:(n_row+1), cols = (n_col-8):n_col, stack = TRUE, gridExpand = TRUE)
  
  addStyle(wb = wb,
           sheet = 1, 
           style = createStyle(numFmt = "0",), 
           rows = 1:(n_row+1), cols = (n_col-10):(n_col-9), stack = TRUE, gridExpand = TRUE)
  
  setColWidths(wb = wb, sheet = 1, cols = 1:n_col, widths = "auto")
  setColWidths(wb = wb, sheet = 1, cols = 2, widths = 20)
  
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}

download_analyses_factor_tables <- function(factor_tables, orig_data, variables, file){
  wb <- createWorkbook()
  
  addWorksheet(wb = wb,
               sheetName = "Frekvencne_tabele",
               gridLines = FALSE)
  
  # starting row = number of rows of previous table
  # + 2 (for the header and to add a empty row)
  # + 1 for the first table
  start_rows <- c(0, cumsum(3 + sapply(factor_tables, nrow)[-length(factor_tables)])) + 2
  
  for(i in seq_along(factor_tables)){
    writeDataTable(wb = wb,
                   sheet = 1,
                   x = factor_tables[[i]],
                   startRow = start_rows[i],
                   withFilter = FALSE,
                   tableStyle = "TableStyleLight9",
                   bandedRows = TRUE,
                   bandedCols = TRUE,
                   headerStyle = createStyle(halign = "center"),
                   keepNA = FALSE)
    
    writeData(wb = wb,
              sheet = 1,
              x = attr(x = orig_data[[variables[[i]]]], which = "label"),
              startCol = 1, startRow = start_rows[i] - 1)
    
    # addStyle(wb = wb,
    #          sheet = 1,
    #          style = createStyle(numFmt = "NUMBER"),
    #          rows = start_rows[i]:(nrow(factor_tables[[i]]) + start_rows[i]),
    #          cols = 3:9,
    #          stack = TRUE, gridExpand = TRUE)
  }
  
  addStyle(wb = wb,
           sheet = 1,
           style = createStyle(numFmt = "NUMBER"),
           rows = 1:(start_rows[length(start_rows)] + nrow(factor_tables[[length(factor_tables)]])),
           cols = 3:9,
           stack = TRUE, gridExpand = TRUE)
  
  setColWidths(wb = wb, sheet = 1, cols = 1:ncol(factor_tables[[1]]), widths = "auto")
  
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}
