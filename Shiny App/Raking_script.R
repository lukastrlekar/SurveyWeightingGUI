# Trim weights after raking iterative procedure
## Weights are trimmed at user-specified values, and then readjusted so that the
## average of weights again equals 1.
trim_weights <- function(weights, lower = -Inf, upper = Inf){
  new_weights <- ifelse(weights <= lower, lower,
                        ifelse(weights >= upper, upper, weights))
  new_weights <- new_weights/mean(new_weights, na.rm = TRUE)
  return(new_weights)
}


# TODO
# design effect preveri izračun in navedi vir

# Kish (1992) estimator for design effect
design_effect <- function(weights) {
  weights <- weights[!is.na(weights)] * sum(!is.na(weights))/sum(weights[!is.na(weights)])
  deff <- sum(weights^2)/sum(weights)
  return(deff)
}

# prepare data for raking and perform raking
perform_weighting <- function(orig_data = NULL,
                              margins_data,
                              all_raking_variables = NULL,
                              case_id = NULL, 
                              lower,
                              upper, ...){
  
  all_raking_variables <- unname(unlist(all_raking_variables))

  # path <- "Test files/Vnos_margin.xlsx"
  # all_raking_variables <- sheet_names
  # path <- "Test files/Vnos margin (16).xlsx"
  # sheet_names <- openxlsx::getSheetNames(path)
  # sheet_names <- sheet_names[sheet_names %in% all_raking_variables]
  # sheet_list <- lapply(sheet_names, function(sn){openxlsx::read.xlsx(path, sheet = sn)})
  
  sheet_names <- margins_data$sheet_names
  sheet_names <- sheet_names[sheet_names %in% all_raking_variables]
  sheet_list <- margins_data$sheet_list
  sheet_list <- sheet_list[sheet_names]
  
  variables <- lapply(sheet_list, FUN = function(x){
    names(x)[1:2][names(x)[1:2] != "Frekvenca"]
  })
  
  selected_data <- orig_data[ ,unlist(variables), drop = FALSE] # unlist use.names ?
  
  selected_data[] <- lapply(selected_data, FUN = function(x){
    droplevels(clean_data(variable = x))
  }) 
  
  two_dimensional_raking_variables <- lapply(variables, FUN = function(x){
    if(length(x) == 2) x 
  })
  
  two_dimensional_raking_variables <- two_dimensional_raking_variables[lengths(two_dimensional_raking_variables) != 0]
  
  for(i in seq_along(two_dimensional_raking_variables)){
    two_dim_names <- paste0(two_dimensional_raking_variables[[i]], collapse = " x ")
    
    selected_data[[two_dim_names]] <- as.factor(paste0(selected_data[[two_dimensional_raking_variables[[i]][1]]],"_",
                                                       selected_data[[two_dimensional_raking_variables[[i]][2]]]))
  }   
  
  # variable_names <- unlist(lapply(variables, FUN = function(x){
  #   if(length(x) == 2){
  #     paste0(x, collapse = " x ")
  #   } else {
  #     x
  #   }
  # }))
  

  popul_margins <- vector(mode = "list", length = length(sheet_names))
  names(popul_margins) <- sheet_names

  for(i in seq_along(popul_margins)){
    df <- sheet_list[[i]]
    
    df <- df[,colnames(df) %in% c(unlist(variables), "Populacijski.(ciljni).%", "Populacijski (ciljni) %"), drop = FALSE]
    df <- na.omit(df)
    n_col <- ncol(df)
    df <- df[-nrow(df),]
    
    df[df[[n_col]] == 0,] <- NA
    df <- df[complete.cases(df[[n_col]]),]
    temp_vec <- df[[n_col]]

    if(n_col == 3){
      names(temp_vec) <- paste0(df[[1]],"_",df[[2]])
    } else{
      names(temp_vec) <- df[[1]]
    }
    popul_margins[[i]] <- temp_vec/100

    if(all(levels(selected_data[[names(popul_margins)[[i]]]]) %in% names(temp_vec)) != TRUE){
      stop("Population and sample levels must match. Variable categories ",
           paste0(levels(selected_data[[names(popul_margins)[[i]]]])[which(levels(selected_data[[names(popul_margins)[[i]]]]) %in% names(temp_vec) == FALSE)], collapse = ", "),
           " are observed in sample but not in population margins.")
    }
  }
  
  if(is.null(case_id)){
    selected_data$caseid <- seq_len(nrow(selected_data))
    
  } else {
    selected_data$caseid <- orig_data[[case_id]]
  } 
  
  # raking
  outsave <- anesrake::anesrake(inputter = popul_margins,
                                dataframe = as.data.frame(selected_data), 
                                caseid = selected_data$caseid,
                                type = "nolim", 
                                verbose = FALSE, ...)
  
  outsave$weightvec <- trim_weights(weights = outsave$weightvec,
                                    lower = lower, upper = upper)
  
  # save caseid name, in case if user changes this input after weighting is run
  outsave$caseid_name <- ifelse(is.null(case_id), "zaporedna_stevilka", case_id)
  
  return(outsave)
}


download_weights <- function(weights_object = NULL,
                             file_type,
                             file_name,
                             separator,
                             decimal,
                             quote_col_names){
  
  caseweights <- data.frame(caseid = weights_object$caseid,
                            weights = weights_object$weightvec)
  
  names(caseweights)[1] <- weights_object$caseid_name
  
  switch(file_type,
         txt = write.table(x = caseweights, file = file_name, sep = separator, dec = decimal, quote = quote_col_names, row.names = FALSE),
         csv = write.table(x = caseweights, file = file_name, sep = separator, dec = decimal, quote = quote_col_names, row.names = FALSE),
         sav = haven::write_sav(data = caseweights, path = file_name))
}


download_weighting_diagnostic <- function(weights_object, file_name, cut_text, iter_cut_text){
  
  diagnostic <- summary(weights_object)
  vec <- weights_object$weightvec
  deff <- design_effect(vec)
  
  wb <- createWorkbook()

  tabela <- rbind("Konvergenca:" = diagnostic$convergence,
                  "Iterativno rezanje uteži:" = iter_cut_text, 
                  "Rezanje uteži po koncu iteracij:" = cut_text,
                  "Uteževalne spremenljivke:" = paste0(diagnostic$raking.variables, collapse = ", "),
                  "Privzete uteži:" = diagnostic$base.weights,
                  "Vzorčni učinek (design effect):" = paste0(round(deff, 5), ". Porast vzorčne variance zaradi uteževanja: ", round((deff-1)*100,2), "%."))
  
  addWorksheet(wb = wb, sheetName = "Povzetek", gridLines = FALSE)
  writeData(wb = wb,
            sheet = 1,
            x = tabela,
            colNames = FALSE, rowNames = TRUE, withFilter = FALSE)
  
  writeData(wb = wb,
            sheet = 1,
            x = "Opisne statistike uteži",
            startCol = 4, startRow = 1)
  
  writeData(wb = wb,
            sheet = 1,
            x = data.frame(t(unclass(summary(vec))), check.names = FALSE),
            startCol = 4, startRow = 3, headerStyle = createStyle(halign = "center", textDecoration = "bold"),
            rowNames = FALSE, colNames = TRUE)
  
  addStyle(wb = wb, sheet = 1,
           style = createStyle(numFmt = "NUMBER", halign = "center"), rows = 4, cols = 4:9, gridExpand = TRUE)
  
  setColWidths(wb = wb,
               sheet = 1,
               cols = 1:2,
               widths = "auto")
  
  setColWidths(wb = wb,
               sheet = 1,
               cols = 4:9,
               widths = 8)
  
  # NOT ABLE TO MAKE IT WORK ON SHINY SERVER (WORKS LOCALLY THOUGH)
  # my_plot <- hist(x = vec,
  #                 xlim = c(0, max(vec) + 1),
  #                 breaks = seq(min(vec), max(vec), by = ((max(vec) - min(vec))/(length(vec)))))
  # 
  # insertPlot(wb = wb,
  #            sheet = 1, startRow = 6, startCol = 4)
  
  all_raking_variables <- weights_object$varsused
  
  for(i in seq_along(all_raking_variables)){
    sheet_name <- substr(all_raking_variables[i], 1, 31)
    
    addWorksheet(wb = wb,
                 sheetName = sheet_name,
                 gridLines = FALSE)
    writeData(wb = wb,
              sheet = sheet_name,
              x = diagnostic[[all_raking_variables[i]]],
              rowNames = TRUE, borders = "all", headerStyle = createStyle(border = "TopBottomLeftRight"))
    setColWidths(wb = wb,
                 sheet = sheet_name,
                 cols = 1:9,
                 widths = "auto")
  }

  saveWorkbook(wb = wb, file = file_name, overwrite = TRUE)
}

