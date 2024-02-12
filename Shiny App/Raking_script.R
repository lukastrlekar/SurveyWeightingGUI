# TODO
# user_na_to_na ni treba povsod
# za survey rakign če je error da se izpiše kaj informativnega in konvergenca


# Trim weights after raking iterative procedure
## Weights are trimmed at user-specified values, and then readjusted so that the
## average of weights again equals 1.
trim_weights <- function(weights, lower = -Inf, upper = Inf){
  weights[weights < lower] <- lower
  weights[weights > upper] <- upper
  
  weights <- weights/mean(weights)
  return(weights)
}

# Kish (1992) estimator for design effect
design_effect <- function(weights) {
  n <- length(weights)
  (n * sum(weights^2))/sum(weights)^2
}

weightassess <- function(inputter, dataframe, weightvec, prevec = NULL) {
  if (is.null(prevec)) {
    prevec <- rep(1, length(weightvec))
  }
  
  prx <- "Neutežen (vzorčni)"
  if (sum(prevec == 1) != length(prevec)) {
    prx <- "Stare uteži (vzorčni)"
  }
  
  out <- list()
  
  for(i in names(inputter)) {
    target <- inputter[[i]] * 100
    levels <- names(target)
    orign <- weighted_table(dataframe[, i], weights = prevec)[levels]
    origpct <- (orign/sum(orign)) * 100
    newn <- weighted_table(dataframe[, i], weights = weightvec)[levels]
    newpct <- (newn/sum(newn)) * 100
    chpct <- newpct - origpct
    rdisc <- target - newpct
    nout <- cbind(orign, origpct, newn, newpct, target, chpct, rdisc)
    nout2 <- rbind.data.frame(nout, Skupaj = apply(nout, 2, function(x) sum(x, na.rm = TRUE)))
    nout2 <- cbind.data.frame(rownames(nout2), nout2)
    colnames(nout2) <- c(i,
                         paste(prx, "N"),
                         paste(prx, "%"),
                         "Utežen N",
                         "Utežen %",
                         "Populacijski (ciljni) %",
                         "Absolutna razlika utežen neutežen % (odst. točke)",
                         "Rezidualna razlika ciljni utežen % (odst. točke)")
    out[[i]] <- nout2
  }
  out
}

# generic summary for raking with survey package
summary.surveyrake <- function(object, ...) {
  namer <- c("convergence", "base.weights", "raking.variables")
  bwstat <- "Using Base Weights Provided"
  if(sum(object$prevec==1) == length(object$prevec)){
    bwstat <- "No Base Weights Were Used"
  }
  part1list <- list(paste("Convergence was achieved"),
                    bwstat,
                    object$varsused)
  names(part1list) <- namer
  part2list <- weightassess(object$targets, object$dataframe, 
                            object$weightvec, object$prevec)
  out <- c(part1list, part2list)
  out
}

# anesrake summary function code with slight modifications
summary.anesrake <- function(object, ...) {
  namer <- c("convergence", "base.weights", "raking.variables")
  bwstat <- "Using Base Weights Provided"
  if(sum(object$prevec==1)==length(object$prevec)){
    bwstat <- "No Base Weights Were Used"
  }
  part1list <- list(paste(object$converge, "after", object$iterations, "iterations"), bwstat, object$varsused)
  names(part1list) <- namer
  part2list <- weightassess(object$targets, object$dataframe, 
                            object$weightvec, object$prevec)
  out <- c(part1list, part2list)
  out
}

# prepare data for raking and perform raking
perform_weighting <- function(orig_data = NULL,
                              margins_data,
                              all_raking_variables = NULL,
                              case_id = NULL, 
                              lower,
                              upper,
                              package = c("anesrake", "survey"),
                              weightvec,
                              epsilon, ...){

  all_raking_variables <- unname(unlist(all_raking_variables))
  
  sheet_names <- margins_data$sheet_names
  sheet_names <- sheet_names[sheet_names %in% all_raking_variables]
  sheet_list <- margins_data$sheet_list
  sheet_list <- sheet_list[sheet_names]
  
  # TODO to se naredi že v serverju se ne rabi še enkrat tu
  variables <- lapply(sheet_list, FUN = function(x){
    names(x)[1:2][names(x)[1:2] != "Frekvenca"]
  })
  
  selected_data <- orig_data[ ,unlist(variables), drop = FALSE]
  
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
      validate("Populacijske in vzorčne kategorije se morajo ujemati. Kategorije ",
               paste0(levels(selected_data[[names(popul_margins)[[i]]]])[which(levels(selected_data[[names(popul_margins)[[i]]]]) %in% names(temp_vec) == FALSE)], collapse = ", "),
               "so prisotne v vzorčnih, ne pa tudi v populacijskih marginah.")
    }
  }
  
  if(is.null(case_id)){
    selected_data$caseid <- seq_len(nrow(selected_data))
    
  } else {
    selected_data$caseid <- orig_data[[case_id]]
  } 
  
  if(package == "anesrake"){
    # raking anesrake
    outsave <- anesrake::anesrake(inputter = popul_margins,
                                  dataframe = as.data.frame(selected_data), 
                                  caseid = selected_data$caseid,
                                  type = "nolim", 
                                  verbose = FALSE,
                                  weightvec = weightvec, ...)
    
    outsave$weightvec <- trim_weights(weights = outsave$weightvec,
                                      lower = lower, upper = upper)
  }
  
  if(package == "survey"){
    
    for(i in seq_along(two_dimensional_raking_variables)){
      two_dim_names_survey <- paste0(two_dimensional_raking_variables[[i]], collapse = " x ")
      selected_data[[gsub(pattern = " ", replacement = ".", x = two_dim_names_survey)]] <- selected_data[[two_dim_names_survey]]
    }
    
    n <- nrow(orig_data)
    
    pop_margins <- lapply(seq_along(popul_margins), FUN = function(i){
      temp_df <- data.frame(v = names(popul_margins[[i]]),
                            Freq = unname(popul_margins[[i]]) * n)
      names(temp_df)[1] <- gsub(pattern = " ", replacement = ".", x = names(popul_margins[i]))  
      temp_df
    })
    
    names(pop_margins) <- vapply(pop_margins, FUN = function(x) names(x)[1], FUN.VALUE = character(1))
    
    sam_margins <- paste0("list(", 
                          paste0("~", names(pop_margins), collapse = ", "),
                          ")")
    
    if(is.null(all_raking_variables)){
      # TODO temporary solution
      # if no variables are present use anesrake because survey returns an error and app crashes, anesrake returns vector of weights 1
      outsave <- anesrake::anesrake(inputter = popul_margins,
                                    dataframe = as.data.frame(selected_data), 
                                    caseid = selected_data$caseid,
                                    type = "nolim", 
                                    verbose = FALSE,
                                    weightvec = weightvec)
      
    } else {
      # Specify raking design
      weight_design <- survey::svydesign(ids = ~1, data = selected_data, weights = weightvec)
      
      # Perform raking and retrieve weights
      raking <- survey::rake(design = weight_design, 
                             sample.margins = eval(parse(text = sam_margins)), 
                             population.margins = pop_margins,
                             control = list(maxit = 1000, epsilon = epsilon))
      
      raking <- survey::trimWeights(design = raking, upper = upper, lower = lower)
      
      outsave <- list(weightvec = weights(raking), 
                      caseid = selected_data$caseid,
                      varsused = sheet_names,
                      survey_design = raking,
                      prevec = weightvec,
                      targets = popul_margins,
                      dataframe = selected_data)
      
      class(outsave) <- "surveyrake"
    }
  }
  
  # save caseid name, in case if user changes this input after weighting is run
  outsave$caseid_name <- ifelse(is.null(case_id), "zaporedna_stevilka", case_id)
  outsave$package <- package
  
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

download_weighting_diagnostic <- function(weights_object,
                                          file_name,
                                          cut_text,
                                          iter_cut_text,
                                          is_online_server){
  
  diagnostic <- summary(weights_object)
  vec <- weights_object$weightvec
  deff <- design_effect(vec)
  
  wb <- createWorkbook()

  tabela <- rbind("Paket za raking:" = weights_object$package,
                  "Konvergenca:" = diagnostic$convergence,
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
  
  # histogram of weights - NOT ABLE TO MAKE IT WORK ON SHINY SERVER (WORKS LOCALLY THOUGH)
  if(!is_online_server){
    my_plot <- hist(x = vec,
                    xlim = c(0, max(vec) + 1),
                    breaks = seq(min(vec), max(vec), by = ((max(vec) - min(vec))/(length(vec)))),
                    xlab = "Uteži", ylab = "Frekvence",
                    main = "Porazdelitev uteži")
    
    insertPlot(wb = wb,
               sheet = 1, startRow = 6, startCol = 4)
  }
  
  all_raking_variables <- weights_object$varsused
  
  for(i in seq_along(all_raking_variables)){
    sheet_name <- substr(all_raking_variables[i], 1, 31)
    
    addWorksheet(wb = wb,
                 sheetName = sheet_name,
                 gridLines = FALSE)
    writeData(wb = wb,
              sheet = sheet_name,
              x = diagnostic[[all_raking_variables[i]]],
              rowNames = FALSE, borders = "all", headerStyle = createStyle(border = "TopBottomLeftRight",
                                                                           halign = "center",
                                                                           textDecoration = "bold"))
    setColWidths(wb = wb,
                 sheet = sheet_name,
                 cols = 1:9,
                 widths = "auto")
    
    addStyle(wb = wb,
             sheet = sheet_name,
             style = createStyle(numFmt = "0"),
             rows = 2:(nrow(diagnostic[[all_raking_variables[i]]])+1), cols = c(2,4),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb,
             sheet = sheet_name,
             style = createStyle(numFmt = "0.00"),
             rows = 2:(nrow(diagnostic[[all_raking_variables[i]]])+1), cols = c(3,5:7),
             gridExpand = TRUE, stack = TRUE)
  }

  saveWorkbook(wb = wb, file = file_name, overwrite = TRUE)
}

